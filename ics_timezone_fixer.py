#!/usr/bin/env python3
"""Скачивает, фильтрует и модифицирует календарь в формате ICS.

Скрипт предназначен для подготовки рабочего календаря Outlook к импорту в Google
Calendar. Outlook записывает таймзоны как Windows-названия (например
``TZID=Central European Standard Time``), которые Google Calendar не распознаёт.
Скрипт по правилам из ``config.json``:

* удаляет ненужные события (``action: delete``);
* подменяет строку ``TZID`` у событий (``action: change_timezone``) — БЕЗ пересчёта
  времени: цифры часов остаются как есть, меняется только ярлык таймзоны;
* применяет глобальный словарь ``tzid_map`` для замены ``TZID`` сразу у всех событий
  (и у соответствующих блоков ``VTIMEZONE``).

Сопоставление событий ведётся по полю ``SUMMARY`` с поддержкой wildcard-масок
(``*`` и ``?``) через модуль ``fnmatch``.
"""

from __future__ import annotations

import argparse
import datetime as dt
import fnmatch
import json
import logging
import os
import sys
import urllib.parse
from pathlib import Path
from typing import Any

import requests
from icalendar import Calendar

logger = logging.getLogger("ics_timezone_fixer")

# Свойства события, у которых таймзона задаётся через параметр TZID.
TZID_PROPERTIES = ("DTSTART", "DTEND", "RECURRENCE-ID")

ENV_SOURCE_URL = "ICS_SOURCE_URL"


# --------------------------------------------------------------------------- #
# CLI и логирование
# --------------------------------------------------------------------------- #
def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    """Разбирает аргументы командной строки."""
    parser = argparse.ArgumentParser(
        description="Фильтрация и правка таймзон в ICS-календаре перед импортом в Google Calendar.",
    )
    parser.add_argument(
        "--config",
        default="config.json",
        help="Путь к файлу конфигурации (по умолчанию: config.json).",
    )
    parser.add_argument(
        "--url",
        default=None,
        help=(
            "URL исходного .ics (http/https) или путь к локальному файлу. "
            f"Перекрывает env {ENV_SOURCE_URL} и source_url из конфига."
        ),
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Путь для сохранения результата (перекрывает output_file из конфига).",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="Уровень логирования (по умолчанию: INFO).",
    )
    return parser.parse_args(argv)


def setup_logging(level: str) -> None:
    """Настраивает базовое логирование."""
    logging.basicConfig(
        level=getattr(logging, level),
        format="%(asctime)s %(levelname)s %(message)s",
        datefmt="%H:%M:%S",
    )


# --------------------------------------------------------------------------- #
# Конфигурация и загрузка
# --------------------------------------------------------------------------- #
def load_config(path: str) -> dict[str, Any]:
    """Читает и валидирует JSON-конфигурацию.

    Возвращает словарь с нормализованными ключами ``output_file``, ``rules``,
    ``tzid_map`` и (опционально) ``source_url``.
    """
    try:
        raw = Path(path).read_text(encoding="utf-8")
    except FileNotFoundError as exc:
        raise SystemExit(f"Файл конфигурации не найден: {path}") from exc

    try:
        config = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise SystemExit(f"Некорректный JSON в {path}: {exc}") from exc

    if not isinstance(config, dict):
        raise SystemExit("Корень config.json должен быть объектом (JSON object).")

    config.setdefault("output_file", "processed_calendar.ics")
    config.setdefault("rules", [])
    config.setdefault("tzid_map", {})
    config.setdefault("source_url", "")

    if not isinstance(config["rules"], list):
        raise SystemExit("Поле 'rules' должно быть списком.")
    if not isinstance(config["tzid_map"], dict):
        raise SystemExit("Поле 'tzid_map' должно быть объектом {старое: новое}.")

    return config


def resolve_url(config: dict[str, Any], args: argparse.Namespace) -> str:
    """Определяет URL источника: --url > env ICS_SOURCE_URL > config.source_url."""
    url = args.url or os.environ.get(ENV_SOURCE_URL) or config.get("source_url")
    if not url:
        raise SystemExit(
            "Не задан источник .ics. Укажите --url, переменную окружения "
            f"{ENV_SOURCE_URL} или source_url в config.json."
        )
    return url


def download_ics(url: str, timeout: int = 30) -> bytes:
    """Скачивает .ics по http(s) или читает локальный файл (для тестов).

    Поддерживает схемы ``http``/``https`` (через requests), а также ``file://``
    и обычные локальные пути.
    """
    parsed = urllib.parse.urlparse(url)

    if parsed.scheme in ("http", "https"):
        logger.info("Скачиваю календарь: %s", url)
        try:
            response = requests.get(url, timeout=timeout)
            response.raise_for_status()
        except requests.RequestException as exc:
            raise SystemExit(f"Ошибка скачивания {url}: {exc}") from exc
        return response.content

    # Локальный файл или file:// URL.
    if parsed.scheme == "file":
        local_path = _file_url_to_path(parsed)
    else:
        local_path = url
    logger.info("Читаю локальный файл: %s", local_path)
    try:
        return Path(local_path).read_bytes()
    except OSError as exc:
        raise SystemExit(f"Не удалось прочитать файл {local_path}: {exc}") from exc


def _file_url_to_path(parsed: urllib.parse.ParseResult) -> str:
    """Преобразует file:// URL в путь файловой системы (с учётом Windows)."""
    path = urllib.parse.unquote(parsed.path)
    # На Windows file:///D:/x -> '/D:/x', убираем ведущий слеш.
    if os.name == "nt" and path.startswith("/") and len(path) > 2 and path[2] == ":":
        path = path[1:]
    return path


def parse_calendar(data: bytes) -> Calendar:
    """Парсит ICS-данные в объект Calendar."""
    try:
        return Calendar.from_ical(data)
    except Exception as exc:  # icalendar бросает разные типы при битом ICS
        raise SystemExit(f"Не удалось распарсить ICS: {exc}") from exc


# --------------------------------------------------------------------------- #
# Сопоставление и модификация
# --------------------------------------------------------------------------- #
def summary_matches(summary: str, pattern: str) -> bool:
    """Регистронезависимое сопоставление SUMMARY с wildcard-маской (* и ?)."""
    return fnmatch.fnmatchcase(summary.lower(), pattern.lower())


def relabel_tzid(event: Calendar, new_tzid: str) -> int:
    """Меняет ярлык TZID у временных полей события, сохраняя цифры времени.

    Возвращает число изменённых свойств. Поля DATE-only (события на весь день) и
    значения в UTC (с суффиксом ``Z``) пропускаются — у них нет TZID.
    """
    changed = 0
    for name in TZID_PROPERTIES:
        prop = event.get(name)
        if prop is None:
            continue
        value = getattr(prop, "dt", None)
        if not isinstance(value, dt.datetime):
            continue  # DATE-only / иные типы — пропускаем
        if prop.to_ical().endswith(b"Z"):
            continue  # время в UTC — TZID отсутствует, не трогаем
        old = prop.params.get("TZID")
        prop.params["TZID"] = new_tzid
        changed += 1
        logger.debug("  %s: TZID %r -> %r", name, old, new_tzid)
    return changed


def apply_rules(event: Calendar, rules: list[dict[str, Any]], stats: dict[str, int]) -> bool:
    """Применяет правила к событию по порядку.

    Возвращает ``True``, если событие нужно удалить. Первое сработавшее правило
    ``delete`` прерывает цепочку; ``change_timezone`` применяется и не прерывает.
    """
    summary = str(event.get("SUMMARY", ""))
    for rule in rules:
        pattern = rule.get("match")
        action = rule.get("action")
        if not pattern or not summary_matches(summary, pattern):
            continue

        if action == "delete":
            logger.debug("delete: %r (маска %r)", summary, pattern)
            stats["deleted"] += 1
            return True

        if action == "change_timezone":
            params = rule.get("parameters") or {}
            new_tz = params.get("new_tz")
            if not new_tz:
                logger.warning(
                    "Правило change_timezone без parameters.new_tz пропущено (маска %r).",
                    pattern,
                )
                continue
            logger.debug("change_timezone: %r -> %s (маска %r)", summary, new_tz, pattern)
            if relabel_tzid(event, new_tz):
                stats["retimed"] += 1
        else:
            logger.warning("Неизвестное действие %r в правиле (маска %r).", action, pattern)

    return False


def apply_tzid_map(cal: Calendar, tzid_map: dict[str, str], stats: dict[str, int]) -> None:
    """Глобально заменяет TZID у всех событий и у блоков VTIMEZONE."""
    if not tzid_map:
        return

    # 1. Ссылки на таймзону в событиях.
    for event in cal.walk("VEVENT"):
        for name in TZID_PROPERTIES:
            prop = event.get(name)
            if prop is None:
                continue
            current = prop.params.get("TZID")
            if current in tzid_map:
                new = tzid_map[current]
                prop.params["TZID"] = new
                stats["mapped"] += 1
                logger.debug("map %s: TZID %r -> %r", name, current, new)

    # 2. Сами блоки VTIMEZONE — чтобы файл остался самосогласованным.
    for vtz in cal.walk("VTIMEZONE"):
        current = str(vtz.get("TZID", ""))
        if current in tzid_map:
            new = tzid_map[current]
            vtz["TZID"] = new
            stats["vtimezones"] += 1
            logger.debug("map VTIMEZONE: TZID %r -> %r", current, new)


def process_calendar(cal: Calendar, config: dict[str, Any]) -> dict[str, int]:
    """Применяет правила и глобальный маппинг к календарю (in-place).

    Удалённые события исключаются из ``cal.subcomponents``. Возвращает статистику.
    """
    rules = config["rules"]
    stats = {"events": 0, "deleted": 0, "retimed": 0, "mapped": 0, "vtimezones": 0}

    kept: list[Any] = []
    for component in cal.subcomponents:
        if component.name != "VEVENT":
            kept.append(component)
            continue
        stats["events"] += 1
        if apply_rules(component, rules, stats):
            continue  # событие удаляется — не добавляем
        kept.append(component)
    cal.subcomponents = kept

    apply_tzid_map(cal, config["tzid_map"], stats)
    return stats


def write_calendar(cal: Calendar, path: str) -> None:
    """Сохраняет календарь в файл."""
    Path(path).write_bytes(cal.to_ical())
    logger.info("Результат сохранён: %s", path)


# --------------------------------------------------------------------------- #
# main
# --------------------------------------------------------------------------- #
def main(argv: list[str] | None = None) -> int:
    """Точка входа: оркестрация всего процесса."""
    args = parse_args(argv)
    setup_logging(args.log_level)

    config = load_config(args.config)
    url = resolve_url(config, args)
    output = args.output or config["output_file"]

    data = download_ics(url)
    cal = parse_calendar(data)

    stats = process_calendar(cal, config)
    write_calendar(cal, output)

    logger.info(
        "Готово: событий обработано %d, удалено %d, change_timezone %d, "
        "TZID-маппинг %d (VTIMEZONE переименовано %d).",
        stats["events"],
        stats["deleted"],
        stats["retimed"],
        stats["mapped"],
        stats["vtimezones"],
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
