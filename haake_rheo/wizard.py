"""Step-by-step entry points; existing argument-based commands remain valid."""
from __future__ import annotations

import argparse
import math
import shlex
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


class Prompts:
    def __init__(self, language="ru"):
        self.language = language

    def text(self, ru, en):
        return ru if self.language == "ru" else en

    def ask(self, ru, en, default=None):
        suffix = f" [{default}]" if default is not None else self.text(" [обязательно]", " [required]")
        while True:
            answer = input(self.text(ru, en) + suffix + ": ").strip()
            if answer:
                return answer
            if default is not None:
                return str(default)
            print(self.text("Введите значение.", "Enter a value."))

    def choice(self, ru, en, choices, default):
        while True:
            value = self.ask(f"{ru} ({'/'.join(choices)})", f"{en} ({'/'.join(choices)})", default)
            for choice in choices:
                if value.casefold() == choice.casefold():
                    return choice
            print(self.text("Допустимые значения: ", "Allowed values: ") + ", ".join(choices))

    def yes(self, ru, en, default=False):
        options = ("да", "нет") if self.language == "ru" else ("yes", "no")
        return self.choice(ru, en, options, options[0 if default else 1]) == options[0]

    def integer(self, ru, en, default, minimum=1):
        while True:
            try:
                value = int(self.ask(ru, en, default))
                if value >= minimum:
                    return value
            except ValueError:
                pass
            print(self.text(f"Введите целое число ≥ {minimum}.", f"Enter an integer ≥ {minimum}."))

    def paths(self, ru, en, default=None, suffixes=None, many=False, exists=True):
        while True:
            raw = self.ask(ru, en, default)
            try:
                # Also accept an unquoted pasted path containing spaces.
                direct = Path(raw).expanduser()
                parts = [raw] if direct.exists() else shlex.split(raw)
                paths = [Path(part).expanduser().resolve() for part in parts]
                if not paths or (not many and len(paths) != 1):
                    raise ValueError("count")
                if exists and any(not path.exists() for path in paths):
                    raise ValueError("missing")
                if suffixes and any(path.suffix.lower() not in suffixes or (exists and not path.is_file()) for path in paths):
                    raise ValueError("suffix")
                return paths if many else paths[0]
            except (ValueError, OSError):
                print(self.text("Проверьте путь и расширение файла. Можно перетащить файл в терминал.",
                                "Check the path and file extension. You can drag a file into the terminal."))

    def output(self, default, source=None):
        while True:
            path = self.paths("Новая итоговая книга", "New output workbook", str(default), {".xlsx"}, exists=False)
            if path.exists() or path == source:
                print(self.text("Этот файл уже существует. Укажите новое имя.", "This file already exists. Choose a new name."))
            else:
                path.parent.mkdir(parents=True, exist_ok=True)
                return path

    def tokens(self, ru, en, default, allowed=None):
        while True:
            values = self.ask(ru, en, default).split()
            if values and (allowed is None or all(v in allowed for v in values)):
                return values
            print(self.text("Проверьте значения.", "Check the values."))

    def numbers(self, ru, en, default, positive=False):
        while True:
            try:
                values = [float(v.replace(",", ".")) for v in self.ask(ru, en, default).split()]
                if values and all(math.isfinite(v) and (not positive or v > 0) for v in values):
                    return values
            except ValueError:
                pass
            print(self.text("Введите числа через пробел.", "Enter numbers separated by spaces."))


def select_workflow(argv, kind):
    """Return None for legacy commands, otherwise (mode, profile, language)."""
    if argv and argv[0] not in {"auto", "custom", "--mode"} and not argv[0].startswith("--mode="):
        return None
    parser = argparse.ArgumentParser(description="auto / custom: B_draft / K_draft; ru / en")
    parser.add_argument("mode", nargs="?", choices=["auto", "custom"])
    parser.add_argument("--mode", dest="mode_flag", choices=["auto", "custom"])
    parser.add_argument("--profile", choices=["B_draft", "K_draft"])
    parser.add_argument("--language", choices=["ru", "en"])
    args = parser.parse_args(argv)
    if args.mode and args.mode_flag and args.mode != args.mode_flag:
        parser.error("Conflicting modes")
    p = Prompts(args.language or "ru")
    mode = args.mode_flag or args.mode or p.choice("Режим / Mode", "Mode / Режим", ["auto", "custom"], "auto")
    if mode == "auto":
        if args.profile or args.language:
            parser.error("--profile / --language require custom")
        return mode, None, "ru"
    profile = args.profile or p.choice("Подход и шаблон / Profile", "Profile / Подход и шаблон", ["B_draft", "K_draft"], "K_draft")
    language = args.language or p.choice("Язык / Language", "Language / Язык", ["en", "ru"], "ru")
    return mode, profile, language


def custom_import(profile, language):
    p = Prompts(language)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    print(p.text("Enter принимает значение в скобках. Ctrl+C — выход.", "Enter accepts the value in brackets. Ctrl+C exits."))
    if profile == "B_draft":
        from .b_draft import import_report, identity
        from .excel_import import collect_csv_source_files
        inputs = p.paths("Папка или RWD-файлы", "Folder or RWD files", many=True)
        recursive = p.yes("Искать во вложенных папках?", "Search subfolders?", True)
        identities = {}
        files, discovered_ascii = collect_csv_source_files(inputs, recursive)
        selected_files = []
        for file in files:
            try:
                identity(file)
            except RuntimeError:
                print(p.text("Уточните условия для файла: ", "Specify conditions for file: ") + str(file))
                action = p.choice("Пропустить (skip) или указать условия (specify)",
                                  "Skip file or specify conditions", ["skip", "specify"], "skip")
                if action == "skip":
                    print(p.text("Пропущен: ", "Skipped: ") + str(file))
                    continue
                material = p.choice("Вещество", "Material", ["Fe3O4", "TiO2"], None)
                concentration = p.numbers("Объемная доля, % (одно число)", "Volume fraction, % (one number)", "", positive=True)
                if len(concentration) != 1 or concentration[0] >= 100:
                    raise RuntimeError(p.text("Нужно одно значение от 0 до 100%.", "Enter one value between 0 and 100%."))
                shaft = p.choice("Вал", "Shaft", ["Alum", "Standart"], None)
                field = p.choice("Поле: control = без магнита", "Field: control = no magnet", ["control", "magnet"], None)
                identities[file] = (material, concentration[0], shaft, field)
            selected_files.append(file)
        if not selected_files:
            raise RuntimeError(p.text("Не выбраны RWD-файлы.", "No RWD files selected."))
        inputs = [*selected_files, *discovered_ascii]
        ascii_paths = []
        if p.yes("Указать ASCII отдельно (иначе поиск рядом с RWD)?", "Select ASCII separately (otherwise search beside RWD)?"):
            ascii_paths = p.paths("ASCII-файлы", "ASCII files", suffixes={".txt", ".asc", ".csv"}, many=True)
        temperatures = p.numbers("Номинальные температуры сегментов по порядку, °C", "Nominal segment temperatures in order, °C", "20 30 40 50 60")
        print(p.text("В эталоне 46nm_3&4.xlsx усредняются точки 50 и 100 каждого сегмента.",
                     "Reference 46nm_3&4.xlsx averages points 50 and 100 of each segment."))
        first = p.integer("Первая точка для среднего", "First point to average", 50)
        second = p.integer("Вторая точка для среднего", "Second point to average", 100)
        if first == second:
            raise RuntimeError(p.text("Для среднего нужны две разные точки.", "The two points must be different."))
        output = p.output(ROOT / "outputs" / f"B_draft_{stamp}_{language}.xlsx")
        import_report(inputs, ascii_paths, recursive, temperatures, (first, second), output, language, identities)
        print(p.text("Создана книга: ", "Workbook created: ") + str(output))
        return 0
    from . import excel_import
    default_template = ROOT / "input" / "may_proj_template_first20_clean.xlsx"
    template = p.paths("Excel-шаблон", "Excel template", str(default_template) if default_template.exists() else None, {".xlsx"})
    ascii_path = p.paths("ASCII-структура", "ASCII structure", suffixes=excel_import.ASCII_SUFFIXES)
    inputs = p.paths("Папка или RWD-файлы", "Folder or RWD files", many=True)
    recursive = p.yes("Искать во вложенных папках?", "Search subfolders?", True)
    series_folders = p.yes("Каждая папка верхнего уровня — отдельная серия?", "Treat each top-level folder as a separate series?")
    sheet = p.ask("Лист для измерений", "Measurement sheet", "input_data")
    output = p.output(ROOT / "outputs" / f"K_draft_{stamp}.xlsx", template)
    include_ascii = p.yes("Брать измерения также из ASCII?", "Include measurements from ASCII too?")
    diagnostics = p.yes("Добавить диагностический лист?", "Add a diagnostics sheet?")
    args = [str(template), str(ascii_path), *map(str, inputs), "--sheet", sheet, "-o", str(output)]
    for flag, enabled in [("--recursive", recursive), ("--series-folders", series_folders),
                          ("--include-ascii-data", include_ascii), ("--include-diagnostics", diagnostics)]:
        if enabled:
            args.append(flag)
    return excel_import.main(args)


def custom_plot(profile, language):
    p = Prompts(language)
    workbook = p.paths("Книга с результатами", "Results workbook", suffixes={".xlsx"})
    default_dir = workbook.with_name(workbook.stem + "_plots_" + datetime.now().strftime("%Y%m%d_%H%M%S"))
    output = p.paths("Папка для графиков", "Plot directory", str(default_dir), exists=False)
    if profile == "B_draft":
        from .b_draft import plot_report
        formats = p.tokens("Форматы: png svg pdf xlsx", "Formats: png svg pdf xlsx", "xlsx", {"png", "svg", "pdf", "xlsx"})
        dpi = p.integer("DPI", "DPI", 300, 72) if "png" in formats else 300
        return plot_report(workbook, output, formats, language, dpi)
    from . import plot
    sheet = p.ask("Лист с данными", "Data sheet", "output_data")
    points = p.integer("Сколько первых точек брать", "Number of first points", 20)
    no_fit = p.yes("Отключить полиномиальные линии?", "Disable polynomial curves?")
    order = 3 if no_fit else p.integer("Степень полинома", "Polynomial degree", 3)
    formats = p.tokens("Форматы: png svg pdf", "Formats: png svg pdf", "png", {"png", "svg", "pdf"})
    metrics = p.tokens("Метрики: g_prime g_double_prime eta tan_delta", "Metrics: g_prime g_double_prime eta tan_delta", " ".join(plot.METRICS), set(plot.METRICS))
    frequencies = None
    if p.yes("Выбрать отдельные частоты?", "Select individual frequencies?"):
        frequencies = p.numbers("Частоты, Гц", "Frequencies, Hz", "0.1 0.5 1 5 10 20 30 50 75 100", positive=True)
    split = p.yes("Разделить графики по смесям?", "Separate plots by mixture?", True)
    series = None
    if split and p.yes("Выбрать отдельные серии?", "Select individual series?"):
        series = p.tokens("Названия серий через пробел", "Series names separated by spaces", "")
    columns = p.integer("Столбцов в легенде", "Legend columns", 1)
    dpi = p.integer("DPI", "DPI", 300, 72)
    args = [str(workbook), "--sheet", sheet, "-o", str(output), "--points", str(points),
            "--poly-order", str(order), "--formats", *formats, "--metrics", *metrics,
            "--legend-columns", str(columns), "--dpi", str(dpi), "--language", language]
    if no_fit:
        args.append("--no-fit")
    if split:
        args.append("--split-by-series")
    if series:
        args.extend(["--series", *series])
    if frequencies:
        args.extend(["--frequencies", *map(str, frequencies)])
    return plot.main(args)


def run(kind, argv=None):
    argv = sys.argv[1:] if argv is None else argv
    from . import excel_import, plot
    engine = excel_import if kind == "import" else plot
    try:
        if argv in (["--help"], ["-h"]):
            print("Modes: auto (legacy wizard), custom (B_draft / K_draft; ru / en).\n"
                  "Examples: python3 haake_cli.py custom; python3 haake_plot.py custom\n"
                  "Custom options: custom --help\nLegacy command options:")
        selection = select_workflow(argv, kind)
        if selection is None:
            return engine.main(argv)
        mode, profile, language = selection
        if mode == "auto":
            return engine.main([])
        return custom_import(profile, language) if kind == "import" else custom_plot(profile, language)
    except (EOFError, KeyboardInterrupt):
        print("\nОстановлено / Cancelled.")
        return 130
    except (RuntimeError, ValueError, OSError) as exc:
        print(f"Ошибка / Error: {exc}", file=sys.stderr)
        return 1
