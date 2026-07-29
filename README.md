# Парсер HAAKE RheoWin `.rwd`

Проект переносит измерения из бинарных файлов RheoWin в подготовленный
Excel-шаблон и строит графики через `matplotlib`.

Рабочий процесс состоит из двух команд:

1. `haake_cli.py` читает `.rwd`, заполняет `input_data` и пересобирает
   `output_data`.
2. `haake_plot.py` читает `output_data` и сохраняет графики в PNG, SVG или PDF.

ASCII-экспорт RheoWin нужен как описание структуры колонок. Измерительные
значения берутся из `.rwd`; ASCII не обязан относиться к тому же эксперименту,
если состав и порядок колонок совпадают.

ASCII можно получить в RheoWin Data Manager экспортом готового `.rwd` или
добавить соответствующий блок в рабочий скрипт Job Manager.

## Структура проекта

```text
haake_parser/
├── haake_cli.py
├── haake_plot.py
├── haake_rheo/
│   ├── __init__.py
│   ├── excel_import.py
│   └── plot.py
├── docs/
│   └── parser_usage_sequence.puml
├── input/
│   ├── may_proj_template_first20_clean.xlsx
│   ├── D_series_report_template_first20_clean.xlsx
│   ├── D_series_template_first20_clean.xlsx
│   └── SCTNA_60_PyrO_40_Katal68_3_template_first20_clean.xlsx
├── tests/
│   ├── test_excel_import.py
│   └── test_plot.py
├── output/
├── requirements.txt
├── .gitignore
└── README.md
```

Папки с исходными экспериментами и содержимое `output/` не хранятся в Git.

## Установка

```bash
cd /Users/artem/Desktop/Output/iam_ras/haake_parser
python3 -m pip install -r requirements.txt
```

## Июльский проект: полный запуск

### 1. Заполнить Excel

```bash
cd /Users/artem/Desktop/Output/iam_ras/haake_parser

python3 haake_cli.py \
  "/Users/artem/Desktop/Output/iam_ras/haake_parser/input/may_proj_template_first20_clean.xlsx" \
  "/Users/artem/Desktop/Output/iam_ras/haake_parser/input/07_july_proj/SCTNA_59_tio2_41_3000nm_catal68_6/0v/SCTNA_59_tio2_41_3000nm_catal68_6_freq=from_01_to_100hz_U=0v_30_points_ver1_good.txt" \
  "/Users/artem/Desktop/Output/iam_ras/haake_parser/input/07_july_proj" \
  --recursive \
  --sheet "input_data" \
  -o "/Users/artem/Desktop/Output/iam_ras/haake_parser/output/07_july_proj.xlsx"
```

В этой команде:

- шаблон задаёт расположение листов и таблиц;
- ASCII задаёт состав бинарных каналов;
- `07_july_proj` является корневой папкой с `.rwd`;
- `--recursive` включает поиск во всех вложенных папках;
- результат сохраняется в `output/07_july_proj.xlsx`.

Файлы и папки можно перетаскивать в терминал. Лучше сохранять результат под
новым именем и не держать эту книгу открытой в Excel во время записи.

### 2. Построить графики

```bash
python3 haake_plot.py \
  "/Users/artem/Desktop/Output/iam_ras/haake_parser/output/07_july_proj.xlsx" \
  --sheet "output_data" \
  -o "/Users/artem/Desktop/Output/iam_ras/haake_parser/output/plots_07_july_proj" \
  --points 20 \
  --poly-order 3 \
  --legend-columns 1 \
  --formats png \
  --split-by-series
```

При `--split-by-series` каждая смесь получает отдельную папку графиков. На
одной картинке остаются только напряжения одной смеси.

## Правила идентификации данных

### Имя `.rwd` имеет приоритет

Состав, напряжение и заявленная частота определяются по имени `.rwd`.

Например:

```text
SCTNA_59_tio2_41_3000nm_catal68_6_freq=from_01_to_100hz_U=2000v_30_points_ver1_good.rwd
```

относится к смеси `SCTNA_59_tio2_41_3000nm_catal68_6` и напряжению `2000 В`.

Внутренние служебные пути RheoWin могут сохранять старое имя эксперимента
после ручного переименования файла. Если внутренний путь не совпадает с именем
`.rwd`, парсер выводит предупреждение, но использует имя файла.

### Частоты приводятся к общей сетке

Для отчётов используется стандартный набор:

```text
0.1, 0.5, 1, 5, 10, 20, 30, 50, 75, 100 Гц
```

Близкие значения канала, например `30.66` и `51.1 Гц`, записываются в блоки
`30` и `50 Гц`. Сам канал частоты всё равно проверяется при разборе `.rwd`.

### Измерения округляются как в ASCII RheoWin

По умолчанию числа приводятся к разрядности, близкой к официальному
ASCII-экспорту. Это устраняет хвосты `float32` вида `319462.0313`, когда
RheoWin показывает `319500`.

## Что делает парсер

1. Проверяет пути к шаблону, ASCII и `.rwd`.
2. Читает заголовки ASCII и строит карту бинарных каналов.
3. Находит одиночные и сборные `.rwd`.
4. Группирует данные по смеси и напряжению из имени файла.
5. Сопоставляет частотные сегменты и приводит частоты к общей сетке.
6. Записывает измерения на `input_data`.
7. Проверяет заполненные блоки на дубли.
8. Пересобирает `output_data` для первых 20 точек.
9. Обновляет `instrument_metadata`.
10. Сохраняет новую Excel-книгу и печатает время выполнения.

Лист `freq_diagrams` парсер не заполняет картинками. Графики сохраняются
отдельным скриптом.

## Интерактивный запуск

Парсер можно запустить без аргументов:

```bash
python3 haake_cli.py
```

Файлы и папки можно перетаскивать в терминал. Для обработки вложенных папок
нужно ответить `y` на вопрос о рекурсивном поиске.

## D31, D33 и D35

Серии можно последовательно записать в один отчёт:

```bash
cd /Users/artem/Desktop/Output/iam_ras/haake_parser

TEMPLATE="/Users/artem/Desktop/Output/iam_ras/haake_parser/input/D_series_report_template_first20_clean.xlsx"
ASCII="/Users/artem/Desktop/Output/iam_ras/haake_parser/input/06_june_proj/d31/2000v/D31_freq=20hz_U=2000v_30_points_ver1.txt"
INPUT_ROOT="/Users/artem/Desktop/Output/iam_ras/haake_parser/input/06_june_proj"
OUTPUT="/Users/artem/Desktop/Output/iam_ras/haake_parser/output/D31_D33_D35_06_june.xlsx"

python3 haake_cli.py series d31 \
  --template "$TEMPLATE" \
  --ascii "$ASCII" \
  --input-root "$INPUT_ROOT" \
  -o "$OUTPUT"

python3 haake_cli.py series d33 \
  --template "$OUTPUT" \
  --ascii "$ASCII" \
  --input-root "$INPUT_ROOT" \
  -o "$OUTPUT"

python3 haake_cli.py series d35 \
  --template "$OUTPUT" \
  --ascii "$ASCII" \
  --input-root "$INPUT_ROOT" \
  -o "$OUTPUT"
```

Графики:

```bash
python3 haake_plot.py \
  "/Users/artem/Desktop/Output/iam_ras/haake_parser/output/D31_D33_D35_06_june.xlsx" \
  --sheet "output_data" \
  -o "/Users/artem/Desktop/Output/iam_ras/haake_parser/output/plots_D31_D33_D35_06_june" \
  --points 20 \
  --poly-order 3 \
  --legend-columns 3 \
  --formats png \
  --split-by-series \
  --series D31 D33 D35
```

## CSV без Excel-шаблона

Для экспериментов с другой структурой можно выгрузить каждый `.rwd` в CSV:

```bash
python3 haake_cli.py csv \
  "/Users/artem/Desktop/Output/iam_ras/data/TiO2_5V_PMS400_95V" \
  --recursive \
  -o "/Users/artem/Desktop/Output/iam_ras/haake_parser/output/TiO2_5V_PMS400_95V_csv" \
  --decimal-comma
```

Кроме таблиц измерений будут созданы:

- `_summary.csv` — результат обработки каждого `.rwd`;
- `_metadata.csv` — метаданные;
- `_mapping.csv` — использованные бинарные каналы.

## Настройки графиков

Количество точек:

```bash
--points 20
```

Степень полинома в координатах `log10(gamma)` и `log10(y)`:

```bash
--poly-order 3
```

Если полином третьего порядка неустойчив, скрипт автоматически пробует более
низкую степень. Отключить сглаживание полностью:

```bash
--no-fit
```

Построить только выбранные показатели:

```bash
--metrics g_prime g_double_prime eta tan_delta
```

Построить только выбранные частоты:

```bash
--frequencies 0.1 0.5 1 5
```

## Проверка перед коммитом

```bash
cd /Users/artem/Desktop/Output/iam_ras/haake_parser
python3 -m unittest discover -s tests -v
python3 -m py_compile haake_cli.py haake_plot.py haake_rheo/*.py
git status --short
```
