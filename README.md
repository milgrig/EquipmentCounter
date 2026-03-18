# Equipment Counter

Автоматический подсчёт оборудования и кабельных длин из инженерных чертежей (DWG/DXF/PDF).

Программа анализирует электротехнические чертежи, находит легенду условных обозначений, подсчитывает оборудование (светильники, щиты, розетки, выключатели) и длины кабелей, генерирует PNG-визуализации с маркерами и формирует ВОР (ведомость объёмов работ) в формате DOCX.

---

## Содержание

- [Требования](#требования)
- [Установка на Windows](#установка-на-windows)
- [Установка на Linux (сервер)](#установка-на-linux-сервер)
- [ODA File Converter](#oda-file-converter)
- [Использование](#использование)
  - [GUI (Windows)](#gui-windows)
  - [Командная строка](#командная-строка)
  - [Примеры](#примеры)
- [Выходные файлы](#выходные-файлы)
- [Структура проекта](#структура-проекта)
- [Решение проблем](#решение-проблем)

---

## Требования

| Компонент | Версия | Обязательно |
|-----------|--------|-------------|
| **Python** | 3.10+ | Да |
| **pip** | актуальный | Да |
| **ODA File Converter** | 25+ | Только для `.dwg` файлов |

### Python-зависимости

| Пакет | Назначение |
|-------|-----------|
| `ezdxf` >= 1.4.0 | Чтение и парсинг DXF файлов |
| `pdfplumber` >= 0.11.0 | Извлечение текста из PDF |
| `openpyxl` >= 3.1.0 | Чтение Excel (для сравнения с эталонами) |
| `matplotlib` >= 3.8.0 | Рендеринг DXF в PNG |
| `PyMuPDF` >= 1.24.0 | Рендеринг PDF страниц |
| `Pillow` >= 10.0.0 | Обработка изображений |
| `python-docx` | Генерация ВОР документов (.docx) |

---

## Установка на Windows

### 1. Установить Python

1. Скачайте Python 3.10+: https://www.python.org/downloads/
2. При установке **обязательно** поставьте галочку **"Add Python to PATH"**
3. Перезагрузите компьютер

Проверка:
```cmd
python --version
```

### 2. Установить ODA File Converter (для DWG)

См. раздел [ODA File Converter](#oda-file-converter).

### 3. Установить зависимости и запустить

**Вариант А — Двойной клик (рекомендуется):**

Дважды кликните на файл `equipment_counter.bat`. Скрипт сам проверит зависимости, установит их при первом запуске и откроет GUI.

**Вариант Б — Из командной строки:**

```cmd
cd C:\путь\к\EquipmentCounter
pip install -r requirements.txt
pip install python-docx
python equipment_gui.py
```

---

## Установка на Linux (сервер)

### 1. Установить Python и системные пакеты

**Ubuntu / Debian:**
```bash
sudo apt update
sudo apt install -y python3 python3-pip python3-venv git
```

**CentOS / RHEL / Fedora:**
```bash
sudo dnf install -y python3 python3-pip git
```

Проверка:
```bash
python3 --version   # должна быть 3.10+
```

### 2. Скопировать проект и настроить окружение

```bash
# Скопировать файлы проекта на сервер (scp, rsync или распаковать zip)
mkdir -p /opt/equipment-counter
cp -r EquipmentCounter/* /opt/equipment-counter/
cd /opt/equipment-counter

# Создать виртуальное окружение
python3 -m venv venv
source venv/bin/activate

# Установить зависимости
pip install --upgrade pip
pip install -r requirements.txt
pip install python-docx
```

### 3. Установить ODA File Converter (для DWG)

На Linux ODA File Converter доступен в виде `.deb` или `.rpm` пакета:

```bash
# Скачать ODA File Converter для Linux
# Перейдите на: https://www.opendesign.com/guestfiles/oda_file_converter
# Выберите вашу ОС (например: "QT5 Linux 64-bit DEB" или "RPM")

# Ubuntu / Debian — установка .deb:
sudo apt install -y ./ODAFileConverter_QT6_lnxX64_8.3dll_25.4.deb

# CentOS / Fedora — установка .rpm:
sudo dnf install -y ./ODAFileConverter_QT6_lnxX64_8.3dll_25.4.rpm
```

После установки убедитесь, что конвертер доступен в PATH:
```bash
which ODAFileConverter
# Если не находит — добавьте путь:
# По умолчанию: /usr/bin/ODAFileConverter или /opt/ODAFileConverter/ODAFileConverter

# Если установился в нестандартное место:
sudo ln -s /opt/ODAFileConverter/ODAFileConverter /usr/local/bin/ODAFileConverter
```

Проверка:
```bash
ODAFileConverter --help
```

> **Примечание:** Для работы ODA на headless Linux-сервере (без графического окружения) может потребоваться виртуальный дисплей:
> ```bash
> sudo apt install -y xvfb
> xvfb-run ODAFileConverter "/input" "/output" "ACAD2018" "DXF" "0" "1" "*.DWG"
> ```

### 4. Запуск на сервере (CLI)

На Linux-сервере без монитора используется командная строка (GUI недоступен):

```bash
cd /opt/equipment-counter
source venv/bin/activate

# Полный pipeline: DWG → DXF → JSON + PNG + VOR
python3 batch_equipment.py "/путь/к/чертежам" --png --vor --keep-converted

# Только парсинг DXF (без конвертации DWG)
python3 batch_equipment.py "/путь/к/dxf_файлам" --no-convert --png

# Один файл
python3 equipment_counter.py "/путь/к/файлу.dxf" --json result.json --csv result.csv --png

# Генерация ВОР отдельно
python3 vor_generator.py "/путь/к/dxf_папке" -o VOR_EO.docx
```

### 5. Запуск как systemd-сервис (опционально)

Если нужно запускать обработку по расписанию:

```bash
# Создать скрипт запуска
cat > /opt/equipment-counter/run.sh << 'EOF'
#!/bin/bash
cd /opt/equipment-counter
source venv/bin/activate
python3 batch_equipment.py "$1" --png --vor --keep-converted 2>&1 | tee -a /var/log/equipment-counter.log
EOF
chmod +x /opt/equipment-counter/run.sh

# Добавить в cron (например, каждый день в 03:00):
# crontab -e
# 0 3 * * * /opt/equipment-counter/run.sh "/путь/к/чертежам"
```

---

## ODA File Converter

[ODA File Converter](https://www.opendesign.com/guestfiles/oda_file_converter) — бесплатная утилита от Open Design Alliance для конвертации проприетарных файлов AutoCAD `.dwg` в открытый формат `.dxf`.

**Нужен только если исходные чертежи в формате `.dwg`.** Если у вас уже есть `.dxf` — конвертер не нужен.

### Скачать

https://www.opendesign.com/guestfiles/oda_file_converter

| Платформа | Формат |
|-----------|--------|
| Windows | `.exe` установщик |
| Linux (Ubuntu/Debian) | `.deb` пакет |
| Linux (CentOS/Fedora) | `.rpm` пакет |

### Где программа ищет конвертер

**Windows:**
```
C:\Program Files\ODA\ODAFileConverter 27.1.0\ODAFileConverter.exe
C:\Program Files\ODA\ODAFileConverter\ODAFileConverter.exe
C:\Program Files (x86)\ODA\ODAFileConverter\ODAFileConverter.exe
```

**Linux:**
```
/usr/bin/ODAFileConverter          (стандартный путь после установки deb/rpm)
/usr/local/bin/ODAFileConverter    (симлинк)
```

Программа также проверяет системный `PATH` (`which ODAFileConverter`).

### Вызов вручную

```bash
ODAFileConverter "/input_dir" "/output_dir" "ACAD2018" "DXF" "0" "1" "*.DWG"
```

| Параметр | Значение |
|----------|----------|
| input_dir | Папка с DWG файлами |
| output_dir | Папка для DXF файлов |
| ACAD2018 | Целевая версия формата |
| DXF | Выходной формат |
| 0 | Не рекурсивно |
| 1 | Audit (проверка целостности) |
| *.DWG | Фильтр файлов |

---

## Использование

### GUI (Windows)

1. Запустите `equipment_counter.bat` или `python equipment_gui.py`
2. Нажмите **"Обзор"** и выберите папку с чертежами
3. Настройте параметры:
   - **Конвертация DWG** — конвертировать DWG в DXF (требуется ODA)
   - **Парсинг DXF** — анализировать DXF файлы
   - **Парсинг PDF** — анализировать PDF файлы
   - **Генерация PNG** — создать визуализации с маркерами оборудования
   - **Генерация ВОР** — создать ведомость объёмов работ (.docx)
4. Нажмите **"Запуск"**
5. Результаты сохраняются в выбранной папке

### Командная строка

#### Полный pipeline (основной режим)

```bash
python batch_equipment.py "путь/к/чертежам" [опции]
```

| Опция | Описание |
|-------|----------|
| `--png` | Генерировать PNG с маркерами оборудования |
| `--png-dpi N` | DPI для PNG (по умолчанию: 200) |
| `--vor` | Генерировать ВОР документ (.docx) |
| `--no-convert` | Пропустить конвертацию DWG → DXF |
| `--no-dxf` | Не парсить DXF файлы |
| `--no-pdf` | Не парсить PDF файлы |
| `--keep-converted` | Сохранить сконвертированные DXF файлы |
| `-o / --output` | Путь для JSON-отчёта |

#### Одиночный файл

```bash
python equipment_counter.py "файл.dxf" --json result.json --csv result.csv --png
```

#### Генерация ВОР отдельно

```bash
python vor_generator.py "путь/к/dxf_папке" -o VOR_EO.docx --project "Название проекта"
```

### Примеры

```bash
# Полная обработка папки с DWG: конвертация + парсинг + PNG + ВОР
python batch_equipment.py "D:\Проект\Чертежи" --png --vor --keep-converted

# Только парсинг существующих DXF файлов, без конвертации
python batch_equipment.py "D:\Проект\DXF" --no-convert --png

# Обработка на Linux-сервере
python3 batch_equipment.py "/data/drawings/section_3" --png --vor

# Один DXF файл → JSON + CSV
python equipment_counter.py "план_освещения.dxf" --json report.json --csv report.csv
```

---

## Выходные файлы

| Файл | Формат | Описание |
|------|--------|----------|
| `equipment_report.json` | JSON | Отчёт по всем файлам: список оборудования с подсчётом и кабели |
| `*.png` | PNG | Визуализация чертежа с цветными маркерами найденного оборудования |
| `ВОР_ЭО.docx` | DOCX | Ведомость объёмов работ (оборудование по категориям высот) |
| `_converted_dxf/` | Папка | Сконвертированные DXF файлы (при `--keep-converted`) |

### Структура equipment_report.json

```json
{
  "файл.dxf": {
    "equipment": [
      {
        "symbol": "1",
        "name": "Светильник SLICK.PRS LED 50 5000K",
        "count": 33,
        "count_ae": 17
      }
    ],
    "cables": [
      {
        "type": "ВВГнг(А)-LSLTx 3x2.5",
        "length_m": 125.0
      }
    ]
  }
}
```

---

## Структура проекта

```
EquipmentCounter/
  equipment_counter.py    # Ядро парсера (DXF/PDF → оборудование + кабели)
  batch_equipment.py      # Пакетный pipeline (сканирование → конвертация → парсинг → отчёт)
  equipment_gui.py        # GUI интерфейс (Tkinter)
  dxf_visualizer.py       # Генерация PNG визуализаций с маркерами
  pdf_overlay.py          # Наложение маркеров на PDF-рендер
  vor_generator.py        # Генерация ВОР документов (.docx)
  equipment_counter.bat   # Windows-лаунчер (двойной клик)
  requirements.txt        # Python-зависимости
  README.md               # Документация
```

### Поддерживаемые форматы

| Формат | Оборудование | Кабели | Примечание |
|--------|-------------|--------|-----------|
| `.dwg` | Да | Да | Требуется ODA File Converter |
| `.dxf` | Да | Да | Нативная поддержка |
| `.pdf` | Да | Нет | Только оборудование из легенды |

---

## Решение проблем

| Проблема | Решение |
|----------|---------|
| `python: command not found` | Установите Python 3.10+ и добавьте в PATH |
| `ODA File Converter не найден` | Установите ODA или используйте готовые DXF файлы |
| Ошибки с кириллицей в путях | Поддерживается нативно, проблем быть не должно |
| Пустой результат (0 оборудования) | Убедитесь, что чертёж содержит блок "Условные обозначения" |
| PNG не генерируются | Добавьте флаг `--png` при запуске |
| `ModuleNotFoundError: docx` | Выполните `pip install python-docx` |
| ODA не работает на headless Linux | Установите `xvfb`: `sudo apt install xvfb`, запуск через `xvfb-run` |
| `matplotlib` ошибки на Linux без GUI | Установите backend: `export MPLBACKEND=Agg` перед запуском |

---

## Лицензия

Проприетарное ПО. Все права защищены.
