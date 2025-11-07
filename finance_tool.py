import os
import re
import sys
import warnings
import time
import winsound
import pandas as pd
from tabulate import tabulate
from rich.console import Console
from rich.table import Table
from rich.prompt import Prompt
from rich.live import Live
from rich.spinner import Spinner
from rich.panel import Panel
import msvcrt

warnings.simplefilter("ignore", UserWarning)
console = Console()

# === Настройки ===
base_dir = r"C:\Users\Исхан\ExcelData"
log_file = "log.txt"

if os.path.exists(log_file):
    os.remove(log_file)


def log(msg):
    """Записывает сообщение в log.txt и в консоль"""
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(msg + "\n")
    console.log(msg)


# === Попытка конвертации .xls → .xlsx (если Excel установлен) ===
def convert_xls_to_xlsx(file_path):
    """Пробует открыть .xls в Excel и сохранить как .xlsx"""
    try:
        import win32com.client
        excel = win32com.client.Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(file_path)
        new_path = file_path + "x"
        wb.SaveAs(new_path, FileFormat=51)  # 51 = xlsx
        wb.Close()
        excel.Quit()
        log(f"✅ Конвертирован в XLSX: {new_path}")
        return new_path
    except Exception as e:
        log(f"⚠️ Не удалось конвертировать {file_path}: {e}")
        return file_path


# === Универсальное чтение Excel (все листы, автоконверсия) ===
def try_read_excel(file_path):
    """Пробует прочитать ВСЕ листы Excel разными движками; при .xls — автоконверт."""
    ext = os.path.splitext(file_path)[1].lower()
    if ext in [".xlsx", ".xlsm"]:
        engines = ["openpyxl"]
    elif ext == ".xls":
        engines = ["xlrd"]
    else:
        engines = ["openpyxl", "xlrd"]

    for eng in engines:
        try:
            df_dict = pd.read_excel(file_path, engine=eng, sheet_name=None, dtype=str)
            if not df_dict:
                raise Exception("Пустая книга (нет листов)")
            df = pd.concat(df_dict.values(), ignore_index=True)
            return df
        except Exception as e:
            log(f"⚠️ {eng} не смог прочитать {file_path}: {e}")

    # попытка автоконвертации .xls → .xlsx
    if ext == ".xls":
        converted = convert_xls_to_xlsx(file_path)
        if converted != file_path and os.path.exists(converted):
            try:
                df_dict = pd.read_excel(converted, engine="openpyxl", sheet_name=None, dtype=str)
                if not df_dict:
                    raise Exception("Пустая книга (нет листов) после конвертации")
                df = pd.concat(df_dict.values(), ignore_index=True)
                return df
            except Exception as e:
                log(f"⚠️ После конвертации не удалось прочитать {converted}: {e}")

    log(f"❌ Не удалось прочитать {file_path}")
    return pd.DataFrame()


# === Надёжное чтение и обработка Excel ===
def smart_read_excel(file_path):
    """Читает Excel, очищает, объединяет, устраняет дубликаты"""
    try:
        df = try_read_excel(file_path)
        if df.empty:
            raise Exception("Файл пустой или не читается")

        df = df.reset_index(drop=True).astype(str)

        # Поиск заголовка по ключевым словам
        header_row = None
        keywords = ["фио", "сотруд", "долж", "подраздел", "остат", "дата", "работ", "совмест"]
        for i, row in df.head(20).iterrows():
            row_str = " ".join(str(x).lower() for x in row.values if x and x != "nan")
            if any(kw in row_str for kw in keywords):
                header_row = i
                break

        # Если не найдено ключевых признаков таблицы — пропускаем файл
        if header_row is None:
            first_block = " ".join(df.head(20).astype(str).stack().str.lower().tolist())
            if not any(k in first_block for k in keywords):
                return pd.DataFrame()

        # Переназначаем заголовки, если нашли
        if header_row is not None:
            new_header = df.iloc[header_row]
            df = df.drop(df.index[:header_row + 1])
            df.columns = new_header
            df = df.reset_index(drop=True)

        df = df.loc[:, ~df.columns.duplicated()]
        df = df.dropna(how="all").dropna(axis=1, how="all")
        df = df[~df.apply(lambda x: x.astype(str).str.lower().str.contains("итого|руковод|отдел|список").any(), axis=1)]

        # Склейка "хвостов" строк
        rows, buffer = [], None
        for _, row in df.iterrows():
            first_cell = str(row.iloc[0]).strip()
            non_empty_cells = sum(bool(str(x).strip()) for x in row)
            if first_cell:
                if buffer is not None:
                    rows.append(buffer)
                buffer = row.copy()
                continue
            if non_empty_cells == 0:
                continue
            if buffer is not None and non_empty_cells <= 2:
                for i in range(len(row)):
                    cell = str(row.iloc[i]).strip()
                    if cell:
                        buffer.iloc[i] = f"{buffer.iloc[i]} {cell}".strip()
            else:
                if buffer is not None:
                    rows.append(buffer)
                buffer = row.copy()
        if buffer is not None:
            rows.append(buffer)

        df_clean = pd.DataFrame(rows, columns=df.columns)
        df_clean = df_clean.apply(lambda col: col.map(lambda x: str(x).strip() if pd.notna(x) else ""))

        for c in df_clean.columns:
            if "дата" in c.lower():
                try:
                    df_clean[c] = pd.to_datetime(df_clean[c], errors="coerce").dt.date.astype(str).replace("NaT", "")
                except:
                    pass

        return df_clean

    except Exception as e:
        log(f"⚠️ Ошибка при чтении {file_path}: {e}")
        return pd.DataFrame()


# === Постраничный вывод таблицы ===
def rich_table(df, page_size=20):
    if df.empty:
        console.print("[red]Нет данных.[/]")
        return
    total = len(df)
    pages = (total // page_size) + (1 if total % page_size else 0)
    for i in range(pages):
        start, end = i * page_size, min((i + 1) * page_size, total)
        chunk = df.iloc[start:end]
        console.print(f"\n[bold cyan]--- Страница {i + 1}/{pages} ({start + 1}–{end} из {total}) ---[/]")
        table = Table(show_header=True, header_style="bold magenta", show_lines=True)
        for col in df.columns:
            table.add_column(col, overflow="fold")
        for _, row in chunk.iterrows():
            table.add_row(*[str(x) for x in row.values])
        console.print(table)
        if i < pages - 1:
            console.print("[dim]Enter — следующая страница, Esc — выход[/]")
            while True:
                key = msvcrt.getch()
                if key == b"\r":
                    break
                elif key == b"\x1b":
                    console.print("[yellow]Выход из просмотра таблицы[/]")
                    return
        else:
            console.print("[green]✅ Конец таблицы[/]")


# === Этап 1. Красивая анимация загрузки + интерактивное меню ===
panel = Panel.fit(
    "[bold cyan]🔍 Инициализация системы анализа Excel-файлов...[/]\n"
    "[dim]Подготовка компонентов, проверка каталогов...[/]\n\n"
    "[bold yellow]📎 Полезные ссылки:[/]\n"
    "  [link=file://C:/Users/Исхан/ExcelData]📂 Открыть папку с Excel-файлами[/link]\n"
    "  [link=file://C:/Users/Исхан/scripts/finance_tool.py]🧠 Открыть исходный скрипт[/link]\n"
    "  [link=file://C:/Users/Исхан/scripts/log.txt]📜 Посмотреть log-файл[/link]\n"
    "\n[dim]Кликните на ссылку мышкой, чтобы открыть.[/]",
    title="[white on blue] FINANCE TOOL [/] ",
    border_style="cyan",
)
console.print()
console.print(panel)
time.sleep(1)

excel_exts = {".xlsx", ".xls", ".xlsm"}
name_keywords = ["6.4", "перечень", "список", "работник", "сотрудник", "остат", "отпуск", "штат"]

target_files = []
phrases = [
    "Ищу файлы с остатками отпусков...",
    "Анализирую структуру каталогов...",
    "Проверяю кодировки Excel...",
    "Отслеживаю скрытые файлы...",
    "Собираю данные по компаниям..."
]

spinner = Spinner("dots", text="Сканирование папки...")
start_time = time.time()

with Live(spinner, console=console, refresh_per_second=10):
    for root, _, files in os.walk(base_dir):
        spinner.text = f"[cyan]{phrases[len(target_files) % len(phrases)]}[/]"
        for file in files:
            ext = os.path.splitext(file)[1].lower()
            if ext in excel_exts:
                lname = file.lower()
                if any(k in lname for k in name_keywords):
                    target_files.append(os.path.join(root, file))
                else:
                    target_files.append(os.path.join(root, file))
        time.sleep(0.05)

elapsed = time.time() - start_time
spinner.text = "[green]✅ Анализ завершён![/]"
time.sleep(0.5)
winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
console.print(f"\n[green]⏱ Время сканирования:[/] {elapsed:.2f} сек\n")

if not target_files:
    console.print("[red]⚠️ Не найдено Excel-файлов[/]")
    sys.exit()

console.print(f"[bold green]📁 Найдено файлов:[/] {len(target_files)}\n")

# === Обработка Excel ===
combined_df = pd.DataFrame()
success, skipped = 0, 0

for file_path in target_files:
    if "6.1" in os.path.basename(file_path).lower():
        log(f"⏩ Пропущен файл (6.1): {file_path}")
        continue

    company = os.path.basename(os.path.dirname(file_path))
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".pdf":
        continue
    df = smart_read_excel(file_path)
    if df.empty:
        skipped += 1
        continue
    df.columns = [str(c).strip() for c in df.columns]
    fio_col = next((c for c in df.columns if "фио" in c.lower() or "сотруд" in c.lower()), df.columns[0])
    df["Компания"] = company
    df["Файл"] = file_path
    df["ФИО"] = df[fio_col].astype(str).str.strip()
    combined_df = pd.concat([combined_df, df], ignore_index=True)
    success += 1

console.print(f"\n[bold cyan]📊 Итог:[/]")
console.print(f"  [green]Успешно прочитано:[/] {success}")
console.print(f"  [yellow]Пропущено (пустые/нерелевантные):[/] {skipped}")
console.print(f"  [white]Всего файлов:[/] {len(target_files)}")

if combined_df.empty:
    console.print("[red]❌ Нет данных для анализа.[/]")
    sys.exit()


# === Команды ===
def print_commands():
    console.print("\n[bold magenta]📋 Команды:[/]")
    console.print("  [cyan]компании[/] — показать список компаний")
    console.print("  [cyan]фио[/] — поиск сотрудника по фамилии / имени")
    console.print("  [cyan]ошибки[/] — показать список файлов с ошибками")
    console.print("  [cyan]выход[/] — завершить\n")


def show_errors():
    """Показать список файлов с ошибками"""
    if not os.path.exists(log_file):
        console.print("[green]✅ Ошибок не найдено.[/]")
        return

    error_files = []
    with open(log_file, "r", encoding="utf-8") as f:
        for line in f:
            if "Ошибка при чтении" in line or "не смог прочитать" in line:
                match = re.search(r"([A-ZА-Яa-zа-я0-9_\\/:.\-\s]+\.xls[x]?)", line)
                if match:
                    path = match.group(1).strip()
                    company = os.path.basename(os.path.dirname(path))
                    error_files.append((company, path))

    if error_files:
        console.print(f"\n[bold red]⚠️ Найдено {len(error_files)} проблемных файлов:[/]")
        table = Table(show_header=True, header_style="bold red", show_lines=True)
        table.add_column("Компания", style="cyan")
        table.add_column("Путь к файлу", style="magenta")
        for comp, path in error_files:
            table.add_row(comp, path)
        console.print(table)
    else:
        console.print("[green]✅ Все файлы успешно прочитаны![/]")


def open_company(company_name):
    """Открытие и отображение данных по компании"""
    company_files = [f for f in target_files if os.path.basename(os.path.dirname(f)) == company_name]
    if not company_files:
        console.print("[red]❌ Файлы компании не найдены.[/]")
        return
    if len(company_files) == 1:
        selected_file = company_files[0]
    else:
        console.print(f"\n[bold green]📂 Найдено {len(company_files)} файлов для {company_name}:[/]")
        for i, f in enumerate(company_files, 1):
            console.print(f"[{i}] {f}")
        try:
            idx = int(Prompt.ask("\nВыберите файл (номер)"))
            selected_file = company_files[idx - 1]
        except:
            console.print("[red]❌ Неверный выбор файла.[/]")
            return

    console.print(f"\n✅ [green]Выбран файл:[/] {selected_file}")
    action = Prompt.ask("Введите действие (1 — открыть, 2 — таблица)", choices=["1", "2"], default="2", show_choices=False)
    df = smart_read_excel(selected_file)
    if action == "1":
        os.startfile(selected_file)
    elif action == "2":
        rich_table(df)


# === Основной цикл ===
print_commands()
companies = sorted(combined_df["Компания"].unique())

while True:
    cmd = Prompt.ask("\n[bold white]Введите команду или название компании[/]").strip().lower()

    if cmd == "компании":
        console.print(f"\n[bold cyan]Компании ({len(companies)}):[/]")
        for i, c in enumerate(companies, 1):
            console.print(f"[{i}] {c}")
        choice = Prompt.ask("\nВведите номер или название компании").strip().lower()
        if choice.isdigit():
            idx = int(choice) - 1
            if 0 <= idx < len(companies):
                open_company(companies[idx])
        elif choice:
            matches = [c for c in companies if choice in c.lower()]
            if matches:
                open_company(matches[0])

    elif cmd.startswith("фио"):
        query = cmd.replace("фио", "").strip().lower()
        if not query:
            query = Prompt.ask("Введите фамилию / имя").strip().lower()
        pattern = re.compile(rf"\b{re.escape(query)}\b", re.IGNORECASE)
        result = combined_df[combined_df["ФИО"].apply(lambda fio: bool(pattern.search(str(fio)) or query in str(fio).lower()))]
        if result.empty:
            console.print("[red]❌ Совпадений нет.[/]")
        else:
            console.print(f"\n[bold green]🔍 Найдено {len(result)} совпадений:[/]")
            rich_table(result[["ФИО", "Компания", "Файл"]])

    elif cmd == "ошибки":
        show_errors()

    elif cmd == "выход":
        console.print("[bold red]Завершение работы.[/]")
        break

    else:
        print_commands()
