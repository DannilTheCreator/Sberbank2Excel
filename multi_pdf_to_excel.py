#!/usr/bin/env python3
"""
Скрипт для конвертации нескольких PDF файлов банка в один Excel файл
Объединяет все операции из разных PDF в один лист или разные листы

Использование:
    python multi_pdf_to_excel.py output.xlsx                  # Все PDF из текущей папки
    python multi_pdf_to_excel.py output.xlsx *.pdf            # Конкретные PDF файлы
    python multi_pdf_to_excel.py -d /path/to/pdfs output.xlsx # Из конкретной папки
    python multi_pdf_to_excel.py --separate output output.xlsx # Разные листы для каждого PDF
"""

import sys
import os
import argparse
import glob
from pathlib import Path
from typing import List, Tuple

# Добавляем путь к модулям
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from Sberbank2Excel.pdf2txtev import pdf_2_txt_file
from Sberbank2Excel.sberbankPDFtext2Excel import sberbankPDFtext2Excel
import pandas as pd


def process_single_pdf(pdf_file: str, temp_dir: str = None) -> Tuple[str, pd.DataFrame]:
    """
    Обрабатывает один PDF файл и возвращает DataFrame

    Args:
        pdf_file: Путь к PDF файлу
        temp_dir: Директория для временных TXT файлов

    Returns:
        Кортеж (имя_файла, DataFrame с данными)
    """
    if temp_dir is None:
        temp_dir = os.path.dirname(pdf_file)

    if not os.path.exists(pdf_file):
        print(f"❌ Файл не найден: {pdf_file}")
        return None, None

    # Генерируем имя для временного TXT файла
    base_name = os.path.basename(pdf_file).replace('.pdf', '')
    txt_file = os.path.join(temp_dir, f"{base_name}_temp.txt")

    try:
        print(f"📄 Обрабатываем: {os.path.basename(pdf_file)}...")

        # Шаг 1: Конвертируем PDF → TXT
        print(f"   1️⃣  PDF → TXT...", end=' ')
        pdf_2_txt_file(pdf_file, txt_file)
        print("✅")

        # Шаг 2: Читаем TXT файл
        print(f"   2️⃣  Парсим TXT...", end=' ')
        with open(txt_file, 'r', encoding='utf-8') as f:
            file_text = f.read()

        # Шаг 3: Используем существующий парсер для получения DataFrame
        from Sberbank2Excel.extractors_generic import determine_extractor_auto

        extractor_type = determine_extractor_auto(file_text)
        actual_extractor = extractor_type(file_text)

        # Получаем все записи
        entries = actual_extractor.split_text_on_entries()
        records = []
        for entry in entries:
            try:
                record = actual_extractor.decompose_entry_to_dict(entry)
                records.append(record)
            except Exception as e:
                print(f"\n   ⚠️  Ошибка при парсинге записи: {e}")
                continue

        # Создаем DataFrame
        df = pd.DataFrame(records)

        # Переименовываем колонки
        columns_info = actual_extractor.get_columns_info()
        df = df.rename(columns=columns_info)
        df = df[list(columns_info.values())]

        # Добавляем колонку с именем файла
        df.insert(0, 'Файл', os.path.basename(pdf_file))

        print("✅")
        print(f"   📊 Операций найдено: {len(df)}")

        # Очищаем временный файл
        try:
            os.remove(txt_file)
        except:
            pass

        return os.path.basename(pdf_file), df

    except Exception as e:
        print(f"\n❌ Ошибка при обработке {pdf_file}: {e}")
        # Очищаем временный файл
        try:
            os.remove(txt_file)
        except:
            pass
        return None, None


def process_multiple_pdfs(pdf_files: List[str],
                          output_file: str,
                          separate_sheets: bool = False,
                          temp_dir: str = None) -> bool:
    """
    Обрабатывает несколько PDF файлов и объединяет в один Excel

    Args:
        pdf_files: Список PDF файлов
        output_file: Имя выходного Excel файла
        separate_sheets: Если True, каждый PDF в отдельном листе; если False, все в одном
        temp_dir: Директория для временных файлов

    Returns:
        True если успешно, False если ошибка
    """
    if not pdf_files:
        print("❌ Не найдено PDF файлов для обработки")
        return False

    print(f"\n{'='*60}")
    print(f"🚀 Начинаем обработку {len(pdf_files)} PDF файлов")
    print(f"{'='*60}\n")

    all_dataframes = {}
    total_records = 0

    # Обрабатываем каждый PDF
    for pdf_file in pdf_files:
        file_name, df = process_single_pdf(pdf_file, temp_dir)

        if df is not None and not df.empty:
            all_dataframes[file_name] = df
            total_records += len(df)

    if not all_dataframes:
        print("\n❌ Не удалось обработать ни один PDF файл")
        return False

    print(f"\n{'='*60}")
    print(f"📝 Объединяем данные...")

    # Объединяем все данные
    try:
        if separate_sheets:
            # Каждый PDF в отдельном листе
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                for file_name, df in all_dataframes.items():
                    # Убираем расширение и ограничиваем длину имени листа (max 31 символ)
                    sheet_name = file_name.replace('.pdf', '')[:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"   ✅ Лист '{sheet_name}': {len(df)} операций")
        else:
            # Все в одном листе
            combined_df = pd.concat(list(all_dataframes.values()), ignore_index=True)

            # Сортируем по дате
            if 'Дата операции' in combined_df.columns:
                combined_df = combined_df.sort_values('Дата операции')

            combined_df.to_excel(output_file, sheet_name='Все операции', index=False)
            print(f"   ✅ Объединено в один лист: {len(combined_df)} операций")

        print(f"\n{'='*60}")
        print(f"✅ Успешно создан файл: {output_file}")
        print(f"   📊 Всего операций: {total_records}")
        print(f"   📄 Количество PDF: {len(all_dataframes)}")
        print(f"   💾 Размер файла: {os.path.getsize(output_file):,} байт")
        print(f"{'='*60}\n")

        return True

    except Exception as e:
        print(f"\n❌ Ошибка при создании Excel файла: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(
        description='Конвертирует несколько PDF выписок банка в один Excel файл',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры использования:

  # Все PDF из текущей папки в один лист:
  python multi_pdf_to_excel.py output.xlsx

  # Все PDF из текущей папки в разные листы:
  python multi_pdf_to_excel.py -s output.xlsx

  # Конкретные PDF файлы:
  python multi_pdf_to_excel.py file1.pdf file2.pdf file3.pdf output.xlsx

  # Из конкретной папки:
  python multi_pdf_to_excel.py -d ./pdfs output.xlsx

  # Из папки в разные листы:
  python multi_pdf_to_excel.py -d ./pdfs -s output.xlsx
        """
    )

    parser.add_argument('files', nargs='*',
                       help='PDF файлы или папка с PDF (если не указано, используются все PDF из текущей папки)')
    parser.add_argument('-o', '--output', dest='output',
                       help='Имя выходного Excel файла (по умолчанию: combined.xlsx)')
    parser.add_argument('-d', '--directory', dest='directory',
                       help='Папка с PDF файлами (вместо указания файлов)')
    parser.add_argument('-s', '--separate', action='store_true',
                       help='Каждый PDF в отдельном листе (по умолчанию: все в одном листе)')
    parser.add_argument('--temp-dir', dest='temp_dir',
                       help='Папка для временных файлов (по умолчанию: текущая папка)')

    args = parser.parse_args()

    # Определяем PDF файлы для обработки
    pdf_files = []

    if args.directory:
        # Ищем все PDF в указанной папке
        pattern = os.path.join(args.directory, '*.pdf')
        pdf_files = sorted(glob.glob(pattern))
        output_file = args.output or 'combined.xlsx'

    elif args.files:
        # Проверяем содержит ли последний файл расширение .xlsx
        if args.files[-1].endswith('.xlsx'):
            output_file = args.files[-1]
            pdf_files = args.files[:-1]
        else:
            pdf_files = args.files
            output_file = args.output or 'combined.xlsx'
    else:
        # Ищем все PDF в текущей папке
        pdf_files = sorted(glob.glob('*.pdf'))
        output_file = args.output or 'combined.xlsx'

    if not pdf_files:
        print("❌ PDF файлы не найдены")
        print("Проверьте:")
        print("  - Находятся ли PDF файлы в текущей папке")
        print("  - Или используйте параметр -d для указания папки")
        parser.print_help()
        sys.exit(1)

    # Убедимся что выходной файл имеет расширение .xlsx
    if not output_file.endswith('.xlsx'):
        output_file += '.xlsx'

    # Проверяем наличие временной директории
    temp_dir = args.temp_dir
    if temp_dir and not os.path.exists(temp_dir):
        os.makedirs(temp_dir)

    # Обрабатываем PDF файлы
    success = process_multiple_pdfs(
        pdf_files,
        output_file,
        separate_sheets=args.separate,
        temp_dir=temp_dir
    )

    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
