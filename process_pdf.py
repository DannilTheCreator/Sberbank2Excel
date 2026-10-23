#!/usr/bin/env python3
"""
Скрипт для полной обработки: PDF → TXT → XLSX
"""
import sys
import os

# Добавляем путь к модулям
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from Sberbank2Excel.pdf2txtev import pdf_2_txt_file
from Sberbank2Excel.sberbankPDFtext2Excel import sberbankPDFtext2Excel

def process_pdf_to_excel(pdf_file, output_file=None):
    """
    Обрабатывает PDF файл:
    1. Конвертирует PDF в TXT
    2. Конвертирует TXT в XLSX
    """
    print(f"📄 Начинаем обработку {pdf_file}...")

    # Step 1: Конвертируем PDF в TXT
    txt_file = pdf_file.replace('.pdf', '.txt')
    print(f"\n1️⃣  Конвертируем PDF → TXT ({txt_file})...")
    pdf_2_txt_file(pdf_file, txt_file)
    print(f"   ✅ TXT файл создан: {txt_file}")

    # Step 2: Конвертируем TXT в XLSX
    if output_file is None:
        output_file = pdf_file.replace('.pdf', '_result')

    print(f"\n2️⃣  Конвертируем TXT → XLSX ({output_file}.xlsx)...")
    sberbankPDFtext2Excel(
        input_txt_file_name=txt_file,
        output_file_name=output_file,
        perform_balance_check=True,
        output_file_type='xlsx'
    )
    print(f"   ✅ XLSX файл создан: {output_file}.xlsx")

    print(f"\n✅ Обработка завершена!")
    return output_file

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Использование: python process_pdf.py <pdf_file> [output_file]")
        sys.exit(1)

    pdf_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None

    if not os.path.exists(pdf_file):
        print(f"❌ Файл не найден: {pdf_file}")
        sys.exit(1)

    process_pdf_to_excel(pdf_file, output_file)
