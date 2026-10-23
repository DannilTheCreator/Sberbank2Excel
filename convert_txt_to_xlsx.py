#!/usr/bin/env python3
import sys
sys.path.insert(0, '/home/daniil/My-Python-projects/Sberbank2Excel/src')

from Sberbank2Excel.sberbankPDFtext2Excel import sberbankPDFtext2Excel

print("🔄 Конвертируем TXT в XLSX (тестируем новый TBANK_DEBIT формат)...")
result_file = sberbankPDFtext2Excel(
    input_txt_file_name='/home/daniil/My-Python-projects/Sberbank2Excel/8542637896.txt',
    output_file_name='/home/daniil/My-Python-projects/Sberbank2Excel/tbank_result',
    format='auto',
    perform_balance_check=True,
    output_file_type='xlsx'
)
print(f"✅ Успешно создан файл: {result_file}.xlsx")
