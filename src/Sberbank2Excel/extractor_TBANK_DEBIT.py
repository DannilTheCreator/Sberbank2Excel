import re
from datetime import datetime
import sys
from decimal import Decimal

from typing import Any

from Sberbank2Excel import exceptions
from Sberbank2Excel.utils import get_decimal_from_money, split_Sberbank_line
from Sberbank2Excel.extractor import Extractor
from Sberbank2Excel import extractors_generic


class TBANK_DEBIT(Extractor):

    def check_specific_signatures(self):

        test1 = re.search(r'сбербанк', self.bank_text, re.IGNORECASE)
        # print(f"{test1=}")

        test2 = re.search(r'Выписка по договору', self.bank_text, re.IGNORECASE)
        # print(f"{test2=}")

        test_ostatok = re.search(r'Баланс на', self.bank_text, re.IGNORECASE)

        if not test1  or not test2 or not test_ostatok:
            raise exceptions.InputFileStructureError("Не найдены паттерны, соответствующие выписке")

    def get_period_balance(self) -> Decimal:
        """
        функция ищет в тексте значения начального и конечного остатков
        и также суммы поступлений и расходов
        используется для контрольной проверки вычислений

        Пример текста
        ----------------------------------------------------------
        Баланс на 09.04.26        0.00 ₽
        •     Поступления       67 975.00 ₽
        •     Расходы   - 67 757.96 ₽
        •     Кэшбэк    62.00 ₽
        Баланс на 08.05.26      279.04 ₽
        ----------------------------------------------------------

        :return: Decimal
        """

        # Extract "Поступления" sum
        res_input = re.search(r'Поступления\s+([\d\s,.-]+)\s*₽', self.bank_text, re.MULTILINE)
        # Extract "Расходы" sum
        res_output = re.search(r'Расходы\s*-?\s*([\d\s,.-]+)\s*₽', self.bank_text, re.MULTILINE)

        if not res_input or not res_output:
            raise exceptions.InputFileStructureError(
                'Не найдена структура с остатками и пополнениями')

        summa_popolneniy = get_decimal_from_money(res_input.group(1).strip())
        summa_rashodov = get_decimal_from_money(res_output.group(1).strip())

        return summa_popolneniy - summa_rashodov

    def split_text_on_entries(self)->list[str]:
        """
        разделяет текстовый файл формата 2610 на отдельные записи

        Файл использует TAB как разделитель колонок:
        ДАТА_ОПЕ\tДАТА_ОБРАБОТКИ\tОПИСАНИЕ\tСУММА\tОСТАТОК

        Описание может переноситься на следующие строки (строки без TAB в начале)
        """
        lines = self.bank_text.split('\n')
        individual_entries = []

        # Найдём строку "Дата и время операции" - с неё начинаются операции
        start_idx = 0
        for i, line in enumerate(lines):
            if 'Дата и время' in line:
                # Пропускаем заголовок (может быть на 1-2 строках)
                start_idx = i + 1
                # Пропускаем строку "операции\tобработки..." если она есть
                if start_idx < len(lines) and '\tобработки' in lines[start_idx]:
                    start_idx += 1
                break

        for line in lines[start_idx:]:
            # Пропускаем пустые строки
            if not line.strip():
                continue

            # Прерываем на конец выписки
            if re.match(r'Баланс\s+на', line):
                break

            # Строки с TAB содержат основные данные операции
            if '\t' in line:
                # Проверяем начинается ли первая колонка с даты
                first_col = line.split('\t')[0].strip()
                if re.match(r'^\d\d\.\d\d\.\d\d', first_col):
                    # Это новая операция
                    individual_entries.append(line)
            else:
                # Это строка продолжения описания (без TAB)
                # Добавляем её к последней операции
                if individual_entries and line.strip():
                    # Помечаем что это продолжение
                    individual_entries[-1] += '\n' + line

        if len(individual_entries) == 0:
            raise exceptions.InputFileStructureError(
                "Не обнаружена ожидаемая структора данных: не найдено ни одной трансакции")

        return individual_entries

    def decompose_entry_to_dict(self, entry:str)->dict:
        """
        Выделяем данные из одной TAB-разделённой записи в dictionary

        Структура (TAB-разделённая):
        ДАТА_ОПЕ\tДАТА_ОБРАБОТКИ\tОПИСАНИЕ\tСУММА_В_ВАЛЮТЕ_СЧЁТА\tОСТАТОК

        Может быть продолжение описания на следующих строках (без TAB)

        Пример:
        25.04.26\t25.04.26\tПополнение. Сбербанк Онлайн\t+ 3 500.00 ₽\t+ 3 500.00 ₽
        25.04.26 17:08\t25.04.26\tОплата в 1841 DETSKIY MIR SANKT-\t1 299.00 ₽\t1 299.00 ₽
        PETERBU RUS
        """
        # Разделяем на основную строку и продолжение
        entry_lines = entry.split('\n')
        main_line = entry_lines[0]

        columns = main_line.split('\t')

        if len(columns) < 4:
            raise exceptions.InputFileStructureError(
                "entry должна иметь как минимум 4 TAB-разделённых колонки\n" + str(entry))

        result: dict = {}

        # Колонка 0: Дата операции с опциональным временем
        op_date_str = columns[0].strip()
        date_match = re.match(r'^(\d\d\.\d\d\.\d\d)(?:\s+(\d\d:\d\d))?', op_date_str)

        if not date_match:
            raise exceptions.InputFileStructureError(
                "Не найдена дата операции в: " + op_date_str)

        date_str = date_match.group(1)
        time_str = date_match.group(2) if date_match.group(2) else "00:00"
        result['operation_date'] = datetime.strptime(date_str + " " + time_str, '%d.%m.%y %H:%M')

        # Колонка 1: Дата обработки
        proc_date_str = columns[1].strip() if len(columns) > 1 else date_str
        if proc_date_str:
            try:
                result['processing_date'] = datetime.strptime(proc_date_str, '%d.%m.%y')
            except:
                result['processing_date'] = result['operation_date']
        else:
            result['processing_date'] = result['operation_date']

        # Колонка 2: Описание (может быть продолжено на следующих строках)
        description = columns[2].strip() if len(columns) > 2 else ""

        # Добавляем продолжение из следующих строк
        if len(entry_lines) > 1:
            for continuation_line in entry_lines[1:]:
                if continuation_line.strip():
                    description += " " + continuation_line.strip()

        result['description'] = description.strip()
        result['category'] = description.split()[0] if description else ""

        # Колонка 3: Сумма в валюте счёта
        value_str = columns[3].strip() if len(columns) > 3 else "0.00"

        # Колонка 4: Остаток
        remainder_str = columns[4].strip() if len(columns) > 4 else value_str

        # Clean and parse currency values
        try:
            result['value_account_currency'] = get_decimal_from_money(value_str, True)
            result['remainder_account_currency'] = get_decimal_from_money(remainder_str)
        except Exception as e:
            raise exceptions.InputFileStructureError(
                f"Ошибка парсинга сумм. value_str='{value_str}', remainder_str='{remainder_str}': {str(e)}\nОпись: {entry}")

        # Default values for optional fields
        result['authorisation_code'] = '-'
        result['operational_currency'] = '₽'
        result['value_operational_currency'] = result['value_account_currency']

        return result

    def get_column_name_for_balance_calculation(self)->str:
        return 'value_account_currency'

    def get_columns_info(self)->dict:
        """
        Returns full column names in the order they shall appear in Excel
        The keys in dictionary shall correspond to keys of the result of the function self.decompose_entry_to_dict()
        """
        return {'operation_date': 'Дата операции',
                'processing_date': 'Дата обработки',
                'authorisation_code': 'Код авторизации',
                'description': 'Описание операции',
                'category': 'Категория',
                'value_account_currency': 'Сумма в валюте счёта',
                'value_operational_currency': 'Сумма в валюте операции',
                'operational_currency': 'Валюта операции',
                'remainder_account_currency': 'Остаток по счёту в валюте счёта'}


if __name__ == '__main__':

    if len(sys.argv) < 2:
        print('Не указано имя текстового файла для проверки экстрактора')
        print(__doc__)

    else:
        extractors_generic.debug_extractor(TBANK_DEBIT,
                                           test_text_file_name=sys.argv[1])
