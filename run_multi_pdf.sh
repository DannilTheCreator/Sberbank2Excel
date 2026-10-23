#!/bin/bash
# Удобный shell скрипт для конвертации нескольких PDF в Excel
# Использование: ./run_multi_pdf.sh [опции] [вывод.xlsx]

# Определяем путь к проекту
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
export PYTHONPATH="$SCRIPT_DIR/src:$PYTHONPATH"

# Проверяем Python
if ! command -v python3 &> /dev/null; then
    echo "❌ Python3 не найден. Пожалуйста, установите Python 3"
    exit 1
fi

# Выводим помощь
if [ "$1" == "--help" ] || [ "$1" == "-h" ]; then
    echo "🚀 Конвертация нескольких PDF в Excel"
    echo ""
    echo "Использование:"
    echo "  $0                              # Все PDF из текущей папки"
    echo "  $0 result.xlsx                  # В файл result.xlsx"
    echo "  $0 -s result.xlsx               # Разные листы для каждого PDF"
    echo "  $0 -d /path/to/pdfs result.xlsx # Из конкретной папки"
    echo ""
    echo "Опции:"
    echo "  -s, --separate    Каждый PDF в отдельном листе"
    echo "  -d, --dir DIR     Папка с PDF файлами"
    echo "  --help            Показать эту справку"
    echo ""
    echo "Примеры:"
    echo "  # Все PDF из текущей папки в один Excel:"
    echo "  $0 combined.xlsx"
    echo ""
    echo "  # Все PDF в разные листы:"
    echo "  $0 -s combined.xlsx"
    echo ""
    echo "  # PDF из папки ./statements:"
    echo "  $0 -d ./statements combined.xlsx"
    exit 0
fi

# Запускаем Python скрипт
python3 "$SCRIPT_DIR/multi_pdf_to_excel.py" "$@"
