"""Короткие сообщения об ошибках для пользователя."""

MSG_WORD_SAVE_FAILED = (
    "Не удалось сохранить документ Word. "
    "Проверьте путь, права доступа и что файл не открыт в другой программе."
)
MSG_EXCEL_SAVE_FAILED = "Не удалось сохранить файл проекта Excel."
MSG_EXCEL_LOAD_FAILED = "Не удалось открыть файл проекта. Проверьте формат и путь."
MSG_MISSING_OPENPYXL = "Для работы с Excel установите пакет: pip install openpyxl"
MSG_MISSING_DOCX = "Для экспорта в Word установите пакет: pip install python-docx"
MSG_MISSING_PILLOW = "Для схем в Word установите пакет: pip install Pillow"
MSG_MISSING_MATPLOTLIB = "Для карт селективности установите пакет: pip install matplotlib"
