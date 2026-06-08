# Electroproject

Десктопное приложение на Python (Tkinter) для инженерных расчётов:

- **Проверка на невозгорание** — кабельные линии, экспорт в Word
- **Расчёт СОПТ** — сопротивления, ТКЗ, карты селективности, экспорт в Word
- **Расчёт ТКЗ 0,4 кВ (ЩСН)** — модель сети из Excel, ТКЗ, чувствительность и карты селективности АВ, экспорт в Word
- **Проверка оборудования** — заглушка (в разработке)

## Требования

- Python 3.10+
- Зависимости из `requirements.txt`

## Запуск из исходников

```bash
pip install -r requirements.txt
python main.py
```

## Сборка EXE (Windows)

```bash
build_windows.bat
```

Скрипт установит зависимости из `requirements.txt` и соберёт exe через `pack.py`.
Запуск после сборки: `dist\RUN_THIS_EXE.bat` или `dist\Electroproject.exe`.

Тот же результат вручную:

```bash
pip install -r requirements.txt
python pack.py
```

Готовые файлы появятся в папке `dist/` (папка создаётся при сборке, в Git не входит).

## Структура проекта

```
electroproject/
├── main.py              # точка входа
├── app/
│   ├── core/            # контракты и реестр плагинов
│   ├── shell/           # главное окно
│   ├── shared/          # общие утилиты (Excel, Word, ошибки)
│   └── calculations/    # модули расчётов (fire_check, sopt, …)
├── requirements.txt
├── pack.py
└── build_windows.bat
```

Проекты пользователя сохраняются в Excel через диалог «Открыть / Сохранить»; последний путь запоминается в `~/.electroproject/`.
