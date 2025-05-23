# Генератор Release Notes из JIRA CSV (v1.0)

Этот скрипт предназначен для автоматической генерации документа "Release Notes" (Заметки о выпуске) в формате `.docx` на основе данных, экспортированных из системы управления задачами JIRA в формате CSV.

## Основные возможности

*   **Чтение данных из CSV**: Импортирует задачи из CSV-файла, выгруженного из JIRA.
*   **Автоматическое определение версий**:
    *   Опционально определяет глобальную версию релиза (например, "2.3.3") из данных CSV, ища строки с маркером `(global)`.
    *   Опционально определяет версии для отдельных компонентов/микросервисов (например, "phobos-AFM: 1.2.0") на основе префиксов в поле `Fix Version/s` и выбирает максимальную версию для каждого.
*   **Гибкая конфигурация полей**: Файл `fields_mapping.json` позволяет детально настроить, какие поля из CSV включать в отчет, как их называть и стилизовать.
*   **Настройка стилей документа**: Файл `word_styles.json` позволяет управлять шрифтами, размерами, цветами и отступами для различных элементов генерируемого Word-документа.
*   **Структурированный отчет**:
    *   Генерирует титульную часть с заголовком и логотипом (опционально).
    *   Включает таблицу с версиями компонентов релиза.
    *   Формирует раздел "Перечень изменений" с группировкой задач по микросервисам и далее по типам задач (Bug, Story, Task и т.д.).
    *   Формирует раздел "Настройки системы" для задач, содержащих инструкции по установке, также сгруппированных по микросервисам.
*   **Сортировка**: Автоматическая сортировка микросервисов, типов задач и самих задач внутри групп (например, по приоритету).
*   **Логирование**: Ведение логов процесса выполнения в консоль и, опционально, в файл для отладки и мониторинга.
*   **Генерация в `.docx`**: Создает готовый к использованию документ Microsoft Word.

## Требования к системе

*   Python 3.8+
*   Основные библиотеки (детали в `requirements.txt`):
    *   `pandas` (для работы с CSV)
    *   `python-docx` (для генерации `.docx` файлов)
    *   `packaging` (для корректного сравнения версий)

## Установка

1.  **Клонируйте репозиторий (или скачайте архив):**
    ```bash
    git clone https://github.com/ваш_юзернейм/имя_вашего_репозитория.git
    cd имя_вашего_репозитория
    ```
2.  **Создайте и активируйте виртуальное окружение (рекомендуется):**
    ```bash
    python -m venv .venv
    ```
    *   Для Windows:
        ```bash
        .venv\Scripts\activate
        ```
    *   Для macOS/Linux:
        ```bash
        source .venv/bin/activate
        ```
3.  **Установите зависимости:**
    ```bash
    pip install -r requirements.txt
    ```

## Структура проекта
├── configs/ # Директория с конфигурационными файлами
│ ├── config.json # Основные настройки скрипта
│ ├── fields_mapping.json # Настройка полей из CSV для отчета
│ └── word_styles.json # Настройка стилей Word-документа
├── data/ # Рекомендуемая директория для входных CSV-файлов
│ └── jira_export.csv # Пример имени вашего CSV-файла
├── output/ # Директория для сгенерированных отчетов (создается автоматически)
├── src/ # Исходный код скрипта
│ ├── init.py
│ ├── config_loader.py
│ ├── csv_parser.py
│ ├── data_processor.py
│ ├── logger_config.py
│ └── report_generator.py
├── assets/ # Рекомендуемая директория для логотипа (если используется)
│ └── logo.png
├── main.py # Главный исполняемый файл скрипта
├── requirements.txt # Список зависимостей Python
└── README.md # Этот файл


## Конфигурация

Все настройки скрипта производятся через JSON-файлы в директории `configs/`.

### 1. `configs/config.json` (Основные настройки)

*   `input_csv_file`: Путь к входному CSV-файлу (относительно корня проекта, например, `"data/jira_export.csv"`).
*   `output_report_file_docx`: Шаблон имени выходного файла `.docx`. Можно использовать плейсхолдер `{global_release_version}` (например, `"output/Release_Notes_{global_release_version}.docx"`).
*   `auto_detect_global_version` (boolean): `true` для автоматического определения глобальной версии релиза из CSV (ищет формат `X.Y.Z (global)`), `false` для использования значения ниже.
*   `global_release_version` (string): Глобальная версия релиза. Используется, если `auto_detect_global_version` равно `false` или если версия не найдена автоматически. Подставляется в имя файла и заголовок отчета.
*   `report_title_template` (string): Шаблон заголовка отчета. Можно использовать плейсхолдер `{global_release_version}`.
*   `logo_path` (string): Путь к файлу логотипа (относительно корня проекта, например, `"assets/logo.png"`). Оставьте пустым `""` или `null`, если логотип не нужен.
*   `csv_encoding` (string): Кодировка вашего CSV-файла (например, `"utf-8"`, `"windows-1251"`).
*   `csv_delimiter` (string): Разделитель полей в CSV-файле (например, `","`, `";"`).
*   `microservice_source_field_csv` (string): Название колонки в CSV, содержащей версии компонентов/микросервисов (например, `"Fix Version/s"`). Если JIRA создает несколько колонок с этим именем (например, "Fix Version/s", "Fix Version/s.1"), укажите здесь базовое имя.
*   `microservice_prefix_mapping` (object): Словарь для сопоставления префиксов версий (из `microservice_source_field_csv`) с полными именами микросервисов. Пример: `{"AM": "phobos-AFM", "IN": "phobos-integration"}`.
*   `auto_detect_component_versions` (boolean): `true` для автоматического определения версий компонентов из CSV для таблицы версий в отчете, `false` для использования списка ниже.
*   `microservices_versions_for_table` (array of objects): Список версий компонентов для таблицы в отчете. Каждый объект: `{"microservice": "Имя Сервиса", "version": "X.Y.Z"}`. Используется, если `auto_detect_component_versions` равно `false`.
*   `sort_microservices_by` (string): Порядок сортировки микросервисов (`"name_asc"` - по алфавиту, `"name_desc"` - в обратном порядке).
*   `sort_issue_types_order` (array of strings): Желаемый порядок отображения типов задач (например, `["Bug", "Story", "Task"]`). Остальные типы будут добавлены по алфавиту.
*   `sort_tasks_within_group_by` (string): `internal_name` поля, по которому будут сортироваться задачи внутри каждой группы (например, `"priority_val"` или `"issue_key"`).
*   `priority_order` (array of strings): Порядок приоритетов от самого высокого к самому низкому (например, `["Highest", "High", "Medium", "Low"]`). Используется, если `sort_tasks_within_group_by` указывает на поле приоритета.
*   `report_section_titles` (object): Тексты для заголовков разделов отчета (например, `"main_changes": "Перечень изменений"`, `"system_setup": "Настройки системы"`, `"no_changes_text": "Изменений нет"`).
*   `links_label_text` (string, опционально): Текст, добавляемый перед ссылками в описании задачи (по умолчанию "реализовано в рамках").

### 2. `configs/fields_mapping.json` (Настройка полей)

Это JSON-массив, где каждый элемент – объект, описывающий одно поле из CSV:

*   `csv_header` (string): Точное название колонки в вашем CSV-файле.
*   `internal_name` (string): Уникальное внутреннее имя для этого поля, используемое в коде. Рекомендуется использовать `snake_case`. Если не указано, используется `csv_header`.
*   `report_label` (string, опционально): Текстовая метка, которая может отображаться перед значением поля в отчете (зависит от настроек стиля поля).
*   `display_in_changes` (boolean, опционально): `true`, если поле должно отображаться в разделе "Перечень изменений". По умолчанию `false`.
*   `changes_order` (integer, опционально): Порядок отображения этого поля относительно других полей для одной задачи в "Перечне изменений". Меньшие числа отображаются раньше.
*   `changes_style` (object, опционально): Настройки стиля для этого поля в "Перечне изменений":
    *   `"bold"` (boolean): Жирный шрифт.
    *   `"italic"` (boolean): Курсив.
    *   `"prefix"` (string): Текст перед значением поля.
    *   `"suffix"` (string): Текст после значения поля (например, `": "`).
    *   `"new_line_before"` (boolean): Начинать ли вывод этого поля с новой строки (в рамках одного элемента списка задачи).
    *   `"multiline"` (boolean): Обрабатывать ли значение поля как многострочный текст (актуально для описаний, инструкций).
*   `display_in_setup` (boolean, опционально): `true`, если поле должно отображаться в разделе "Настройки системы". По умолчанию `false`.
*   `setup_order` (integer, опционально): Порядок отображения в "Настройках системы".
*   `setup_style` (object, опционально): Аналогично `changes_style`, но для раздела "Настройки системы".

**Важно:** Для корректной работы скрипта `fields_mapping.json` должен содержать описания для ключевых полей, таких как "Issue key", "Summary", "Issue Type", "Priority", "Custom field (Description for the customer)", "Custom field (Инструкция по установке)", "Links" (или как они называются в вашем CSV), даже если вы не планируете их все отображать напрямую. Скрипт использует `internal_name` этих полей для своей внутренней логики. Также не забудьте добавить специальное поле `"internal_name": "task_report_text"` (без `csv_header`), если хотите управлять его отображением через этот конфиг.

### 3. `configs/word_styles.json` (Настройка стилей документа)

Определяет шрифты, размеры, отступы и цвета для различных элементов Word-документа.

*   `fonts`: Объект со шрифтами для `default`, `title`, `heading1`, `heading2`, `heading3`.
*   `font_sizes`: Объект с размерами шрифтов в пунктах (числа) для `default`, `title`, `heading1`, `heading2`, `heading3`, `task_item`, `table_header`, `table_content`.
*   `paragraph_spacing`: Объект с отступами после параграфов в пунктах (числа) для `after_title`, `after_heading1`, `after_heading2`, `after_heading3`, `list_item_before`, `list_item_after`.
*   `colors_hex`: Объект с цветами в HEX-формате (без символа `#`), например, `"table_header_background": "D9E1F2"`.
*   `table_properties`: Объект с настройками таблиц, например, `"width_col1_percent": 40`.

## Подготовка CSV-файла

*   Экспортируйте необходимые задачи из JIRA в формат CSV.
*   **Кодировка**: Убедитесь, что кодировка файла соответствует значению `csv_encoding` в `config.json` (рекомендуется UTF-8).
*   **Разделитель**: Убедитесь, что разделитель полей соответствует `csv_delimiter` в `config.json`.
*   **Названия колонок**: Названия колонок в CSV должны точно совпадать со значениями `csv_header` в файле `fields_mapping.json` для тех полей, которые вы хотите использовать.
*   **Поле с версиями (`Fix Version/s`)**: Если одна задача затрагивает несколько версий, и JIRA при экспорте создает несколько колонок с одинаковым базовым именем (например, "Fix Version/s", "Fix Version/s.1", "Fix Version/s.2"), то в `config.json` в поле `microservice_source_field_csv` укажите это базовое имя (т.е. "Fix Version/s").

## Запуск скрипта

1.  Убедитесь, что вы находитесь в **корневой директории проекта**.
2.  Активируйте ваше виртуальное окружение (если вы его используете).
3.  Выполните команду:
    ```bash
    python main.py
    ```
4.  Сгенерированный `.docx` файл будет сохранен в директорию, указанную в `output_report_file_docx` в `config.json` (по умолчанию, это папка `output/` в корне проекта).
5.  Логи выполнения будут выводиться в консоль. Если в `src/logger_config.py` указан `LOG_FILE_NAME`, логи также будут записываться в соответствующий файл (по умолчанию, `release_notes_generator.log` в корне проекта). Уровень логирования настраивается в `src/logger_config.py`.

## Устранение распространенных проблем

*   **`FileNotFoundError`**: Проверьте правильность путей к CSV-файлу, файлу логотипа в `config.json` и путей к конфигурационным файлам в `src/config_loader.py` (хотя последние должны работать с текущей структурой). Убедитесь, что файлы действительно существуют по указанным путям.
*   **`ImportError`**: Убедитесь, что вы правильно установили все зависимости из `requirements.txt` в активном виртуальном окружении. Запускайте скрипт (`python main.py`) из корневой директории проекта.
*   **Неправильное определение микросервисов или версий**:
    *   Проверьте `microservice_prefix_mapping` в `config.json`.
    *   Проверьте формат данных в колонке `Fix Version/s` вашего CSV-файла.
    *   Установите `LOG_LEVEL_CONSOLE` и `LOG_LEVEL_FILE` в `logging.DEBUG` в файле `src/logger_config.py` для получения более детальной информации о процессе извлечения.
*   **Проблемы со шрифтами в Word-документе**: Убедитесь, что шрифты, указанные в `word_styles.json`, установлены в вашей операционной системе.
*   **Не отображаются какие-то поля задачи в отчете**: Проверьте соответствующие флаги (`display_in_changes`, `display_in_setup`) и порядок (`changes_order`, `setup_order`) в `fields_mapping.json`.

---