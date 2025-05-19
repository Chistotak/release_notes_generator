import os
import sys
import json
from collections import OrderedDict

# --- Настройка sys.path и импорты ---
current_main_dir = os.path.dirname(os.path.abspath(__file__))
src_dir_path = os.path.join(current_main_dir, "src")
if src_dir_path not in sys.path:
    sys.path.insert(0, src_dir_path)

from src import logger_config

logger = logger_config.setup_logger(__name__)

from src import config_loader
from src import csv_parser
from src import data_processor
from src import report_generator


def pretty_print_json_for_debug(data, indent=2, ensure_ascii=False):
    if data:
        logger.debug(f"Содержимое JSON данных:\n{json.dumps(data, indent=indent, ensure_ascii=ensure_ascii)}")
    else:
        logger.debug("Нет данных JSON для вывода.")


if __name__ == "__main__":
    logger.info("--- Запуск JiraCsvReleaseNotesGenerator ---")
    project_root = os.path.dirname(os.path.abspath(__file__))
    config_directory_name = "configs"

    logger.info("--- Этап 1: Загрузка конфигураций ---")
    main_cfg, fields_cfg, styles_cfg = config_loader.get_all_configs(config_dir_name=config_directory_name)
    if not (main_cfg and fields_cfg and styles_cfg):
        logger.critical("--- Ошибка загрузки конфигураций. Завершение работы. ---");
        exit(1)
    logger.info("Все конфигурационные файлы успешно загружены.")

    logger.info("--- Этап 2: Чтение CSV-файла ---")
    input_csv_relative_path = main_cfg.get('input_csv_file')
    if not input_csv_relative_path:
        logger.critical("Путь к CSV не указан в config.json. Завершение.");
        exit(1)
    csv_full_path = os.path.join(project_root, input_csv_relative_path)
    # ... (остальной код чтения CSV как был) ...
    jira_dataframe_raw = csv_parser.load_csv_to_dataframe(file_path=csv_full_path,
                                                          encoding=main_cfg.get('csv_encoding', 'utf-8'),
                                                          delimiter=main_cfg.get('csv_delimiter', ','))
    if jira_dataframe_raw is None or jira_dataframe_raw.empty: logger.critical("Ошибка CSV. Завершение."); exit(1)
    logger.info(f"CSV успешно загружен. Строк: {len(jira_dataframe_raw)}")

    data_proc_shared_config = {
        "microservice_source_field_csv": main_cfg.get('microservice_source_field_csv'),
        "microservice_prefix_mapping": main_cfg.get('microservice_prefix_mapping', {})
    }

    global_release_ver = main_cfg.get('global_release_version', "N/A")
    if main_cfg.get("auto_detect_global_version", False):
        logger.info("--- Авто-определение глобальной версии ---")
        detected_gv = data_processor.detect_global_release_version(jira_dataframe_raw.copy(), {
            "microservice_source_field_csv": data_proc_shared_config["microservice_source_field_csv"]})
        if detected_gv:
            global_release_ver = detected_gv
        else:
            logger.warning(f"Не удалось авто-определить глоб.версию. Используется: '{global_release_ver}'")
    else:
        logger.info(f"Используется глоб.версия из config: '{global_release_ver}'")

    microservice_versions_for_table_data = []
    if main_cfg.get("auto_detect_component_versions", False):
        logger.info("--- Авто-определение версий компонентов ---")
        microservice_versions_for_table_data = data_processor.detect_component_versions_from_data(
            jira_dataframe_raw.copy(), data_proc_shared_config)
    else:
        microservice_versions_for_table_data = main_cfg.get('microservices_versions_for_table', [])
        logger.info("Используются версии компонентов из config.")

    logger.info("--- Этапы 3, 4: Обработка данных задач ---")
    processed_df_before_grouping = data_processor.process_initial_data(
        df_raw=jira_dataframe_raw.copy(),
        fields_mapping_config=fields_cfg,
        microservice_config=data_proc_shared_config,
        main_app_config=main_cfg  # <--- ПЕРЕДАЕМ main_cfg СЮДА
    )
    if processed_df_before_grouping.empty:
        logger.critical("Ошибка process_initial_data. Завершение.");
        exit(1)
    logger.info(f"Данные после process_initial_data. Задач: {len(processed_df_before_grouping)}")

    logger.info("--- Этап 5: Группировка для 'Перечня изменений' ---")
    sort_options_config = {  # ... (как было) ...
        "sort_microservices_by": main_cfg.get('sort_microservices_by', "name_asc"),
        "sort_issue_types_order": main_cfg.get('sort_issue_types_order', []),
        "sort_tasks_within_group_by": main_cfg.get('sort_tasks_within_group_by', "issue_key"),
        "priority_order": main_cfg.get('priority_order', [])
    }
    grouped_and_sorted_tasks_for_changes = data_processor.group_and_sort_tasks(
        processed_df=processed_df_before_grouping.copy(), sort_config=sort_options_config,
        fields_mapping_config=fields_cfg
    )
    if not grouped_and_sorted_tasks_for_changes:
        logger.warning("Нет данных для 'Перечня изменений'.")
        grouped_and_sorted_tasks_for_changes = OrderedDict()
    else:
        logger.info("Группировка для 'Перечня изменений' завершена.")

    logger.info("--- Этап 8: Подготовка для 'Настроек системы' ---")
    grouped_tasks_for_setup_section = data_processor.prepare_setup_instructions_data(
        processed_df_with_tasks=processed_df_before_grouping.copy(), fields_mapping_config=fields_cfg,
        sort_config=sort_options_config
    )
    if not grouped_tasks_for_setup_section:
        logger.warning("Нет данных для 'Настроек системы'.")
        grouped_tasks_for_setup_section = OrderedDict()
    else:
        logger.info("Подготовка для 'Настроек системы' завершена.")

    logger.info("--- Этапы 6, 7, 8: Генерация Word-документа ---")
    report_title_template = main_cfg.get('report_title_template',
                                         "Отчет по релизу версия {global_release_version}")  # ... (как было)
    report_title = report_title_template.format(global_release_version=global_release_ver)
    # ... (остальной код генерации отчета как был, он уже использует logger) ...
    logo_relative_path = main_cfg.get('logo_path')
    logo_full_abs_path = None
    if logo_relative_path:
        logo_full_abs_path = os.path.join(project_root, logo_relative_path)
        if not os.path.exists(logo_full_abs_path):
            logger.warning(f"Файл логотипа не найден: {logo_full_abs_path}");
            logo_full_abs_path = None

    output_file_template = main_cfg.get('output_report_file_docx', "output/ReleaseNotes_default.docx")
    safe_gv_filename = str(global_release_ver).replace("/", "-").replace("\\", "-").replace(":", "-").replace("*",
                                                                                                              "-").replace(
        "?", "-").replace("\"", "").replace("<", "").replace(">", "").replace("|", "").strip()
    if not safe_gv_filename: safe_gv_filename = "UNKNOWN_VERSION"; logger.warning(
        f"Глоб.версия ('{global_release_ver}') пустая. Имя файла: '{safe_gv_filename}'.")
    output_filename_docx = output_file_template.format(global_release_version=safe_gv_filename)
    output_full_path_docx = os.path.join(project_root, output_filename_docx)

    report_generator.generate_report_docx(
        output_filename=output_full_path_docx, report_title_text=report_title,
        logo_full_path=logo_full_abs_path, microservice_versions_list=microservice_versions_for_table_data,
        word_styles_config=styles_cfg, grouped_data_for_changes=grouped_and_sorted_tasks_for_changes,
        grouped_data_for_setup=grouped_tasks_for_setup_section, main_config_for_titles=main_cfg,
        fields_mapping_for_details=fields_cfg
    )
    logger.info(f"--- Генерация контента завершена. Документ: {output_full_path_docx} ---")