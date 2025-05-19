import pandas as pd
from collections import OrderedDict
from packaging.version import parse as parse_version
from . import logger_config  # Относительный импорт

logger = logger_config.setup_logger(__name__)

GLOBAL_VERSION_IDENTIFIER = "(GLOBAL)"


class _FieldNames:
    """Внутренний класс для хранения и доступа к стандартизированным именам колонок."""

    def __init__(self, fields_mapping_config: list):
        # Вспомогательная функция для поиска internal_name
        def get_name(standard_csv_header: str, default_internal: str, alt_internals: list = None):
            if alt_internals is None: alt_internals = []
            for field_spec in fields_mapping_config:
                # Сначала проверяем internal_name, если он явно задан и совпадает
                spec_internal = field_spec.get('internal_name')
                if spec_internal and (spec_internal == default_internal or spec_internal in alt_internals):
                    # Если нашли по internal_name, проверяем, есть ли для него csv_header,
                    # чтобы убедиться, что это поле вообще может прийти из CSV.
                    # Но для определения имени колонки это не так важно, как сам internal_name.
                    return spec_internal

                # Затем проверяем csv_header
                spec_csv_header = field_spec.get('csv_header', '').lower()
                if spec_csv_header == standard_csv_header.lower():
                    return field_spec.get('internal_name',
                                          field_spec.get('csv_header'))  # Возвращаем internal_name или csv_header

            logger.debug(
                f"Для стандартного заголовка '{standard_csv_header}' не найдено явного internal_name в маппинге. "
                f"Используется значение по умолчанию: '{default_internal}'.")
            return default_internal

        self.key = get_name('issue key', 'issue_key')
        self.summary = get_name('summary', 'summary_text', ['summary'])
        self.customer_desc = get_name('custom field (description for the customer)', 'description_for_customer')
        self.links = get_name('links', 'links_text')
        self.issue_type = get_name('issue type', 'type', ['issue_type'])
        self.priority = get_name('priority', 'priority_val', ['priority'])
        self.setup_instructions = get_name('custom field (инструкция по установке)', 'setup_instructions')
        self.fix_versions_display = get_name('fix version/s',
                                             'fix_versions_display_all')  # Для отображения сырых версий


def extract_single_microservice_and_version(version_code_str: str, prefix_mapping: dict):
    if not isinstance(version_code_str, str) or not version_code_str.strip():
        return None, None
    vc_str = version_code_str.strip()
    logger.debug(f"[extract_ms_ver] Вход: '{vc_str}'")
    for prefix, full_name in prefix_mapping.items():
        if vc_str.upper().startswith(prefix.upper()):
            if len(vc_str) > len(prefix):
                version_part = vc_str[len(prefix):].strip()
                logger.debug(
                    f"  [extract_ms_ver] Префикс '{prefix}' для '{vc_str}' -> Сервис: '{full_name}', Версия: '{version_part}'")
                return full_name, version_part
            else:
                logger.debug(f"  [extract_ms_ver] Префикс '{prefix}' для '{vc_str}', но версия отсутствует.")
                return full_name, None
    if vc_str and GLOBAL_VERSION_IDENTIFIER not in vc_str.upper():
        logger.debug(f"  [extract_ms_ver] Префикс не найден для '{vc_str}'")
    return None, None


def detect_component_versions_from_data(df_raw: pd.DataFrame, microservice_config: dict) -> list:
    if df_raw is None or df_raw.empty: return []
    fix_versions_base_csv_header = microservice_config.get('microservice_source_field_csv')
    prefix_mapping = microservice_config.get('microservice_prefix_mapping', {})
    if not fix_versions_base_csv_header:
        logger.warning("(detect_comp_ver): 'microservice_source_field_csv' не указан.")
        return []
    version_columns_in_df = sorted([col for col in df_raw.columns if col.startswith(fix_versions_base_csv_header)])
    if not version_columns_in_df:
        logger.warning(f"(detect_comp_ver): Колонки версий ('{fix_versions_base_csv_header}*') не найдены.")
        return []

    component_versions = {}
    for index, row in df_raw.iterrows():
        for v_col_name in version_columns_in_df:
            version_code_str = str(row.get(v_col_name, ""))
            if version_code_str:
                service_name, version_str = extract_single_microservice_and_version(version_code_str, prefix_mapping)
                if service_name and version_str:
                    try:
                        cleaned_version_str = version_str.split(" ")[0].split("(")[0]
                        current_version_obj = parse_version(cleaned_version_str)
                        if service_name not in component_versions or current_version_obj > component_versions[
                            service_name]:
                            component_versions[service_name] = current_version_obj
                    except Exception as e:
                        logger.warning(
                            f"(detect_comp_ver): Не удалось распарсить версию '{version_str}' (очищенная: '{cleaned_version_str}') для '{service_name}'. Ошибка: {e}")

    detected_versions_list = [{'microservice': name, 'version': str(ver_obj)} for name, ver_obj in
                              component_versions.items()]
    detected_versions_list.sort(key=lambda x: x['microservice'])
    if detected_versions_list:
        logger.info(f"Авто-определено версий компонентов для таблицы: {len(detected_versions_list)} шт.")
    else:
        logger.warning("Не удалось авто-определить версии компонентов для таблицы.")
    return detected_versions_list


def detect_global_release_version(df_raw: pd.DataFrame, microservice_config: dict) -> str | None:
    if df_raw is None or df_raw.empty: return None
    fix_versions_base_csv_header = microservice_config.get('microservice_source_field_csv')
    if not fix_versions_base_csv_header:
        logger.warning("(detect_global_ver): 'microservice_source_field_csv' не указан.")
        return None
    version_columns_in_df = sorted([col for col in df_raw.columns if col.startswith(fix_versions_base_csv_header)])
    if not version_columns_in_df:
        logger.warning(f"(detect_global_ver): Колонки версий ('{fix_versions_base_csv_header}*') не найдены.")
        return None

    found_global_versions = set()
    for index, row in df_raw.iterrows():
        for v_col_name in version_columns_in_df:
            version_code_str = str(row.get(v_col_name, "")).strip()
            if GLOBAL_VERSION_IDENTIFIER in version_code_str.upper():
                version_part = version_code_str.upper().split(GLOBAL_VERSION_IDENTIFIER)[0].strip()
                if version_part:
                    cleaned_version_part = version_part.replace('(', '').replace(')', '').strip()
                    if cleaned_version_part and any(char.isdigit() for char in cleaned_version_part):
                        found_global_versions.add(cleaned_version_part)
                    else:
                        logger.debug(
                            f"(detect_global_ver): Найдена '{version_code_str}', но извлеченная часть '{cleaned_version_part}' не версия.")
    if not found_global_versions:
        logger.warning(f"Глобальная версия релиза с '{GLOBAL_VERSION_IDENTIFIER}' не найдена.")
        return None
    if len(found_global_versions) > 1:
        first_global_ver = sorted(list(found_global_versions))[0]
        logger.warning(f"Найдено несколько глоб. версий: {found_global_versions}. Использована: {first_global_ver}")
        return first_global_ver
    global_version = found_global_versions.pop()
    logger.info(f"Авто-определена глобальная версия релиза: {global_version}")
    return global_version


def prepare_task_description_text(row: pd.Series, field_names: _FieldNames,
                                  links_label: str = "реализовано в рамках") -> str:
    description_text = ""
    customer_description = row.get(field_names.customer_desc, "")
    summary = row.get(field_names.summary, "")
    links_text_val = row.get(field_names.links, "")

    if customer_description and str(customer_description).strip():
        description_text = str(customer_description).strip()
    elif summary and str(summary).strip():
        description_text = str(summary).strip()
    else:
        description_text = "Нет описания."
    if links_text_val and str(links_text_val).strip():
        description_text += f" ({links_label}: {str(links_text_val).strip()})"
    return description_text


def process_initial_data(df_raw: pd.DataFrame, fields_mapping_config: list, microservice_config: dict,
                         main_app_config: dict) -> pd.DataFrame:  # Добавлен main_app_config
    if df_raw is None or df_raw.empty:
        logger.warning("(PIDs): Входной DataFrame пуст.")
        return pd.DataFrame()

    logger.info("Начало process_initial_data: отбор полей, извлечение МС, подготовка текста.")
    field_names = _FieldNames(fields_mapping_config)

    fix_versions_base_csv_header = microservice_config.get('microservice_source_field_csv')
    if not fix_versions_base_csv_header:
        logger.error("(PIDs): 'microservice_source_field_csv' не указан.")
        return pd.DataFrame()
    version_columns_in_df = sorted([col for col in df_raw.columns if col.startswith(fix_versions_base_csv_header)])
    if not version_columns_in_df: logger.warning(
        f"(PIDs): Колонки версий ('{fix_versions_base_csv_header}*') не найдены.")

    columns_to_process_map = {}
    original_headers_for_selection = []
    for spec in fields_mapping_config:
        csv_h = spec.get('csv_header')
        internal_n = spec.get('internal_name', csv_h)
        if csv_h and csv_h != fix_versions_base_csv_header:
            columns_to_process_map[csv_h] = internal_n
            if csv_h in df_raw.columns: original_headers_for_selection.append(csv_h)

    unique_original_headers = list(
        OrderedDict.fromkeys(h for h in original_headers_for_selection if h in df_raw.columns))

    if not unique_original_headers:
        key_csv_h = next((spec.get('csv_header') for spec in fields_mapping_config if
                          spec.get('internal_name') == field_names.key and spec.get('csv_header') in df_raw.columns),
                         field_names.key if field_names.key in df_raw.columns else None)
        if key_csv_h:
            processed_df = df_raw[[key_csv_h]].copy()
        else:
            logger.error(f"(PIDs): Ключ. колонка ('{field_names.key}') не найдена."); processed_df = pd.DataFrame(
                index=df_raw.index)
    else:
        processed_df = df_raw[unique_original_headers].copy()

    rename_map = {csv_col: internal_n for csv_col, internal_n in columns_to_process_map.items() if
                  csv_col in processed_df.columns}
    processed_df.rename(columns=rename_map, inplace=True)
    logger.debug(f"(PIDs): Колонки после отбора и переименования: {list(processed_df.columns)}")

    prefix_mapping = microservice_config.get('microservice_prefix_mapping', {})
    service_names_for_rows = []
    if version_columns_in_df:
        for idx, raw_row in df_raw.iterrows():
            current_row_services = set()
            for v_col in version_columns_in_df:
                val = str(raw_row.get(v_col, ""))
                if val:
                    service_name, _ = extract_single_microservice_and_version(val, prefix_mapping)
                    if service_name: current_row_services.add(service_name)
            service_names_for_rows.append(sorted(list(current_row_services)))
    else:
        service_names_for_rows = [[] for _ in range(len(processed_df))]

    if len(service_names_for_rows) == len(processed_df):
        processed_df['identified_microservices'] = service_names_for_rows
    else:
        logger.error(
            f"(PIDs): Несовпадение длин ({len(service_names_for_rows)} vs {len(processed_df)}) при доб. identified_microservices.")
        processed_df['identified_microservices'] = (service_names_for_rows + [[] for _ in range(len(processed_df))])[
                                                   :len(processed_df)]
    logger.info("Извлечение ИМЕН микросервисов для задач завершено.")

    links_label_text = main_app_config.get("links_label_text", "реализовано в рамках")  # Получаем из main_app_config
    processed_df['task_report_text'] = processed_df.apply(
        lambda row: prepare_task_description_text(row, field_names, links_label=links_label_text),
        axis=1
    )
    logger.info("Подготовка текста задачи для отчета завершена.")

    required_cols_map = {
        "issue type": field_names.issue_type, "priority": field_names.priority,
        "custom field (инструкция по установке)": field_names.setup_instructions,
        "summary": field_names.summary, "issue key": field_names.key
    }
    for std_csv_h, internal_n_val in required_cols_map.items():
        if internal_n_val and internal_n_val not in processed_df.columns:
            original_csv_h = next((spec.get('csv_header') for spec in fields_mapping_config if
                                   spec.get('internal_name') == internal_n_val), std_csv_h)
            if original_csv_h in df_raw.columns:
                processed_df[internal_n_val] = df_raw[original_csv_h]
                logger.debug(f"Добавлена колонка '{internal_n_val}' из CSV-колонки '{original_csv_h}'.")
            elif internal_n_val in df_raw.columns:
                processed_df[internal_n_val] = df_raw[internal_n_val]
                logger.debug(f"Добавлена колонка '{internal_n_val}' (т.к. имя совпадает с CSV).")
            else:
                logger.warning(
                    f"Не удалось добавить необходимую колонку '{internal_n_val}' (из CSV '{original_csv_h}' или '{std_csv_h}')")

    if field_names.fix_versions_display and version_columns_in_df:
        spec_for_fix_ver_display = next((s for s in fields_mapping_config if
                                         s.get('internal_name') == field_names.fix_versions_display or s.get(
                                             'csv_header') == fix_versions_base_csv_header), None)
        if spec_for_fix_ver_display and spec_for_fix_ver_display.get('include_in_task_details', False):
            processed_df[field_names.fix_versions_display] = df_raw.apply(lambda r: ", ".join(
                filter(None, [str(r.get(vc, "")) for vc in version_columns_in_df if str(r.get(vc, "")).strip()])),
                                                                          axis=1)
            logger.info(f"Добавлена колонка '{field_names.fix_versions_display}' с объединенными сырыми версиями.")

    logger.info("Финальная обработка данных (process_initial_data) завершена.")
    return processed_df


def group_and_sort_tasks(processed_df: pd.DataFrame, sort_config: dict, fields_mapping_config: list) -> OrderedDict:
    if processed_df.empty:
        logger.warning("(group_sort): Входной DataFrame пуст.")
        return OrderedDict()
    logger.info("Начало group_and_sort_tasks: группировка и сортировка для 'Перечня изменений'.")

    field_names = _FieldNames(fields_mapping_config)
    issue_type_col = field_names.issue_type
    priority_gs = field_names.priority
    key_col_gs = field_names.key

    if not issue_type_col or issue_type_col not in processed_df.columns:
        logger.error(f"(group_sort): Колонка типа задачи ('{issue_type_col}') отсутствует.");
        return OrderedDict()
    if not priority_gs or priority_gs not in processed_df.columns:
        logger.warning(
            f"(group_sort): Колонка приоритета ('{priority_gs}') отсутствует. Сортировка по приоритету не будет применена.");
        priority_gs = None

    expanded_tasks_list = []
    for index, task_row in processed_df.iterrows():
        task_dict = task_row.to_dict()
        microservices = task_row.get('identified_microservices', [])
        if not microservices: continue
        for ms_name in microservices:
            expanded_tasks_list.append({'microservice': ms_name, 'task_data': task_dict.copy()})
    if not expanded_tasks_list: logger.warning("(group_sort): Нет задач после расширения по МС."); return OrderedDict()

    temp_grouped = {}
    for item in expanded_tasks_list:
        ms_name, task_data = item['microservice'], item['task_data']
        task_type = task_data.get(issue_type_col, "Неизвестный тип")
        temp_grouped.setdefault(ms_name, {}).setdefault(task_type, []).append(task_data)

    priority_order_list = sort_config.get('priority_order', [])
    priority_map = {p.lower(): i for i, p in enumerate(priority_order_list)}
    sort_tasks_by_internal_name = sort_config.get('sort_tasks_within_group_by', key_col_gs)
    if sort_tasks_by_internal_name not in processed_df.columns:
        logger.warning(
            f"Поле для сортировки '{sort_tasks_by_internal_name}' отсутствует в DataFrame. Сортировка по '{key_col_gs}'.")
        sort_tasks_by_internal_name = key_col_gs

    for ms_name, types_dict in temp_grouped.items():
        for task_type, tasks_list in types_dict.items():
            # Проверка, что tasks_list не пустой перед доступом к tasks_list[0]
            if not tasks_list: continue

            if priority_gs and priority_map and sort_tasks_by_internal_name == priority_gs:
                tasks_list.sort(
                    key=lambda t: priority_map.get(str(t.get(priority_gs, "")).lower(), len(priority_map)))
            elif sort_tasks_by_internal_name in tasks_list[0]:
                tasks_list.sort(key=lambda t: str(t.get(sort_tasks_by_internal_name, '')))
            elif key_col_gs in tasks_list[0]:
                tasks_list.sort(key=lambda t: str(t.get(key_col_gs, '')))
            # Если ни одно из полей сортировки не найдено, задачи останутся в исходном порядке для этой группы

    issue_type_order_list = sort_config.get('sort_issue_types_order', [])
    sorted_grouped_tasks = OrderedDict()
    sorted_microservice_names = sorted(temp_grouped.keys())
    if sort_config.get('sort_microservices_by') == 'name_desc': sorted_microservice_names.reverse()
    for ms_name in sorted_microservice_names:
        types_dict = temp_grouped[ms_name]
        sorted_types_dict = OrderedDict()
        for ordered_type in issue_type_order_list:
            if ordered_type in types_dict: sorted_types_dict[ordered_type] = types_dict.pop(ordered_type)
        for remaining_type in sorted(types_dict.keys()): sorted_types_dict[remaining_type] = types_dict[remaining_type]
        sorted_grouped_tasks[ms_name] = sorted_types_dict
    logger.info("Группировка и сортировка задач для 'Перечня изменений' завершена.")
    return sorted_grouped_tasks


def prepare_setup_instructions_data(processed_df_with_tasks: pd.DataFrame, fields_mapping_config: list,
                                    sort_config: dict) -> OrderedDict:
    if processed_df_with_tasks.empty:
        logger.warning("(prepare_setup_data): Входной DataFrame пуст.")
        return OrderedDict()
    logger.info("Начало prepare_setup_data: подготовка данных для 'Настроек системы'.")

    field_names = _FieldNames(fields_mapping_config)
    setup_instructions_col = field_names.setup_instructions
    summary_col_setup = field_names.summary
    key_col_setup = field_names.key
    sort_field_tasks_setup = sort_config.get('sort_tasks_within_group_by', key_col_setup)
    priority_for_sort_setup = field_names.priority

    if not key_col_setup or key_col_setup not in processed_df_with_tasks.columns:
        logger.error(f"(prepare_setup_data): Колонка ключа ('{key_col_setup}') отсутствует.");
        return OrderedDict()
    if not setup_instructions_col or setup_instructions_col not in processed_df_with_tasks.columns:
        logger.warning(f"(prepare_setup_data): Колонка инструкций ('{setup_instructions_col}') отсутствует.");
        return OrderedDict()
    if not summary_col_setup or summary_col_setup not in processed_df_with_tasks.columns:
        logger.warning(
            f"(prepare_setup_data): Колонка Summary ('{summary_col_setup}') отсутствует. Используется 'Без заголовка'.")
        # Не добавляем колонку в df, просто используем значение по умолчанию в get()

    df_with_setup = processed_df_with_tasks[
        processed_df_with_tasks[setup_instructions_col].fillna('').astype(str).str.strip() != ''
        ].copy()

    if df_with_setup.empty: logger.warning("(prepare_setup_data): Нет задач с инструкциями."); return OrderedDict()
    logger.info(f"Найдено задач с инструкциями по установке: {len(df_with_setup)}")

    expanded_setup_tasks_list = []
    for index, task_row in df_with_setup.iterrows():
        task_data_for_setup = {
            'issue_key': task_row.get(key_col_setup),
            'summary': task_row.get(summary_col_setup, "Без заголовка"),
            'setup_instructions': task_row.get(setup_instructions_col),
            sort_field_tasks_setup: task_row.get(sort_field_tasks_setup)
        }
        microservices = task_row.get('identified_microservices', [])
        if not microservices: continue
        for ms_name in microservices:
            expanded_setup_tasks_list.append({'microservice': ms_name, 'task_data': task_data_for_setup.copy()})

    if not expanded_setup_tasks_list: logger.warning(
        "(prepare_setup_data): Нет задач после расширения по МС."); return OrderedDict()

    temp_grouped_setup = {}
    for item in expanded_setup_tasks_list:
        ms_name, task_data = item['microservice'], item['task_data']
        temp_grouped_setup.setdefault(ms_name, []).append(task_data)

    priority_order_list = sort_config.get('priority_order', [])
    priority_map = {p.lower(): i for i, p in enumerate(priority_order_list)}

    for ms_name, tasks_list in temp_grouped_setup.items():
        if not tasks_list: continue  # Пропускаем, если список задач пуст

        # Проверяем, существует ли поле для сортировки в первой задаче (предполагая однородность)
        if sort_field_tasks_setup and sort_field_tasks_setup in tasks_list[0] and tasks_list[0].get(
                sort_field_tasks_setup) is not None:
            if priority_map and priority_for_sort_setup and sort_field_tasks_setup == priority_for_sort_setup:
                tasks_list.sort(
                    key=lambda t: priority_map.get(str(t.get(sort_field_tasks_setup, "")).lower(), len(priority_map)))
            else:
                tasks_list.sort(key=lambda t: str(t.get(sort_field_tasks_setup, '')))
        elif key_col_setup in tasks_list[0]:
            logger.debug(
                f"Сортировка для {ms_name} в Настройках по ключу, т.к. поле '{sort_field_tasks_setup}' не найдено/пусто.")
            tasks_list.sort(key=lambda t: t.get(key_col_setup, ''))

    sorted_grouped_setup_tasks = OrderedDict()
    sorted_microservice_names_setup = sorted(temp_grouped_setup.keys())
    if sort_config.get('sort_microservices_by') == 'name_desc':
        sorted_microservice_names_setup.reverse()
    for ms_name in sorted_microservice_names_setup: sorted_grouped_setup_tasks[ms_name] = temp_grouped_setup[ms_name]
    logger.info("Подготовка данных для раздела 'Настройки системы' завершена.")
    return sorted_grouped_setup_tasks