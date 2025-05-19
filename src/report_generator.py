from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
from . import logger_config

# Импортируем _FieldNames, если он нужен здесь напрямую, или передаем конкретные internal_names
# Для упрощения, будем ожидать, что task_dict уже содержит поля с нужными internal_names
# из data_processor._FieldNames.
# import data_processor # Чтобы получить доступ к data_processor._FieldNames (но это создаст цикл. импорт)

logger = logger_config.setup_logger(__name__)

# --- Константы для стилей по умолчанию ---
DEFAULT_FONT = "Calibri"
DEFAULT_FONT_SIZE_PT_VAL = 11
TITLE_FONT_SIZE_PT_VAL = 18
H1_FONT_SIZE_PT_VAL = 14
H2_FONT_SIZE_PT_VAL = 13
H3_FONT_SIZE_PT_VAL = 12
TASK_ITEM_FONT_SIZE_PT_VAL = 10


def set_run_font(run, font_name=None, size_pt_val=None, bold=None, italic=None, color_rgb=None):
    if font_name:
        run.font.name = font_name
        r = run._element
        r.rPr.rFonts.set(qn('w:eastAsia'), font_name)
        r.rPr.rFonts.set(qn('w:cs'), font_name)
    if size_pt_val is not None:
        try:
            actual_size = Pt(int(size_pt_val))
            run.font.size = actual_size
        except ValueError:
            logger.warning(f"Неверное значение для размера шрифта: {size_pt_val}.")
        except Exception as e:
            logger.warning(f"Ошибка при установке размера шрифта ({size_pt_val}): {e}")
    if bold is not None: run.font.bold = bold
    if italic is not None: run.font.italic = italic
    if color_rgb is not None and isinstance(color_rgb, RGBColor): run.font.color.rgb = color_rgb


def add_styled_paragraph(document, text="", style_name=None, font_name=None, size_pt_val=None,
                         bold=None, italic=None, align=None,
                         space_after_pt=None, space_before_pt=None):  # Убрал is_list_item и list_style для общности
    p = document.add_paragraph()  # Сначала создаем пустой параграф
    if text:  # Добавляем текст, если он есть, это будет первый run
        run = p.add_run(text)
        if font_name or size_pt_val is not None or bold is not None or italic is not None:
            set_run_font(run, font_name, size_pt_val, bold, italic)

    if style_name:
        try:
            p.style = style_name
        except Exception as e:
            logger.warning(f"Не удалось применить стиль параграфа '{style_name}': {e}")

    pf = p.paragraph_format
    if align: pf.alignment = align
    if space_after_pt is not None: pf.space_after = Pt(int(space_after_pt))
    if space_before_pt is not None: pf.space_before = Pt(int(space_before_pt))
    return p


def _apply_field_style_to_run(run, field_style: dict, default_font_name: str, default_font_size: int):
    """Применяет стили из field_style к объекту run."""
    set_run_font(
        run,
        font_name=field_style.get('font_name', default_font_name),
        size_pt_val=field_style.get('font_size', default_font_size),
        bold=field_style.get('bold', False),
        italic=field_style.get('italic', False)
        # Можно добавить цвет, если он будет в field_style
        # color_rgb=RGBColor.from_string(field_style.get('color_hex')) if field_style.get('color_hex') else None
    )


def _add_task_fields_to_paragraph(paragraph, task_dict: dict, fields_to_display: list,
                                  default_font_name: str, default_font_size: int,
                                  style_key_in_field_spec: str):
    """
    Добавляет отформатированные поля задачи в существующий параграф.
    Каждое поле (или его часть: префикс, лейбл, значение, суффикс) добавляется как отдельный run.

    :param style_key_in_field_spec: Ключ для получения стиля из field_spec (например, 'changes_style' или 'setup_style')
    """
    first_field_component_added_to_paragraph = not paragraph.runs  # True, если параграф пока пуст

    for field_spec in fields_to_display:
        internal_name = field_spec.get('internal_name')
        value = task_dict.get(internal_name)

        if value is None or (isinstance(value, str) and not value.strip()):
            logger.debug(f"Поле '{internal_name}' пустое или None, пропускаем.")
            continue

        value_str = str(value).strip()

        style_options = field_spec.get(style_key_in_field_spec, {})

        # Новая строка перед этим полем, ЕСЛИ это не первый компонент в параграфе
        if not first_field_component_added_to_paragraph and style_options.get('new_line_before', False):
            paragraph.add_run().add_break()

        current_field_text_parts = []  # Собираем части текущего поля (префикс, лейбл, значение, суффикс)

        # Префикс
        prefix = style_options.get('prefix', '')
        if prefix:
            current_field_text_parts.append({'text': prefix, 'style': style_options})

        # Лейбл
        label = field_spec.get('report_label', '')
        if label:
            current_field_text_parts.append({'text': label + (" " if value_str else ""),
                                             'style': style_options})  # Пробел после лейбла, если есть значение

        # Значение поля (с обработкой многострочности)
        if style_options.get('multiline', False) and '\n' in value_str:
            lines = value_str.splitlines()
            for i, line in enumerate(lines):
                if i > 0:  # Для всех строк, кроме первой в многострочном значении, добавляем перенос
                    # Добавляем предыдущую часть как отдельный run, если она была
                    if current_field_text_parts:
                        for part in current_field_text_parts:
                            run = paragraph.add_run(part['text'])
                            _apply_field_style_to_run(run, part['style'], default_font_name, default_font_size)
                        current_field_text_parts = []  # Очищаем для следующей части поля
                        first_field_component_added_to_paragraph = False  # Следующий run будет в том же параграфе

                    paragraph.add_run().add_break()  # Перенос перед следующей строкой многострочного значения
                current_field_text_parts.append({'text': line, 'style': style_options})
        else:
            current_field_text_parts.append({'text': value_str, 'style': style_options})

        # Суффикс
        suffix = style_options.get('suffix', '')
        if suffix:
            current_field_text_parts.append({'text': suffix, 'style': style_options})

        # Добавляем все собранные части текущего поля в параграф
        for part in current_field_text_parts:
            run = paragraph.add_run(part['text'])
            _apply_field_style_to_run(run, part['style'], default_font_name, default_font_size)
            first_field_component_added_to_paragraph = False  # После добавления первого run, флаг сбрасывается


# --- Функции create_title_section, create_microservices_version_table (без изменений из Итерации 7 с фиксом таблицы) ---
def create_title_section(document, report_title, logo_path, styles_config):
    logger.info("Создание титульной секции...")  # ... (код как в предыдущем ответе)
    fonts = styles_config.get('fonts', {});
    font_sizes = styles_config.get('font_sizes', {})
    para_spacing = styles_config.get('paragraph_spacing', {})
    title_font_name = fonts.get('title', DEFAULT_FONT)
    title_font_size_val = font_sizes.get('title', TITLE_FONT_SIZE_PT_VAL)
    if logo_path and os.path.exists(logo_path):
        try:
            document.add_picture(logo_path, width=Inches(1.5)); logger.info(f"Логотип добавлен: {logo_path}")
        except Exception as e:
            logger.error(f"Не удалось добавить логотип {logo_path}: {e}", exc_info=True)
    elif logo_path:
        logger.warning(f"Файл логотипа не найден: {logo_path}")
    add_styled_paragraph(document, text=report_title, font_name=title_font_name, size_pt_val=title_font_size_val,
                         bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                         space_after_pt=para_spacing.get('after_title', 24), space_before_pt=12)
    logger.info(f"Заголовок отчета: '{report_title}'")


def create_microservices_version_table(document, versions_data, styles_config):
    """
    Создает таблицу с версиями микросервисов в документе.
    Упрощенная установка ширины колонок, если свойства секции недоступны.
    """
    if not versions_data:
        logger.warning("Нет данных для таблицы версий микросервисов. Таблица не будет создана.")
        return

    logger.info("Создание таблицы версий микросервисов...")

    fonts = styles_config.get('fonts', {})
    font_sizes = styles_config.get('font_sizes', {})
    colors = styles_config.get('colors_hex', {})
    table_props = styles_config.get('table_properties', {})

    default_font = fonts.get('default', DEFAULT_FONT)
    heading2_font_name = fonts.get('heading2', default_font)
    heading2_font_size = font_sizes.get('heading2', H2_FONT_SIZE_PT_VAL)
    table_header_font_size_val = font_sizes.get('table_header', 10)
    table_content_font_size_val = font_sizes.get('table_content', 10)
    table_header_bg_color_hex = colors.get('table_header_background', "D9D9D9")

    add_styled_paragraph(
        document,
        text="Версии компонентов релиза:",
        font_name=heading2_font_name,
        size_pt_val=heading2_font_size,
        bold=True,
        space_after_pt=6,
        space_before_pt=12
    )

    try:
        table = document.add_table(rows=1, cols=2)
        table.style = 'TableGrid'
    except Exception as e:
        logger.error(f"Не удалось создать таблицу: {e}", exc_info=True)
        return

    header_texts = ['Микросервис', 'Версия']
    hdr_cells = table.rows[0].cells
    for i, cell_obj in enumerate(hdr_cells):
        try:
            cell_obj.text = header_texts[i]
            p = cell_obj.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for old_run in p.runs: p._element.remove(old_run._element)
            run = p.add_run(header_texts[i])
            set_run_font(run, font_name=default_font, size_pt_val=table_header_font_size_val, bold=True)
            if table_header_bg_color_hex:
                shd = OxmlElement('w:shd')
                shd.set(qn('w:val'), 'clear');
                shd.set(qn('w:color'), 'auto');
                shd.set(qn('w:fill'), table_header_bg_color_hex)
                cell_obj._tc.get_or_add_tcPr().append(shd)
        except Exception as e:
            logger.warning(f"Ошибка при стилизации заголовка таблицы '{header_texts[i]}': {e}", exc_info=True)

    for item in versions_data:
        try:
            row_cells = table.add_row().cells
            cell_texts = [item.get('microservice', ''), item.get('version', '')]
            for i, cell_obj in enumerate(row_cells):
                cell_obj.text = cell_texts[i]
                p = cell_obj.paragraphs[0]
                for old_run in p.runs: p._element.remove(old_run._element)
                run = p.add_run(cell_texts[i])
                set_run_font(run, font_name=default_font, size_pt_val=table_content_font_size_val)
        except Exception as e:
            logger.warning(f"Ошибка при добавлении строки данных в таблицу ({item}): {e}", exc_info=True)

    # Установка ширины колонок - УПРОЩЕННЫЙ ВАРИАНТ
    try:
        col1_w_percent = table_props.get('width_col1_percent', 30)
        col2_w_percent = table_props.get('width_col2_percent', 70)

        # Используем фиксированную предполагаемую общую ширину для расчета
        # Это менее точно, но должно работать всегда.
        # Стандартная страница A4 имеет ширину ~8.27 дюйма. Поля по 1 дюйму = ~6.27 полезной.
        # Округлим до 6.0 для простоты, или можно сделать это настраиваемым.
        assumed_total_width_inches = 6.0
        logger.debug(f"Используется предполагаемая общая ширина для таблицы: {assumed_total_width_inches:.2f} дюймов.")

        col0_width_inches = assumed_total_width_inches * (col1_w_percent / 100.0)
        col1_width_inches = assumed_total_width_inches * (col2_w_percent / 100.0)

        logger.debug(
            f"Расчетная ширина колонки 0: {col0_width_inches:.2f} дюймов, колонки 1: {col1_width_inches:.2f} дюймов.")

        if table.columns and len(table.columns) == 2:
            if col0_width_inches > 0.1:
                table.columns[0].width = Inches(col0_width_inches)
            else:
                logger.warning(
                    f"Рассчитанная ширина для колонки 0 ({col0_width_inches:.2f} дюймов) слишком мала, ширина не установлена.")

            if col1_width_inches > 0.1:
                table.columns[1].width = Inches(col1_width_inches)
            else:
                logger.warning(
                    f"Рассчитанная ширина для колонки 1 ({col1_width_inches:.2f} дюймов) слишком мала, ширина не установлена.")
        else:
            logger.warning("Не удалось получить доступ к колонкам таблицы для установки ширины.")

    except Exception as e:
        logger.error(f"Критическая ошибка при установке ширины колонок таблицы: {e}", exc_info=True)

    document.add_paragraph()
    logger.info("Таблица версий микросервисов создана.")

def _get_internal_name_from_mapping(fields_mapping_config: list, standard_csv_header: str, default_internal: str, alt_internals: list = None) -> str:
    """
    Ищет internal_name в fields_mapping_config для стандартного заголовка CSV или имени по умолчанию.
    Возвращает найденный internal_name или default_internal.
    """
    if alt_internals is None: alt_internals = []
    for field_spec in fields_mapping_config:
        # Сначала проверяем по internal_name
        spec_internal = field_spec.get('internal_name')
        if spec_internal and (spec_internal == default_internal or spec_internal in alt_internals):
            return spec_internal
        # Затем по csv_header
        spec_csv_header = field_spec.get('csv_header', '').lower()
        if spec_csv_header == standard_csv_header.lower():
            # Если csv_header совпал, возвращаем его internal_name, или сам csv_header, если internal_name нет
            return field_spec.get('internal_name', field_spec.get('csv_header'))
    return default_internal


def create_changes_section(document, grouped_data, styles_config, fields_mapping_config, report_config):
    logger.info("Создание раздела 'Перечень изменений'...")

    fonts = styles_config.get('fonts', {});
    font_sizes = styles_config.get('font_sizes', {})
    para_spacing = styles_config.get('paragraph_spacing', {});
    section_titles = report_config.get('report_section_titles', {})
    default_font = fonts.get('default', DEFAULT_FONT)
    h1_font = fonts.get('heading1', default_font);
    h1_size = font_sizes.get('heading1', H1_FONT_SIZE_PT_VAL)
    h2_font = fonts.get('heading2', default_font);
    h2_size = font_sizes.get('heading2', H2_FONT_SIZE_PT_VAL)
    h3_font = fonts.get('heading3', default_font);
    h3_size = font_sizes.get('heading3', H3_FONT_SIZE_PT_VAL)
    task_font_size = font_sizes.get('task_item', TASK_ITEM_FONT_SIZE_PT_VAL)

    if not grouped_data:
        logger.warning("Нет данных для 'Перечня изменений'.")
        add_styled_paragraph(document, text=report_config.get('report_section_titles', {}).get('no_changes_text',
                                                                                               "Изменений нет."),
                             font_name=default_font, size_pt_val=task_font_size, italic=True, space_before_pt=6,
                             space_after_pt=6)
        return

    changes_title = section_titles.get('main_changes', "Перечень изменений")
    add_styled_paragraph(document, text=changes_title, font_name=h1_font, size_pt_val=h1_size, bold=True,
                         space_before_pt=para_spacing.get('after_title', 24),
                         space_after_pt=para_spacing.get('after_heading1', 12))

    fields_for_changes_display = sorted(
        [f for f in fields_mapping_config if f.get('display_in_changes', False)],
        key=lambda f: f.get('changes_order', 99)
    )
    if not fields_for_changes_display:
        logger.warning(
            "Не настроены поля для отображения в 'Перечне изменений' (display_in_changes: true). Задачи могут быть не отображены или отображены некорректно.")
        # В этом случае _add_task_fields_to_paragraph не будет вызван, и задачи не будут детализированы.
        # Можно добавить фоллбэк, если это критично.

    for ms_name, types_dict in grouped_data.items():
        add_styled_paragraph(document, text=ms_name, font_name=h2_font, size_pt_val=h2_size, bold=True,
                             space_before_pt=para_spacing.get('after_heading1', 12),
                             space_after_pt=para_spacing.get('after_heading2', 8))

        for issue_type, tasks_list in types_dict.items():
            add_styled_paragraph(document, text=issue_type, font_name=h3_font, size_pt_val=h3_size, italic=True,
                                 space_before_pt=para_spacing.get('after_heading2', 8),
                                 space_after_pt=para_spacing.get('after_heading3', 4))

            if not tasks_list:
                try:
                    p_no = document.add_paragraph(style='ListBullet')
                except KeyError:
                    p_no = document.add_paragraph()
                set_run_font(p_no.add_run("Нет задач этого типа."), font_name=default_font, size_pt_val=task_font_size,
                             italic=True)
                pf = p_no.paragraph_format;
                pf.space_before = Pt(0);
                pf.space_after = Pt(int(para_spacing.get('list_item_after', 6)))
                continue

            for task_dict in tasks_list:
                try:
                    p_task = document.add_paragraph(style='List Number')
                except KeyError:
                    logger.warning("Стиль 'List Number' не найден, используется 'ListBullet'.")
                    try:
                        p_task = document.add_paragraph(style='ListBullet')
                    except KeyError:
                        logger.warning("Стиль 'ListBullet' не найден."); p_task = document.add_paragraph()

                if fields_for_changes_display:
                    _add_task_fields_to_paragraph(p_task, task_dict, fields_for_changes_display,
                                                  default_font, task_font_size, 'changes_style')
                else:  # Фоллбэк, если fields_for_changes_display пуст (маловероятно, если есть task_report_text)
                    # Этот фоллбэк теперь менее нужен, если task_report_text всегда включен через fields_mapping
                    key_internal_name_fb = _get_internal_name_from_mapping(fields_mapping_config, 'issue key',
                                                                           'issue_key')
                    key_val = task_dict.get(key_internal_name_fb, "")
                    desc_val = task_dict.get('task_report_text', "Нет текста.")
                    if key_val:  # Выводим ключ только если он есть
                        set_run_font(p_task.add_run(f"{key_val}: "), font_name=default_font, size_pt_val=task_font_size,
                                     bold=True)
                    set_run_font(p_task.add_run(desc_val), font_name=default_font, size_pt_val=task_font_size)

                pf = p_task.paragraph_format
                pf.space_before = Pt(int(para_spacing.get('list_item_before', 0)))
                pf.space_after = Pt(int(para_spacing.get('list_item_after', 6)))
    logger.info("Раздел 'Перечень изменений' создан.")


def create_setup_section(document, grouped_setup_data, styles_config, fields_mapping_config, report_config):
    logger.info("Создание раздела 'Настройки системы'...")

    fonts = styles_config.get('fonts', {});
    font_sizes = styles_config.get('font_sizes', {})
    para_spacing = styles_config.get('paragraph_spacing', {});
    section_titles = report_config.get('report_section_titles', {})
    default_font = fonts.get('default', DEFAULT_FONT)
    h1_font = fonts.get('heading1', default_font);
    h1_size = font_sizes.get('heading1', H1_FONT_SIZE_PT_VAL)
    h2_font = fonts.get('heading2', default_font);
    h2_size = font_sizes.get('heading2', H2_FONT_SIZE_PT_VAL)
    task_font_size = font_sizes.get('task_item', TASK_ITEM_FONT_SIZE_PT_VAL)

    if not grouped_setup_data:
        logger.info("Нет данных для раздела 'Настройки системы'. Раздел будет пропущен.")
        return

    setup_title = section_titles.get('system_setup', "Настройки системы")
    add_styled_paragraph(document, text=setup_title, font_name=h1_font, size_pt_val=h1_size, bold=True,
                         space_before_pt=para_spacing.get('after_title', 24),
                         space_after_pt=para_spacing.get('after_heading1', 12))

    fields_for_setup_display = sorted(
        [f for f in fields_mapping_config if f.get('display_in_setup', False)],
        key=lambda f: f.get('setup_order', 99)
    )
    if not fields_for_setup_display:
        logger.warning(
            "Не настроены поля для отображения в 'Настройках системы' (display_in_setup: true). Задачи могут быть не отображены или отображены некорректно.")

    for ms_name, tasks_list in grouped_setup_data.items():
        add_styled_paragraph(document, text=ms_name, font_name=h2_font, size_pt_val=h2_size, bold=True,
                             space_before_pt=para_spacing.get('after_heading1', 12),
                             space_after_pt=para_spacing.get('after_heading2', 8))

        if not tasks_list:
            try:
                p_no = document.add_paragraph(style='ListBullet')
            except KeyError:
                p_no = document.add_paragraph()
            set_run_font(p_no.add_run("Нет инструкций по настройке для этого компонента."), font_name=default_font,
                         size_pt_val=task_font_size, italic=True)
            pf = p_no.paragraph_format;
            pf.space_before = Pt(0);
            pf.space_after = Pt(int(para_spacing.get('list_item_after', 6)))
            continue

        for task_dict in tasks_list:
            try:
                p_task = document.add_paragraph(style='ListBullet')
            except KeyError:
                logger.warning("Стиль 'ListBullet' не найден."); p_task = document.add_paragraph()

            if fields_for_setup_display:
                _add_task_fields_to_paragraph(p_task, task_dict, fields_for_setup_display,
                                              default_font, task_font_size, 'setup_style')
            else:  # Фоллбэк, если поля не настроены
                key_internal_name_fb = _get_internal_name_from_mapping(fields_mapping_config, 'issue key', 'issue_key')
                summary_internal_name_fb = _get_internal_name_from_mapping(fields_mapping_config, 'summary',
                                                                           'summary_text', ['summary'])
                setup_instr_internal_name_fb = _get_internal_name_from_mapping(fields_mapping_config,
                                                                               'custom field (инструкция по установке)',
                                                                               'setup_instructions')

                key_val = task_dict.get(key_internal_name_fb, "")
                summary_val = task_dict.get(summary_internal_name_fb, "")
                instr_val = task_dict.get(setup_instr_internal_name_fb, "Инструкций нет.")

                header_parts_fb = []
                if key_val: header_parts_fb.append(key_val)
                if summary_val: header_parts_fb.append(summary_val)
                header_text_fb = ": ".join(header_parts_fb) if header_parts_fb else "Инструкция"

                set_run_font(p_task.add_run(f"{header_text_fb}\n"), font_name=default_font, size_pt_val=task_font_size,
                             bold=True)
                set_run_font(p_task.add_run(instr_val), font_name=default_font, size_pt_val=task_font_size)

            pf = p_task.paragraph_format
            pf.space_before = Pt(int(para_spacing.get('list_item_before', 0)))
            pf.space_after = Pt(int(para_spacing.get('list_item_after', 6)))
    logger.info("Раздел 'Настройки системы' создан.")


def generate_report_docx(output_filename, report_title_text, logo_full_path,
                         microservice_versions_list, word_styles_config,
                         grouped_data_for_changes,
                         grouped_data_for_setup,
                         main_config_for_titles,
                         fields_mapping_for_details
                         ):
    logger.info(f"Начало генерации DOCX отчета: {output_filename}")
    doc = Document()

    create_title_section(doc, report_title_text, logo_full_path, word_styles_config)
    create_microservices_version_table(doc, microservice_versions_list, word_styles_config)
    create_changes_section(doc, grouped_data_for_changes, word_styles_config, fields_mapping_for_details,
                           main_config_for_titles)
    create_setup_section(doc, grouped_data_for_setup, word_styles_config, fields_mapping_for_details,
                         main_config_for_titles)

    try:
        output_dir = os.path.dirname(output_filename)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir);
            logger.info(f"Создана директория: {output_dir}")
        doc.save(output_filename)
        logger.info(f"Отчет успешно сохранен: {output_filename}")
    except Exception as e:
        logger.error(f"Не удалось сохранить документ {output_filename}: {e}", exc_info=True)