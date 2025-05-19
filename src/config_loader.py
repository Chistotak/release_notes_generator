import json
import os
from . import logger_config  # Относительный импорт для logger_config.py в той же директории

logger = logger_config.setup_logger(__name__)


def load_json_config(file_path: str):
    """
    Загружает конфигурацию из JSON-файла.

    :param file_path: Абсолютный путь к JSON-файлу.
    :return: Словарь с конфигурацией или None в случае ошибки.
    """
    logger.debug(f"Попытка загрузки JSON из: {file_path}")
    if not os.path.exists(file_path):
        logger.error(f"Конфигурационный файл не найден: {file_path}")
        return None
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            config_data = json.load(f)
        logger.info(f"Конфигурация успешно загружена из: {file_path}")
        return config_data
    except json.JSONDecodeError as e:
        logger.error(f"Не удалось декодировать JSON из файла {file_path}: {e}")
        return None
    except Exception as e:
        logger.error(f"Произошла непредвиденная ошибка при чтении файла {file_path}: {e}", exc_info=True)
        return None


def get_all_configs(config_dir_name: str = "configs"):
    """
    Загружает все основные конфигурационные файлы (config.json, fields_mapping.json, word_styles.json).
    Предполагается, что этот модуль находится в подпапке 'src/', а директория
    с конфигурациями (указанная в config_dir_name) находится в корне проекта.

    :param config_dir_name: Имя директории с конфигурационными файлами,
                           расположенной в корне проекта (например, "configs").
    :return: Кортеж (main_config, fields_mapping, word_styles).
             Если какой-либо файл не удалось загрузить, соответствующий элемент будет None.
    """
    logger.info(f"Загрузка всех конфигураций из директории '{config_dir_name}' в корне проекта.")

    # Определяем корень проекта: config_loader.py (этот файл) -> src/ -> корень_проекта/
    current_module_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(current_module_dir)

    absolute_config_dir_path = os.path.join(project_root, config_dir_name)

    if not os.path.isdir(absolute_config_dir_path):
        logger.error(f"Директория с конфигурациями не найдена: {absolute_config_dir_path}")
        return None, None, None

    main_config_path = os.path.join(absolute_config_dir_path, "config.json")
    fields_mapping_path = os.path.join(absolute_config_dir_path, "fields_mapping.json")
    word_styles_path = os.path.join(absolute_config_dir_path, "word_styles.json")

    logger.debug(f"Абсолютный путь к config.json: {main_config_path}")
    logger.debug(f"Абсолютный путь к fields_mapping.json: {fields_mapping_path}")
    logger.debug(f"Абсолютный путь к word_styles.json: {word_styles_path}")

    main_config = load_json_config(main_config_path)
    fields_mapping = load_json_config(fields_mapping_path)
    word_styles = load_json_config(word_styles_path)

    if main_config is None or fields_mapping is None or word_styles is None:
        logger.warning("Не удалось загрузить один или несколько основных конфигурационных файлов.")
    else:
        logger.info("Все основные конфигурационные файлы успешно обработаны.")

    return main_config, fields_mapping, word_styles