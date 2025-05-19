import pandas as pd
import os
from . import logger_config # Относительный импорт

logger = logger_config.setup_logger(__name__)

def load_csv_to_dataframe(file_path: str, encoding: str = 'utf-8', delimiter: str = ','):
    """
    Загружает данные из CSV-файла в pandas DataFrame.
    Пустые значения читаются как пустые строки (а не NaN).

    :param file_path: Абсолютный или относительный путь к CSV-файлу.
                      (main.py должен передавать абсолютный путь).
    :param encoding: Кодировка CSV-файла.
    :param delimiter: Разделитель полей в CSV-файле.
    :return: pandas DataFrame с данными или None в случае ошибки.
    """
    logger.debug(f"Попытка загрузки CSV из: {file_path} (кодировка: {encoding}, разделитель: '{delimiter}')")
    if not os.path.exists(file_path):
        logger.error(f"CSV-файл не найден: {file_path}")
        return None
    try:
        # keep_default_na=False и na_filter=False нужны, чтобы пустые строки читались как "", а не NaN
        df = pd.read_csv(file_path, encoding=encoding, delimiter=delimiter, keep_default_na=False, na_filter=False)
        logger.info(f"CSV-файл успешно загружен: {file_path}. Обнаружено строк: {len(df)}, колонок: {len(df.columns)}")
        logger.debug(f"Имена колонок в DataFrame: {list(df.columns)}")
        return df
    except FileNotFoundError: # Хотя эта проверка уже есть выше, pandas может выдать свою специфическую ошибку
        logger.error(f"FileNotFoundError при попытке чтения pandas: {file_path}", exc_info=True)
        return None
    except pd.errors.EmptyDataError:
        logger.error(f"CSV-файл пуст: {file_path}")
        return None
    except pd.errors.ParserError as e:
        logger.error(f"Не удалось разобрать (ошибка парсинга) CSV-файл {file_path}: {e}")
        return None
    except Exception as e:
        logger.error(f"Произошла непредвиденная ошибка при чтении CSV-файла {file_path}: {e}", exc_info=True)
        return None