import logging
import sys
import os  # Добавлен для определения пути к лог-файлу относительно корня проекта

# Уровни логирования
LOG_LEVEL_CONSOLE = logging.INFO  # Уровень для вывода в консоль
LOG_LEVEL_FILE = logging.DEBUG  # Уровень для вывода в файл (может быть более детальным)

LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
LOG_DATE_FORMAT = '%Y-%m-%d %H:%M:%S'

# Имя файла лога (относительно корня проекта)
# Если None, логирование в файл отключено.
LOG_FILE_NAME = "release_notes_word.log"
# LOG_FILE_NAME = None # Раскомментировать, чтобы отключить логирование в файл

# Словарь для хранения уже настроенных логгеров, чтобы избежать дублирования обработчиков
_configured_loggers = {}


def setup_logger(name="AppLogger", level_console=LOG_LEVEL_CONSOLE, level_file=LOG_LEVEL_FILE,
                 log_file_name=LOG_FILE_NAME):
    """
    Настраивает и возвращает логгер.
    Если логгер с таким именем уже настроен, возвращает его.
    """
    if name in _configured_loggers:
        return _configured_loggers[name]

    logger = logging.getLogger(name)
    logger.setLevel(min(level_console, level_file))  # Устанавливаем общий минимальный уровень для логгера

    # Предотвращаем дублирование, если другой код уже добавил обработчики (маловероятно при такой схеме)
    if logger.hasHandlers():
        logger.handlers.clear()

    formatter = logging.Formatter(LOG_FORMAT, datefmt=LOG_DATE_FORMAT)

    # 1. Обработчик для вывода в консоль
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(level_console)  # Индивидуальный уровень для консоли
    ch.setFormatter(formatter)
    logger.addHandler(ch)

    # 2. Обработчик для вывода в файл (если указано имя файла)
    actual_log_file_path = None
    if log_file_name:
        try:
            # Определяем путь к лог-файлу относительно корня проекта
            # Предполагаем, что logger_config.py находится в src/
            current_script_dir = os.path.dirname(os.path.abspath(__file__))
            project_root = os.path.dirname(current_script_dir)  # Поднимаемся из src в корень
            actual_log_file_path = os.path.join(project_root, log_file_name)

            # Создаем директорию для лога, если она не существует (на случай если log_file_name содержит папки)
            log_dir = os.path.dirname(actual_log_file_path)
            if log_dir and not os.path.exists(log_dir):
                os.makedirs(log_dir)
                # Используем print, т.к. логгер еще не полностью настроен для файла
                print(f"INFO: Создана директория для лог-файла: {log_dir}")

            fh = logging.FileHandler(actual_log_file_path, encoding='utf-8', mode='a')  # mode='a' для дозаписи
            fh.setLevel(level_file)  # Индивидуальный уровень для файла
            fh.setFormatter(formatter)
            logger.addHandler(fh)
            # Используем print для первого сообщения о лог-файле, т.к. logger может быть еще не готов для файла
            print(
                f"INFO: Логирование в файл настроено: {actual_log_file_path} (Уровень: {logging.getLevelName(level_file)})")
        except Exception as e:
            print(
                f"КРИТИЧЕСКАЯ ОШИБКА: Не удалось настроить файловый логгер для '{actual_log_file_path or log_file_name}': {e}",
                file=sys.stderr)

    logger.propagate = False  # Отключаем передачу сообщений родительским логгерам
    _configured_loggers[name] = logger

    # Первое сообщение от только что настроенного логгера
    # logger.debug(f"Логгер '{name}' настроен. Консоль: {logging.getLevelName(level_console)}, Файл: {actual_log_file_path if actual_log_file_path else 'Отключено'} ({logging.getLevelName(level_file) if actual_log_file_path else ''})")
    return logger


if __name__ == '__main__':
    # Тестирование конфигурации логгера
    # Этот блок покажет, как логгеры с разными именами и настройками могут работать

    # Настраиваем корневой логгер или логгер по умолчанию для приложения
    # Этот вызов настроит и вернет логгер. Последующие вызовы setup_logger('AppLogger') вернут этот же экземпляр.
    app_logger_test = setup_logger('AppLogger', level_console=logging.INFO, level_file=logging.DEBUG,
                                   log_file_name="app_test.log")
    app_logger_test.info("Это сообщение от AppLogger (info).")
    app_logger_test.debug("Это сообщение от AppLogger (debug, должно быть в файле, но не в консоли по умолч. INFO).")

    # Пример логгера для конкретного модуля
    module_logger_test = setup_logger('MyModuleTest', level_console=logging.DEBUG,
                                      log_file_name="module_test.log")  # Файл будет свой
    module_logger_test.debug("Это сообщение от MyModuleTest (debug).")
    module_logger_test.info("Это сообщение от MyModuleTest (info).")
    module_logger_test.warning("Это сообщение от MyModuleTest (warning).")

    # Если вызвать setup_logger снова с тем же именем, он вернет существующий экземпляр
    same_app_logger = setup_logger('AppLogger')
    same_app_logger.info("Это еще одно сообщение от AppLogger (info), через тот же экземпляр.")

    print(f"\nТестовые логи были выведены в консоль.")
    print(f"Проверьте файлы 'app_test.log' и 'module_test.log' в корневой папке проекта.")
    print(f"Для 'app_test.log' уровень DEBUG, для 'module_test.log' уровень DEBUG для файла.")
    print(f"В консоли для AppLogger должен быть уровень INFO, для MyModuleTest - DEBUG.")