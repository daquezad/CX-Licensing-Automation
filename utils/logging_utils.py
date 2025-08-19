import os
import logging


def setup_logging(log_dir: str | None = None, log_filename: str = "compare_excels.log"):
    # Configure root logger to write everything to a file only
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)

    # Remove any existing handlers to avoid duplicates/console output
    for handler in list(root_logger.handlers):
        root_logger.removeHandler(handler)

    formatter = logging.Formatter("%(asctime)s %(levelname)s %(name)s - %(message)s")
    logfile_path = os.path.join(log_dir or os.getcwd(), log_filename)
    file_handler = logging.FileHandler(logfile_path, mode='w', encoding='utf-8')
    file_handler.setFormatter(formatter)
    file_handler.setLevel(logging.DEBUG)
    root_logger.addHandler(file_handler)

    root_logger.debug(f"Logging to file: {logfile_path}")
    return root_logger


