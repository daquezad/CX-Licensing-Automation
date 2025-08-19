import os
import shutil
import logging


def ensure_clean_dir(directory_path: str) -> None:
    """Ensure directory exists and is empty.

    - Creates the directory if missing.
    - Removes all files, symlinks, and subdirectories if it exists.
    - Logs warnings for entries that cannot be deleted and re-raises unexpected errors.
    """
    os.makedirs(directory_path, exist_ok=True)
    logger = logging.getLogger(__name__)
    try:
        with os.scandir(directory_path) as it:
            for entry in it:
                path = entry.path
                try:
                    if entry.is_file() or entry.is_symlink():
                        os.unlink(path)
                    elif entry.is_dir():
                        shutil.rmtree(path)
                except Exception as exc:
                    logger.warning("Failed to delete %s: %s", path, exc)
    except Exception:
        # Let caller decide how to handle overall failure
        raise


