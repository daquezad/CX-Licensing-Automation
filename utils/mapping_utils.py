import os
import json
import logging


def load_pid_to_skus_map(mapping_path: str | None) -> dict[str, list[str]]:
    """Load PID→SKUs mapping from an external JSON file.

    Expected JSON structure:
    {
      "AIR-DNA-E": ["AIR-DNA-E-T"],
      "DNA-P-T2-E-5Y": ["DSTACK-T2-E"]
    }
    """
    logger = logging.getLogger(__name__)
    if mapping_path is None:
        default_path = os.path.join(os.getcwd(), "sku_map.json")
        if os.path.exists(default_path):
            mapping_path = default_path
        else:
            logger.info("No external mapping provided; proceeding without PID→SKU exceptions.")
            return {}

    try:
        with open(mapping_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            raise ValueError("Mapping JSON must be an object mapping strings to list[str]")
        normalized: dict[str, list[str]] = {}
        for key, value in data.items():
            if isinstance(value, list):
                normalized[str(key).strip()] = [str(v).strip() for v in value]
            elif isinstance(value, str):
                normalized[str(key).strip()] = [value.strip()]
            else:
                logger.warning("Ignoring invalid mapping value for key '%s': %r", key, value)
        logger.info("Loaded %d PID→SKU exception entries from %s", len(normalized), mapping_path)
        return normalized
    except FileNotFoundError:
        logger.warning("Mapping file not found: %s. Continuing without exceptions.", mapping_path)
        return {}
    except Exception:
        logger.exception("Failed to load mapping file: %s. Continuing without exceptions.", mapping_path)
        return {}


def get_valid_sku_matches(cssm_matches, pre_ea_migrated_pid_str: str, pid_to_skus_map: dict[str, list[str]]):
    """Return CSSM rows that match the PID directly or via the exception map.

    - First tries a direct SKU match on the PRE-EA migrated PID.
    - If none, looks up mapped exception SKUs from pid_to_skus_map and matches any of them.
    """
    # Try direct match
    sku_match = cssm_matches[cssm_matches['SKU'] == pre_ea_migrated_pid_str]
    # Try mapped exceptions if no direct match
    if sku_match.empty and pre_ea_migrated_pid_str in pid_to_skus_map:
        valid_skus = pid_to_skus_map[pre_ea_migrated_pid_str]
        sku_match = cssm_matches[cssm_matches['SKU'].isin(valid_skus)]
    return sku_match


