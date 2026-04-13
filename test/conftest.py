import json
import sys

import pytest
import yaml


def del_xls_management_imports():
    """Assure fresh import removing current loaded imports"""
    to_remove = [name for name in sys.modules if name.startswith("xls_management")]
    for name in to_remove:
        del sys.modules[name]

def get_test_input(file_path:str, load_method) -> dict[str,any]:
    with open(file_path, "r", encoding='utf8') as f:
        result = load_method(f)
        inputsList = list([tuple((value for key,value in item.items() if key != "id")) for item in result["argvalues"]])
        ids = []
        if "ids" in result.keys():
            ids = result["ids"]
        else:
            i: int = 0
            for item in result["argvalues"]:
                myDict: dict[str,any]
                ids.append(item.get("id", f"use case {i}"))
                i += 1
        return {
            "argnames": result["argnames"],
            "ids": ids,
            "argvalues": inputsList
        }

def parametrize_from_json(file_path:str):
    data = get_test_input(file_path, json.load)
    return pytest.mark.parametrize(**data)

def parametrize_from_yaml(file_path:str):
    data = get_test_input(file_path, yaml.safe_load)
    return pytest.mark.parametrize(**data)
