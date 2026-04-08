from pathlib import Path

import olefile


def is_ole(file_path: Path|str) -> str:
    if isinstance(file_path, Path):
        file_path = f'{file_path}'
    if olefile.isOleFile(file_path):
        return f"{file_path} is ole"
    return f"{file_path} is ole"
    