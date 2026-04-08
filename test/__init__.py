import os
import sys
from pathlib import Path

# Get absolute paths for src and tests
working_path = Path(__file__).parent.parent
BASE_DIR = working_path.as_uri()
SRC_PATH = os.path.join(BASE_DIR, "src")
TESTS_PATH = os.path.join(BASE_DIR, "test")

# Add them to sys.path if not already present
for path in (SRC_PATH, TESTS_PATH):
    if path not in sys.path:
        sys.path.insert(0, path)

