# jav_trials

## Repository Structure

High-level overview of the project layout with key files and directories.

```
.
├─ LICENSE                      # Project license
├─ pyproject.toml               # Project and dependency config
├─ README.md                    # Project overview and docs
├─ script                       # Binary files building root
│  └─ data/                     # Extra data required to build a binary
├─ src/
│  └─ xls_management/           # Main package implementing XLS utilities
│     ├─ __init__.py            # Package initializer
│     ├─ config.py              # Configuration helpers
│     ├─ is_ole.py              # Detect OLE/legacy XLS files
│     ├─ workbook.py            # Workbook read/write helpers
│     ├─ ate/                   # ATE-related modules
│     │  ├─ __init__.py
│     │  ├─ data.py             # ATE data models/utilities
│     │  ├─ project.py          # Project-level ATE helpers
│     │  └─ tracking.py         # ATE tracking utilities
│     ├─ shell/                 # CLI / shell helpers
│     │  ├─ __init__.py
│     │  └─ ate.py
│     ├─ tui/                   # Terminal UI components
│     │  ├─ file_picker.py
│     │  ├─ main.py
│     │  ├─ msgbox.py
│     │  └─ project_form.py
│     └─ utils/                 # Support utilities used across package
│        ├─ __init__.py
│        ├─ tools.py              # tools helpers used in tests and modules
│        └─ color.py            # Terminal color helpers
└─ test/                        # Unit tests and fixtures
    ├─ test_config.py
    ├─ test_ole.py
    ├─ test_workbook.py
    └─ test_ate/
        └─ test_om/
            └─ test_db_info.py
```

- `LICENSE`, `pyproject.toml`, `README.md`: project metadata and dependency configuration.
- `src/xls_management/`: main package implementing XLS management utilities and tools.
- `src/xls_management/ate/`: ATE-related modules and domain-specific subpackage `om`.
- `src/xls_management/tui/`: terminal UI components and forms.
- `src/xls_management/utils/`: helper utilities used across the package.
- `test/`: unit tests and test fixtures mirroring the package structure


Notes:
- Use `src/xls_management/workbook.py` for primary XLS operations.
- Tests mirror the source layout under `test/`; run them with your chosen test runner.

# Usage
## Binary
Binary file should be copied into 
   - the user folder where it should be run; or
   - a folder in the windows path
   - e.g. copy it to %OneDrive%\vw
Adding the folder to the windows PATH is strongly recommended; so it can be called from 
Open a cmd window and run the command
`xls_shell`
   
## Developers
cmd
A python virtual environment is recommended
### Virtual environment creation
python -m xls
### Virtual environment activation
xls\Scripts\activate.bat
### Package installation
A package manager should make the python package available in a shared folder or a pip index service.
Given a xls_management-x.y.z-py3-none-any.whl --being x y and z version number  package has been placed in repo\dist folder --e.g. xls_management-1.0.0-py3-none-any.whl 
pip install repo\dist\

## Package management
This section is only required to provide a xls_management-x.y.z-py3-none-any.whl package file for the installation; if a package file was provided go directly to [Binary](#Binary)
cmd
A python virtual environment is recommended
### Virtual environment creation
python -m build
### Virtual environment activation
build\Scripts\activate.bat
### Install required packages
pip install build
### Clone the repository
mkdir vw
cd vw
git clone https://github.com/joseayude/jav_trials.git repo
cd repo
git checkout tracking_debuging
### build the package
python -m build .
## Binary building
### Prerequisites
A package of the the desired version should have been
    - built as specified in [Package mangagement](#package-management)
    - installed in xls environment as specified in [Package installation](#package-installation)
    - package pyinstaller should be installed in xls virtual envirnomet. It can be installed using pip
    `pip install pyinstaller`
### Build the binary
Go to script folder
`cd script`
Execute next command
`pyinstaller --onefile --add-data "data:xls_management\tui" xls_shell.py`
