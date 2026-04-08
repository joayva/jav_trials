import os
import platform
from pathlib import Path

ROOTPATH = Path(__file__).parent
home_var:str
if platform.system():
    home_var='USERPROFILE'
else:
    home_var='HOME'
HOMEPATH = Path(os.getenv(home_var,'.'))
