import sys
from unittest.mock import patch
from test import working_path


def test_init(tmp_path) -> None:
    to_remove = [name for name in sys.modules if name.startswith("xls_management")]
    for name in to_remove:
        del sys.modules[name]
    fake_root = 'C:\\fake_root'
    with patch('os.getenv', return_value=fake_root):
        from xls_management import WORKPATH, HOMEPATH, ROOTPATH
        assert str(WORKPATH) == fake_root
        assert str(HOMEPATH) == fake_root
        assert (
            ROOTPATH.parent.name == 'site-packages' or
            ROOTPATH.parent == working_path / 'src'
        )
    