import pytest

from xls_management.utils.sensitive_dict import SensitiveDict


def test_sensitive_dict():

    d:SensitiveDict = SensitiveDict({"a":"a_value_0"})
    assert isinstance(d,SensitiveDict)
    assert d["a"] == "a_value_0"
    with pytest.raises(KeyError) as e:
        d["a"] = "a_value_1"
        assert e.message == 'a key already exists'
    assert d["a"] == "a_value_0"
    with pytest.raises(KeyError):
        result = d["b"]
    d["b"] = "b_value_0"
    assert d["b"]  == "b_value_0"
    with pytest.raises(KeyError) as e:
        d["b"] = "b_value_1"
        assert e.message == 'b key already exists'
    assert d["b"]  == "b_value_0"