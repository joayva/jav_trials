import pytest

from xls_management.utils.dict_first import DictFirst


def test_dict_first():

    d:DictFirst = DictFirst({"a":"a_value_0"})
    assert isinstance(d,DictFirst)
    assert d["a"] == "a_value_0"
    d["a"] = "a_value_1"
    assert d["a"] == "a_value_0"
    with pytest.raises(KeyError):
        result = d["b"]
    d["b"] = "b_value_0"
    assert d["b"]  == "b_value_0"
    d["b"] = "b_value_1"
    assert d["b"]  == "b_value_0"