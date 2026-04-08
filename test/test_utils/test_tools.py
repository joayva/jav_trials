import re

import pytest

from xls_management import HOMEPATH
from xls_management.utils.tools import (all_in_sequence, col_data_from,
                                        col_name_from, get_slices,
                                        list_from_comma_separated_str)


def test_all_in_sequence():
    a3_0 = ('one', 'two', 'three')
    a3_1 = ('one', 'two', 'three')
    
    assert all_in_sequence(a3_0, a3_1) is True

    
    a2_0 = ('two', 'one')
    assert all_in_sequence(a2_0, a3_0) is True
    assert all_in_sequence(a3_0, a2_0) is False
    a0:tuple = ()
    
    assert all_in_sequence(a0, a2_0) is True
    assert all_in_sequence(a3_0, a0) is False

def test_get_comma_separated_list():   
    assert list_from_comma_separated_str('one, two, three') == ['one', 'two', 'three']
    assert list_from_comma_separated_str(' one , two , three ') == ['one', 'two', 'three']
    assert list_from_comma_separated_str('one,?two,three') == ['one', 'two', 'three']
    assert list_from_comma_separated_str('one,\t?two ,th ree\n') == ['one', 'two', 'three']
    assert list_from_comma_separated_str('one,\t\t\t\t?two ,th ree\n') == ['one', 'two', 'three']

def test_re_sub():
    assert re.sub(r'[\?r]','','1234') == '1234'
    assert re.sub(r'[\?r]','','1234?') == '1234'
    assert re.sub(r'[\?r]','','r1234') == '1234'
    assert re.sub(r'[\?r]','','1234r\r\n5678r\r\n5679r') == '1234\r\n5678\r\n5679'
    assert re.sub(r'[\?r]','','1234?\r\n5678?\r\n5679r') == '1234\r\n5678\r\n5679'

    assert re.sub(r'USE( )?CASE', 'USE-CASE', 'MY USE CASE') == 'MY USE-CASE'
    assert re.sub(r'USE( )?CASE', 'USE-CASE', 'MY USECASE') == 'MY USE-CASE'
    assert re.sub(r'USE( )?CASE', 'USE-CASE', 'MY USE-CASE') == 'MY USE-CASE'

    assert re.sub(r'\r',r'\\r', '\r') == '\\r'

def test_dict_iterator():
    my_dict = {'a':'uno','b':'dos', 'c':'tres'}

    my_iter = iter(my_dict.items())
    assert next(my_iter) == ('a','uno')
    assert next(my_iter) == ('b','dos')
    for v in my_iter:
        assert v == ('c','tres')
    
    my_iter = iter(my_dict.values())
    assert next(my_iter) == 'uno'
    assert next(my_iter) == 'dos'
    for v in my_iter:
        assert v == 'tres'
    
    my_iter = iter(my_dict.keys())
    assert next(my_iter) == 'a'
    assert next(my_iter) == 'b'
    for v in my_iter:
        assert v == 'c'

def test_col_name_from():
    assert col_name_from(0) == 'A'
    assert col_name_from(25) == 'Z'
    assert col_name_from(6) == 'G'
    assert col_name_from(26) == 'AA'
    assert col_name_from(51) == 'AZ'
    assert col_name_from(52) == 'BA'
    assert col_name_from(77) == 'BZ'
    assert col_name_from(78) == 'CA'
    assert col_name_from(900) == 'AHQ'

def test_get_slices():
    assert list(get_slices(0,0,5)) == []
    assert list(get_slices(0,1,5)) == [(0,0,1)]
    assert list(get_slices(0,4,5)) == [(0,0,4)]
    assert list(get_slices(0,5,5)) == [(0,0,5)]
    assert list(get_slices(0,10,5)) == [(0,0,5), (1,5,10)]
    assert list(get_slices(0,11,5)) == [(0,0,5), (1,5,10), (2,10,11)]

def test_unic_join():
    from xls_management.utils.tools import unic_join

    assert unic_join('-', []) == ''
    assert unic_join('-', [1]) == '1'
    assert unic_join('-', [1,1,1]) == '1'
    assert unic_join('-', [1,2,1]) == '1-2'
    assert unic_join('-', [1,1,2]) == '1-2'
    assert unic_join(", ",['one',]) == 'one'
    assert unic_join(", ",['one','one']) == 'one'
    assert unic_join(", ",['one','one','two']) == 'one, two'
    assert unic_join(", ",['one','two','one']) == 'one, two'
    # No side efects here
    assert unic_join(", ",['anyone', 'tomorrow', 'one', 'to', 'row']) == 'anyone, tomorrow, one, to, row'
    assert unic_join(", ",['one', 'to', 'anyone', 'tomorrow', 'row']) == 'one, to, anyone, tomorrow, row'

def test_lazy_join():
    from xls_management.utils.tools import lazy_join


    assert lazy_join(", ",['one']) == 'one'
    assert lazy_join(", ",['one','one']) == 'one'
    assert lazy_join(", ",['one','one','two']) == 'one, two'
    assert lazy_join(", ",['one','two','one']) == 'one, two'
    # side effects: any item contained in another item could be not added
    assert lazy_join(", ",['anyone', 'tomorrow', 'one', 'to', 'row']) == 'anyone, tomorrow'
    assert lazy_join(", ",['one', 'to', 'anyone', 'tomorrow', 'row']) == 'one, to, anyone, tomorrow'


def test_index():
    values = ["one","three"]
    with pytest.raises(ValueError) as exinfo:
        _ = values.index("two")
    assert str(exinfo.value) == "'two' is not in list"
    assert values.index('one') == 0
    assert values.index('three') == 1

def test_col_data_from():
    from xls_management.utils.tools import col_data_from
    values = [35.1, 10.2, 40.5]
    expected = [('A',35.1), ('B',10.2), ('C',40.5)]
    got = []
    for v in col_data_from(values):
        got.append(v)
    assert got == expected
