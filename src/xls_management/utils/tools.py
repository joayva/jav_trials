import re
from typing import Any


def all_in_sequence(sequence_0:list|tuple, sequence_1:list|tuple) -> bool:
    """
    Returns True if each value in sequence_0 is in sequence_1; elsewhere False
    """
    for value in sequence_0:
        if value not in sequence_1:
            return False
    return True

#   Function EinlesenGetrennteWerteKomma(ByVal lahIDs As String) As Collection
    #Moved to utils as it could be used in multiple places.
def list_from_comma_separated_str(comma_separated_string:str) -> list:
    """
    Given commma separated string is cleaned and converted to list.
    The cleaning includes removing "?" and whitespace characters.
    
    :param comma_separated_string: A string with comma separated values, 
                which may include "?" characters and whitespace.
    :type comma_separated_string: str
    :return: list of cleaned values from the input string.
    :rtype: list
    """
    cleaned_string = re.sub(r'[\?\s]','',comma_separated_string)
    return cleaned_string.split(',')
#       Dim idCollection As Collection
#       Dim subStrings() As String
#       Set idCollection = New Collection
#       Dim newString As String
#       Dim x As Integer
#       
#       newString = Replace(lahIDs, "?", "")
#       subStrings = Split(newString, ",")
#       For x = LBound(subStrings) To UBound(subStrings)
#           idCollection.Add Trim(subStrings(x))
#       Next
#       Set EinlesenGetrennteWerteKomma = idCollection
#       End Function

def col_name_from(index:int):
    col_name = ''
    while index >= 0:
        modulus = index % 26
        col_name = f'{chr(modulus+65)}{col_name}'
        index = index // 26 -1
    return col_name

def get_slices(start, top, size):
    i = 0
    first = start
    for second in range(first+size,top,size):
        yield i, first, second
        first = second
        i += 1
    if first < top:
        yield i, first, top

def unic_join(str_value:str, input_list:list|tuple):
    if len(input_list) == 0:
        return ''
    values = list(input_list[:1])
    for v in input_list[1:]:
        if v not in values:
            values.append(v)
    return str_value.join([str(v) for v in values])

def lazy_join(str_value:str, input_list:list[str]|tuple[str, ...]):
    if len(input_list) == 0:
        return ''
    str_list = input_list[0]
    for v in input_list[1:]:
        if v not in str_list:
            str_list += f'{str_value}{v}'
    return str_list

def col_data_from(data:list[Any]|tuple[Any]):
    for i, v in enumerate(data):
        yield col_name_from(i), v 
