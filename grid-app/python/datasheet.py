"""
Here the `cell` and the `sheet` functions are defined. They rely on many helper
functions for better usability, such as autocasting etc.
"""

import json
import re
import pandas as pd
import traceback

sheet_data = {}

real_print = print


def print(text):
    """
    Encapsulate the printed text with tags that supports parsing
    """
    if not isinstance(text, str):
        text = str(text)

    real_print("#INTERPRETER#" + text + "#ENDPARSE#", end='', flush=True)


def getReferenceRowIndex(reference):
    """
    TODO: Write documentation, provide examples
    """
    return int(re.findall(r'\d+', reference)[0])


def getReferenceColumnIndex(reference):
    """
    TODO: Write documentation, provide examples
    """
    return letterToIndex(''.join(re.findall(r'[a-zA-Z]', reference)))


def letterToIndex(letters):
    """
    TODO: Write documentation, provide examples
    """
    columns = len(letters) - 1
    total = 0
    base = 26

    for x in letters:
        number = ord(x)-64
        total += number * int(base**columns)
        columns -= 1
    return total


def indexToLetters(index):
    """
    TODO: Write documentation, provide examples
    """
    base = 26
    # start at the base that is bigger and work your way down
    leftOver = index
    columns = []

    while leftOver > 0:
        remainder = leftOver % base
        if remainder == 0:
            remainder = base

        columns.insert(0, int(remainder))
        leftOver = (leftOver - remainder) / base

    buff = ""

    for x in columns:
        buff += chr(x + 64)

    return buff


def cell_range_to_indexes(cell_range):
    """
    TODO: Write documentation, provide examples
    """
    references = []

    cells = cell_range.split(":")

    cell1Row = getReferenceRowIndex(cells[0])
    cell2Row = getReferenceRowIndex(cells[1])

    cell1Column = getReferenceColumnIndex(cells[0])
    cell2Column = getReferenceColumnIndex(cells[1])

    for x in range(cell1Column, cell2Column+1):
        columnReferences = []
        for y in range(cell1Row, cell2Row+1):
            columnReferences.append(indexToLetters(x) + str(y))
        references.append(columnReferences)

    return references


def has_number(s):
    """
    Return whether any of the items in a list are a digit.
    """
    return any(i.isdigit() for i in s)


def convert_to_json_string(element):
    """
    TODO: Write documentation
    """
    if isinstance(element, str):
        # string meant as string, escape
        element = element.replace("\n", "")

        # if data is string without starting with =, add escape quotes
        if len(element) > 0 and element[0] == '=':
            return element[1:]
        else:
            return "\"" + element + "\""
    else:
        return format(element, '.12f')


def df_to_list(df, include_headers=True):
    """
    Args:
        df (pd.DataFrame): The DataFrame to turn into a list
        include_headers (bool): Include the column names
    """
    columns = list(df.columns.values)
    data = []
    column_length = 0
    for column in columns:
        column_data = df[column].tolist()

        if include_headers:
            column_data = [column] + column_data

        column_length = len(column_data)
        data = data + column_data
    return (data, column_length, len(columns))


def _write_to_sheet(cell_range, data, include_headers, sheet_index):
    """
    An internal helper function for `sheet()` regarding the read operation.

    Args:
        cell_range: The range to use
        data: The data to write
        include_headers (bool): Include column names
        sheet_index: The index of the sheet
    """

    # convert numpy to array
    data_type_string = str(type(data))
    if data_type_string == "<class 'numpy.ndarray'>":
        data = data.tolist()

    if data_type_string == "<class 'pandas.core.series.Series'>":
        data = data.to_frame()
        data_type_string = str(type(data))

    if data_type_string == "<class 'pandas.core.frame.DataFrame'>":
        df_tuple = df_to_list(data, include_headers)
        data = df_tuple[0]
        column_length = df_tuple[1]
        column_count = df_tuple[2]

        # create cell_range
        if not has_number(cell_range):
            cellColumnLetter = cell_range
            cellColumnEndLetter = indexToLetters(
                letterToIndex(cellColumnLetter) + column_count - 1)
            cell_range = (cellColumnLetter + "1:" +
                          cellColumnEndLetter + str(column_length))
        else:
            cellColumnLetter = indexToLetters(
                getReferenceColumnIndex(cell_range))
            startRow = getReferenceRowIndex(cell_range)
            cellColumnEndLetter = indexToLetters(
                letterToIndex(cellColumnLetter) + column_count - 1)
            cell_range = (cellColumnLetter + str(startRow) + ":" +
                          cellColumnEndLetter + str(column_length +
                                                    startRow - 1))

    # always convert cell to range
    if ":" not in cell_range:
        if not has_number(cell_range):
            if type(data) is list:
                cell_range = (cell_range + "1:"
                              + cell_range + str(len(data)))
            else:
                cell_range = cell_range + "1:" + cell_range + "1"
        else:
            cell_range = cell_range + ":" + cell_range

    if type(data) is list:
        newList = list(map(convert_to_json_string, data))
        arguments = ['RANGE', 'SETLIST', cell_range, str(sheet_index)]
        # append list
        arguments = arguments + newList
        json_object = {'arguments': arguments}
        json_string = ''.join(['#PARSE#', json.dumps(json_object),
                               '#ENDPARSE#'])
        real_print(json_string, flush=True, end='')

    else:
        data = convert_to_json_string(data)

        data = {'arguments': ['RANGE', 'SETSINGLE', cell_range,
                              str(sheet_index), ''.join(["=", str(data)])]}
        data = ''.join(['#PARSE#', json.dumps(data), '#ENDPARSE#'])
        real_print(data, flush=True, end='')


def _read_from_sheet(cell_range, sheet_index):
    """
    An internal helper function for `sheet()` regarding the read operation.

    Args:
        cell_range: The range to read
        sheet_index: The index of the sheet

    Returns:
        A pandas DataFrame with the data
    """

    # convert non-range to range for get operation
    if ":" not in cell_range:
        cell_range = ':'.join([cell_range, cell_range])

    # in blocking fashion get latest data of range from Go
    real_print("#DATA#" + str(sheet_index) + '!' +
               cell_range + '#ENDPARSE#', end='', flush=True)

    command_buffer = ""
    while True:
        code_input = input("")
        # when empty line is found, execute code
        if code_input == "":
            try:
                exec(command_buffer, globals(), globals())
            except Exception:
                traceback.print_exc()
            break
        else:
            command_buffer += code_input + "\n"

    # if everything works, the exec command has filled sheet_data
    # with the appropriate data - return data range as arrays
    column_references = cell_range_to_indexes(cell_range)

    result = []
    for column in column_references:
        column_data = []
        for reference in column:
            column_data.append(sheet_data[str(sheet_index) + '!' + reference])
        result.append(column_data)

    return pd.DataFrame(data=result).transpose()


def sheet(cell_range, data=None, headers=False, sheet_index=0):
    """
    This reads from or writes to a cell range.

    >> sheet("A1", pd.DataFrame(data=[1,2,3])
    This displays 1, 2, and 3 in the first column (column A)

    >> print(sheet("A1"))
    This shows the value in the first row in the first column.

    """
    if data is not None:
        _write_to_sheet(cell_range, data, headers, sheet_index)
    else:
        _read_from_sheet(cell_range, sheet_index)


def cell(cell, value=None):
    """
    This reads from or writes to a specific cell.

    TODO: Note concrete examples
    """
    if value is not None:
        sheet(cell, data=value)
    else:
        cell_range = ':'.join([cell, cell])
        return sheet(cell_range)
