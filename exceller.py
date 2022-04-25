import zipfile
import xml.dom.minidom
import re
import argparse
import olefile
from math import pow as power

SHARED_STRINGS_FILE = 'sharedStrings.xml'
ROW_INDEX = "row_index"
COLUMN_INDEX = "column_index"
STRING_INDEX = "string_index"
IS_STRING_LABEL = "is_string"
STRING_CONTENT = "string"
    

def extract_cell_strings(opened_ooxml, str_path_to_excel_file):
    files_in_opened_ooxml_list = opened_ooxml.namelist()
    matches = [match for match in files_in_opened_ooxml_list if SHARED_STRINGS_FILE in match]
    if len(matches) == 0:
        print(f"There is no {SHARED_STRINGS_FILE} file in {str_path_to_excel_file}, could not complete the analysis")
        return None
    # There is supposed to be only one sharedStrings.xml in an Excel file.
    xml_content = xml.dom.minidom.parseString(opened_ooxml.read(matches[0]))
    texts_list = []
    for si in xml_content.getElementsByTagName("si"):
        for t in si.getElementsByTagName("t"):
            texts_list.insert(len(texts_list), t.firstChild.nodeValue)
    return texts_list


def create_cells_dict(opened_ooxml):
    files_in_opened_ooxml_list = opened_ooxml.namelist()
    cells_sheet_files_list = []
    # Cells mappings are supposed to be found in xl/worksheets
    for file_path in files_in_opened_ooxml_list:
        matches_list = re.findall(f'^xl/worksheets/sheet[1-9][0-9]*\.xml$', file_path)
        if len(matches_list) > 0:
            cells_sheet_files_list += matches_list
    sheets_and_content_dict = dict()
    for worksheet in cells_sheet_files_list:
        xml_content = xml.dom.minidom.parseString(opened_ooxml.read(worksheet))
        sheets_and_content_dict[worksheet] = create_sheet_cells_dict(xml_content)
    if len(sheets_and_content_dict) == 0:
        print("There are no worksheet XMLs to work with")
        return None
    return sheets_and_content_dict


def create_sheet_cells_dict(xml_content):

    # c stands for cell, that is the tag used by this XML file to identify a cell
    all_cells_list = []
    for c in xml_content.getElementsByTagName("c"):
        cell_info_dict = dict()
        #The cells are represented by a letter, which represents the column and a number, which represents a row, e.g: C23
        cell_letter_and_number_format = c.getAttribute('r')
        if len(cell_letter_and_number_format) > 0 and len(c.getElementsByTagName("v")) > 0:
            cell_column_as_letter = re.findall('^[A-Z]+', cell_letter_and_number_format)[0]
            cell_row_as_number = int(re.findall('[0-9]+$', cell_letter_and_number_format)[0])
            cell_column_as_int = cell_column_letter_to_number(cell_column_as_letter)
            cell_info_dict[ROW_INDEX] = cell_row_as_number
            cell_info_dict[COLUMN_INDEX] = cell_column_as_int
            # v stands for value, that is the tag used in this XML file to represent the index of the 'si' entry in
            # sharedStrings.xml that contains the matching string
            # There is supposed to be only one v tag for each cell
            cell_value_type = c.getAttribute('t')
            # If the cell contains a string, its 'c' entry will contain a 't' value, which will be set to 's' (t='s')
            if len(cell_value_type) > 0 and cell_value_type == 's':
                cell_info_dict[STRING_INDEX] = int(c.getElementsByTagName("v")[0].firstChild.nodeValue)
                cell_info_dict[IS_STRING_LABEL] = True
            # Note: even though it says 'STRING_INDEX' it will be used later as the actual value of the cell, if the
            # cell contains a value from a different type
            else:
                cell_info_dict[STRING_INDEX] = c.getElementsByTagName("v")[0].firstChild.nodeValue
                cell_info_dict[IS_STRING_LABEL] = False
            all_cells_list.insert(len(all_cells_list), cell_info_dict)
    return all_cells_list


def match_cells_to_strings(strings_list, sheets_cells_dict):
    for sheet_name in sheets_cells_dict.keys():
        for cell_dict in sheets_cells_dict[sheet_name]:
            # If the cell contains a non string value, the value is stored in cell_dict[STRING_INDEX]
            if not cell_dict[IS_STRING_LABEL]:
                cell_dict[STRING_CONTENT] = cell_dict[STRING_INDEX]
                continue
            cell_dict[STRING_CONTENT] = strings_list[cell_dict[STRING_INDEX]]
    return sheets_cells_dict


def cell_column_letter_to_number(cell_column_as_letter):
    abc = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    int_val = 0
    index_in_reverse = 0
    for letter in cell_column_as_letter[::-1]:
        int_val += (abc.find(letter.upper()) + 1) * power(26, index_in_reverse)
        index_in_reverse += 1
    return int(int_val)


def parse_arguments():
    parser = argparse.ArgumentParser(
        description='Replace cells calls in VBA, e.g Cells(112, 2) with their content by retrieving it '
                    'from default XML files in Excel')
    parser.add_argument('--excel_file', type=str, help='Path to Excel file')
    parser.add_argument('--vba_file', type=str, help='Path to the VBA file extracted from the provided Excel file')
    parser.add_argument('--edited_vba_file', type=str, help='The path to save the newly created VBA file')
    args = parser.parse_args()
    return args


def replace_cell_funcs_with_cell_content(opened_vba_file, opened_new_vba_file, sheets_cells_dict):
    replace_cell_funcs_in_vba = re.findall(f'Replace\( *Cells\([0-9]+, [0-9]+\), [\"\'].*[\"\'], *[\"\'][a-zA-Z0-9_\-]*[\"\']\)',
                                           opened_vba_file, re.IGNORECASE)

    # TODO: Add check for Worksheets
    #if len(sheet_cells_funcs_in_vba) > 0:
    #    relevant_sheets = []
    #sheet_cells_funcs_in_vba = re.findall(f'Worksheets\([\"\']Sheet[0-9]*[\"\']\)Cells\([0-9]+, [0-9]+\)',
    #                                      opened_vba_file, re.IGNORECASE)
    for replace_func in replace_cell_funcs_in_vba:
        #The function will look like this: Replace(Cells(109, 8), "ghwuy", "")
        # We need to replace the string "ghwuy" (in this example) with "" in the string that is stored in the Cell H109
        split_by_comma = replace_func.split(',')
        # No need to check if there are at least 4 values in 'split_by_comma' because the regex pattern to which all
        # strings in 'replace_cell_funcs_in_vba' ensures this is true
        str_to_replace = re.findall('\w+', split_by_comma[2])

        if len(str_to_replace) == 0:
            str_to_replace = ''
        else:
            str_to_replace = str_to_replace[0]
        str_to_replace_with = re.findall('\w+', split_by_comma[3])
        if len(str_to_replace_with) == 0:
            str_to_replace_with = ''
        else:
            str_to_replace_with = str_to_replace_with[0]

        # Again. no need to check the findall array length
        cell_func = re.findall(f'Cells\([0-9]+, [0-9]+\)', replace_func, re.IGNORECASE)[0]
        cell_content_str = find_matching_string(cell_func=cell_func, sheets_cells_dict=sheets_cells_dict)
        if cell_content_str is None:
            continue
        cell_content_str = cell_content_str.replace(str_to_replace, str_to_replace_with)
        opened_vba_file = opened_vba_file.replace(replace_func, f'\"{cell_content_str}\"')
    cells_funcs_in_vba = re.findall(f'Cells\([0-9]+, [0-9]+\)', opened_vba_file, re.IGNORECASE)

    for cell_func in cells_funcs_in_vba:
        cell_content_str = find_matching_string(cell_func=cell_func, sheets_cells_dict=sheets_cells_dict)
        if cell_content_str is None:
            continue
        opened_vba_file = opened_vba_file.replace(cell_func, f'\"{cell_content_str}\"')
    join_transpose_statements = re.findall(f'Join\(\[TRANSPOSE\([a-z]+[0-9]+\:[a-z]+[0-9]+\)\] *, *[\"\'][\w_\*\$\#\@\%]*[\"\']\)',
                                           opened_vba_file, re.IGNORECASE)

    for join_transpose_statement in join_transpose_statements:
        cells = cells_range_to_list(join_transpose_statement)
        joined_str = ""
        delimiter = re.findall('[\"\'][\w_\*\$#@%]*[\"\']\)', join_transpose_statement, re.IGNORECASE)
        if len(delimiter) == 0:
            continue
        delimiter = delimiter[0].strip('\'")')

        for cell in cells:
            str_match = find_matching_string(cell_func=cell, sheets_cells_dict=sheets_cells_dict)
            if str_match is not None:
                joined_str += delimiter + str_match
        opened_vba_file = opened_vba_file.replace(join_transpose_statement, f'\"{joined_str}\"')

    opened_new_vba_file.write(opened_vba_file)
    opened_new_vba_file.close()


def cells_range_to_list(str_cells_range):
    str_range = re.findall(f'[a-z]+[0-9]+\:[a-z]+[0-9]+', str_cells_range, re.IGNORECASE)
    if len(str_range) == 0:
        return None
    str_range = str_range[0]
    # There is no need to check the regex matches array length, since the regex used to find 'str_range' insures the
    # following regex matches will return a value
    first_cell = str_range.split(':')[0]
    last_cell = str_range.split(':')[1]
    first_letter = re.findall(f'[a-z]+', first_cell, re.IGNORECASE)[0]
    first_row_index = int(re.findall(f'[0-9]+', first_cell)[0])
    last_letter = re.findall(f'[a-z]+', last_cell, re.IGNORECASE)[0]
    last_row_index = int(re.findall(f'[0-9]+', last_cell)[0])
    cells = []
    if first_letter == last_letter:
        letter_num_index = cell_column_letter_to_number(first_letter)
        # Even though instinctively the first cell's row index should be lower than the last one's, it can be the other
        # way around, but the transpose outcome will be the same
        lower_row_index = min(first_row_index, last_row_index)
        higher_row_index = max(first_row_index, last_row_index)
        for row_index in range(lower_row_index, higher_row_index+1):
            cells.append(f"{row_index}, {letter_num_index}")
        return cells
    if first_row_index == last_row_index:
        first_letter_index = cell_column_letter_to_number(first_letter)
        last_letter_index = cell_column_letter_to_number(last_letter)
        for column_index in range(min(first_letter_index, last_letter_index), max(first_letter_index, last_letter_index) +1):
            cells.append(f"{last_row_index}, {column_index}")
        return cells
    return cells


def find_matching_string(cell_func, sheets_cells_dict):
    row_and_column = re.findall(f'[0-9]+', cell_func)
    if len(row_and_column) < 2:
        return None
    row_index = int(row_and_column[0])
    column_index = int(row_and_column[1])
    for sheet_cells_list in sheets_cells_dict.values():
        for cell_details_dict in sheet_cells_list:
            if (cell_details_dict[ROW_INDEX] == row_index) and (cell_details_dict[COLUMN_INDEX] == column_index):
                return cell_details_dict[STRING_CONTENT]
    return None


def validate_ooxml(str_path_to_file):
    ooxml_magic = b'\x50\x4b'
    with open(str_path_to_file, 'rb') as opened_file:
        if not opened_file.read(len(ooxml_magic)) == ooxml_magic:
            return False
    # The magic check is not enough, since OOXML, APK and other files share a magic
    try:
        zip_file = zipfile.ZipFile(str_path_to_file, 'r')
    except zipfile.BadZipfile:
        return False
    try:
        if not '[Content_Types].xml' in zip_file.namelist():
            return False
        zip_file.open('[Content_Types].xml').close()
    except:
        return False
    return True


def validate_ole(str_path_to_file):
    ole_header = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
    # Valid Excel OLE streams to look for
    valid_ole_streams = (['VBA'], ['PROJECT'], ['_VBA_PROJECT'], ['Workbook'], ['Book'], ['Equation Native'],
                         ['\x04JSRV_SegmentInformation'], ['\x05HwpSummaryInformation'])
    special_stream_name = 'equation native'
    try:
        ole = olefile.OleFileIO(str_path_to_file)
        if not ole.header_signature == ole_header:
            print(f"This is the wrong header: {ole.header_signature}")
            return False
        # ole.listdir() will return a list of lists, we need to check if any of these lists are in 'valid_ole_streams'
        for stream in ole.listdir():
            if stream in valid_ole_streams or stream[0].lower() == special_stream_name:
                return True
        return False
    except Exception as e:
        print(f"OLE exception:\n{e}")
        return False
    return False


def ooxml_main(str_path_to_excel_file, str_path_to_vba_file, str_path_to_edited_vba_file):
    opened_ooxml = zipfile.ZipFile(str_path_to_excel_file)
    strings_list = extract_cell_strings(opened_ooxml, str_path_to_excel_file)
    if strings_list is None:
        return
    sheets_cells_dict = create_cells_dict(opened_ooxml)
    if sheets_cells_dict is None:
        return
    sheets_cells_content_dict = match_cells_to_strings(strings_list=strings_list, sheets_cells_dict=sheets_cells_dict)
    opened_vba_file = open(str_path_to_vba_file, 'rt').read()
    opened_new_vba_file = open(str_path_to_edited_vba_file, 'wt')
    replace_cell_funcs_with_cell_content(opened_vba_file, opened_new_vba_file, sheets_cells_content_dict)


def main(str_path_to_excel_file, str_path_to_vba_file, str_path_to_edited_vba_file):
    if not validate_ooxml(str_path_to_excel_file):
        if not validate_ole(str_path_to_excel_file):
            print("Not a valid OOXML/ OLE, cannot proceed")
            return
        else:
            print("A valid OLE was provided, but unfortunately, the script does not support OLE files at this moment")
    else:
        ooxml_main(str_path_to_excel_file, str_path_to_vba_file, str_path_to_edited_vba_file)

        
if __name__ == "__main__":
    args = parse_arguments()
    excel_path = args.excel_file
    vba_path = args.vba_file
    edited_vba_file = args.edited_vba_file
    main(str_path_to_excel_file=excel_path, str_path_to_vba_file=vba_path, str_path_to_edited_vba_file=edited_vba_file)
