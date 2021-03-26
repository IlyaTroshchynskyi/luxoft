import pdfplumber
import pandas as pd
from operator import itemgetter
import re
import openpyxl



def define_pages_of_content(path):
    """
    Specifies the pages where the table of contents is located. This is necessary it order to parse the
     table of contents and get the pages of edit checks
      :param path: path to file
      :return: returns tuple start and end table of contents
    """
    start_content = 0
    end_content = 0
    with pdfplumber.open(path) as pdf:

        for num_page in range(11):
            for page in pdf.pages[num_page].extract_text().split('\n'):
                for line in page.split('\n'):
                    if 'TABLE OF CONTENTS' in line:
                        start_content = num_page
                    elif 'Introduction' in line and num_page > start_content:
                        end_content = num_page

    return start_content, end_content


def parse_content_table(path, pages):
    """
     Parses the table of contents. The extract_words function returns a dictionary.
     Finding the number of page  in the table of contents, subtract
     one because for function "0" this is the first page.
     :param path:
     :param pages: path - path to the file, pages - the beginning and end of the table
     :return: return a dictionary with a table of contents for two portfolio levels
    """
    with pdfplumber.open(path) as pdf:

        first_page = pdf.pages[pages[0]:pages[1]] # извлекаю данные из оглавления
        content_table = {'PORTFOLIO_LEVEL_END': pdf.pages[-1].page_number}
        for page in first_page:
            for dictionary in page.extract_words(keep_blank_chars=True):
                for key, value in dictionary.items():
                    if key == 'text':
                        if 'ACCOUNT LEVEL' in value:
                            content_table['ACCOUNT_LEVEL'] = int(value.split(' ')[-2]) - 1
                        elif 'PORTFOLIO LEVEL' in value:
                            content_table['ACCOUNT_LEVEL_END'] = int(value.split(' ')[-2]) - 1
                            content_table['PORTFOLIO_LEVEL'] = int(value.split(' ')[-2]) - 1
    return content_table


def parse_edits(path, content_table, edits_level):
    """
    First, we collect all the lines that we need in the array_lines array. Next, we find the number of lines
    which you need to skip , since our page with edits starts with the text that we
    not needed. Next, we find the name of the mdrm and break it into a number and a name and insert it into
    a new filter_data and then if the line starts with a new type of check, then we simply insert it, otherwise we will
    concatenate it to the previous one.
    filter_data [line] [num_line] [filter_data [line] [num_line] .find (':') + 1:]. lstrip () this record means write
    to a specific line to a specific cell a line that starts with a colon until the end of the line
    : param path: path to file
    : param content_table: dictionary of contents
    : param edits_level: the type of edits to be parsed
    : return:
    """
    array_lines = []

    with pdfplumber.open(path) as pdf:
        for page in range(content_table[edits_level], content_table[edits_level+'_END']):
            text = pdf.pages[page].extract_text()
            for x in text.split('\n'):
                if x != ' ' and x.rstrip(' ') != 'CONFIDENTIAL'\
                        and not x.startswith('FR Y-14M – Credit Cards'):
                    array_lines.append([x.rstrip(' ')])


    match_mdrm = re.compile(r'^\d+\.  ')
    match_type_edits = re.compile(r'^\d+\)  ')
    match_other = re.compile(r'^[A-Za-z)(]+|^\d+\)* |^\d+%')
    match_data_type = re.compile(r'^\d+\)\s+Data Type Check')
    match_missing = re.compile(r'^\d+\)\s+Missing Check')
    match_invalid = re.compile(r'^\d+\)\s+Invalid Check')
    match_integrity = re.compile(r'^\d+\)\s+Integrity Check')
    match_distribution = re.compile(r'^\d+\)\s+Distribution')
    match_integrity_num = re.compile(r'^[ivx]+\. ')
    sub_type_invalid = re.compile(r'^\d+ – |^\d+ ‐ |^\d+ - ')
    footnotes = re.compile(r'^[0-9]{1,2}[A-Z]{1}[a-z]{1,}|^[0-9]{1,2} [A-Z]{1}[a-z]{1,}')

    skip_rows = 0
    for number, row in enumerate(array_lines):
        if row[0].split(' ')[0] == '1.':
            skip_rows = number
            break
    filter_data = []
    parent_line = 0
    parent_line_join = 0
    data_for_join = []
    invalid_key = 0
    array_footnotes = []
    for number_line in range(skip_rows, len(array_lines)):

        if match_mdrm.match(array_lines[number_line][0]):
            array_footnotes.clear()
            filter_data.append([array_lines[number_line][0]])
            parent_line = len(filter_data) - 1

            data_for_join.append([array_lines[number_line][0]])
            parent_line_join = len(data_for_join)-1


        elif match_data_type.match(array_lines[number_line][0]):
            filter_data[parent_line].append(array_lines[number_line][0])
            array_footnotes.clear()

        elif match_missing.match(array_lines[number_line][0]):
            filter_data.append([filter_data[parent_line][0]])
            parent_line = len(filter_data) - 1
            filter_data[parent_line].append(array_lines[number_line][0])
            array_footnotes.clear()

        elif match_invalid.match(array_lines[number_line][0]):
            filter_data.append([filter_data[parent_line][0]])
            parent_line = len(filter_data) - 1
            filter_data[parent_line].append(array_lines[number_line][0])
            array_footnotes.clear()

            data_for_join[parent_line_join].append\
                (array_lines[number_line][0][array_lines[number_line][0].find(':') + 1:].lstrip())
            invalid_key = number_line

        elif sub_type_invalid.match(array_lines[number_line][0]):
            filter_data.append([filter_data[parent_line][0]])
            parent_line = len(filter_data) - 1
            filter_data[parent_line].append('Invalid Check:' + array_lines[number_line][0])
            data_for_join[parent_line_join][-1] += ' ' + array_lines[number_line][0]
            array_footnotes.clear()

        elif match_integrity.match(array_lines[number_line][0]):
            continue

        elif match_integrity_num.match(array_lines[number_line][0]):
            filter_data.append([filter_data[parent_line][0]])
            parent_line = len(filter_data) - 1
            filter_data[parent_line].append(array_lines[number_line][0])
            array_footnotes.clear()

        elif match_distribution.match(array_lines[number_line][0]):
            filter_data.append([filter_data[parent_line][0]])
            parent_line = len(filter_data) - 1
            filter_data[parent_line].append(array_lines[number_line][0])
            array_footnotes.clear()

        elif footnotes.match(array_lines[number_line][0]):
            array_footnotes.append(array_lines[number_line][0])


        elif match_other.match(array_lines[number_line][0]) and len(filter_data[parent_line]) >=2 and not array_footnotes:
            filter_data[parent_line][-1] += ' ' + array_lines[number_line][0]
            array_footnotes.clear()
            if number_line - invalid_key == 1:
                data_for_join[parent_line_join][-1] += ' ' + array_lines[number_line][0]
                invalid_key += 1


    df = pd.DataFrame(filter_data)

    def define_type_ch(cell):
        if 'Data Type Check' in cell:
            return 'Data Type'
        elif 'Missing Check' in cell:
            return 'Missing'
        elif 'Invalid Check' in cell:
            return 'Invalid'
        elif 'Distribution' in cell:
            return 'Distribution'
        else:
            return 'Integrity'

    def number_type(cell):
        num_type = {'Data Type': '1', 'Missing': '2','Invalid': '3', 'Integrity': '4', 'Distribution': '5'}
        return num_type.get(cell, '')

    def filter_descript(cell):
        type_checks = ['Data Type Check', 'Missing Check', 'Invalid Check', 'Distribution', 'Integrity Check']
        for type_ch in type_checks:
            if type_ch in cell:
                return cell[cell.find(':')+1:].lstrip()

        for item in cell.split('.'):
            if len(item) > 15:
                return item.lstrip()+'.'

    def number_integrity(cell):
        roman = {'i': '1', 'ii': '2', 'iii': '3', 'iv': '4', 'v': '5', 'vi': '6', 'vii': '7', 'viii': '8', 'ix': '9',
                 'x': '10', 'xi': '11', 'xii': '12', 'xiii': '13', 'xiv': '14',
                 'xv': '15', 'xvi': '16', 'xvii': '17', 'xviii': '18', 'xix': '19', 'xx': '20'}
        if match_integrity_num.match(cell):
            for item in cell.split('.'):
                return roman.get(item.lstrip(), '')

    def convert_mdrm_name(cell):
        return cell.lstrip(' ').replace('–', '').replace('/ ', ''). \
            replace('/', '_').replace(' - ', '_').replace('  ', '_').replace('-', ' ').replace(' ', '_').lower()

    cut_mdrms = {'average_daily_balance_(adb)': 'average_daily_balance_adb',
                 'sop_03_03_(or_purchased_credit_deteriorated_status)': 'sop_03_03',
                 'fee_income_other_fee_income': 'fee_income_other_fee'
                 }

    def clean_mdrm_name(cell):

        """Clean mdrm name. It will be node name"""
        r = re.compile(r'[A-Za-z]_[0-9]')
        if r.match(cell[len(cell) - 3:]):
            return cell[:len(cell) - 2]
        else:
            return cut_mdrms.get(cell, cell)

    df[['number', 'mdrm_name']] = df[0].str.split('.', expand=True)
    df['mdrm_name'] = df['mdrm_name'].str.lstrip()
    df['mdrm_name_clean'] = df['mdrm_name'].apply(convert_mdrm_name)
    df['mdrm_name_clean'] = df['mdrm_name_clean'].apply(clean_mdrm_name)
    df['type'] = df[1].apply(define_type_ch)
    df['num_type'] = df['type'].apply(number_type)
    df['number_integrity'] = df[1].apply(number_integrity)
    df['edit_number'] = df[['number','num_type', 'number_integrity']].apply(lambda x: '.'.join(x.dropna()), axis=1)
    df['description'] = df[1].apply(filter_descript)
    df['level'] = 'Account' if edits_level == 'ACCOUNT_LEVEL' else 'Portfolio'

    df_djoin = pd.DataFrame(data_for_join)
    df_djoin[['number', 'mdrm_name']] = df_djoin[0].str.split('.', expand=True)
    df_djoin['type'] = 'Invalid'

    df = pd.merge(df, df_djoin, how='left', left_on=['number','type'], right_on=['number','type'])
    df.drop(columns=['0_x','1_x','0_y', 'mdrm_name_y' ], inplace=True)
    df.to_excel('EDIT_CHECKS_' + edits_level + '.xlsx', index=False)


if __name__ == '__main__':

    try:
        path = input('Enter path to file \n')
        edits_level = input('Enter level of edit checks from this list: ACCOUNT_LEVEL,PORTFOLIO_LEVEL\n')
        pages = define_pages_of_content(path)
        edits_level = [sch for sch in edits_level.split(',')]
        content = parse_content_table(path, pages)

        for level in edits_level:
            parse_edits(path, content_table=content, edits_level=level)
        input('Press any button to end this program\n')
    except Exception as ex:
        print(f'You have got error {ex}')