import pdfplumber
import pandas as pd
import re


def parse_user_guide(path, start_page, end_page):

    mdrms = []
    with pdfplumber.open(path) as pdf:
        for page in range(start_page-1, end_page):
            for line in pdf.pages[page].extract_text().split('\n'):
                md = re.findall(r'[A-Z]{3,6}[0-9]{2,5}', line)
                if md:
                    mdrms.extend(md)

    df = pd.DataFrame(mdrms)
    df.to_excel('Elec_sub.xlsx', index=False)


if __name__ == '__main__':
    try:
        path = input('Enter path to user guide\n')
        start_page = int(input('Enter start page\n'))
        end_page = int(input('Enter end page\n'))
        parse_user_guide(path, start_page, end_page)
    except Exception as ex:
        print(f'You have got error {ex}')
    input('Press any button to end this program\n')
