import xml.etree.ElementTree as ET
import pandas as pd
from styleframe import StyleFrame, Styler, utils
import re

templates_name = {
    "retail_intcard": "ICC_XML_Template",
    "retail_inthe": "IHE_XML_Template",
    "retail_intlothcons": "IOC_XML_Template",
    "retail_intsb": "ISB_XML_Template",
    "retail_intauto": "RIAL_XML_Template",
    "retail_intfm": "RIFM_XML_Template",
    "retail_student": "RSL_XML_Template",
    "retail_auto": "USAL_XML_Template",
    "retail_usothcons": "USOC_XML_Template",
    "retail_ussb": "USSB_XRETAIL_USOTHCONSML_Template",
    "cil": "CORP_XML_Template",
    "cre": "CRE_XML_Template"
}


def define_sheet_name(path):
    name = re.sub(r'_version[0-9]{1,2}\.xsd', '', path.lower())
    return templates_name.get(name, 'XML_Template')


def get_root_collection(path):

    tree = ET.parse(path)
    root_collection = []
    for element in tree.findall('./'):
        root_collection.append(element.attrib)
    return root_collection[-1].get('name','')


def filter_element(path):
    tree = ET.parse(path)
    array_valid_element = []
    for element in tree.findall('.//'):
        if element.tag.endswith('element') or element.tag.endswith('attribute'):
            array_valid_element.append(element)
    return array_valid_element


def define_length_column(df):
    return max(df.apply(len))


def create_template(array_valid_element, root_last, sheet_name):
    template = []
    collection = ''

    for element in array_valid_element:
        if len(element.attrib) == 1 and 'ref' not in element.attrib:
            collection = element.attrib.get('name', '')
        else:
            if element.tag.endswith('element') and 'ref' not in element.attrib:
                template.append([element.get('name', ''), root_last + '.' + collection+'.',collection, 'ELEMENT'])
            elif 'ref' not in element.attrib:
                if element.attrib.get('name') in ['DATA_ASOF_TSTMP', 'LAST_ASOF_TSTMP']:
                    template.append([element.get('name', ''), root_last + '/', collection, 'ATTRIBUTE'])
                else:
                    template.append([element.get('name', ''), root_last + '.' + collection + '/',
                                     collection,  'ATTRIBUTE'])

    df = pd.DataFrame(template, columns= ['mdrm', 'xml_path', 'collection', 'type'])
    excel_writer = StyleFrame.ExcelWriter('XML_Template.xlsx')
    sf = StyleFrame(df).\
        apply_column_style(['mdrm', 'xml_path', 'collection', 'type'],
                           styler_obj=Styler(font_size=11, font_color=None,
                                             horizontal_alignment=utils.horizontal_alignments.left,
                                             wrap_text=False, shrink_to_fit=False))

    sf.set_column_width_dict({'mdrm': define_length_column(df['mdrm'])+8,
                              'xml_path':define_length_column(df['xml_path'])+10,
                              'collection': define_length_column(df['collection'])+10,
                              'type': define_length_column(df['type'])+5})
    sf.to_excel(excel_writer=excel_writer, sheet_name=sheet_name)

    excel_writer.save()


if __name__ == '__main__':
    try:
        path = input('Enter path to xsd schema\n')
        root_collection = get_root_collection(path)
        array = filter_element(path)
        sheet_name = define_sheet_name(path)
        create_template(array, root_collection, sheet_name)
    except Exception as e:
        print(e)

    input('Press any button to end this program\n')