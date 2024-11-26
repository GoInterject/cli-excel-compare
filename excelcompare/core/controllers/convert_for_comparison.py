import os
import xml.etree.ElementTree as ET
import xml.dom.minidom
import zipfile
from pathlib import Path
from excelcompare.core.utils.pretty_print import XMLPrettyPrint
from typing import Dict, List


worksheets_folder_path = "xl/worksheets"
worksheets_rels_folder_path = "xl/worksheets/_rels"
workbook_styles_path = "xl/styles.xml"
workbook_sharedStrings_path = "xl/sharedStrings.xml"
workbook_xml_path = "xl/workbook.xml"
data_folder = "data_and_files/"

workbook_namespaces: Dict = {
    "": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    "xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
    "xr3": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"
}

workbook_sheet_rels_namespaces: Dict = {
    "": "http://schemas.openxmlformats.org/package/2006/relationships"
}

def get_xml_tree_from_zip_file(zip_ref: zipfile.ZipFile, file_path: str):
    # Read the XML file
    with zip_ref.open(file_path, 'r') as fp: # , encoding='utf-8'
        xml_tree = ET.parse(fp)
    return xml_tree


def save_sheet_for_compare_view(zip_ref: zipfile.ZipFile, file_path: str, res_sheetxml_file_path: Path, file_list: List):
    print(f'Processing {file_path}')
    sheet_tree = get_xml_tree_from_zip_file(zip_ref, file_path)
    sheet_root = sheet_tree.getroot()

    # Registered namespaces explicitly to preserve prefixes
    for prefix, uri in workbook_namespaces.items():
        ET.register_namespace(prefix, uri)

    sharedStrings_tree = get_xml_tree_from_zip_file(zip_ref, workbook_sharedStrings_path)
    sharedStrings_root = sharedStrings_tree.getroot()
    # file_path = Path(file_path)

    all_si = sharedStrings_root.findall('si', namespaces=workbook_namespaces)

    for sheetData in sheet_root.findall('sheetData', namespaces=workbook_namespaces):
        for row in sheetData.findall('row', namespaces=workbook_namespaces):
            for c in row.findall('c', namespaces=workbook_namespaces):
                for v in c.findall('v', namespaces=workbook_namespaces):
                    v_text = v.text
                    if not v_text.isdigit():
                        continue
                    elem_number = int(v_text)
                    len_all_si = len(all_si)
                    if (elem_number + 1) > len_all_si:
                        continue
                    si_element = all_si[elem_number]
                    t_element = si_element.find('t', namespaces=workbook_namespaces)
                    if t_element.attrib:
                        v.attrib.update(t_element.attrib)
                    v.text = t_element.text



    styles_tree = get_xml_tree_from_zip_file(zip_ref, workbook_styles_path)
    styles_root = styles_tree.getroot()
    mc_namespace = workbook_namespaces['mc']
    mc_ignorable = f'{{{mc_namespace}}}Ignorable'
    if mc_ignorable in styles_root.attrib:
        del styles_root.attrib[mc_ignorable]
    # xml_content = file.read()
    sheet_root.append(styles_root)

    workbook_sheet_rels_path = worksheets_rels_folder_path + "/" + Path(file_path).name + ".rels"
    if file_list and workbook_sheet_rels_path in file_list:
        rels_tree = get_xml_tree_from_zip_file(zip_ref, workbook_sheet_rels_path)
        rels_root = rels_tree.getroot()
        sheet_root.append(rels_root)
    
    for relData in rels_root.findall('Relationship', namespaces=workbook_sheet_rels_namespaces):
        if not('TargetMode' in relData.attrib and relData.attrib['TargetMode'] == 'External'):
            parent_path = os.path.abspath("")
            rel_path = Path(worksheets_folder_path)/Path(relData.attrib['Target'])
            rel_path = rel_path.resolve()
            rel_path = rel_path.relative_to(parent_path)
            rel_tree = get_xml_tree_from_zip_file(zip_ref, rel_path.as_posix())
            rel_root = rel_tree.getroot()
            sheet_root.append(rel_root)

    xml_str = ET.tostring(sheet_root, 'utf-8')
    # Pretty print the XML
    pretty_xml = XMLPrettyPrint().pretty_print_xml(xml_str)

    res_sheetxml_file_path.parent.mkdir(parents = True, exist_ok = True)

    # sheet_tree.write(res_sheetxml_file_path, xml_declaration=True, encoding="UTF-8") # sheet_root.write()
    # Write the pretty-printed XML back to the file
    with open(res_sheetxml_file_path, 'w', encoding='utf-8') as file:
        file.write(pretty_xml)



def convert_sheets_and_save(zip_file_path):
    """Extract sheets files in folder of ZIP file and save."""
    excel_path = Path(zip_file_path)
    excel_name = excel_path.stem
    workbook_folder = excel_path.parent.joinpath(excel_name)

    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        # Get the list of file names
        file_list = zip_ref.namelist()
        
        for file_path in file_list:
            if (file_path.startswith(worksheets_folder_path)
                and not file_path.endswith('/')
                and file_path.endswith('.xml')
                ):
                # print(file_name)
                res_sheetxml_file_path = workbook_folder.joinpath(Path(file_path).name)
                save_sheet_for_compare_view(zip_ref, file_path, res_sheetxml_file_path, file_list)
    return True
    # try:
    # except Exception as e:
    #     print(f"Error converting and saving sheets of Excel file {zip_file_path}: {e}")
    #     return False

