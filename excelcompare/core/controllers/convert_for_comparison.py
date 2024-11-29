from excelcompare.core.utils.pretty_print import XMLPrettyPrint
from excelcompare.core.models.convert_for_comparison import ExcelParameters
import os
import xml.etree.ElementTree as ET
import xml.dom.minidom
import zipfile
from pathlib import Path
from typing import List

class ExcelConverter:
    def get_xml_tree_from_zip_file(self, zip_ref: zipfile.ZipFile, file_path: str):
        # Read the XML file
        with zip_ref.open(file_path, 'r') as fp: # , encoding='utf-8'
            xml_tree = ET.parse(fp)
        return xml_tree


    def save_sheet_for_compare_view(self, zip_ref: zipfile.ZipFile, file_path: str, res_sheetxml_file_path: Path, file_list: List):
        print(f'Processing {file_path}')
        sheet_tree = self.get_xml_tree_from_zip_file(zip_ref, file_path)
        sheet_root = sheet_tree.getroot()

        # Registered namespaces explicitly to preserve prefixes
        for prefix, uri in ExcelParameters.workbook_namespaces.items():
            ET.register_namespace(prefix, uri)

        sharedStrings_tree = self.get_xml_tree_from_zip_file(zip_ref, ExcelParameters.workbook_sharedStrings_path)
        sharedStrings_root = sharedStrings_tree.getroot()
        # file_path = Path(file_path)

        all_si = sharedStrings_root.findall('si', namespaces=ExcelParameters.workbook_namespaces)

        for sheetData in sheet_root.findall('sheetData', namespaces=ExcelParameters.workbook_namespaces):
            for row in sheetData.findall('row', namespaces=ExcelParameters.workbook_namespaces):
                for c in row.findall('c', namespaces=ExcelParameters.workbook_namespaces):
                    for v in c.findall('v', namespaces=ExcelParameters.workbook_namespaces):
                        v_text = v.text
                        if not v_text.isdigit():
                            continue
                        elem_number = int(v_text)
                        len_all_si = len(all_si)
                        if (elem_number + 1) > len_all_si:
                            continue
                        si_element = all_si[elem_number]
                        t_element = si_element.find('t', namespaces=ExcelParameters.workbook_namespaces)
                        if t_element.attrib:
                            v.attrib.update(t_element.attrib)
                        v.text = t_element.text



        styles_tree = self.get_xml_tree_from_zip_file(zip_ref, ExcelParameters.workbook_styles_path)
        styles_root = styles_tree.getroot()
        mc_namespace = ExcelParameters.workbook_namespaces['mc']
        mc_ignorable = f'{{{mc_namespace}}}Ignorable'
        if mc_ignorable in styles_root.attrib:
            del styles_root.attrib[mc_ignorable]
        # xml_content = file.read()
        sheet_root.append(styles_root)

        workbook_sheet_rels_path = ExcelParameters.worksheets_rels_folder_path + "/" + Path(file_path).name + ".rels"
        if file_list and workbook_sheet_rels_path in file_list:
            rels_tree = self.get_xml_tree_from_zip_file(zip_ref, workbook_sheet_rels_path)
            rels_root = rels_tree.getroot()
            sheet_root.append(rels_root)
        
        for relData in rels_root.findall('Relationship', namespaces=ExcelParameters.workbook_sheet_rels_namespaces):
            if not('TargetMode' in relData.attrib and relData.attrib['TargetMode'] == 'External'):
                parent_path = os.path.abspath("")
                rel_path = Path(ExcelParameters.worksheets_folder_path)/Path(relData.attrib['Target'])
                rel_path = rel_path.resolve()
                rel_path = rel_path.relative_to(parent_path)
                rel_tree = self.get_xml_tree_from_zip_file(zip_ref, rel_path.as_posix())
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



    def convert_sheets_and_save(self, zip_file_path):
        """Extract sheets files in folder of ZIP file and save."""
        excel_path = Path(zip_file_path)
        excel_name = excel_path.stem
        workbook_folder = excel_path.parent.joinpath(excel_name)

        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            # Get the list of file names
            file_list = zip_ref.namelist()
            
            for file_path in file_list:
                if (file_path.startswith(ExcelParameters.worksheets_folder_path)
                    and not file_path.endswith('/')
                    and file_path.endswith('.xml')
                    ):
                    # print(file_name)
                    res_sheetxml_file_path = workbook_folder.joinpath(Path(file_path).name)
                    self.save_sheet_for_compare_view(zip_ref, file_path, res_sheetxml_file_path, file_list)
        return True
        # try:
        # except Exception as e:
        #     print(f"Error converting and saving sheets of Excel file {zip_file_path}: {e}")
        #     return False

