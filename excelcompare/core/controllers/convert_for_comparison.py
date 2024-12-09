from excelcompare.core.utils.pretty_print import XMLPrettyPrint
from excelcompare.core.models.convert_for_comparison import ExcelParameters
from excelcompare.core.controllers.informer import Informer
import os
import xml.etree.ElementTree as ET
import xml.dom.minidom
import zipfile
import re
from pathlib import Path
from typing import List, Dict, Union

class ExcelConverter:

    def __init__(self):
        self._c_informer = Informer()
        self._m_excel_parameters = ExcelParameters()
        self.excel_paths_usage: Dict = {}

    def get_xml_tree_from_zip_file(self, zip_ref: zipfile.ZipFile, file_path: str):
        # Read the XML file
        try:
            with zip_ref.open(file_path, 'r') as fp: # , encoding='utf-8'
                xml_tree = ET.parse(fp)
            self.excel_paths_usage_set(file_path)
            return xml_tree
        except ET.ParseError:
            return False  # If a ParseError occurs, it's not an XML file
        except Exception as e:
            print(f"An error occurred: {e}")
            return False  # For any other errors (e.g., file not found, permission issues)


    def save_sheet_for_compare_view(self, zip_ref: zipfile.ZipFile, file_path: str, res_sheetxml_file_path: Path, file_list: List):
        print(f'Processing {file_path}')
        sheet_tree = self.get_xml_tree_from_zip_file(zip_ref, file_path)
        sheet_root = sheet_tree.getroot()

        # Registered namespaces explicitly to preserve prefixes
        for prefix, uri in self._m_excel_parameters.workbook_namespaces.items():
            ET.register_namespace(prefix, uri)

        sharedStrings_tree = self.get_xml_tree_from_zip_file(zip_ref, self._m_excel_parameters.workbook_sharedStrings_path)
        sharedStrings_root = sharedStrings_tree.getroot()
        # file_path = Path(file_path)

        all_si = sharedStrings_root.findall('si', namespaces=self._m_excel_parameters.workbook_namespaces)

        for sheetData in sheet_root.findall('sheetData', namespaces=self._m_excel_parameters.workbook_namespaces):
            for row in sheetData.findall('row', namespaces=self._m_excel_parameters.workbook_namespaces):
                for c in row.findall('c', namespaces=self._m_excel_parameters.workbook_namespaces):
                    for v in c.findall('v', namespaces=self._m_excel_parameters.workbook_namespaces):
                        v_text = v.text
                        if not v_text or not v_text.isdigit():
                            continue
                        elem_number = int(v_text)
                        len_all_si = len(all_si)
                        if (elem_number + 1) > len_all_si:
                            continue
                        si_element = all_si[elem_number]
                        t_element = si_element.find('t', namespaces=self._m_excel_parameters.workbook_namespaces)
                        if t_element.attrib:
                            v.attrib.update(t_element.attrib)
                        v.text = t_element.text


        styles_tree = self.get_xml_tree_from_zip_file(zip_ref, self._m_excel_parameters.workbook_styles_path)
        styles_root = styles_tree.getroot()
        mc_namespace = self._m_excel_parameters.workbook_namespaces['mc']
        mc_ignorable = f'{{{mc_namespace}}}Ignorable'
        if mc_ignorable in styles_root.attrib:
            del styles_root.attrib[mc_ignorable]
        # xml_content = file.read()
        sheet_root.append(styles_root)

        workbook_sheet_rels_path = self._m_excel_parameters.worksheets_rels_folder_path + "/" + Path(file_path).name + ".rels"
        if file_list and workbook_sheet_rels_path in file_list:
            rels_tree = self.get_xml_tree_from_zip_file(zip_ref, workbook_sheet_rels_path)
            rels_root = rels_tree.getroot()
            sheet_root.append(rels_root)
        
        for relData in rels_root.findall('Relationship', namespaces=self._m_excel_parameters.workbook_sheet_rels_namespaces):
            if not('TargetMode' in relData.attrib and relData.attrib['TargetMode'] == 'External'):
                parent_path = os.path.abspath("")
                rel_path = Path(self._m_excel_parameters.worksheets_folder_path)/Path(relData.attrib['Target'])
                rel_path = rel_path.resolve()
                rel_path = rel_path.relative_to(parent_path)
                rel_tree = self.get_xml_tree_from_zip_file(zip_ref, rel_path.as_posix())
                if rel_tree:
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

    def search_path(self, pattern_text: str, string: str):
        pattern_text = pattern_text.replace("\\", r"\\")
        pattern_text = pattern_text.replace("[", r"\[")
        pattern_text = pattern_text.replace("]", r"\]")
        return re.search("^" + pattern_text, string)

    def is_exclude(self, exclude_lst, path: str):
        for exclude_name in exclude_lst:
            if self.search_path(exclude_name, path):
                return True
        return False

    def ignore_check(self, excel_paths_usage_dict: Dict, xml_file_path: str):
        if self.is_exclude(self._m_excel_parameters.ignored_paths_in_excel, xml_file_path):
            excel_paths_usage_dict[xml_file_path]["usage"] = True

    def check_existion_missed_files(self, excel_paths_usage_dict: Dict):
        msg = ""
        for file_path in excel_paths_usage_dict.keys():
            if not ("usage" in excel_paths_usage_dict[file_path] and excel_paths_usage_dict[file_path]["usage"]):
                msg += f'"{file_path}",\n'
        if msg:
            msg = "These files were missed by the converter: {\n" + msg + "}"
            self._c_informer.print_info(self._c_informer.info_level_error, msg)


    def convert_sheets_and_save(self, zip_file_path):
        """Extract sheets files in folder of ZIP file and save."""
        excel_path = Path(zip_file_path)
        excel_name = excel_path.stem
        workbook_folder = excel_path.parent.joinpath(excel_name)

        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            # Get the list of file names
            file_list = zip_ref.namelist()

            usage_dict: Dict = {"usage": False}
            for key in file_list:
                self.excel_paths_usage[key] = usage_dict.copy()

            for file_path in file_list:
                self.ignore_check(self.excel_paths_usage, file_path)
                if (file_path.startswith(self._m_excel_parameters.worksheets_folder_path)
                    and not file_path.endswith('/')
                    and file_path.endswith('.xml')
                    ):
                    res_sheetxml_file_path = workbook_folder.joinpath(Path(file_path).name)
                    self.save_sheet_for_compare_view(zip_ref, file_path, res_sheetxml_file_path, file_list)
        self.check_existion_missed_files(self.excel_paths_usage)
        return True
        # try:
        # except Exception as e:
        #     print(f"Error converting and saving sheets of Excel file {zip_file_path}: {e}")
        #     return False

    def excel_paths_usage_set(self, used_path: Union[str, Path]):
        if isinstance(used_path, Path):
            used_path = used_path.as_posix()
        self.excel_paths_usage.update({used_path: {"usage": True}})
