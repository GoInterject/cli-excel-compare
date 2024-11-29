from typing import Dict

class ExcelParameters:

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


    ignored_paths_in_excel = {
        "[Content_Types].xml",
        "_rels/.rels",
        "xl/workbook.xml",
        "xl/_rels/workbook.xml.rels",
        "xl/theme/", # "theme1.xml"
        "xl/styles.xml",
        "xl/sharedStrings.xml",
        "docProps/core.xml",
        "docProps/app.xml",
        "docProps/custom.xml",



        # "xl/worksheets/sheet1.xml",
        # "xl/worksheets/sheet2.xml",
        # "xl/worksheets/_rels/sheet1.xml.rels",
        # "xl/worksheets/_rels/sheet2.xml.rels",
        # "xl/tables/table1.xml",
        # "xl/tables/table2.xml",
    }
