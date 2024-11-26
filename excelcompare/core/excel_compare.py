from excelcompare.core.controllers.convert_for_comparison import convert_sheets_and_save
from excelcompare.core.controllers.informer import Informer
from pathlib import Path
import os
import subprocess
import sys


class ExcelComparator:

    def check_sheets_count(self, first_excel_sheet_files, second_excel_sheet_files):
        
        if first_excel_sheet_files != second_excel_sheet_files:
            msg = "Sheets count is not equal"
            self._c_informer = Informer()
            self._c_informer.print_info(self._c_informer.info_level_error, msg)
            difference_sheets = [item for item in first_excel_sheet_files + second_excel_sheet_files if item not in first_excel_sheet_files or item not in second_excel_sheet_files]
            print(difference_sheets)


    def excel_compare(self, excel_first_file_path, excel_second_file_path):
        res1 = convert_sheets_and_save(excel_first_file_path)
        res2 = convert_sheets_and_save(excel_second_file_path)
        if not (res1 and res2):
            return False
        excel_first_file_path = Path(excel_first_file_path)
        excel_second_file_path = Path(excel_second_file_path)
        excel_sheets_first_dir_path = excel_first_file_path.parent.joinpath(excel_first_file_path.stem)
        excel_sheets_second_dir_path = excel_second_file_path.parent.joinpath(excel_second_file_path.stem)

        # List files in the directory
        first_excel_sheet_files = os.listdir(excel_sheets_first_dir_path)
        first_excel_sheet_files = [f for f in first_excel_sheet_files if os.path.isfile(os.path.join(excel_sheets_first_dir_path, f))]

        # List files in the directory
        second_excel_sheet_files = os.listdir(excel_sheets_second_dir_path)
        second_excel_sheet_files = [f for f in second_excel_sheet_files if os.path.isfile(os.path.join(excel_sheets_second_dir_path, f))]

        # second_excel_sheet_files.append("sheet3.xml")
        self.check_sheets_count(first_excel_sheet_files, second_excel_sheet_files)

        print(f'Running compare Excel files')
        for first_excel_sheet_file in first_excel_sheet_files:
            first_excel_sheet_file_path = excel_sheets_first_dir_path.joinpath(first_excel_sheet_file)
            second_excel_sheet_file_path = excel_sheets_second_dir_path.joinpath(first_excel_sheet_file)
            if not second_excel_sheet_file_path.exists():
                continue
            
            command_for_compare = f'code --diff "{first_excel_sheet_file_path.absolute()}" "{second_excel_sheet_file_path.absolute()}"'

            subprocess.run(["powershell", "-c", command_for_compare])

        return True

