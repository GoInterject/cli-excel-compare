from excelcompare.core.models.unpack_zip import ZipParameters
from excelcompare.core.utils.pretty_print import XMLPrettyPrint
import zipfile
import os
import xml.etree.ElementTree as ET


class UnpackZip:

    def extract_and_convert_files(self, zip_file_path, extract_folder):

        # Create a directory to extract the files
        os.makedirs(extract_folder, exist_ok=True)

        # Step 1: Unpack the ZIP archive
        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            zip_ref.extractall(extract_folder)

        # Specify the directory containing XML files
        XMLPrettyPrint().convert_xml_files_in_folder(extract_folder)


    def list_files_in_zip(self, zip_file_path, folder_name):
        """List all files in a ZIP file."""
        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            # Get the list of file names
            file_list = zip_ref.namelist()
            
            # Print each file name
            for file_path in file_list:
                if file_path.startswith(folder_name) and not file_path.endswith('/'):
                    print(f'"{file_path}",')

