
# cli framework - https://pypi.org/project/typer/
import typer

from excelcompare.core.utils.unpack_zip import UnpackZip
from excelcompare.core.models.unpack_zip import ZipParameters
from pathlib import Path


# ::SETUP -------------------------------------------------------------------------- #
app = typer.Typer(add_completion=False, 
                  no_args_is_help=True)

@app.command(
    context_settings={"allow_extra_args": True, "ignore_unknown_options": True}
)
def extract_xml_files(path: str):
    extract_folder = ZipParameters.main_data_folder + "/" + ZipParameters.folder_for_extracted_files + "/" + Path(path).stem
    UnpackZip().extract_and_convert_files(path, extract_folder)

@app.command(
    context_settings={"allow_extra_args": True, "ignore_unknown_options": True}
)
def list_files_in_zip(path: str):
    zip_file_path: str = path
    folder_name: str = ""
    if '>' in path:
        zip_file_path, folder_name = path.split('>', 1)
    UnpackZip().list_files_in_zip(zip_file_path, folder_name)
