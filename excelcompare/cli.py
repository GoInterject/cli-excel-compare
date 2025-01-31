
# cli framework - https://pypi.org/project/typer/
import typer
from excelcompare.core.excel_compare import ExcelComparator
from excelcompare.core.excel_convertor import ExcelConvertor
import excelcompare.cli_tools as cli_tools

# ::SETUP -------------------------------------------------------------------------- #
app = typer.Typer(add_completion=False, 
                  no_args_is_help=True)
# app.add_typer(cli_tools.app, name="tools", short_help="tools for work with excel")

# ::CLI ---------------------------------------------------------------------------- #
@app.command()
def excelcompare(excel1: str, excel2: str):
    """ compare two excel files in VSCode diff """
    excel_cmp = ExcelComparator()
    return excel_cmp.excel_compare(excel1, excel2)

@app.command(
    context_settings={"allow_extra_args": True, "ignore_unknown_options": True}
)
def tojson(excel_file: str, output: str): # "output" parameter is output folder for file JSON.
    """ Excel convert to JSON file """
    ExcelConvertor.run_excel_convert_to_json(excel_file, output)

# ::EXECUTE ------------------------------------------------------------------------ #
def main():
    app()


if __name__ == "__main__":
    main()
