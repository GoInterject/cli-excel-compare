
# cli framework - https://pypi.org/project/typer/
import typer
from excelcompare.core.excel_compare import ExcelComparator
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


# ::EXECUTE ------------------------------------------------------------------------ #
def main():
    app()


if __name__ == "__main__":
    main()
