from win32com.client import Dispatch
from pathlib import Path


def convert_xlsx_to_pdf(xlsx_file: str | Path, save_file: str | Path) -> None:
    # It seems like win32com or 'Excel.Application' or something in between
    # has trouble with paths that are not escaped. Converting the path to a
    # `pathlib.Path` instance and then converting it back to a `str` will
    # give it the necessary escapes.
    if not isinstance(xlsx_file, Path):
        xlsx_file = Path(xlsx_file)

    if not isinstance(save_file, Path):
        save_file = Path(save_file)

    excel = Dispatch('Excel.Application')
    excel.Visible = False

    workbook = excel.Workbooks.Open(str(xlsx_file))

    # Excel will not write over an existing file and will instead throw
    # an error, so it's necessary to delete the file if it exists.
    if save_file.exists():
        save_file.unlink()


    # Some settings to make the output look better
    for i in range(workbook.Worksheets.Count):
        worksheet = workbook.Worksheets[i]

        worksheet.PageSetup.Zoom = False
        worksheet.PageSetup.FitToPagesTall = 1
        worksheet.PageSetup.FitToPagesWide = 1

    workbook.ExportAsFixedFormat(0, save_file)

    workbook.Close()
    excel.Quit()
