from win32com.client import Dispatch
from pathlib import Path


def convert_docx_to_pdf(docx_file: str | Path, save_file: str | Path) -> None:
    # It seems like win32com or 'Word.Application' or something in between
    # has trouble with paths that are not escaped. Converting the path to a
    # `pathlib.Path` instance and then converting it back to a `str` will
    # give it the necessary escapes.
    if not isinstance(docx_file, Path):
        docx_file = Path(docx_file)

    if not isinstance(save_file, Path):
        save_file = Path(save_file)

    word = Dispatch('Word.Application')
    word.Visible = False

    document = word.Documents.Open(str(docx_file))

    # Word will not write over an existing file and will instead throw
    # an error, so it's necessary to delete the file if it exists.
    if save_file.exists():
        save_file.unlink()

    document.SaveAs(save_file, 17)

    document.Close()
    word.Quit()
