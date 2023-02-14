from PyPDF2 import PdfMerger

import os

def workbooks_to_pdf(students):
    import jpype
    import asposecells
    jpype.startJVM()
    from asposecells.api import Workbook, PdfSaveOptions

    pdfOptions = PdfSaveOptions()
    pdfOptions.setOnePagePerSheet(True)
    pdfOptions.setPageIndex(1)

    merger = PdfMerger()
    paths = []
    for student in students:
        path = student["path"]
        workbook = Workbook(path)
        pdf_path = path.replace(".xlsx", ".pdf")
        workbook.save(pdf_path, pdfOptions)
        paths.append(pdf_path)

    jpype.shutdownJVM()

    for path in paths:
        merger.append(path)
    enclosing_folder, _ = os.path.split(path)

    merger.write(os.path.join(enclosing_folder, "COMPLETE.pdf"))
    merger.close()

