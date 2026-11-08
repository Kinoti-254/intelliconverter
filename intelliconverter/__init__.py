"""IntelliConverter: analyses an Excel sheet, recommends a page layout and exports a PDF.

Modules:
    cli      - command-line options and error messages
    errors   - errors reported to the user without a traceback
    reader   - checks the file and reads a sheet into a SheetModel
    numberformat - shows values the way Excel does (number and date formats)
    textmetrics  - measures and wraps text in the PDF font
    planner  - the layout engine: paper, orientation, font, widths, pages
    report   - prints the layout report, then exports the PDF
    render   - draws the PDF exactly as planned, with ReportLab
    measure  - converts Excel column widths into inches (kept for a future
               "preserve Excel widths" mode)
"""
