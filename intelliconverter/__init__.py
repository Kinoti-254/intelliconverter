"""IntelliConverter: analyses an Excel sheet, recommends a page layout and exports a PDF.

Modules:
    cli      - command-line options and error messages
    errors   - errors reported to the user without a traceback
    reader   - checks the file and reads a sheet into a SheetModel
    measure  - converts Excel column widths into inches
    layout   - current page-size candidates and the layout score
    textmetrics - measures and wraps text in the PDF font
    planner  - new layout engine (previewed with --plan)
    report   - prints the layout analysis report, then exports the PDF
    render   - builds the PDF with ReportLab
"""
