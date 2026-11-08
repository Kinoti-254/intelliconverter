"""IntelliConverter: analyses an Excel sheet, recommends a page layout and exports a PDF.

Modules:
    reader   - opens a workbook
    measure  - converts Excel column widths into inches
    layout   - page-size candidates and the layout score
    report   - prints the layout analysis report, then exports the PDF
    render   - builds the PDF with ReportLab
"""
