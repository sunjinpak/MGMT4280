#!/usr/bin/env python3
"""Fix table column widths in Syllabus Word document."""

from docx import Document
from docx.shared import Inches, Pt, Twips
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

def set_table_borders(table):
    """Add black borders to a table."""
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else parse_xml(r'<w:tblPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')

    tblBorders = parse_xml(
        r'<w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        r'<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        r'</w:tblBorders>'
    )

    tblPr.append(tblBorders)
    if tbl.tblPr is None:
        tbl.insert(0, tblPr)

def set_column_width(column, width):
    """Set width for a table column."""
    for cell in column.cells:
        cell.width = width

def fix_syllabus():
    """Fix the Syllabus Word document."""
    doc = Document("Syllabus.docx")

    for table in doc.tables:
        # Add borders to all tables
        set_table_borders(table)

        # Check if this is the Weekly Schedule table (2 columns: Class, Contents)
        if len(table.columns) == 2:
            # Check first row header
            first_cell_text = table.rows[0].cells[0].text.strip().lower()
            if 'class' in first_cell_text or 'week' in first_cell_text:
                # This is likely the Weekly Schedule table
                # Set column widths: Class = 1.5 inches, Contents = 5.5 inches
                for row in table.rows:
                    row.cells[0].width = Inches(1.5)
                    row.cells[1].width = Inches(5.5)
                print("Fixed Weekly Schedule table column widths")

    doc.save("Syllabus.docx")
    print("Saved: Syllabus.docx")

if __name__ == "__main__":
    fix_syllabus()
