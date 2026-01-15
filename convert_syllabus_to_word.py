#!/usr/bin/env python3
"""Convert Syllabus markdown to Word with proper table formatting."""

import re
import subprocess
import os
from docx import Document
from docx.shared import Cm, Pt
from docx.oxml import parse_xml
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

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

def preprocess_markdown(input_file, output_file):
    """Preprocess markdown to handle line breaks in tables properly."""
    with open(input_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # Remove the div wrapper for black-border-table (not needed for Word)
    content = re.sub(r'<div class="black-border-table" markdown="1">\s*\n?', '', content)
    content = re.sub(r'\n?</div>', '', content)

    # Remove Jekyll front matter
    content = re.sub(r'^---\n.*?\n---\n', '', content, flags=re.DOTALL)

    # Remove navigation links at top
    content = re.sub(r'^\[Home\]\(index\)\n+', '', content)
    content = re.sub(r'^\*\*\[Download Word Document\].*?\*\*\n+', '', content)

    # In table cells, replace <br> with a special marker that will survive pandoc
    # Use Unicode line separator character
    content = content.replace('<br>', '⏎')

    # Replace horizontal rules within table cells (-----) with newline marker
    # This pattern matches table cells containing -----
    def replace_dashes_in_tables(match):
        cell_content = match.group(0)
        # Replace ----- patterns with nothing (they're just visual separators)
        cell_content = re.sub(r'-{5,}', '', cell_content)
        return cell_content

    # Process each table row
    lines = content.split('\n')
    processed_lines = []
    for line in lines:
        if line.startswith('|') and '-----' in line and not re.match(r'^\|[-:\s|]+\|$', line):
            # This is a table data row with dashes - remove them
            line = re.sub(r'-{5,}', '', line)
        processed_lines.append(line)
    content = '\n'.join(processed_lines)

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(content)

def postprocess_word(docx_file):
    """Post-process Word document to fix table formatting."""
    doc = Document(docx_file)

    # Add borders to all tables
    for table in doc.tables:
        set_table_borders(table)

    # Find the Weekly Schedule table (2 columns: Class, Contents)
    for table in doc.tables:
        if len(table.columns) == 2:
            first_cell = table.rows[0].cells[0].text.strip().lower()
            if 'class' in first_cell:
                print(f"Found Weekly Schedule table with {len(table.rows)} rows")

                for row_idx, row in enumerate(table.rows):
                    # Set column widths
                    row.cells[0].width = Cm(3)
                    row.cells[1].width = Cm(15)

                    # Process each cell
                    for cell in row.cells:
                        # Get the current text with the marker
                        full_text = cell.text

                        if '⏎' in full_text:
                            # Clear all paragraphs in the cell
                            for p in cell.paragraphs:
                                p.clear()

                            # Split by the marker and rebuild with line breaks
                            parts = full_text.split('⏎')

                            # Use the first paragraph
                            para = cell.paragraphs[0]

                            for i, part in enumerate(parts):
                                part = part.strip()
                                if not part:
                                    continue

                                if i > 0 and para.runs:
                                    # Add a line break before this part
                                    run = para.add_run()
                                    run.add_break()

                                # Add the text
                                para.add_run(part)

                print("Fixed Weekly Schedule table")

    doc.save(docx_file)
    print(f"Saved: {docx_file}")

def main():
    os.chdir('/tmp/MGMT4280')

    # Step 1: Preprocess markdown
    print("Step 1: Preprocessing markdown...")
    preprocess_markdown('syllabus.md', 'syllabus_temp.md')

    # Step 2: Convert to Word with pandoc
    print("Step 2: Converting to Word with pandoc...")
    subprocess.run([
        'pandoc', 'syllabus_temp.md',
        '-o', 'Syllabus.docx',
        '--from=markdown+raw_html'
    ], check=True)

    # Step 3: Post-process Word document
    print("Step 3: Post-processing Word document...")
    postprocess_word('Syllabus.docx')

    # Cleanup
    os.remove('syllabus_temp.md')
    print("Done!")

if __name__ == "__main__":
    main()
