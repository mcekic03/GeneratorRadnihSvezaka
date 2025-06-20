import os
import textwrap
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import Image
from reportlab.lib.enums import TA_CENTER, TA_RIGHT

def generate_pdf_testing_test1(data, pdf_path):
    # Register fonts
    font_path_regular = os.path.join(os.path.dirname(__file__), 'static/fonts', 'DejaVuSans.ttf')
    font_path_bold = os.path.join(os.path.dirname(__file__), 'static/fonts', 'DejaVuSans-Bold.ttf')
    pdfmetrics.registerFont(TTFont('DejaVuSans', font_path_regular))
    pdfmetrics.registerFont(TTFont('DejaVuSans-Bold', font_path_bold))

    # Define styles
    styles = getSampleStyleSheet()  
    style_normal = styles['Normal']
    style_normal.fontName = 'DejaVuSans'
    style_normal.fontSize = 9
    style_normal.leading = 12

    style_heading1 = styles['Heading1']
    style_heading1.fontName = 'DejaVuSans-Bold'
    style_heading1.fontSize = 14
    style_heading1.spaceAfter = 12

    style_heading2 = styles['Heading2']
    style_heading2.fontName = 'DejaVuSans-Bold'
    style_heading2.fontSize = 12
    style_heading2.spaceAfter = 6

    style_heading2_nobold = styles['Heading2']
    style_heading2_nobold.fontName = 'DejaVuSans'
    style_heading2_nobold.fontSize = 12
    style_heading2_nobold.spaceAfter = 6

    # Add a specific style for sheet title
    style_sheet_title = ParagraphStyle(
        name='SheetTitle',
        parent=style_heading2,  # Inherit from style_heading2 to keep other parameters
        fontSize=16,  # Only change the font size
        fontName='DejaVuSans-Bold'  # Explicitly set bold font
    )

    style_heading_center = ParagraphStyle(
    name='HeadingCenter',
    parent=styles['Heading1'],
    alignment=TA_CENTER,
    fontSize=16,
    spaceAfter=12,
    )

    # Add a custom style for section headers (manual bold)
    style_section_header = ParagraphStyle(
        name='SectionHeader',
        parent=styles['Heading2'],
        fontName='DejaVuSans-Bold',
        fontSize=14,
        spaceAfter=8,
        leading=18,
    )

    # Create a PDF document
    doc = SimpleDocTemplate(pdf_path, pagesize=A4, 
                          leftMargin=0.5*inch, rightMargin=0.5*inch,
                          topMargin=0.5*inch, bottomMargin=0.5*inch)
    story = []

    # Predefined sections
    sections = [
        "Osnovni podaci",
        "Ukupan broj predmeta na kojima je nastavnik angažovan",
        "Ukupno opterećenje",
        "Predmeti na kojima je saradnik angažovan",
        "Članstvo u komisijama (timovima)",
        "Ostala zaduženja",
    ]

    for sheet_name, section_content in data.items():

        # Initialize name variables
        name_data = {"first_name": "", "last_name": ""}
        rows_processed = 0
        
        # Extract professor's name from "Osnovni podaci" section
        for row in section_content.get("Osnovni podaci", []):
            if rows_processed >= 2:  # We only need 2 rows (Ime and Prezime)
                break
            
            if len(row) > 1:  # Ensure row has both label and value
                label = str(row[0]).strip().lower()
                value = str(row[1]).strip()
                
                if "ime" in label and not name_data["first_name"]:
                    name_data["first_name"] = value
                    rows_processed += 1
                elif "prezime" in label and not name_data["last_name"]:
                    name_data["last_name"] = value
                    rows_processed += 1

        # Combine names for heading
        full_name = f"{name_data['first_name']} {name_data['last_name']}".strip()

        heading_text = f"<b>{full_name if full_name else 'Osnovni podaci'}</b>"

        logo_path = os.path.join(os.path.dirname(__file__), 'static/images/akademija-logo-boja-latinica 1.png')  # <-- change if needed
        try:
            logo = Image(logo_path, width=150, height=60)  # You can adjust size
        except:
            logo = Paragraph("", style_normal)

        heading_title = Paragraph(heading_text, style_heading_center)
        sheet_title = Paragraph(sheet_name, style_sheet_title)  # Remove <b> tags since we're using bold font

        header_table = Table(
            [[logo, heading_title, sheet_title]],
            colWidths=[doc.width * 0.2, doc.width * 0.6, doc.width * 0.2]
        )
        header_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (0, 0), 'LEFT'),
            ('ALIGN', (1, 0), (1, 0), 'CENTER'),
            ('ALIGN', (2, 0), (2, 0), 'RIGHT'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
        ]))
        story.append(header_table)
        story.append(Spacer(1, 12))

        for section in sections:
            items = section_content.get(section, [])
            if not items:
                continue

            story.append(Paragraph(f"<b>{section}</b>", style_section_header))

            if section == "Osnovni podaci":
                # Always take first 3 rows for Osnovni podaci
                filtered_items = items[:3]
                
                # Format as label-value pairs
                data_table = []
                for cols in filtered_items:
                    col_b = cols[0] if len(cols) > 0 else ""  # Column B
                    col_c = cols[1] if len(cols) > 1 else ""  # Column C
                    
                    clean_b = str(col_b).strip()
                    clean_c = str(col_c).strip()
                    
                    if clean_b or clean_c:  # Show row even if one is empty
                        data_table.append([
                            Paragraph(f"<b>{clean_b}:</b>", style_heading2),
                            Paragraph(clean_c, style_heading2)
                        ])

                if data_table:
                    total_width = doc.width - inch
                    table = Table(data_table, 
                                  colWidths = [total_width*0.4, total_width*0.6])
                    table.setStyle(TableStyle([
                        ('BOX', (0,0), (-1,-1), 1, colors.black),  # Outer border
                        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                        ('PADDING', (0,0), (-1,-1), 6),
                        ('GRID', (0,0), (-1,-1), 1, colors.lightgrey)
                    ]))
                    story.append(table)

            elif section == "Ostalo":
                # Single column layout for Ostalo (combines B and C)
                content_paragraphs = []
                for cols in items:
                    cols = list(cols) + [''] * (2 - len(cols))
                    col_b, col_c = cols[:2]
                    
                    if col_b:
                        clean_b = str(col_b).replace("Kratak opis (max 30 reči)", "").strip()
                        if clean_b:
                            wrapped = "<br/>".join(textwrap.wrap(clean_b, width=120))
                            content_paragraphs.append(Paragraph(f"{wrapped}", style_heading2))
                    
                    if col_c:
                        clean_c = str(col_c).replace("Kratak opis (max 30 reči)", "").strip()
                        if clean_c:
                            wrapped = "<br/>".join(textwrap.wrap(clean_c, width=120))
                            content_paragraphs.append(Paragraph(f"{wrapped}", style_heading2))
                    
                    content_paragraphs.append(Spacer(1, 6))

                if content_paragraphs:
                    # Remove last spacer if it exists
                    if isinstance(content_paragraphs[-1], Spacer):
                        content_paragraphs.pop()
                    
                    # Create bordered box
                    table = Table([[content_paragraphs]], colWidths=[doc.width - inch])
                    table.setStyle(TableStyle([
                        ('BOX', (0,0), (-1,-1), 1, colors.black),
                        ('PADDING', (0,0), (-1,-1), 6),
                    ]))
                    story.append(table)
            else:
                # Two-column layout with grid between B and C
                data_table = []
                for cols in items:
                    cols = list(cols) + [''] * (2 - len(cols))
                    col_b, col_c = cols[:2]

                    if not col_c or not str(col_c).strip():
                        continue
                    
                    if col_b or col_c:
                        clean_b = str(col_b).replace("Kratak opis (max 30 reči)", "").strip()
                        clean_c = str(col_c).replace("Kratak opis (max 30 reči)", "").strip()
                        left_content = Paragraph(f"{clean_b}" if clean_b else "", style_heading2)
                        right_content = Paragraph(f"{clean_c}" if clean_c else "", style_heading2)
                        data_table.append([left_content, right_content])

                if data_table:
                    col_width = doc.width - inch
                    col_b_width = col_width * 0.55
                    col_c_width = col_width * 0.45
                    table = Table(data_table, colWidths=[col_b_width, col_c_width])
                    table.setStyle(TableStyle([
                        ('BOX', (0,0), (-1,-1), 1, colors.black),  # Outer border
                        ('VALIGN', (0,0), (-1,-1), 'TOP'),
                        ('PADDING', (0,0), (-1,-1), 6),
                        ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),  # Grid between all cells
                    ]))
                    story.append(table)

            story.append(Spacer(1, 12))

    doc.build(story)

def process_excel(filepath):
    try:
        sheet_names = ["Osnovni podaci"]
        sheets = pd.read_excel(filepath, sheet_name=sheet_names, header=None)
        extracted_data = {}
        sections = [
            "Osnovni podaci",
            "Ukupan broj predmeta na kojima je nastavnik angažovan",
            "Ukupno opterećenje",
            "Predmeti na kojima je saradnik angažovan",
            "Članstvo u komisijama (timovima)",
            "Ostala zaduženja",
        ]

        for sheet, df in sheets.items():
            print(f"Processing sheet: {sheet}")
            df = df.map(lambda x: str(x).strip() if pd.notna(x) and str(x).strip() != "" else "")
            
            section_data = {section: [] for section in sections}
            current_section = None
            row_counter = 0

            for row in df.values.tolist():
                row_cells = [cell for cell in row if cell]
                if not row_cells:
                    continue

                # FIRST 3 ROWS = "Osnovni podaci"
                if row_counter < 3:
                    cols = tuple(row[i] if i < len(row) else "" for i in range(1, 5))  # Columns B-E
                    section_data["Osnovni podaci"].append(cols)
                    row_counter += 1
                    continue

                # Process remaining rows normally
                first_cell = row_cells[0].lower()
                if first_cell in [s.lower() for s in sections[1:]]:  # Skip "Osnovni podaci"
                    current_section = sections[[s.lower() for s in sections].index(first_cell)]
                    continue

                if current_section:
                    cols = tuple(row[i] if i < len(row) else "" for i in range(1, 5))
                    section_data[current_section].append(cols)

            extracted_data[sheet] = section_data

        return extracted_data
    except Exception as e:
        print(f"Error processing Excel file: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    # File processing
    filepath = "Novi_Izveštaj o radu_za_Nastavnike_Natasa_Bogdanovic.xlsx"
    extracted_data = process_excel(filepath)
    
    # Generate PDF
    generate_pdf_testing_test1(extracted_data, "OPtest16.pdf")
    print("PDF generated successfully")