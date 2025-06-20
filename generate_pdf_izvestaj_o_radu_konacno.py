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
from reportlab.lib.enums import TA_CENTER

def generate_pdf_testing_test1(data, pdf_path, professor_name):
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

    style_sheetName = styles['Heading1']
    style_sheetName.fontName = 'DejaVuSans-Bold'
    style_sheetName.fontSize = 14
    style_sheetName.spaceAfter = 12

    style_heading_center = ParagraphStyle(
    name='HeadingCenter',
    parent=styles['Heading1'],
    alignment=TA_CENTER,
    fontSize=16,
    spaceAfter=12,
)

    # Create a PDF document
    doc = SimpleDocTemplate(pdf_path, pagesize=A4, 
                          leftMargin=0.5*inch, rightMargin=0.5*inch,
                          topMargin=0.5*inch, bottomMargin=0.5*inch)
    story = []

    # Predefined sections
    sections = [
        "Kvalitet nastavnog procesa",
        "Rad sa Studentima",
        "Podizanje kvaliteta ustanove",
        "Jačanje kapaciteta i imidža ustanove",
        "Ostalo"
    ]

    for sheet_name, section_content in data.items():
        logo_path = os.path.join(os.path.dirname(__file__), 'static/images/akademija-logo-boja-latinica 1.png')  # <-- change if needed
        try:
            logo = Image(logo_path, width=150, height=60)  # You can adjust size
        except:
            logo = Paragraph("", style_normal)

        heading_title = Paragraph("<b>Izveštaj o Radu</b>", style_heading_center)
        professor_title = Paragraph(f"<b>{professor_name}</b>", style_heading_center)
        sheet_title = Paragraph(sheet_name.upper(), style_sheetName)

        header_table = Table(
            [[logo, heading_title, sheet_title],
             ['', professor_title, '']],
            colWidths=[doc.width * 0.2, doc.width * 0.6, doc.width * 0.2]
        )
        header_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (0, 0), 'LEFT'),
            ('ALIGN', (1, 0), (1, -1), 'CENTER'),
            ('ALIGN', (2, 0), (2, 0), 'RIGHT'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
        ]))
        story.append(header_table)
        story.append(Spacer(1, 12))

        for section in sections:
            items = section_content.get(section, [])
            if not items:
                continue

            filtered_items = []
            for cols in items:
                cols = list(cols) + [''] * (4 - len(cols))
                col_c = cols[1]
                if col_c and str(col_c).strip():
                    filtered_items.append(cols)

            if not filtered_items:
                continue

            story.append(Paragraph(section, style_heading2))

            if section == "Rad sa Studentima":
                # Four-column layout (unchanged)
                data_table = []
                for cols in items:
                    cols = list(cols) + [''] * (4 - len(cols))
                    col_b, col_c, col_d, col_e = cols[:4]
                    
                    row_content = []
                    for i, col in enumerate([col_b, col_c, col_d, col_e]):
                        if col and str(col).strip():
                            clean_col = str(col).replace("Kratak opis (max 30 reči)", "").strip()
                            if clean_col:
                                wrapped = "<br/>".join(textwrap.wrap(clean_col, width=40))
                                row_content.append(Paragraph(f"<b>{chr(66+i)}:</b> {wrapped}", style_normal))
                            else:
                                row_content.append(Paragraph("", style_normal))
                        else:
                            row_content.append(Paragraph("", style_normal))
                    
                    data_table.append(row_content)

                if data_table:
                    col_width = doc.width - inch
                    col_widths=[
                        col_width * 0.35, #B width
                        col_width * 0.35, #C width
                        col_width * 0.15, #D width
                        col_width * 0.15 #E width
                    ]
                    table = Table(data_table, colWidths=col_widths)
                    table.setStyle(TableStyle([
                        ('BOX', (0,0), (-1,-1), 1, colors.black),
                        ('VALIGN', (0,0), (-1,-1), 'TOP'),
                        ('PADDING', (0,0), (-1,-1), 6),
                        ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
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
                            content_paragraphs.append(Paragraph(f"{wrapped}", style_normal))
                    
                    if col_c:
                        clean_c = str(col_c).replace("Kratak opis (max 30 reči)", "").strip()
                        if clean_c:
                            wrapped = "<br/>".join(textwrap.wrap(clean_c, width=120))
                            content_paragraphs.append(Paragraph(f"{wrapped}", style_normal))
                    
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
                        left_content = Paragraph(f"{clean_b}" if clean_b else "", style_normal)
                        right_content = Paragraph(f"{clean_c}" if clean_c else "", style_normal)
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

def process_excel(filepath, selected_sheet=None):
    try:
        # Get all sheets from the Excel file
        xls = pd.ExcelFile(filepath)
        all_sheets = xls.sheet_names
        
        # Filter out special sheets
        excluded_sheets = ["Analiza nastave", "Evidencija drzanja nastave", "Osnovni podaci", "PadajucaLista"]
        available_sheets = [sheet for sheet in all_sheets if sheet not in excluded_sheets]
        
        if not available_sheets:
            raise ValueError("No monthly report sheets found in the Excel file")
        
        # If no sheet is selected, use the first available one
        sheet_to_process = selected_sheet if selected_sheet in available_sheets else available_sheets[0]
        
        # Read the selected sheet
        sheets = pd.read_excel(filepath, sheet_name=[sheet_to_process], header=None)
        extracted_data = {}

        # Read "Osnovni podaci" sheet for professor's name
        osnovni_podaci = pd.read_excel(filepath, sheet_name="Osnovni podaci", header=None)
        
        # Extract professor's name
        ime_mask = osnovni_podaci[1] == "Ime"
        prezime_mask = osnovni_podaci[1] == "Prezime"
        
        if not ime_mask.any() or not prezime_mask.any():
            raise ValueError('Could not find "Ime" or "Prezime" in "Osnovni podaci" sheet')
            
        ime_row = ime_mask.idxmax()
        prezime_row = prezime_mask.idxmax()
        
        professor_name = f"{str(osnovni_podaci.iloc[ime_row, 2]).strip()} {str(osnovni_podaci.iloc[prezime_row, 2]).strip()}"
        
        # Process the selected sheet
        for sheet, df in sheets.items():
            sections = [
                "Kvalitet nastavnog procesa",
                "Rad sa Studentima",
                "Podizanje kvaliteta ustanove",
                "Jačanje kapaciteta i imidža ustanove",
                "Ostalo"
            ]
            sections_lower = [s.lower() for s in sections]
            
            df = df.map(lambda x: str(x).strip() if pd.notna(x) and str(x).strip() != "" else "")
            section_data = {section: [] for section in sections}
            current_section = None
            last_ostalo_index = -1

            # Find last "Ostalo" occurrence
            for i, row in enumerate(df.values.tolist()):
                row_cells = [cell for cell in row if cell]
                if row_cells and row_cells[0].lower() == "ostalo":
                    last_ostalo_index = i

            # Process data - now collecting all 4 columns (B, C, D, E)
            for i, row in enumerate(df.values.tolist()):
                row_cells = [cell for cell in row if cell]
                if not row_cells:
                    continue

                first_cell = row_cells[0].lower()
                if first_cell in sections_lower:
                    if first_cell == "ostalo" and i == last_ostalo_index:
                        current_section = "Ostalo"
                    elif first_cell == "ostalo":
                        if current_section:
                            # Collect all 4 columns
                            cols = (
                                row[1] if len(row) > 1 else "",
                                row[2] if len(row) > 2 else "",
                                row[3] if len(row) > 3 else "",
                                row[4] if len(row) > 4 else ""
                            )
                            section_data[current_section].append(cols)
                    else:
                        current_section = sections[sections_lower.index(first_cell)]
                    continue

                if current_section:
                    # Collect all 4 columns for all rows
                    cols = (
                        row[1] if len(row) > 1 else "",
                        row[2] if len(row) > 2 else "",
                        row[3] if len(row) > 3 else "",
                        row[4] if len(row) > 4 else ""
                    )
                    section_data[current_section].append(cols)

            extracted_data[sheet] = section_data

        return extracted_data, professor_name, available_sheets
    except Exception as e:
        print(f"Error processing Excel file: {e}")
        import traceback
        traceback.print_exc()
        return None, None, []

if __name__ == "__main__":
    # File processing
    filepath = "Novi_Izveštaj o radu_za_Nastavnike_Natasa_Bogdanovic.xlsx"
    
    try:
        extracted_data, professor_name, available_sheets = process_excel(filepath)
        if extracted_data and professor_name:
            generate_pdf_testing_test1(extracted_data, "test42.pdf", professor_name)
            print("PDF generated successfully")
            print("Available sheets:", available_sheets)
    except Exception as e:
        print(f"Error processing file or generating PDF: {e}")