import os
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import Image
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfbase.pdfmetrics import stringWidth

def is_header(row):
    non_empty = row.dropna()
    # First table headers
    table1_keywords = {"Predmeti", "predavanja", "(blank)", "Grand Total"}
    # Second table headers
    table2_keywords = {"Prosečan broj studenata", "Vrsta nastave"}
    
    row_values = {str(cell).strip() for cell in non_empty}
    # Check if this row contains enough header keywords for either table
    is_table1_header = len(row_values.intersection(table1_keywords)) >= 2
    is_table2_header = len(row_values.intersection(table2_keywords)) >= 1
    
    return is_table1_header or is_table2_header

def is_metadata_row(row):
    months = {'jan', 'feb', 'mar', 'apr', 'maj', 'jun', 'jul', 'avg', 'sep', 'okt', 'nov', 'dec'}
    non_empty = [str(cell).strip().lower() for cell in row.dropna()]
    if not non_empty:
        return True
    if any(cell.isdigit() and len(cell) == 4 for cell in non_empty):  # year
        return True
    if any(cell in months for cell in non_empty):
        return True
    if any(cell in {'niš'} for cell in non_empty):
        return True
    return False

def extract_tables_heuristic(df):
    tables = []
    current_table = None
    current_headers = None
    
    for i in range(len(df)):
        row = df.iloc[i]
        row_values = [str(val).strip() if pd.notna(val) else "" for val in row]
        
        # Skip empty rows
        if not any(row_values):
            if current_table is not None:
                if len(current_table) > 0:
                    tables.append(pd.DataFrame(current_table, columns=current_headers))
                current_table = None
                current_headers = None
            continue
        
        # Check if this is a header row
        if is_header(row):
            # If we have a current table, save it before starting new one
            if current_table is not None and len(current_table) > 0:
                tables.append(pd.DataFrame(current_table, columns=current_headers))
            
            # Start new table
            non_empty_indices = [i for i, val in enumerate(row_values) if val]
            current_headers = [row_values[i] for i in non_empty_indices]
            current_table = []
            continue
        
        # If we have a current table, add data row
        if current_headers is not None and not is_metadata_row(row):
            # Extract values for current columns
            row_data = []
            for header in current_headers:
                # Find the value in the row that best matches this column
                value = next((val for val in row_values if val), "")
                row_data.append(value)
            if any(row_data):  # Only add non-empty rows
                current_table.append(row_data)
    
    # Don't forget to add the last table if exists
    if current_table is not None and len(current_table) > 0:
        tables.append(pd.DataFrame(current_table, columns=current_headers))
    
    return tables

def analyze_sheet_structure(filepath, sheet_name="Analiza nastave"):
    try:
        print(f"\n=== Analyzing {sheet_name} Sheet Structure ===")
        df = pd.read_excel(
            filepath,
            sheet_name=sheet_name,
            header=None,
            engine='openpyxl'
        )
        print("\nRaw DataFrame shape:", df.shape)
        print("\nFirst 20 rows of data:")
        for idx, row in df.head(20).iterrows():
            non_empty = [str(cell) for cell in row if pd.notna(cell)]
            if non_empty:
                print(f"Row {idx}: {non_empty}")
        return df
    except Exception as e:
        print(f"Error analyzing sheet: {e}")
        return None

def find_header_and_cols(df, col_names):
    for r_idx, row in df.iterrows():
        row_values = [str(v).strip() for v in row]
        if all(name in row_values for name in col_names):
            col_indices = {name: row_values.index(name) for name in col_names}
            return r_idx, col_indices
    return -1, None

def process_analiza_nastave(filepath):
    try:
        print("\n=== Starting Data Processing (Un-Pivot Method) ===")
        df = pd.read_excel(
            filepath,
            sheet_name="Analiza nastave",
            header=None,
            engine='openpyxl'
        )

        # -- Table 1: Broj časova nastave --
        t1_header_names = ["Predmeti", "predavanja", "(blank)", "Grand Total"]
        t1_header_row, t1_cols = find_header_and_cols(df, t1_header_names)
        if t1_header_row == -1: raise ValueError("Table 1 headers not found")
        
        table1_data = [["Predmeti", "predavanja", "(blank)", "Grand Total"]]
        
        # Context tracking
        current_location, current_year, current_month = "", "", ""
        months = {'jan', 'feb', 'mar', 'apr', 'maj', 'jun', 'jul', 'avg', 'sep', 'okt', 'nov', 'dec'}
        context_values = {"Niš", "2025", "2024", "jan", "okt", "nov", "dec", ""}

        for r_idx in range(t1_header_row + 1, len(df)):
            row = df.iloc[r_idx]
            # Stop Table 1 extraction at the first empty row
            if row.isnull().all() or all(str(cell).strip() == "" for cell in row):
                break

            # Find context, which might be in various columns (like Table 2)
            for cell_val in row:
                s_cell_val = str(cell_val).strip()
                l_cell_val = s_cell_val.lower()
                if l_cell_val == 'niš': current_location = s_cell_val
                elif s_cell_val.isdigit() and len(s_cell_val) == 4: current_year = s_cell_val
                elif l_cell_val in months: current_month = s_cell_val

            # Check for subject and data
            subject = str(row.iloc[t1_cols["Predmeti"]]).strip()
            predavanja_val = row.iloc[t1_cols["predavanja"]]

            # Accept both numeric and string numbers
            is_valid_predavanja = False
            if pd.api.types.is_number(predavanja_val):
                is_valid_predavanja = True
            else:
                try:
                    if str(predavanja_val).strip() and str(predavanja_val).strip().lower() != 'nan':
                        float(predavanja_val)
                        is_valid_predavanja = True
                except Exception:
                    pass

            if subject and subject not in context_values and subject.lower() != "nan" and is_valid_predavanja:
                predmeti_str = f"{current_location} - {current_year} - {current_month} - {subject}"
                data_row = [
                    predmeti_str,
                    predavanja_val,
                    row.iloc[t1_cols["(blank)"]],
                    row.iloc[t1_cols["Grand Total"]]
                ]
                table1_data.append(data_row)

        # -- Table 2: Prosečan broj studenata --
        t2_header_names = ["Predmeti", "Prosečan broj studenata"]
        t2_header_row, t2_cols = find_header_and_cols(df, t2_header_names)
        if t2_header_row == -1: raise ValueError("Table 2 headers not found")

        table2_data = [["Predmeti", "Prosečan broj studenata"]]
        
        current_location, current_year, current_month = "", "", ""

        for r_idx in range(t2_header_row + 1, len(df)):
            row = df.iloc[r_idx]
            # Stop Table 2 extraction at the first empty row
            if row.isnull().all() or all(str(cell).strip() == "" for cell in row):
                break

            # Find context, which might be in various columns
            for cell_val in row:
                s_cell_val = str(cell_val).strip()
                l_cell_val = s_cell_val.lower()
                if l_cell_val == 'niš': current_location = s_cell_val
                elif s_cell_val.isdigit() and len(s_cell_val) == 4: current_year = s_cell_val
                elif l_cell_val in months: current_month = s_cell_val

            # Check for subject and data
            subject = str(row.iloc[t2_cols["Predmeti"]]).strip()
            prosecan_val = row.iloc[t2_cols["Prosečan broj studenata"]]

            # Only add if subject is not empty, not nan, and not a context value
            if subject and subject not in context_values and subject.lower() != "nan" and pd.api.types.is_number(prosecan_val):
                predmeti_str = f"{current_location} - {current_year} - {current_month} - {subject}"
                data_row = [predmeti_str, prosecan_val]
                table2_data.append(data_row)
        
        tables_data = {
            'table1': table1_data,
            'table2': table2_data
        }
        
        print(f"\nExtracted tables successfully. Showing contents:")
        for idx, t in enumerate([table1_data, table2_data], 1):
            print(f"\nTable {idx} (first 5 rows):")
            for row in t[:6]:
                print(row)
                
        # Read professor name
        osnovni_podaci = pd.read_excel(filepath, sheet_name="Osnovni podaci", header=None)
        ime_mask = osnovni_podaci[1] == "Ime"
        prezime_mask = osnovni_podaci[1] == "Prezime"
        if not ime_mask.any() or not prezime_mask.any():
            raise ValueError('Could not find "Ime" or "Prezime" in "Osnovni podaci" sheet')
            
        ime_row = ime_mask.idxmax()
        prezime_row = prezime_mask.idxmax()
        professor_name = f"{str(osnovni_podaci.iloc[ime_row, 2]).strip()} {str(osnovni_podaci.iloc[prezime_row, 2]).strip()}"
        
        return tables_data, professor_name
        
    except Exception as e:
        print(f"Error processing Excel file: {e}")
        import traceback
        traceback.print_exc()
        return None, None

def generate_pdf_an(tables_data, pdf_path, professor_name):
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

    style_heading_center = ParagraphStyle(
        name='HeadingCenter',
        parent=styles['Heading1'],
        fontName='DejaVuSans-Bold',
        fontSize=16,
        alignment=TA_CENTER,
        spaceAfter=12,
    )

    style_sheet_title = ParagraphStyle(
        name='SheetTitle',
        parent=styles['Heading2'],
        fontName='DejaVuSans-Bold',
        fontSize=14
    )

    # Create cell styles
    cell_style = ParagraphStyle(
        'CellStyle',
        parent=style_normal,
        fontSize=9,
        leading=12,
        spaceBefore=3,
        spaceAfter=3,
    )

    header_style = ParagraphStyle(
        'HeaderStyle',
        parent=style_normal,
        fontName='DejaVuSans-Bold',
        fontSize=9,
        leading=12,
        spaceBefore=3,
        spaceAfter=3,
        textColor=colors.white
    )

    # Create PDF document
    doc = SimpleDocTemplate(pdf_path, pagesize=A4,
                          leftMargin=0.5*inch, rightMargin=0.5*inch,
                          topMargin=0.5*inch, bottomMargin=0.5*inch)
    story = []

    # Create header
    logo_path = os.path.join(os.path.dirname(__file__), 'static/images/akademija-logo-boja-latinica 1.png')
    try:
        logo = Image(logo_path, width=150, height=60)
    except:
        logo = Paragraph("", style_normal)

    professor_title = Paragraph(f"<b>{professor_name}</b>", style_heading_center)
    sheet_title = Paragraph("<b>Analiza nastave</b>", style_sheet_title)

    # Single row header with three columns
    header_table = Table(
        [[logo, professor_title, sheet_title]],
        colWidths=[doc.width * 0.2, doc.width * 0.6, doc.width * 0.2]
    )
    header_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),    # Logo aligned left
        ('ALIGN', (1, 0), (1, 0), 'CENTER'),  # Professor name centered
        ('ALIGN', (2, 0), (2, 0), 'RIGHT'),   # Sheet title aligned right
        ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 20))

    # Create first table (Broj časova nastave)
    if tables_data['table1']:
        table1_title = Paragraph("<b>ANALIZA REALIZACIJE NASTAVE PO PREDMETIMA</b>", style_heading_center)
        story.append(table1_title)
        story.append(Spacer(1, 12))

        # Process table data
        table1_processed = []
        for row_idx, row in enumerate(tables_data['table1']):
            processed_row = []
            for col_idx, cell in enumerate(row):
                style = header_style if row_idx == 0 else cell_style
                cell_text = str(cell) if pd.notna(cell) else ""
                processed_row.append(Paragraph(cell_text, style))
            table1_processed.append(processed_row)

        # Calculate column widths
        col_widths = [doc.width * 0.4] + [doc.width * 0.2] * (len(table1_processed[0]) - 1)
        
        table1 = Table(table1_processed, colWidths=col_widths)
        table1.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1242F1')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),  # Left align first column
            ('ALIGN', (1, 0), (-1, -1), 'CENTER'),  # Center align other columns
            ('FONTNAME', (0, 0), (-1, 0), 'DejaVuSans-Bold'),
            ('FONTNAME', (0, 1), (-1, -1), 'DejaVuSans'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('BOX', (0, 0), (-1, -1), 1, colors.black),
            ('LEFTPADDING', (0, 0), (-1, -1), 5),
            ('RIGHTPADDING', (0, 0), (-1, -1), 5),
        ]))
        story.append(table1)
        story.append(Spacer(1, 20))

    # Create second table (Prosečan broj studenata)
    if tables_data['table2']:
        table2_title = Paragraph("<b>ANALIZA PRISUTNOSTI STUDENATA</b>", style_heading_center)
        story.append(table2_title)
        story.append(Spacer(1, 12))

        # Process table data
        table2_processed = []
        for row_idx, row in enumerate(tables_data['table2']):
            processed_row = []
            for col_idx, cell in enumerate(row):
                style = header_style if row_idx == 0 else cell_style
                cell_text = str(cell) if pd.notna(cell) else ""
                processed_row.append(Paragraph(cell_text, style))
            table2_processed.append(processed_row)

        # Calculate column widths
        col_widths = [doc.width * 0.7, doc.width * 0.3]  # 70% for names, 30% for numbers
        
        table2 = Table(table2_processed, colWidths=col_widths)
        table2.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1242F1')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),  # Left align first column
            ('ALIGN', (1, 0), (1, -1), 'CENTER'),  # Center align second column
            ('FONTNAME', (0, 0), (-1, 0), 'DejaVuSans-Bold'),
            ('FONTNAME', (0, 1), (-1, -1), 'DejaVuSans'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('BOX', (0, 0), (-1, -1), 1, colors.black),
            ('LEFTPADDING', (0, 0), (-1, -1), 5),
            ('RIGHTPADDING', (0, 0), (-1, -1), 5),
        ]))
        story.append(table2)
    
    doc.build(story)

# File processing
filepath = "Novi_Izveštaj o radu_za_Nastavnike_Natasa_Bogdanovic - Copy.xlsx"

# First, analyze the structure
analysis_df = analyze_sheet_structure(filepath)

# Then process the data and generate PDF
tables_data, professor_name = process_analiza_nastave(filepath)

if tables_data and professor_name:
    generate_pdf_an(tables_data, "ANtest49.pdf", professor_name)
    print("PDF generated successfully!")
else:
    print("Failed to generate PDF due to data processing errors.") 