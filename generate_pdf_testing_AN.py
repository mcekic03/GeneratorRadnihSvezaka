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
from process_student_data import process_student_attendance

def analyze_sheet_structure(filepath, sheet_name="Analiza nastave"):
    try:
        print(f"\n=== Analyzing {sheet_name} Sheet Structure ===")
        # Read the sheet
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

def process_analiza_nastave(filepath):
    try:
        print("\n=== Starting Data Processing ===")
        df = pd.read_excel(
            filepath,
            sheet_name="Analiza nastave",
            header=None,
            engine='openpyxl'
        )
        
        print("\nDataFrame Shape:", df.shape)
        
        # Initialize table 1 data
        table1_data = []  # For "Broj časova nastave"
        
        # Find the header rows and column positions
        header_row = None
        subheader_row = None
        column_positions = {}
        
        # First find the main header row
        for idx, row in df.iterrows():
            row_values = [str(val).strip() if pd.notna(val) else "" for val in row]
            if 'Broj časova nastave' in row_values:
                header_row = idx
                subheader_row = idx + 1
                break
        
        if header_row is None:
            raise ValueError("Could not find header row")
            
        # Get the actual headers and their positions from the subheader row
        subheader_values = [str(val).strip() if pd.notna(val) else "" for val in df.iloc[subheader_row]]
        for col_idx, value in enumerate(subheader_values):
            if value in ['predavanja', '(blank)', 'Grand Total']:
                column_positions[value] = col_idx
        
        # Initialize headers for table 1
        table1_headers = ["Predmeti", "predavanja", "(blank)", "Grand Total"]
        
        # Add headers to table
        table1_data.append(table1_headers)
        
        # Initialize tracking variables
        current_location = None
        current_year = None
        current_month = None
        months = ['jan', 'feb', 'mar', 'apr', 'maj', 'jun', 'jul', 'avg', 'sep', 'okt', 'nov', 'dec']
        
        # Process data rows for table 1
        for idx, row in df.iterrows():
            if idx <= subheader_row:  # Skip header rows
                continue
                
            row_values = [str(val).strip() if pd.notna(val) else "" for val in row]
            if not any(row_values):  # Skip empty rows
                continue
            
            first_value = next((val for val in row_values if val), "")
            
            # Check for location (e.g., 'Niš')
            if first_value == 'Niš':
                current_location = first_value
                continue
                
            # Check for year
            if first_value.isdigit() and len(first_value) == 4:
                current_year = first_value
                continue
                
            # Check for month
            if first_value.lower() in months:
                current_month = first_value
                continue
                
            # Process subject rows
            if first_value and not first_value.isdigit() and first_value.lower() not in months:
                subject_name = first_value
                
                # Extract values for table 1
                values = {}
                for col_name, col_idx in column_positions.items():
                    if col_idx < len(row_values):
                        val = row_values[col_idx]
                        try:
                            # Try to convert to float to validate numeric value
                            float(str(val).replace(',', '.'))
                            values[col_name] = val
                        except (ValueError, TypeError):
                            values[col_name] = ""
                
                # Only add if we have any valid values
                if any(values.values()):
                    prefix = f"{current_location}-" if current_location else ""
                    full_name = f"{prefix}{current_year}-{current_month}-{subject_name}"
                    row_data = [full_name]
                    for col in ['predavanja', '(blank)', 'Grand Total']:
                        row_data.append(values.get(col, ""))
                    table1_data.append(row_data)
        
        if len(table1_data) <= 1:  # Only headers
            raise ValueError("No data found for table 1")
        
        # Process table 2 data using the separate module
        table2_data = process_student_attendance(filepath)
        if not table2_data:
            raise ValueError("No data found for table 2")
            
        # Store the processed tables
        tables_data = {
            'table1': table1_data,
            'table2': table2_data
        }
        
        print("\nProcessed Table 1 (Broj časova nastave):")
        for row in table1_data:
            print(row)
        
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
                processed_row.append(Paragraph(str(cell), style))
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
                processed_row.append(Paragraph(str(cell), style))
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
filepath = "Novi_Izveštaj o radu_za_Nastavnike_Natasa_Bogdanovic.xlsx"

# First, analyze the structure
analysis_df = analyze_sheet_structure(filepath)

# Then process the data and generate PDF
tables_data, professor_name = process_analiza_nastave(filepath)

if tables_data and professor_name:
    generate_pdf_an(tables_data, "ANtest28.pdf", professor_name)
    print("PDF generated successfully!")
else:
    print("Failed to generate PDF due to data processing errors.") 