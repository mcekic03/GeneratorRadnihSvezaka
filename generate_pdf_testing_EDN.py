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
import datetime
import traceback

def generate_pdf_edn(data, pdf_path, professor_name):
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

    # Create cell style with better spacing
    cell_style = ParagraphStyle(
        'CellStyle',
        parent=style_normal,
        fontSize=9,
        leading=12,
        spaceBefore=3,
        spaceAfter=3,
    )

    # Create header cell style
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
    doc = SimpleDocTemplate(pdf_path, pagesize=landscape(A4),
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
    sheet_title = Paragraph("<b>Evidencija držanja nastave</b>", style_sheet_title)

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
    story.append(Spacer(1, 12))

    # Process and add table data
    if data:
        try:
            # Debug print for data
            print("Data received for table:")
            for row in data[:5]:  # Print first 5 rows
                print(row)

            # First pass: measure content width in each column
            num_cols = len(data[0])
            col_max_widths = [0] * num_cols
            
            for row in data:
                for i, cell in enumerate(row):
                    # Convert to string and calculate width
                    content = str(cell)
                    # Split content into lines and find the widest
                    lines = content.split('\n')
                    for line in lines:
                        width = stringWidth(line, 'DejaVuSans', 9)
                        col_max_widths[i] = max(col_max_widths[i], width)

            # Add padding to the calculated widths
            col_max_widths = [width + 12 for width in col_max_widths]

            # Calculate available width
            available_width = doc.width - inch

            # Generate proportions based on number of columns
            if num_cols > 0:
                base_proportion = 1.0 / num_cols
                min_proportions = []
                max_proportions = []
                
                for i in range(num_cols):
                    if i == 0:  # Date column
                        min_proportions.append(base_proportion * 0.8)  # Increase minimum width for date column
                        max_proportions.append(base_proportion * 1.0)
                    elif i == num_cols // 2:  # Middle column
                        min_proportions.append(base_proportion * 1.5)
                        max_proportions.append(base_proportion * 2.0)
                    else:  # Other columns
                        min_proportions.append(base_proportion)
                        max_proportions.append(base_proportion * 1.2)
                
                # Normalize proportions
                total_min = sum(min_proportions)
                total_max = sum(max_proportions)
                min_proportions = [p/total_min for p in min_proportions]
                max_proportions = [p/total_max for p in max_proportions]

            # Calculate minimum and maximum widths based on proportions
            min_widths = [available_width * prop for prop in min_proportions]
            max_widths = [available_width * prop for prop in max_proportions]

            # Adjust column widths to stay within min/max bounds
            for i in range(len(col_max_widths)):
                if i == 0:  # Date column
                    # Set a fixed width for the date column that's enough for YYYY-MM-DD format
                    col_max_widths[i] = stringWidth('0000-00-00', 'DejaVuSans', 9) + 24  # Add extra padding
                else:
                    col_max_widths[i] = max(min_widths[i], min(max_widths[i], col_max_widths[i]))

            # Convert data to use Paragraphs for text wrapping
            wrapped_data = []
            for row_idx, row in enumerate(data):
                wrapped_row = []
                for cell in row:
                    if row_idx == 0:  # Only first row gets bold style
                        wrapped_row.append(Paragraph(str(cell), header_style))
                    else:
                        wrapped_row.append(Paragraph(str(cell), cell_style))
                wrapped_data.append(wrapped_row)

            # Create table with wrapped data and calculated widths
            table = Table(wrapped_data, repeatRows=1, colWidths=col_max_widths)
            table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (-1, -1), 'DejaVuSans'),
                ('FONTNAME', (0, 0), (-1, 0), 'DejaVuSans-Bold'),  # Bold only first row
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('BOX', (0, 0), (-1, -1), 1, colors.black),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
                # Background color for header row
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1242F1')),
                ('PADDING', (0, 0), (-1, -1), 4),
                ('LEFTPADDING', (0, 0), (-1, -1), 3),
                ('RIGHTPADDING', (0, 0), (-1, -1), 3),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ]))
            story.append(table)
        except Exception as e:
            print(f"Error processing table data: {e}")
            raise

    doc.build(story)

# File processing
def process_excel(filepath):
    try:
        print("\n=== Starting Excel Processing ===")
        # Read the specific sheet
        df = pd.read_excel(
            filepath,
            sheet_name="Evidencija drzanja nastave",
            header=None,
            engine='openpyxl'
        )

        print("\nRaw DataFrame first few rows:")
        print(df.head())
        print("\nColumn 0 data type:", df[0].dtype)
        
        # Clean and process data
        data = []
        for idx, row in enumerate(df.values.tolist()):
            # Skip completely empty rows
            if not any(pd.notna(cell) for cell in row):
                continue
                
            # Process each row
            processed_row = []
            for i, cell in enumerate(row):
                if i == 0:  # First column (date column)
                    print(f"\nProcessing cell in first column: {cell}, Type: {type(cell)}")
                    if pd.notna(cell):
                        if isinstance(cell, (pd.Timestamp, datetime.datetime)):
                            # Format date as YYYY-MM-DD without time
                            formatted_date = cell.strftime('%Y-%m-%d')
                            print(f"Formatted date: {formatted_date}")
                            processed_row.append(formatted_date)
                        else:
                            # If it's a header or other text, keep it as is
                            print(f"Non-date value: {cell}")
                            processed_row.append(str(cell))
                    else:
                        print("Empty cell")
                        processed_row.append("")
                else:
                    # Handle other cells normally
                    processed_row.append(str(cell) if pd.notna(cell) else "")
            
            print(f"\nProcessed row {idx}: {processed_row}")
            data.append(processed_row)
        
        print("\nFinal processed data (first few rows):")
        for row in data[:5]:
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
        
        return data, professor_name

    except Exception as e:
        print(f"Error processing Excel file: {e}")
        traceback.print_exc()  # Print the full error traceback
        return None, None

# Generate PDF
filepath = "Novi_Izveštaj o radu_za_Nastavnike_Natasa_Bogdanovic.xlsx"
table_data, professor_name = process_excel(filepath)

if table_data and professor_name:
    generate_pdf_edn(table_data, "EDNtest12.pdf", professor_name)
    print("PDF generated successfully!")
else:
    print("Failed to generate PDF due to data processing errors.")