from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import os
import cyrtranslit

font_path_regular = os.path.join(os.path.dirname(__file__), 'fonts', 'micross-regular.ttf')
pdfmetrics.registerFont(TTFont('Microsoft Sans Serif', font_path_regular))

font_path_bold = os.path.join(os.path.dirname(__file__), 'fonts', 'arial-bold.ttf')
pdfmetrics.registerFont(TTFont('Microsoft Sans Serif Bold', font_path_bold))

font_path_italic = os.path.join(os.path.dirname(__file__), 'fonts', 'arial-corsivo-2.ttf')
pdfmetrics.registerFont(TTFont('Microsoft Sans Serif Italic', font_path_italic))

def to_cyrilic(text):
    return cyrtranslit.to_cyrillic(text, "sr")

def generate_pdf(dataframe, filename):

    dataframe.rename(columns={"Ukupno casova": "Norma časova"}, inplace=True)
    dataframe.rename(columns={"Naziv Predmeta": "Naziv predmeta"}, inplace=True)
    dataframe.rename(columns={"Nedeljni Broj Časova": "Nedeljni broj časova"}, inplace=True)
    dataframe.rename(columns={"Broj Grupa": "Broj grupa"}, inplace=True)
    dataframe.rename(columns={"Tip Predavanja": "Tip predavanja"}, inplace=True)

    doc = SimpleDocTemplate(filename, pagesize=A4)
    elements = []
    spacer = Spacer(0, 0)
    spacer1 = Spacer(0, 10)
    styles = getSampleStyleSheet()

    # Definisanje fonta
    title_style = styles["BodyText"]
    title_style.fontName = "Microsoft Sans Serif"

    
    bold_style = styles["BodyText"]
    bold_style.fontName = "Microsoft Sans Serif"

    bold_style_header = styles["BodyText"]
    bold_style_header.fontName = "Microsoft Sans Serif"
    bold_style_header.alignment = TA_CENTER

    styles = {
        "HeaderStyle": ParagraphStyle(
            name="HeaderStyle",
            fontName="Microsoft Sans Serif",
            fontSize=10,
            alignment=TA_CENTER, 
            textColor="black",
        ),
        "BodyText": ParagraphStyle(
            name="BodyText",
            fontName="Microsoft Sans Serif",
            fontSize=10,
            textColor="black",
        ),
        "BodyTextBold": ParagraphStyle(
            name="BodyTextBold",
            fontName="Microsoft Sans Serif Bold",
            fontSize=10,
            textColor="black",
        ),
        "BodyTextItalic": ParagraphStyle(
            name="BodyTextItalic",
            fontName="Microsoft Sans Serif Italic",
            fontSize=10,
            textColor="black",
        ),
        "TitleStyle": ParagraphStyle(
            name="SpecialCell",
            fontName="Microsoft Sans Serif",
            fontSize=18,
            alignment=TA_LEFT,
            textColor="black",
        ),
    }

    # Izdvajanje Imena profesora i Pozicije
    professor_name = dataframe.iloc[0]["Ime Predavača"] 
    position = dataframe.iloc[0]["Pozicija"]

    # Pozicija
    if position == "profesor":
        position_text = "Професор"
    elif position == "asistent":
        position_text = "Асистент"
    else:
        position_text = "Pozicija Nepoznata"

    title_text = f"{position_text} {professor_name}"

    columns_to_exclude = ["Pozicija", "Ime Predavača"]

    filtered_dataframe = dataframe.drop(columns = columns_to_exclude)

    title_table = Table(
    [[Paragraph(to_cyrilic((title_text)), styles["TitleStyle"])]],
    colWidths=[580]  # Set the width to match the table's starting point
    )
    title_table.setStyle(
        TableStyle([
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),  # Align left
            ("FONTNAME", (0, 0), (-1, -1), "Microsoft Sans Serif"),
            ("FONTSIZE", (0, 0), (-1, -1), 14),  # Adjust font size as needed
            ("BOTTOMPADDING", (0, 0), (-1, -1), 10),  # Add space below the text
        ])
    )
    elements.append(title_table)
    elements.append(Spacer(0, 10))

    # Odvajamo df u osnovne i master
    osnovne_data = dataframe[dataframe['Tip studija'] == 'osnovne']
    master_data = dataframe[dataframe['Tip studija'] == 'master']

    # Izdvajanje imena kolona i wrapped text feature
    columns = filtered_dataframe.columns.tolist()
    columns_wrapped = [Paragraph(to_cyrilic(col), bold_style_header) for col in columns]
    columns_wrapped = [columns_wrapped]

    #Konvertovanje redova u Paragraphs za word wrapping
    # data_wrapped_osnovne = [
    #     [Paragraph(to_cyrilic(str(cell)), styles["BodyText"]) for cell in row]
    #     for row in osnovne_data.drop(columns = columns_to_exclude).values
    # ]

    data_wrapped_osnovne = [
    [
        Paragraph(to_cyrilic(str(cell)), styles["BodyTextBold"] if idx == 0 else styles["BodyText"])
        for idx, cell in enumerate(row)
    ]
    for row in osnovne_data.drop(columns=columns_to_exclude).values
    ]

    # data_wrapped_master = [
    #     [Paragraph(to_cyrilic(str(cell)), styles["BodyText"]) for cell in row]
    #     for row in master_data.drop(columns = columns_to_exclude).values
    # ]

    data_wrapped_master = [
    [
        Paragraph(to_cyrilic(str(cell)), styles["BodyTextBold"] if idx == 0 else styles["BodyText"])
        for idx, cell in enumerate(row)
    ]
    for row in master_data.drop(columns=columns_to_exclude).values
    ]

    # Kombinovanje header-a i podataka
    table_data_osnovne = data_wrapped_osnovne

    table_data_master = data_wrapped_master

    column_table = Table(columns_wrapped, colWidths=[180, 90, 80, 50, 80, 50, 50])
    column_table.setStyle(
        TableStyle(
            [
                    ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                    ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.white),  # Header background
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                    ("GRID", (0, 0), (-1, -1), 1, colors.black),
            ]
        )
    )

    elements.append(column_table)
    elements.append(spacer)

    # Kreiranje glavne tabele
    if not osnovne_data.empty:
        table = Table(table_data_osnovne, colWidths=[180, 90, 80, 50, 80, 50, 50])
        table.setStyle(
            TableStyle(
                [
                    # Header Row: Center horizontally and vertically
                    ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                    ("VALIGN", (0, 0), (-1, 0), "BOTTOM"),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.white),  # Header background
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),  # Header text color

                    # Data Rows: Left-align horizontally, center vertically
                    ("ALIGN", (0, 1), (-1, -1), "LEFT"),
                    ("VALIGN", (0, 1), (-1, -1), "BOTTOM"),
                    ("BACKGROUND", (0, 1), (-1, -1), colors.white),
                    ("TEXTCOLOR", (0, 1), (-1, -1), colors.black),

                    # Grid lines for the whole table
                    ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ]
            )
        )
        elements.append(table)
        elements.append(spacer)

        # Sumiranje ukupnog "Opterećenje" za osnovne
        total_load_osnovne = osnovne_data["Norma časova"].sum()

        # Tabela sum Opterećenja za osnovne
        summary_table_osnovne = Table(
            [
                ["Оптерећење на основним студијама", f"{total_load_osnovne:.2f}"]
            ],
            colWidths=[530, 50]
        )
        summary_table_osnovne.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.lightblue),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                    ("ALIGN", (0, 0), (-1, -1), "RIGHT"),
                    ("FONTNAME", (0, 0), (0, 0), "Microsoft Sans Serif Italic"),
                    ("FONTNAME", (0, 1), (-1, -1), "Microsoft Sans Serif"),
                    ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ]
            )
        )
        elements.append(summary_table_osnovne)
        elements.append(spacer)
    if not master_data.empty:
        elements.append(column_table)

    #Proverava da li profesor ima podatke sa master studija, ako nema ne generiše se tabela sa master predmetima
    if not master_data.empty:
        # Kreiranje glavne mater tabele
        table = Table(table_data_master, colWidths=[180, 90, 80, 50, 80, 50, 50])
        table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.white),  # Header background color
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),  # Header text color
                    ("ALIGN", (0, 0), (-1, 0), "CENTER"),  # Center align column headers horizontally
                    ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),  # Center align column headers vertically
                    ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ]
            )
        )
        elements.append(table)
        elements.append(spacer)

        # Računanje "Opterećenje" za master
        total_load_master = master_data["Norma časova"].sum()

        # Sumuiranje ukupnog Opterećenja za master
        summary_table_master = Table(
            [
                ["Оптерећење на мастер студијама", f"{total_load_master:.2f}"]
            ],
            colWidths=[530, 50]
        )
        summary_table_master.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.lightcoral),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                    ("ALIGN", (0, 0), (-1, -1), "RIGHT"),
                    ("FONTNAME", (0, 0), (0, 0), "Microsoft Sans Serif Italic"),
                    ("FONTNAME", (0, 1), (-1, -1), "Microsoft Sans Serif"),
                    ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ]
            )
        )
        elements.append(summary_table_master)
        elements.append(spacer)

    # Sumiranje ukupnog Opterećenja
    if "Norma časova" in dataframe.columns:
        total_load = dataframe["Norma časova"].sum()
        total_table = Table([["Укупно оптерећење", total_load]], colWidths=[530, 50])
        total_table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.lightgreen),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                    ("ALIGN", (0, 0), (-1, -1), "RIGHT"),
                    ("FONTNAME", (0, 0), (0, 0), "Microsoft Sans Serif Bold"),
                    ("FONTNAME", (0, 0), (-1, 0), "Microsoft Sans Serif Bold"),
                    ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ]
            )
        )
        elements.append(total_table)

    # Bulid-ovanje PDF-a
    doc.build(elements)
