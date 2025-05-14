import openpyxl
from openpyxl.styles import Alignment

def create_excel_template_rotated_headers(filename="template_rotated.xlsx"):
    """Creates an Excel template with rotated headers in Sheet 2."""

    workbook = openpyxl.Workbook()

    # Sheet 1: Bestellblatt_Kopf
    sheet1 = workbook.active
    sheet1.title = "Bestellblatt_Kopf"

    sheet1["A1"] = "D&M KG"
    sheet1["A2"] = "Auf den Dorfwiesen 1-5"
    sheet1["A3"] = "56204 Hillscheid"
    sheet1["D2"] = "Bestellblatt Fehro.AR"
    sheet1["E2"] = "AV-Import"
    sheet1["G1"] = "[Date Placeholder, e.g., 26.02.2025 - optional]"
    sheet1["A6"] = "Eingabe Kundennummern"
    sheet1["C6"] = "Kundenadr."
    sheet1["E6"] = "Rechnungsadr."
    sheet1["G6"] = "Lieferadr."
    sheet1["I6"] = "AufBestAdr."
    sheet1["A9"] = "Bestellinformationen"
    sheet1["A12"] = "Sonderausführung"
    sheet1["E12"] = "Hinweistext"
    sheet1["A14"] = "Angaben Verschattungselemente"
    sheet1["A16"] = "Farben Verschattungselemente"
    sheet1["A18"] = "Farben Insektenschutzelemente"
    sheet1["A20"] = "Hinweise zur getroffenen Farbauswahl:"
    sheet1["C20"] = "Farbauswahl Behang nur bei ALU-Rollladen möglich."
    sheet1["C21"] = "Farbauswahl Reviblende nur bei ALU möglich."
    sheet1["C22"] = "Farbauswahl Führungsschiene nur bei ALU möglich."
    sheet1["A24"] = "Bestellblatt_Kopf [Optional Footer Text]"

    # Sheet 2: Bestellblatt_Positionen
    sheet2 = workbook.create_sheet("Bestellblatt_Positionen")

    headers = [
        "IfdNr_1", "PosNr_31", "Geschoss", "PosBeschreibung_10", "PosBeschreibung2_49",
        "ArtikelNummerKunde_50", "Fenstertyp_9", "Fensteraufteilung", "Fensterart", None,
        "Fensteröffnung_32", None, "Fenstergeometrie", "Konstruktion", None, "BehangTyp",
        None, None, None, None, "Antrieb", None, "Fuehrungsschiene", "ReviblendeArt",
        None, "Standardausführung_15", None, "Fensterbankart", None, "Anzahl_Links_13",
        "Anzahl_Rechts_14", None, "FeBreite_11", "FeHoehe_12", "Maßbezug", "ISS_Ausführung",
        "ISS_Behindertengerecht_40", None, "ISS_Anzahl_Links_41", None, None,
        "ISS_Anzahl_Rechts_42", "EinzelteilTyp", "EinzelteilArt", "EinzelteilAnzahl_25",
        None, None, None, None, None, None, "Schallschutz_48"
    ]

    sheet2.append(headers)

    # Rotate headers in Sheet 2
    for cell in sheet2[1]:
        if cell.value is not None:
            cell.alignment = Alignment(textRotation=90)

    workbook.save(filename)
    print(f"Excel template '{filename}' created successfully.")

create_excel_template_rotated_headers()