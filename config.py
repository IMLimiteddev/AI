
# config.py
import pathlib
import re

# --- File Paths ---
BASE_DIR = pathlib.Path(__file__).parent
# INPUT_PDF_FILENAME is now defined in main.py

# --- Central Configuration Data (Keep for reference if desired) ---
# This section is mostly for reference now, defaults are defined directly below
CONFIG_REFERENCE = {
 
        "default_values": {
        "kundennummer": "2144",  # Default customer number
        "rechnungsadr": "58",  # Default billing address
        "lieferadr": "58",  # Default delivery address
        "aufbestadr": "58",  # Default order address
        "besteller": "PS",  # Default orderer
        "sonderausfuehrung": "Nein",  # Default: No special execution
        "behangart": "Rollladen",  # Default: Roller shutter
        "kurbelstange": "grau",  # Default: Crank rod color
        "anschlagstopfen": "schwarz",  # Default: Stop buffer color
        "standardausfuehrung": "Ja",  # Default: Standard execution
        "fuehrungsschiene": "ALU95",  # Default: Guide rail type
        "reviblendeart": "ALU",  # Default: Revision panel type
        "massbezug": "Fenstermaße",  # Default: Measurement reference
        "fenstergeometrie": "Standard",  # Default: Window geometry
        "fensterart": "PVC",  # Default: Window type
        "fensteroeffnung": "Ja",  # Default: Window opening
        "posbeschreibung": "0",  # Default: Position description
        "posbeschreibung2": "0",  # Default: Position description 2
        "artikelnummerkunde": "0",  # Default: Customer article number
        "fensteraufteilung": "1-tlg",  # Default: Window division
        "behangtyp": "Rollladen Alu DM40",  # Default: Curtain type
        "anzahl_links": "0",  # Default: Number left
        "anzahl_rechts": "0",  # Default: Number right
        "schallschutz": "Nein",  # Default: No soundproofing
        "antrieb": "Becker Motor",  # Default: Drive type
        "fensterbankart": "Alubank",  # Default: Window sill type
        "iss_ausfuehrung": "0",  # Default: ISS execution
        "iss_behindertengerecht": "0",  # Default: ISS disabled-friendly
        "iss_anzahl_links": "0",  # Default: ISS number left
        "iss_anzahl_rechts": "0",  # Default: ISS number right
        "einzelteiltyp": "0",  # Default: Individual part type
        "einzelteilart": "0",  # Default: Individual part type
        "einzelteilanzahl": "0",  # Default: Individual part count
    },

    "color_mapping": {
        "anthrazit": "anthrazit",
        "silberfarbig": "silber",
        "weißaluminium matt": "hwf9006",
        "anthrazit matt": "hwf7016",
        "weiß": "RAL9016",
        "grau": "RAL7016",
        "braun": "RAL8017",
        "silber": "RAL9006",
    },

    "fensterbankart_mapping": {
        "Alubank": "Alubank",
        "Komfortschwelle": "Komfortschwelle",
        "Alu": "Alubank",
        "Kunststoff": "Kunststoff",
        "Kunststoffbank": "Kunststoff",
        "Aluminium": "Alubank",
    },

    "antrieb_mapping": {
        "Becker Motor": "Becker Motor",
        "NHK-Kit": "Becker NHK-Kit",
        "motor": "Becker Motor",
        "gurt": "Gurt",
        "kette": "Becker NHK-Kit",
    },

    "geschoss_mapping": {
        "EG": "EG",
        "OG": "OG",
        "DG": "DG",
        "eg": "EG",
        "og": "OG",
        "dg": "DG",
    },

    "konstruktion_mapping": {
        "k4709": "k4709",
        "k4710": "k4710",
    },

    "reviblende_mapping": {
        "weiß": "weiß",
        "anthrazit": "anthrazit",
    },

    "fuehrungsschiene_mapping": {
        "ALU95": "ALU95",
        "ALU75": "ALU75",
        "alu95": "ALU95",
        "alu75": "ALU75",
    },

    "anschlagstopfen_mapping": {
        "grau": "grau",
        "schwarz": "schwarz",
    },

    "endleiste_mapping": {
        "silber": "silber",
        "weiß": "weiß",
    },

    "behang_mapping": {
        "anthrazit": "anthrazit",
        "silber": "silber",
        "weißaluminium matt": "hwf9006",
        "anthrazit matt": "hwf7016",
    },

    "maßbezug_mapping": {
        "Fenstermaße": "Fenstermaße",
        "Lichtmaße": "Lichtmaße",
    },

    "standardausfuehrung_15_mapping": {
        "Ja": "Ja",
        "Nein": "Nein",
    },

    "fenstergeometrie_mapping": {
        "Standard": "Standard",
        "Sonderform": "Sonderform",
    },

    "fensterart_mapping": {
        "PVC": "PVC",
        "Holz": "Holz",
        "Aluminium": "Aluminium",
    },

    "fensteraufteilung_mapping": {
        "1-tlg": "1-tlg",
        "2-tlg": "2-tlg",
        "3-tlg": "3-tlg",
    },
}

# --- Kopf Sheet Defaults (Define directly) ---
KOPF_DEFAULTS = {
    "Bestelldatum": None, # Placeholder
    "Auftragsname": None, # Placeholder
    "Kunden-Auftrags-Nr": None, # Placeholder
    "Wunsch-Liefertermin": None, # Placeholder
    "Besteller": "PS",
    "Sonderausführung": "Ja", # Changed based on target Excel
    "Hinweistext": "",
    "BehangArt": "Rollladen",
    "Kurbelstange": "grau",
    "Farben_Behang": "silber",
    "Farben_Anschlagstopfen": "schwarz", # From your updated default
    "Farben_Endleiste": "hwf9006",
    "Farben_Aussenkasten": "",
    "Farben_Reviblende": "hwf9006",
    "Farben_Fuehrungsschiene": "hwf9006",
    "Farben_Insekt_Element": "",
    "Farben_Insekt_Endleiste": "",
    "Farben_Insekt_Fuehrungsschiene": "",
}

# --- Positionen Sheet Defaults (Define directly) ---
POS_DEFAULTS = {
 
    "lfdNr_1": None, "PosNr_31": None, "Geschoss": "0",
    "PosBeschreibung_10": "0", "PosBeschreibung2_49": "0", "ArtikelNummerKunde_50": "0",
    "Fenstertyp_9": "0", "Fensteraufteilung": "1-tlg", "Fensterart": "PVC",
    "Fensteröffnung_32": "Ja", "Fenstergeometrie": "Standard", "Konstruktion": "",
    "BehangTyp": "Rollladen Alu DM40", "Schallschutz_48": "Nein", "Antrieb": "Becker Motor",
    "Fuehrungsschiene": "ALU95", "ReviblendeArt": "0", "Standardausführung_15": "Ja",
    "Fensterbankart": "Alubank", "Anzahl_Links_13": "0", "Anzahl_Rechts_14": "0",
    "FeBreite_11": "", "FeHoehe_12": "", "Maßbezug": "Fenstermaße",
    "ISS_Ausführung": "0", "ISS_Behindertengerecht_40": "0", "ISS_Anzahl_Links_41": "0",
    "ISS_Anzahl_Rechts_42": "0", "EinzelteilTyp": "0", "EinzelteilArt": "0",
    "EinzelteilAnzahl_25": "0",
}

# --- Mappings (Use selected mappings from central config or define specifically) ---
GESCHOSS_MAP = CONFIG_REFERENCE["geschoss_mapping"] # Keep using numeric mapping '0', '1', '2'
ANTRIEB_MAP = CONFIG_REFERENCE["antrieb_mapping"]
FENSTERTYP_MAP = { re.compile(r'Fenster\s+KU\s+weiß', re.IGNORECASE): 'PVC', }
COLOR_MAP_BEHANG_KOPF = CONFIG_REFERENCE["color_mapping"]

# --- PDF Parsing Keywords/Regex ---
PDF_MARKERS = {
    "KD_AUFTRAG": r"KD-Auftrag:\s*(\d+)",
    "BESTNR": r"Bestnr\.\s*(\d+)",
    "VOM_DATE": r"vom\s*(\d{2}\.\d{2}\.\d{4})",
    "LIEFERTERMIN": r"Liefertermin:\s*Tag\s*(\d{2}\.\d{2}\.\d{4})",
    "POS_START": r"[_]*Pos\.[_* ]+Material[_* ]+B[_* ]*e[_* ]*z[_* ]*e[_* ]*i", # Keep corrected header regex

    # --- FINAL ATTEMPT: MOST FLEXIBLE POS_ITEM REGEX ---
    # ^ : Start of line
    # (00\d{3}) : Capture Pos (Group 1)
    # \s* : ZERO or more whitespace
    # (\S+) : Capture Material (one or more non-space chars) (Group 2) - Assumes Material has NO spaces
    # \s* : ZERO or more whitespace
    # (.*) : Capture Description (Group 3)
    "POS_ITEM": r"^(00\d{3})\s*(\S+)\s*(.*)",
    # --- END REGEX CHANGE ---

    # --- Detail Markers ---
    "Fensternummer": r"Fensternummer:\s*(\d+)",
    "Breite": r"Breite\s+in\s+mm:\s*(\d+)",
    "LaengeFS": r"Länge\s+Führungsschiene:\s*(\d+)",
    "WinkelFS": r"Abschnittwinkel\s+Führungsschien:\s*(\d+)",
    "Geschoss": r"Geschoss:\s*(\S.*)",
    #"AntriebSeite": r"Antriebsseite:\s*(Links|Rechts|Beidseitig)",
    "Antriebsseite": r"Antriebsseite:\s*(Links|Rechts|Beidseitig)", 
    "Farbe": r"Farbe:\s*(\S.*)",
    "BehangArt": r"Behangart:\s*(\S.*)",
    "Kurbelstange": r"Kurbelstange:\s*(\S.*)",
    "Notkurbel": r"Rollladen\s+Notkurbel:\s*(links|rechts)",
    "Antrieb": r"Rollladen\s+Antrieb:\s*(\S.*)",
    "Panzer": r"Rollladen\s+Panzer:\s*(\S.*)",
    "FehroFS": r"Fehro\s+Führungsschiene:\s*(\S.*)",
    "Endschiene": r"Rollladen/Jalousie\s+Endschiene:\s*(\S.*)",
    "Revision": r"Revisionsblende:\s*(\S.*)",
    "Zeichnung": r"Zeichnungsnummer:\s*(\S.*)",
    "Fensterbank": r"Fensterbank\s+außen:\s*(\S.*)",
    "FensterDesc": r"Fenster\s+KU\s+weiß\s+(\S.*)",
}

# --- Excel Cell Mappings for Kopf Sheet DATA ---
# CORRECTED NAME: KOPF_MAP_DATA_CELLS
KOPF_MAP_DATA_CELLS = {
    "Bestelldatum": "C10", "Auftragsname": "D10", "Kunden-Auftrags-Nr": "E10",
    "Wunsch-Liefertermin": "F10", "Besteller": "G10",  # Add Besteller mapping
    "Sonderausführung": "C12", "Hinweistext": "F12",  # Check F12 is correct cell
    "BehangArt": "C14", "Kurbelstange": "D14", "Farben_Behang": "C16",
    "Farben_Anschlagstopfen": "D16", "Farben_Endleiste": "E16", "Farben_Aussenkasten": "F16",
    "Farben_Reviblende": "G16", "Farben_Fuehrungsschiene": "H16",  # Check H16 is correct
    "Farben_Insekt_Element": "C18", "Farben_Insekt_Endleiste": "D18",
    "Farben_Insekt_Fuehrungsschiene": "E18",
    # NOTE: Static hints like Hinweis_Farbauswahl_1/2/3 are handled in excel_writer now
}

# --- Excel Column Definitions for Positionen Sheet ---
# CORRECTED NAME: POS_COLS_DEFS
# Values MUST be (Header Text, Column Index) tuples
# Header Text should match the desired header in the Excel output
# Column Index MUST be correct for the Excel layout
POS_COLS_DEFS = {
    "lfdNr_1": ("lfdNr_1", 1),
    "PosNr_31": ("PosNr_31", 2),
    "Geschoss": ("Geschoss", 3),
    "PosBeschreibung_10": ("PosBeschreibung_10", 4),
    "PosBeschreibung2_49": ("PosBeschreibung2_49", 5),
    "ArtikelNummerKunde_50": ("ArtikelNummerKunde_50", 6),
    "Fenstertyp_9": ("Fenstertyp_9", 7),
    "Fensteraufteilung": ("Fensteraufteilung", 8),
    "Fensterart": ("Fensterart", 9),
    # Col 10 is empty
    "Fensteröffnung_32": ("Fensteröffnung_32", 10),
    # Col 12 is empty
    "Fenstergeometrie": ("Fenstergeometrie", 11),
    "Konstruktion": ("Konstruktion", 12),
    # Cols 15-20 empty or different headers
    "BehangTyp": ("BehangTyp", 13),  # Note: Excel sheet header is BehangTyp
    # Col 22 empty
    "Schallschutz_48": ("Schallschutz_48", 14),
    "Antrieb": ("Antrieb", 15),
    # Col 25 empty
    "Fuehrungsschiene": ("Fuehrungsschiene", 16),
    # Col 27 empty
    "ReviblendeArt": ("ReviblendeArt", 17),

    # Col 29 empty
    "Standardausführung_15": ("Standardausführung_15", 18),
    "Fensterbankart": ("Fensterbankart", 19),  # Note: Excel sheet header is Fensterbankart
    # Col 32 empty
    "Anzahl_Links_13": ("Anzahl_Links_13", 20),
    "Anzahl_Rechts_14": ("Anzahl_Rechts_14", 21),
    "FeBreite_11": ("FeBreite_11", 22),
    "FeHoehe_12": ("FeHoehe_12", 23),
    "Maßbezug": ("Maßbezug", 24),
    # Col 38 empty
    "ISS_Ausführung": ("ISS_Ausführung", 25),
    # Col 40, 41 empty
    "ISS_Behindertengerecht_40": ("ISS_Behindertengerecht_40", 26),
    "ISS_Anzahl_Links_41": ("ISS_Anzahl_Links_41", 27),
    "ISS_Anzahl_Rechts_42": ("ISS_Anzahl_Rechts_42", 28),
    "EinzelteilTyp": ("EinzelteilTyp", 29),
    # Col 46 empty
    "EinzelteilArt": ("EinzelteilArt", 30),
    # Col 48 empty
    "EinzelteilAnzahl_25": ("EinzelteilAnzahl_25", 31),
}

# --- PDF Writer Configuration ---
# (Keep PDF_TITLE, PDF_FONT etc. the same)
PDF_TITLE = "Bestellblatt Positionen"
PDF_FONT = "Helvetica"
PDF_FONT_SIZE_HEADER = 8  # For rotated headers in table
PDF_FONT_SIZE_DATA = 7  # For table data
# Use POS_COLS_DEFS to dynamically generate PDF column config if possible,
# or define PDF_COL_CONFIG manually ensuring keys match POS_COLS_DEFS keys
# Manual definition for potentially different display names/widths:
PDF_COL_CONFIG = {
    "lfdNr_1": ("LfNr", 8),
    "PosNr_31": ("PNr", 8),
    "Geschoss": ("G", 6),
    # Skipping PosBeschreibung etc. for PDF unless explicitly needed
    "Fenstertyp_9": ("Typ", 10),
    "Fensteraufteilung": ("Aufteilung", 15),
    "Fensterart": ("Art", 15),
    "Fensteröffnung_32": ("Öff", 6),
    "Fenstergeometrie": ("Geom", 12),
    "Konstruktion": ("Konstr", 12),
    "BehangTyp": ("Behang", 25),
    "Antrieb": ("Antrieb", 20),
    "Fuehrungsschiene": ("FS", 10),
    "Standardausführung_15": ("Std", 6),
    "Fensterbankart": ("FB", 15),
    "Anzahl_Links_13": ("L", 5),
    "Anzahl_Rechts_14": ("R", 5),
    "FeBreite_11": ("Breite", 10),
    "FeHoehe_12": ("Höhe", 10),
    "Maßbezug": ("Maß", 15),
    "Schallschutz_48": ("Schall", 8),  # Added Schallschutz
    # Add other POS_COLS_DEFS keys here if they should appear in the PDF
}


