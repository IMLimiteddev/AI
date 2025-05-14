
import logging
from fpdf import FPDF
from datetime import datetime
from typing import Dict, List
import math
# Make sure these are imported and available
from config import KOPF_DEFAULTS, POS_COLS_DEFS, KOPF_MAP_DATA_CELLS # KOPF_MAP_DATA_CELLS might not be directly needed here, but KOPF_DEFAULTS is

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s') # Added line number

# --- PDF Configuration ---
PDF_FONT = "Helvetica"
PDF_MARGIN = 10
PDF_LINE_HEIGHT = 5
PDF_TABLE_HEADER_HEIGHT = 65  # Increased to allow wrapped rotated text
PDF_TABLE_ROW_HEIGHT = 6 # This will be used as the *line height* within multi_cell
PAGE_WIDTH_L = 297
PAGE_HEIGHT_L = 210
PRINTABLE_WIDTH_L = PAGE_WIDTH_L - 2 * PDF_MARGIN

# --- UNCOMMENTED: Map Kopf data keys to their labels for the PDF ---
KOPF_PDF_LABELS = {
    "Bestelldatum": "Bestelldatum:",
    "Auftragsname": "Auftragsname:",
    "Kunden-Auftrags-Nr": "Kunden-Auftrags-Nr:",
    "Wunsch-Liefertermin": "Wunsch-Liefertermin:",
    "Besteller": "Besteller:",  # Add Besteller label
    "Sonderausführung": "Sonderausführung:",
    "Hinweistext": "Hinweistext:",
    "BehangArt": "BehangArt:",
    "Kurbelstange": "Kurbelstange:",
    "Farben_Behang": "Behang:",  # Use shorter labels like Excel sheet
    "Farben_Anschlagstopfen": "Anschlagstopfen:",
    "Farben_Endleiste": "Endleiste:",
    "Farben_Aussenkasten": "Aussenkasten:",
    "Farben_Reviblende": "Reviblende:",
    "Farben_Fuehrungsschiene": "Führungsschiene:",
    "Farben_Insekt_Element": "Element:",
    "Farben_Insekt_Endleiste": "Endleiste:",
    "Farben_Insekt_Fuehrungsschiene": "Führungsschiene:",
    "Bestelldatum": "Bestelldatum:",
    "Auftragsname": "Auftragsname:",
    "KundenAuftragsNr": "Kunden-Auftrags-Nr:",
    "WunschLiefertermin": "Wunsch-Liefertermin:",
    
}

# --- UNCOMMENTED: Define approximate grid/column starts for Kopf page ---
# These help align items visually, mimicking Excel columns
COL1_X = PDF_MARGIN
COL2_X = PDF_MARGIN + 50
COL3_X = PDF_MARGIN + 85
COL4_X = PDF_MARGIN + 135
COL5_X = PDF_MARGIN + 180
COL6_X = PDF_MARGIN + 210
# --- END UNCOMMENTED ---


class PDFWithHeaderFooter(FPDF):
    def header(self):
        self.set_font(PDF_FONT, 'I', 8)
        page_w = self.w - 2 * self.l_margin
        self.cell(page_w / 2, 5, "Bestellblatt Fehro.AR AV-Import", 0, 0, 'L')
        self.cell(page_w / 2, 5, datetime.now().strftime('%d.%m.%Y'), 0, 1, 'R')
        self.set_y(self.t_margin + 5)

    def footer(self):
        self.set_y(-15)
        self.set_font(PDF_FONT, 'I', 8)
        self.cell(0, 10, f'Seite {self.page_no()}/{{nb}}', 0, 0, 'C')

    def word_wrap(self, text, width):
        # Ensure text is a string before processing
        text = str(text) if text is not None else ""
        words = text.split()
        wrapped_text = ""
        current_line = ""
        for word in words:
            # Check width before adding the word and a potential space
            test_line = current_line + word + " "
            if self.get_string_width(test_line.strip()) < width:
                current_line += word + " "
            else:
                # If adding the word exceeds width, finalize the current line
                # But only if current_line is not empty (handles very long words)
                if current_line:
                    wrapped_text += current_line.strip() + "\n"
                # Start the new line with the current word
                current_line = word + " "
        # Add the last line
        wrapped_text += current_line.strip()
        # Return space-separated for rotated header context if needed,
        # but for multi_cell, newline is usually better. Let's keep newline.
        # return wrapped_text.replace("\n", " ") # Keep as newline for multi_cell
        return wrapped_text


    # --- UNCOMMENTED AND MOVED INSIDE CLASS: Page 1: Kopf Data (Revised Layout) ---
    def draw_kopf_page(self, kopf_data: Dict):
        """Draws the first page resembling Bestellblatt_Kopf in Landscape"""
        self.add_page(orientation='L')
        self.set_font(PDF_FONT, '', 10)

        y_pos = self.get_y() + 5  # Starting Y

        # --- Row 1: Company Info ---
        self.set_font(PDF_FONT, 'B', 10)
        self.set_xy(COL1_X, y_pos);
        self.cell(40, PDF_LINE_HEIGHT, "D&M KG")
        self.set_font(PDF_FONT, '', 9)
        self.set_xy(COL4_X, y_pos);
        self.cell(40, PDF_LINE_HEIGHT, "Bestellblatt Fehro.AR")
        y_pos += PDF_LINE_HEIGHT
        self.set_xy(COL1_X, y_pos);
        self.cell(40, PDF_LINE_HEIGHT, "Auf den Dorfwiesen 1-5")
        self.set_xy(COL4_X + 20, y_pos);
        self.cell(40, PDF_LINE_HEIGHT, "AV-Import")
        y_pos += PDF_LINE_HEIGHT
        self.set_xy(COL1_X, y_pos);
        self.cell(40, PDF_LINE_HEIGHT, "56204 Hillscheid")
        y_pos += PDF_LINE_HEIGHT * 2  # Add space

        # --- Row 2: Address Headers (Simplified - Labels only) ---
        self.set_font(PDF_FONT, '', 8)
        self.set_xy(COL1_X, y_pos);
        self.cell(45, PDF_LINE_HEIGHT, "Eingabe Kundennummern")
        self.set_xy(COL2_X, y_pos);
        self.cell(30, PDF_LINE_HEIGHT, "Kundenadr.")
        self.set_xy(COL3_X, y_pos);
        self.cell(35, PDF_LINE_HEIGHT, "Rechnungsadr.")
        self.set_xy(COL4_X, y_pos);
        self.cell(30, PDF_LINE_HEIGHT, "Lieferadr.")
        self.set_xy(COL5_X, y_pos);
        self.cell(30, PDF_LINE_HEIGHT, "AufBestAdr.")
        y_pos += PDF_LINE_HEIGHT * 1.5  # Space + line height
        
        self.set_xy(COL2_X, y_pos);
        self.cell(30, PDF_LINE_HEIGHT, str(kopf_data.get("Kundenadr.", "2144")), border='B')
        self.set_xy(COL3_X, y_pos);
        self.cell(35, PDF_LINE_HEIGHT, str(kopf_data.get("Rechnungsadr.", "58")), border='B')
        self.set_xy(COL4_X, y_pos);
        self.cell(30, PDF_LINE_HEIGHT, str(kopf_data.get("Lieferadr.", "58")), border='B')
        self.set_xy(COL5_X, y_pos);
        self.cell(30, PDF_LINE_HEIGHT, str(kopf_data.get("AufBestAdr.", "58")), border='B')
        y_pos += PDF_LINE_HEIGHT * 2  # Space + line height

        # --- Row 3: Bestellinformationen ---
        self.set_font(PDF_FONT, 'B', 10)
        self.set_xy(COL1_X, y_pos);
        self.cell(40, PDF_LINE_HEIGHT, "Bestellinformationen")
        y_pos += PDF_LINE_HEIGHT
        self.set_font(PDF_FONT, '', 9)
        # Labels first
        self.set_xy(COL2_X, y_pos);
        self.cell(30, PDF_LINE_HEIGHT, KOPF_PDF_LABELS["Bestelldatum"])
        self.set_xy(COL3_X, y_pos);
        self.cell(35, PDF_LINE_HEIGHT, KOPF_PDF_LABELS["Auftragsname"])
        self.set_xy(COL4_X, y_pos);
        self.cell(40, PDF_LINE_HEIGHT, KOPF_PDF_LABELS["Kunden-Auftrags-Nr"])
        self.set_xy(COL5_X, y_pos);
        self.cell(40, PDF_LINE_HEIGHT, KOPF_PDF_LABELS["Wunsch-Liefertermin"])
        self.set_xy(COL6_X, y_pos);
        self.cell(30, PDF_LINE_HEIGHT, KOPF_PDF_LABELS.get("Besteller", "Besteller:"))  # Use .get in case key missing
        y_pos += PDF_LINE_HEIGHT
        # Data below labels
        self.set_xy(COL2_X, y_pos);
        self.cell(30, PDF_LINE_HEIGHT, str(kopf_data.get("Bestelldatum", "")), border='B')
        self.set_xy(COL3_X, y_pos);
        self.cell(35, PDF_LINE_HEIGHT, str(kopf_data.get("Auftragsname", "")), border='B')
        self.set_xy(COL4_X, y_pos);
        self.cell(40, PDF_LINE_HEIGHT, str(kopf_data.get("Kunden-Auftrags-Nr", "")), border='B')
        self.set_xy(COL5_X, y_pos);
        self.cell(40, PDF_LINE_HEIGHT, str(kopf_data.get("Wunsch-Liefertermin", "")), border='B')
        self.set_xy(COL6_X, y_pos);
        self.cell(30, PDF_LINE_HEIGHT, str(kopf_data.get("Besteller", "")), border='B')
        y_pos += PDF_LINE_HEIGHT * 1.5

        # --- Row 4: Sonderausführung ---
        self.set_font(PDF_FONT, 'B', 10)
        self.set_xy(COL1_X, y_pos);
        self.cell(40, PDF_LINE_HEIGHT, "Sonderausführung")
        y_pos += PDF_LINE_HEIGHT
        self.set_font(PDF_FONT, '', 9)
        self.set_xy(COL2_X, y_pos);
        self.cell(30, PDF_LINE_HEIGHT, "J/N") # Label for Sonderausführung value
        self.set_xy(COL4_X, y_pos);
        self.cell(40, PDF_LINE_HEIGHT, KOPF_PDF_LABELS["Hinweistext"])
        y_pos += PDF_LINE_HEIGHT
        self.set_xy(COL2_X, y_pos);
        self.cell(30, PDF_LINE_HEIGHT, str(kopf_data.get("Sonderausführung", "")), border='B')
        # Hinweistext can be long, use multi_cell
        hinweis_start_y = y_pos
        self.set_xy(COL4_X, y_pos);
        # Ensure Hinweistext is treated as string
        hinweis_text = str(kopf_data.get("Hinweistext", ""))
        self.multi_cell(PRINTABLE_WIDTH_L - COL4_X, PDF_LINE_HEIGHT, hinweis_text, border='B', align='L')
        # Update y_pos based on potential multi_cell height
        # self.get_y() returns the position *after* the multi_cell
        y_pos = max(y_pos + PDF_LINE_HEIGHT, self.get_y())
        self.set_y(y_pos) # Ensure cursor is below the longest cell
        y_pos += PDF_LINE_HEIGHT * 0.5 # Add small gap

        # --- Row 5: Angaben Verschattungselemente ---
        self.set_font(PDF_FONT, 'B', 10)
        self.set_xy(COL1_X, y_pos);
        self.cell(40, PDF_LINE_HEIGHT, "Angaben Verschattungselemente")
        y_pos += PDF_LINE_HEIGHT
        self.set_font(PDF_FONT, '', 9)
        self.set_xy(COL2_X, y_pos);
        self.cell(30, PDF_LINE_HEIGHT, KOPF_PDF_LABELS["BehangArt"])
        self.set_xy(COL3_X, y_pos);
        self.cell(35, PDF_LINE_HEIGHT, KOPF_PDF_LABELS["Kurbelstange"])
        y_pos += PDF_LINE_HEIGHT
        self.set_xy(COL2_X, y_pos);
        self.cell(30, PDF_LINE_HEIGHT, str(kopf_data.get("BehangArt", "")), border='B')
        self.set_xy(COL3_X, y_pos);
        self.cell(35, PDF_LINE_HEIGHT, str(kopf_data.get("Kurbelstange", "")), border='B')
        y_pos += PDF_LINE_HEIGHT * 1.5

        # --- Row 6: Farben Verschattungselemente ---
        self.set_font(PDF_FONT, 'B', 10)
        self.set_xy(COL1_X, y_pos);
        self.cell(40, PDF_LINE_HEIGHT, "Farben Verschattungselemente")
        y_pos += PDF_LINE_HEIGHT
        self.set_font(PDF_FONT, '', 9)
        # Labels
        self.set_xy(COL2_X, y_pos);
        self.cell(30, PDF_LINE_HEIGHT, KOPF_PDF_LABELS["Farben_Behang"])
        self.set_xy(COL3_X, y_pos);
        self.cell(35, PDF_LINE_HEIGHT, KOPF_PDF_LABELS["Farben_Anschlagstopfen"])
        self.set_xy(COL4_X, y_pos);
        self.cell(40, PDF_LINE_HEIGHT, KOPF_PDF_LABELS["Farben_Endleiste"])
        self.set_xy(COL5_X, y_pos);
        self.cell(30, PDF_LINE_HEIGHT, KOPF_PDF_LABELS["Farben_Aussenkasten"])
        self.set_xy(COL6_X, y_pos);
        self.cell(30, PDF_LINE_HEIGHT, KOPF_PDF_LABELS["Farben_Reviblende"])
        # Need another column for Führungsschiene
        COL7_X = COL6_X + 35
        self.set_xy(COL7_X, y_pos);
        self.cell(0, PDF_LINE_HEIGHT, KOPF_PDF_LABELS["Farben_Fuehrungsschiene"]) # Use 0 width to extend to margin
        y_pos += PDF_LINE_HEIGHT
        # Data
        self.set_xy(COL2_X, y_pos);
        self.cell(30, PDF_LINE_HEIGHT, str(kopf_data.get("Farben_Behang", "")), border='B')
        self.set_xy(COL3_X, y_pos);
        self.cell(35, PDF_LINE_HEIGHT, str(kopf_data.get("Farben_Anschlagstopfen", "")), border='B')
        self.set_xy(COL4_X, y_pos);
        self.cell(40, PDF_LINE_HEIGHT, str(kopf_data.get("Farben_Endleiste", "")), border='B')
        self.set_xy(COL5_X, y_pos);
        self.cell(30, PDF_LINE_HEIGHT, str(kopf_data.get("Farben_Aussenkasten", "")), border='B')
        self.set_xy(COL6_X, y_pos);
        self.cell(30, PDF_LINE_HEIGHT, str(kopf_data.get("Farben_Reviblende", "")), border='B')
        self.set_xy(COL7_X, y_pos);
        self.cell(0, PDF_LINE_HEIGHT, str(kopf_data.get("Farben_Fuehrungsschiene", "")), border='B') # Use 0 width
        y_pos += PDF_LINE_HEIGHT * 1.5

        # --- Row 7: Farben Insektenschutz ---
        self.set_font(PDF_FONT, 'B', 10)
        self.set_xy(COL1_X, y_pos);
        self.cell(40, PDF_LINE_HEIGHT, "Farben Insektenschutzelemente")
        y_pos += PDF_LINE_HEIGHT
        self.set_font(PDF_FONT, '', 9)
        # Labels
        self.set_xy(COL2_X, y_pos);
        self.cell(30, PDF_LINE_HEIGHT, KOPF_PDF_LABELS["Farben_Insekt_Element"])
        self.set_xy(COL3_X, y_pos);
        self.cell(35, PDF_LINE_HEIGHT, KOPF_PDF_LABELS["Farben_Insekt_Endleiste"])
        self.set_xy(COL4_X, y_pos);
        self.cell(40, PDF_LINE_HEIGHT, KOPF_PDF_LABELS["Farben_Insekt_Fuehrungsschiene"])
        y_pos += PDF_LINE_HEIGHT
        # Data
        self.set_xy(COL2_X, y_pos);
        self.cell(30, PDF_LINE_HEIGHT, str(kopf_data.get("Farben_Insekt_Element", "")), border='B')
        self.set_xy(COL3_X, y_pos);
        self.cell(35, PDF_LINE_HEIGHT, str(kopf_data.get("Farben_Insekt_Endleiste", "")), border='B')
        self.set_xy(COL4_X, y_pos);
        self.cell(40, PDF_LINE_HEIGHT, str(kopf_data.get("Farben_Insekt_Fuehrungsschiene", "")), border='B')
        y_pos += PDF_LINE_HEIGHT * 2

        # --- Hinweise Section ---
        self.set_font(PDF_FONT, 'B', 9)  # Bold label
        self.set_xy(COL1_X, y_pos);
        self.cell(0, PDF_LINE_HEIGHT, "Hinweise zur getroffenen Farbauswahl:", ln=1)
        y_pos += PDF_LINE_HEIGHT
        self.set_font(PDF_FONT, '', 9)  # Normal text
        self.set_x(COL2_X)  # Indent the hints
        self.cell(0, PDF_LINE_HEIGHT, "- Farbauswahl Behang nur bei ALU-Rollladen möglich.", ln=1)
        self.set_x(COL2_X)
        self.cell(0, PDF_LINE_HEIGHT, "- Farbauswahl Reviblende nur bei ALU möglich.", ln=1)
        self.set_x(COL2_X)
        self.cell(0, PDF_LINE_HEIGHT, "- Farbauswahl Führungsschiene nur bei ALU möglich.", ln=1)

        # --- Footer Text ---
        # Position near bottom - check page height if needed
        self.set_y(PAGE_HEIGHT_L - PDF_MARGIN - 10)  # Position near bottom
        self.set_x(COL1_X)
        self.set_font(PDF_FONT, '', 8)
        self.cell(40, PDF_LINE_HEIGHT, "Bestellblatt_Kopf")
    # --- END UNCOMMENTED METHOD ---

    # =========================================================================
    # === START: draw_positionen_pages MOVED INSIDE THE CLASS ===
    # =========================================================================
    def draw_positionen_pages(self, positions_data: List[Dict]):
        if not positions_data:
            self.add_page(orientation='L')
            self.set_font(PDF_FONT, 'B', 12)
            self.cell(0, 20, "Keine Positionen gefunden.", 0, 1, 'C')
            return

        self.add_page(orientation='L')
        self.ln(5) # Add some space below header

        # Sort definitions based on column index specified in config
        sorted_pos_defs = sorted(POS_COLS_DEFS.items(), key=lambda item: item[1][1])
        headers = [header for key, (header, idx) in sorted_pos_defs]
        data_keys = [key for key, (header, idx) in sorted_pos_defs]

        # Define column widths - THESE MUST MATCH THE ORDER AND NUMBER OF headers/data_keys
        # Based on the POS_COLS_DEFS structure provided (31 columns defined)
        col_widths = [
            6, 6, 7, 8, 10, 10, 8, 12, 12, # lfdNr_1 to Fensterart (9 cols)
            6,                          # Fensteröffnung_32 (1 col)
            10, 10,                     # Fenstergeometrie, Konstruktion (2 cols)
            20,                         # BehangTyp (1 col)
            8, 18,                      # Schallschutz_48, Antrieb (2 cols)
            10,                         # Fuehrungsschiene (1 col)
            10,                         # ReviblendeArt (1 col)
            7, 15,                      # Standardausführung_15, Fensterbankart (2 cols)
            5, 5, 8, 8, 16,             # Anzahl_Links_13 to Maßbezug (5 cols)
            10,                         # ISS_Ausführung (1 col)
            8, 8, 8,                    # ISS_Behindertengerecht_40 to ISS_Anzahl_Rechts_42 (3 cols)
            10,                         # EinzelteilTyp (1 col)
            10,                         # EinzelteilArt (1 col)
            8                           # EinzelteilAnzahl_25 (1 col)
        ] # Total = 31 widths

        # --- Sanity Checks ---
        if len(headers) != len(col_widths):
            logging.error(f"Header count ({len(headers)}) does not match column width count ({len(col_widths)}). PDF table layout will be incorrect.")
            # Option: Truncate or pad to prevent index errors, but layout will be wrong.
            # For now, let it proceed but log the error. User needs to fix config/widths.
            min_len = min(len(headers), len(col_widths))
            headers = headers[:min_len]
            data_keys = data_keys[:min_len]
            col_widths = col_widths[:min_len]
            logging.warning(f"Proceeding with {min_len} columns.")


        total_w = sum(col_widths)
        logging.info(f"Using Positionen table width: {total_w}mm (Printable: {PRINTABLE_WIDTH_L}mm)")
        scale_factor = 1.0
        if total_w > PRINTABLE_WIDTH_L:
            logging.warning("Table width exceeds printable area. Scaling columns.")
            scale_factor = PRINTABLE_WIDTH_L / total_w
            col_widths = [w * scale_factor for w in col_widths]
            total_w = sum(col_widths) # Recalculate total width after scaling
            logging.info(f"Scaled Positionen table width to: {total_w}mm")

        def draw_rotated_table_header():
            self.set_font(PDF_FONT, 'B', 7) # Use a small font for headers
            self.set_line_width(0.2)
            header_start_y = self.get_y()
            current_x = PDF_MARGIN

            for i, header in enumerate(headers):
                width = col_widths[i]
                center_x = current_x + width / 2
                # Adjust rotation_y to be near the bottom of the header cell space
                rotation_y = header_start_y + PDF_TABLE_HEADER_HEIGHT - 2

                # Wrap header text if it's too long for the rotated height
                # Estimate max width based on header height (minus margins)
                max_header_text_width = PDF_TABLE_HEADER_HEIGHT - 4
                wrapped_header = self.word_wrap(header, max_header_text_width)

                with self.rotation(angle=90, x=center_x, y=rotation_y):
                    # Adjust text position slightly for better centering after rotation
                    # The x-coordinate in rotated context corresponds to vertical position
                    # The y-coordinate corresponds to horizontal position (relative to rotation point)
                    # We want to center the text block vertically within the header height
                    text_height = self.get_string_width(wrapped_header.split('\n')[0]) # Approx height after rotation
                    start_x_rotated = center_x - (text_height / 2) # Adjust vertical start
                    # Use multi_cell for wrapped text
                    # Width of multi_cell is the available vertical space (header height)
                    # Height of multi_cell lines (e.g., 2.5 or 3)
                    self.set_xy(start_x_rotated, rotation_y) # Set position in rotated context
                    self.multi_cell(w=PDF_TABLE_HEADER_HEIGHT, h=3, txt=wrapped_header, align='C') # Use keywords for clarity

                # Draw vertical lines for the cell borders
                self.line(current_x, header_start_y, current_x, header_start_y + PDF_TABLE_HEADER_HEIGHT)
                current_x += width

            # Draw the rightmost vertical line and horizontal lines
            self.line(current_x, header_start_y, current_x, header_start_y + PDF_TABLE_HEADER_HEIGHT) # Right border
            self.line(PDF_MARGIN, header_start_y, PDF_MARGIN + total_w, header_start_y) # Top border
            self.line(PDF_MARGIN, header_start_y + PDF_TABLE_HEADER_HEIGHT, PDF_MARGIN + total_w, header_start_y + PDF_TABLE_HEADER_HEIGHT) # Bottom border

            # Set Y position for the first data row
            self.set_y(header_start_y + PDF_TABLE_HEADER_HEIGHT + 1) # Add small gap
            self.set_x(PDF_MARGIN)

        # --- Draw Header ---
        draw_rotated_table_header()

        # --- Draw Data Rows ---
        self.set_font(PDF_FONT, '', PDF_TABLE_ROW_HEIGHT - 1) # Font for data rows
        self.set_line_width(0.1) # Thinner lines for data rows

        for row_data in positions_data:
            # --- Calculate Max Height Needed for this Row ---
            if (
                row_data == positions_data[-1] and
                not row_data.get("FeBreite_11") and
                not row_data.get("FeHoehe_12")
            ):
            
                logging.info("Skipping last row due to empty 'FeBreite_11' and 'FeHoehe_12'.")
                continue
            row_start_y = self.get_y()
            max_cell_height_needed = PDF_TABLE_ROW_HEIGHT # Minimum height

            # Pre-calculate height needed for each cell in the row
            for i, key in enumerate(data_keys):
                if i < len(col_widths):
                    width = col_widths[i]
                    value = str(row_data.get(key, ''))
                    # Attempt to encode to handle potential unicode issues gracefully
                    try:
                        value.encode('latin-1')
                    except UnicodeEncodeError:
                        value = value.encode('latin-1', 'replace').decode('latin-1')

                    # Estimate number of lines needed by multi_cell
                    # Use keyword args for clarity in dry_run as well
                    lines = self.multi_cell(w=width, h=PDF_TABLE_ROW_HEIGHT, txt=value, border=0, align='L', dry_run=True, output='LINES')
                    cell_height = len(lines) * PDF_TABLE_ROW_HEIGHT
                    max_cell_height_needed = max(max_cell_height_needed, cell_height)

            # --- Check for Page Break BEFORE drawing the row ---
            if row_start_y + max_cell_height_needed > (self.h - self.b_margin):
                self.add_page(orientation='L')
                draw_rotated_table_header() # Redraw header on new page
                self.set_font(PDF_FONT, '', PDF_TABLE_ROW_HEIGHT - 1) # Reset font
                row_start_y = self.get_y() # Update starting Y for the row

            # --- Draw the actual row cells ---
            current_x = PDF_MARGIN
            for i, key in enumerate(data_keys):
                if i < len(col_widths):
                    self.set_xy(current_x, row_start_y) # Reset Y for each cell in the row
                    width = col_widths[i]
                    value = str(row_data.get(key, ''))
                    try:
                        value.encode('latin-1')
                    except UnicodeEncodeError:
                        value = value.encode('latin-1', 'replace').decode('latin-1')

                    # --- FIX: Use keyword 'h' for total height, 'max_line_height' for line height ---
                    # Remove the positional 'h' (PDF_TABLE_ROW_HEIGHT)
                    self.multi_cell(
                        w=width,                    # Cell width
                        txt=value,                  # Text content
                        border='LR',                # Left-Right border
                        align='L',                  # Left align
                        max_line_height=PDF_TABLE_ROW_HEIGHT, # Height of each line
                        h=max_cell_height_needed    # TOTAL height of the cell
                    )
                    # --- END FIX ---
                    current_x += width

            # --- Move Y position down by the height of the row and draw bottom border ---
            self.set_y(row_start_y + max_cell_height_needed)
            self.line(PDF_MARGIN, self.get_y(), PDF_MARGIN + total_w, self.get_y())
    # =========================================================================
    # === END: draw_positionen_pages MOVED INSIDE THE CLASS ===
    # =========================================================================


# --- Main Function to Generate PDF ---
def write_combined_pdf(mapped_data, output_directory, base_filename):
    """
    Generates a multi-page PDF with Kopf and Positionen data using the new structure.
    """
    kopf_data = mapped_data.get("kopf")
    positions_data = mapped_data.get("positionen")

    if not kopf_data:
        logging.warning("No Kopf data found to generate PDF.")
        return None
    # No need to check positions_data here, draw_positionen_pages handles empty list

    try:
        pdf = PDFWithHeaderFooter(orientation='L', unit='mm', format='A4') # START Landscape
        pdf.set_auto_page_break(auto=True, margin=PDF_MARGIN)
        pdf.alias_nb_pages() # Enable page numbering {nb}

        # --- Page 1: Kopf (Landscape) ---
        # This method MUST exist in the PDFWithHeaderFooter class
        pdf.draw_kopf_page(kopf_data)

        # --- Page 2+: Positionen ---
        # This method MUST exist in the PDFWithHeaderFooter class
        pdf.draw_positionen_pages(positions_data) # <<< THIS CALL WILL NOW WORK

        # --- Save PDF ---
        output_filename = f"{base_filename}.pdf"
        output_path = output_directory / output_filename
        pdf.output(str(output_path))

        logging.info(f"Successfully generated Combined PDF: {output_path}")
        return str(output_path)

    except Exception as e:
        # Log the specific error and traceback
        logging.error(f"Error generating Combined PDF: {e}", exc_info=True) # Keep exc_info=True
        return None

