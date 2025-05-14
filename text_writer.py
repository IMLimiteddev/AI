
# text_writer.py
import logging
import csv
import pathlib
from datetime import datetime
from typing import Dict, List, Any
import os
import re
from config import KOPF_DEFAULTS, POS_COLS_DEFS

KOPF_TXT_LABELS = {
    'Kundennummer': 'Kundennummer',
    'Rechnungsadr': 'Rechnungsadr',
    'Lieferadr': 'Lieferadr',
    'AufBestAdr': 'AufBestAdr',
    'VomDate': 'VomDate', # Mapped from Bestelldatum
    'Auftragsname': 'Auftragsname', # Mapped from Auftragsname
    'Kunden-Auftrag-Nr': 'Kunden-Auftrag-Nr', # Mapped from Kunden-Auftrags-Nr
    'Liefertermin': 'Liefertermin', # Mapped from Wunsch-Liefertermin
    'Besteller': 'Besteller',
    'Panzer': 'Panzer', # Color mapped from Kopf Farben_Behang
    'Anschlag': 'Anschlag', # Color mapped from Kopf Farben_Anschlagstopfen (or default)
    'Endschiene': 'Endschiene', # Color mapped from Kopf Farben_Endleiste
    'Revision': 'Revision', # Color mapped from Kopf Farben_Reviblende
    'Kurbel': 'Kurbel', # Color mapped from Kopf Kurbelstange (or default)
    'FehroFS': 'FehroFS', # Color mapped from Kopf Farben_Fuehrungsschiene
    'Besonderheiten': 'Besonderheiten', # Generated based on IS Rollo / RAL MP
    'LKW': 'LKW', # Static 'LKW'
    'Sonder': 'Sonder', # 'Ja'/'Nein' based on width check
    'SonderText': 'SonderText', # Generated message if Sonder is 'Ja'
    # Fields for IS Rollo colors in header (if needed, map from Kopf Farben_Insekt_*)
    'IS_Endschiene': 'IS_Endschiene',
    'IS_Fuehrungsschiene': 'IS_Fuehrungsschiene',
    'IS_Element': 'IS_Element'

}

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s')

# Define the delimiter
DELIMITER = ';'

# --- Helper function to safely get value ---
def safe_get(data_dict: Dict, key: str, default: Any = '') -> Any:
    """Safely gets a value from dict, returning default if key missing or value is None/empty."""
    val = data_dict.get(key, default)
    return val if val is not None and val != '' else default

# --- Helper function to check for IS Rollo in any position ---
def check_order_has_is_rollo(positions: List[Dict]) -> bool:
    # Check all positions *including* the last one for the header logic
    for pos in positions:
        fehro_fs_text = safe_get(pos, 'FehroFS', '')
        if isinstance(fehro_fs_text, str) and "Insektenschutzrollo Fehro: Ja" in fehro_fs_text:
            return True
    return False

# --- Helper function to check width threshold (>= 2396 triggers Sonder) ---
def check_order_is_sonder(positions: List[Dict]) -> (bool, str):
    max_width_pos_id = None; max_width = 0; sonder = False; sonder_text = "0"
    # Check all positions *including* the last one for the header logic
    for pos in positions:
        breite_str = safe_get(pos, 'FeBreite_11', '0'); pos_id = safe_get(pos, 'Pos', 'UNKNOWN')
        try:
            width = int(breite_str) if breite_str else 0
            if width >= 2396:
                sonder = True
                if max_width_pos_id is None or width > max_width : max_width = width; max_width_pos_id = pos_id
        except (ValueError, TypeError): continue
    if sonder and max_width_pos_id != 'UNKNOWN':
        fenster_nr_match = re.match(r'0*(\d+)', str(max_width_pos_id))
        display_pos_num = fenster_nr_match.group(1) if fenster_nr_match else max_width_pos_id
        sonder_text = f"Pos.{display_pos_num}>{max_width-1}mm"
    return sonder, sonder_text

# --- Helper function to format date dd.mm.yyyy ---
def format_date_dmy_txt(date_str):
    if not date_str: return ""
    try: dt_obj = datetime.strptime(str(date_str), "%d.%m.%Y"); return dt_obj.strftime("%d.%m.%Y")
    except (ValueError, TypeError): return str(date_str)

# --- Helper to extract specific color codes or names ---
def get_color_code(raw_text: Any) -> str | None:
    if not isinstance(raw_text, str): return None
    raw_text_lower = raw_text.lower(); match = re.search(r'(hwf\d+)', raw_text_lower)
    if match: return match.group(1)
    if 'anthrazit matt' in raw_text_lower: return 'hwf7016'; # ... add other color text mappings ...
    if 'anthrazit' in raw_text_lower: return 'anthrazit';
    if 'grau' in raw_text_lower: return 'grau';
    return None

# --- Helper to extract Konstruktion number ---
def get_konstruktion_code(zeich_text: Any, default: str = '0') -> str:
    if not isinstance(zeich_text, str): return default
    match = re.search(r'R\d+\/(\d+)', zeich_text); return match.group(1) if match else default

# --- Main TXT Writing Function ---
def write_auftrag_export_txt(mapped_data: Dict, output_directory: pathlib.Path, base_filename: str) -> str | None:
    """
    Writes mapped data to a semicolon-delimited TXT file (data rows only)
    applying logic from Translation.xlsx and incorporating specific corrections.
    Skips the last position item.
    """
    kopf = mapped_data.get("kopf", {})
    positions = mapped_data.get("positionen", [])
    if not kopf and not positions: logging.warning("No Kopf/Pos data for TXT."); return None

    logging.info(f"Kopf Data: {kopf}")

    try:
        output_filename = f"{base_filename}.txt"
        output_path = output_directory / output_filename
        logging.info(f"Writing Auftrag Export TXT (Translation.xlsx logic + fixes): {output_path}")

        # Header logic should still consider ALL positions (including the last one)
        order_has_is_rollo = check_order_has_is_rollo(positions)
        is_sonder, sonder_text_generated = check_order_is_sonder(positions)
        first_pos = positions[0] if positions else {}

        # --- Prepare Header Data Row (Applying Corrections) ---
        header_data_row = []
       # actual_panzer_color = get_color_code(safe_get(first_pos, 'Panzer')) or 'anthrazit'
        actual_panzer_color = get_color_code(kopf.get('Farben_Behang', ''))
        actual_fuehrung_color = get_color_code(kopf.get('Farben_Fuehrungsschiene', ''))
        actual_endschiene_color = get_color_code(kopf.get('Farben_Endleiste', ''))
        actual_revision_color = get_color_code(kopf.get('Farben_Reviblende', '')) 
        
 # Prioritize colors from the mapped Kopf data
#         # Provide sensible defaults if Kopf colors are missing
        actual_panzer_color = get_color_code(kopf.get('Farben_Behang')) or KOPF_DEFAULTS.get("Farben_Behang", 'silber')
        actual_fuehrung_color = get_color_code(kopf.get('Farben_Fuehrungsschiene')) or KOPF_DEFAULTS.get("Farben_Fuehrungsschiene", '0')
        actual_endschiene_color = get_color_code(kopf.get('Farben_Endleiste')) or KOPF_DEFAULTS.get("Farben_Endleiste", '0')
        actual_revision_color = get_color_code(kopf.get('Farben_Reviblende')) or KOPF_DEFAULTS.get("Farben_Reviblende", '0')
        actual_anschlag_color = get_color_code(kopf.get('Farben_Anschlagstopfen')) or KOPF_DEFAULTS.get("Farben_Anschlagstopfen", 'grau')
        actual_kurbel_color = get_color_code(kopf.get('Kurbelstange')) or KOPF_DEFAULTS.get("Kurbelstange", 'grau')
        # Get IS colors from Kopf (used if order_has_is_rollo)
        is_endschiene_color = get_color_code(kopf.get('Farben_Insekt_Endleiste')) or actual_endschiene_color # Fallback to standard
        is_fuehrung_color = get_color_code(kopf.get('Farben_Insekt_Fuehrungsschiene')) or actual_fuehrung_color # Fallback to standard
        is_element_color = get_color_code(kopf.get('Farben_Insekt_Element')) or actual_panzer_color # Fallback to standard

        header_data_row.append(kopf.get( 'Kundennummer', '2144')) # 1
        header_data_row.append(kopf.get( 'Rechnungsadr', '58')) # 2
        header_data_row.append(kopf.get( 'Lieferadr', '58')) # 3
        header_data_row.append(kopf.get( 'AufBestAdr', '58')) # 4
        header_data_row.append(format_date_dmy_txt(kopf.get( 'Bestelldatum', ''))) # 5 Use VomDate
        header_data_row.append(kopf.get( 'Auftragsname', '')) # 6
        header_data_row.append(kopf.get( 'Kunden-Auftrags-Nr', '')) # 7 Use Bestnr <<< CORRECTED KEY
        header_data_row.append(format_date_dmy_txt(kopf.get( 'Wunsch-Liefertermin', '')))# 8 Use Liefertermin <<< CORRECTED KEY
        header_data_row.append(kopf.get('Besteller', '')) # 9

        header_data_row.append(actual_panzer_color); header_data_row.append(actual_anschlag_color); # 10, 11
        header_data_row.append(actual_fuehrung_color); header_data_row.extend(['0'] * 5); # 12-17
        header_data_row.append(actual_endschiene_color); header_data_row.append('0'); # 18, 19
        header_data_row.append(actual_kurbel_color); header_data_row.extend(['0'] * 2); # 20-22
        header_data_row.append(actual_endschiene_color if order_has_is_rollo else '0'); # 23
        header_data_row.append(actual_fuehrung_color if order_has_is_rollo else '0');   # 24
        header_data_row.append(actual_panzer_color if order_has_is_rollo else '0');      # 25

        besonderheiten_parts = []
        if order_has_is_rollo: besonderheiten_parts.append("mit IS-Rollo")
        fs_code = get_color_code(safe_get(first_pos, 'Farben_Fuehrungsschiene')); el_code = get_color_code(safe_get(first_pos, 'Farben_Endleist')); rev_code = get_color_code(safe_get(first_pos, 'Farben_Reviblende'))
        if any(c and c not in ['hwf9006', 'hwf7016'] for c in [fs_code, el_code, rev_code]): besonderheiten_parts.append("+ RAL MP")
        header_data_row.append(" ".join(besonderheiten_parts) if besonderheiten_parts else '0') # 26

        header_data_row.append('LKW'); header_data_row.append('Ja' if is_sonder else 'Nein'); # 27, 28
        header_data_row.append(sonder_text_generated if is_sonder else '0'); # 29
        header_data_row.extend(['0'] * 5); header_data_row.append(actual_revision_color); # 30-35

        # --- Write to File ---
        with open(output_path, 'w', newline='', encoding='utf-8') as txtfile:
            writer = csv.writer(txtfile, delimiter=DELIMITER, quoting=csv.QUOTE_MINIMAL, lineterminator='\n')
            writer.writerow([str(x) for x in header_data_row]) # Write Header Row

            # --- MODIFICATION HERE: Loop through positions[:-1] ---
            # Write Position Data Rows (excluding the last item)
            logging.info(f"Processing {len(positions)} position rows for TXT output.")
            for i, pos in enumerate(positions):
                # Check if "FeBreite_11" and "FeHoehe_12" are empty for the current position
                if not safe_get(pos, 'FeBreite_11') and not safe_get(pos, 'FeHoehe_12'):
                    if i == len(positions) - 1:  # Skip only if it's the last row
                        logging.info(f"Skipping last row due to empty 'FeBreite_11' and 'FeHoehe_12'.")
                        continue
                
            # --- END MODIFICATION ---
                pos_data_row = []
                # Determine if THIS specific position has IS Rollo
                pos_has_is_rollo = isinstance(safe_get(pos, 'FehroFS', ''), str) and "Insektenschutzrollo Fehro: Ja" in safe_get(pos, 'FehroFS', '')
                winkel_fs_raw = safe_get(pos, 'WinkelFS_raw', '')
                # Get the Antriebsseite stored by the data_mapper
                antrieb_seite_val = safe_get(pos, 'Antriebsseite', '')

                # --- Field mapping applying Translation.xlsx logic + specific fixes ---
                pos_data_row.append(str(i + 1)) # 1 Importzeilen-Num (1, 2, 3...)
                pos_data_row.append('13')       # 2 Kernwand ('13')
                # 3 Geschoss (Dynamic)
                geschoss_text = safe_get(pos, 'Geschoss', '').upper(); geschoss_code = '0'
                if 'DACHGESCHOSS' in geschoss_text: geschoss_code = '4'
                elif 'OBERGESCHOSS' in geschoss_text: geschoss_code = '3'
                elif 'ERDGESCHOSS' in geschoss_text: geschoss_code = '2'
                elif 'KG' in geschoss_text: geschoss_code = '1'
                pos_data_row.append(geschoss_code)

                pos_data_row.append('1'); pos_data_row.append('1'); # 4, 5 ('1')
                # 6 Fensteraufteilung <<< CORRECTED >>>
                # Use the value stored by data_mapper
                is_beidseitig = isinstance(antrieb_seite_val, str) and antrieb_seite_val.strip().lower() == 'beidseitig'
                pos_data_row.append('2' if is_beidseitig else '1')

                pos_data_row.append('6' if pos_has_is_rollo else '5') # 7 Material FS (Dynamic 5/6)
                pos_data_row.append('1'); # 8 ('1')
                pos_data_row.extend(['0'] * 2)  # 9, 10 ('0')
                pos_data_row.append(safe_get(pos, 'FeBreite_11')) # 11 (Dynamic)
                pos_data_row.append(safe_get(pos, 'FeHoehe_12'))  # 12 (Dynamic)
                # 13/14 Bedienung L/R (Dynamic - VALUE FROM DATA_MAPPER)
                links = safe_get(pos, 'Anzahl_Links_13', '0')
                rechts = safe_get(pos, 'Anzahl_Rechts_14', '0')
                pos_data_row.append(links); pos_data_row.append(rechts);

                pos_data_row.append('Ja'); pos_data_row.append('0'); # 15, 16
                # 17 Konstruktion (Dynamic)
                pos_data_row.append(get_konstruktion_code(safe_get(pos, 'Zeichnung', ''), default='0'))
                pos_data_row.append('2') # 18 Behang ('2')
                # 19 Antrieb (Dynamic)
                antrieb_text = safe_get(pos, 'Antrieb', ''); antrieb_code = '0'
                if 'Motor Becker E03' in antrieb_text: antrieb_code = '23'
                elif 'Motor Becker E22 mit NHK-Kit3' in antrieb_text: antrieb_code = '24'
                pos_data_row.append(antrieb_code)

                pos_data_row.append('0'); pos_data_row.append('13'); pos_data_row.append('9'); # 20, 21, 22
                pos_data_row.extend(['0'] * 8) # 23-30 ('0')
                # 31 Positionsnummerierung (0, 1, 2...)
                pos_nr_31_value = safe_get(pos, 'PosNr_31', str(i + 1)) # Get mapped value, fallback to index+1
                pos_data_row.append(pos_nr_31_value)

                pos_data_row.append('Ja'); # 32 ('Ja')
                pos_data_row.extend(['0'] * 3) # 33-35 ('0')

                # 36-42 ISS Fields (Dynamic 0/1)
                iss_flag = '1' if pos_has_is_rollo else '0'
                pos_data_row.extend([iss_flag] * 3); # 36, 37, 38
                pos_data_row.extend(['0'] * 2); # 39, 40
                pos_data_row.append(links if pos_has_is_rollo else '0');  # 41
                pos_data_row.append(rechts if pos_has_is_rollo else '0'); # 42

               
                # 43 Fensterbankart (optional) <<< SWAPPED >>>
                pos_data_row.append('0')
                # 44,45,46 Mehrpreis Typ/Anzahl/Preis
                pos_data_row.extend(['0'] * 3) # ('0')
                # 47 Mehrpreisposition-Art <<< SWAPPED >>>
                mehrpreis_art = '0' # Default
                if winkel_fs_raw == '0': mehrpreis_art = '9' # Corrected value for Winkel 0
                elif winkel_fs_raw == '5': mehrpreis_art = '2' # Corrected value for Winkel 5
                pos_data_row.append(mehrpreis_art)
                # --- END SWAP ---

                pos_data_row.append('Nein'); # 48 ('Nein')
                pos_data_row.extend(['0'] * 2); # 49,50 ('0')
                pos_data_row.append('0' if pos_has_is_rollo else '1'); # 51 ReviblendeArt (Dynamic 0/1)

                writer.writerow([str(x) for x in pos_data_row]) # Write row

        logging.info(f"Successfully generated Auftrag Export TXT file (Translation.xlsx logic + Col Fixes, last row skipped): {output_path}")
        return str(output_path)

    except Exception as e:
        logging.error(f"Error writing Auftrag Export TXT file (Translation.xlsx logic + Col Fixes): {e}", exc_info=True)
        return None