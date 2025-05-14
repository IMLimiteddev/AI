
# data_mapper.py
import re
import sys
from datetime import datetime
from config import (
    KOPF_DEFAULTS, POS_DEFAULTS, ANTRIEB_MAP,
    FENSTERTYP_MAP, COLOR_MAP_BEHANG_KOPF
)

# --- Helper Functions ---

def format_date_dmy(date_str, input_format="%d.%m.%Y", output_format="%d.%m.%Y"):
    """Formats a date string, returns original or None on error."""
    if not date_str: return None
    try:
        dt_obj = datetime.strptime(str(date_str), input_format)
        return dt_obj.strftime(output_format)
    except ValueError:
        # print(f"Warning: Could not parse date '{date_str}' with format '{input_format}'", file=sys.stderr)
        return date_str
    except TypeError:
        # print(f"Warning: Invalid type for date formatting: {type(date_str)}", file=sys.stderr)
        return date_str

def _extract_color_code(text_line):
    """Extracts color codes like hwfXXXX or specific names, returns lowercase."""
    if not text_line: return None
    # Prioritize hwf code
    match_hwf = re.search(r'(hwf\d+)', text_line, re.IGNORECASE)
    if match_hwf: return match_hwf.group(1).lower()

    # If no hwf code, check for keywords (case-insensitive)
    if isinstance(text_line, str):
        text_lower = text_line.lower()
        if 'anthrazit matt' in text_lower: return 'hwf7016'
        if 'weißaluminium matt' in text_lower: return 'hwf9006'
        # Add other specific text mappings that should resolve to standard codes
        if 'anthrazit' in text_lower: return 'anthrazit' # Keep non-hwf codes if needed
        if 'silber' in text_lower: return 'silber'
        if 'grau' in text_lower: return 'grau'
        if 'weiß' in text_lower: return 'ral9016' # Assuming weiß maps to this standard

        # If it's likely a RAL code or other non-standard code, return it cleaned
        match_ral = re.search(r'(ral\s*\d+)', text_lower)
        if match_ral: return match_ral.group(1).replace(" ", "") # e.g., "ral7016"

        # Fallback: return the cleaned text if it doesn't match known patterns
        # This helps identify potentially non-standard codes in the check later
        cleaned_text = text_lower.strip()
        if cleaned_text: return cleaned_text

    return None # Return None if input is not string or no code/text found

# --- Main Mapping Function ---

def map_data_to_template(extracted_data):
    """
    Maps the raw extracted data to the structure required by the Excel template.
    - Sets Positionsnummerierung (PosNr_31) based on PDF description, formatted.
    - Generates Hinweistext based on Length, Beidseitig, and Color conditions.
    - Sets Geschoss to "0".
    - Includes address code defaults.
    """
    if not extracted_data or "positions" not in extracted_data:
        print("Error: Cannot map data - no extracted data or positions provided.", file=sys.stderr)
        return None

    kopf_data = KOPF_DEFAULTS.copy()
    positionen_data = []
    # Dictionary to store conditions met for each position number (PosNr_31)
    # Value is a set of strings: {"Kombi ändern", "mehrpreisRAL"}
    hinweis_conditions_met = {}

    # --- Map Kopf Data (Dates, Order Numbers, Address Codes) ---
    header = extracted_data.get("header", {})
    kopf_data["Auftragsname"] = header.get("KdAuftrag")
    kopf_data["Kunden-Auftrags-Nr"] = header.get("Bestnr")
    kopf_data["Bestelldatum"] = format_date_dmy(header.get("VomDate"), input_format="%d.%m.%Y")
    kopf_data["Wunsch-Liefertermin"] = format_date_dmy(header.get("Liefertermin"), input_format="%d.%m.%Y")
    # Add address codes from defaults (can be overridden if parsed from PDF later)
    kopf_data["Kundennummer"] = KOPF_DEFAULTS.get("kundennummer", "2144")
    kopf_data["Rechnungsadr"] = KOPF_DEFAULTS.get("rechnungsadr", "58")
    kopf_data["Lieferadr"] = KOPF_DEFAULTS.get("lieferadr", "58")
    kopf_data["AufBestAdr"] = KOPF_DEFAULTS.get("aufbestadr", "58")


    # --- Determine Kopf Colors ---
    if extracted_data.get("positions"):
        first_main_pos = next((pos for pos in extracted_data["positions"] if "Führungsschiene Alu Paarweise" not in pos.get("InitialBeschreibung", "")), None)
        if not first_main_pos and extracted_data["positions"]: first_main_pos = extracted_data["positions"][0]
        if first_main_pos:
            # Use _extract_color_code for consistency, fallback to default from kopf_data
            kopf_data["Farben_Behang"] = _extract_color_code(first_main_pos.get("Panzer")) or kopf_data["Farben_Behang"]
            kopf_data["Farben_Endleiste"] = _extract_color_code(first_main_pos.get("Endschiene")) or kopf_data["Farben_Endleiste"]
            kopf_data["Farben_Reviblende"] = _extract_color_code(first_main_pos.get("Revision")) or kopf_data["Farben_Reviblende"]
            kopf_data["Farben_Fuehrungsschiene"] = _extract_color_code(first_main_pos.get("FehroFS")) or kopf_data["Farben_Fuehrungsschiene"]
            # Anschlagstopfen and Kurbelstange often have fixed defaults or simple text
            kopf_data["Farben_Anschlagstopfen"] = first_main_pos.get("Anschlagstopfen", kopf_data["Farben_Anschlagstopfen"]) # Assuming Anschlagstopfen is parsed
            kopf_data["Kurbelstange"] = first_main_pos.get("Kurbelstange", kopf_data["Kurbelstange"]) # Assuming Kurbelstange is parsed


    # --- Map Positionen Data ---
    for i, pos_raw in enumerate(extracted_data["positions"]):
        lfd_nr = i + 1
        pos_mapped = POS_DEFAULTS.copy() # Gets Geschoss="0" from default

        try:
            # --- Basic Info & PosNr_31 Formatting ---
            pos_mapped["lfdNr_1"] = str(lfd_nr)
            beschreibung_pos_nr_raw = pos_raw.get("BeschreibungPosNr")
            pos_nr_31_value = str(lfd_nr) # Default fallback
            if beschreibung_pos_nr_raw and isinstance(beschreibung_pos_nr_raw, str):
                beschreibung_pos_nr_cleaned = beschreibung_pos_nr_raw.rstrip('_')
                try:
                    # Attempt to convert to int for sorting/checking range
                    num_val = int(beschreibung_pos_nr_cleaned)
                    # Keep as simple number string if 1-9, otherwise keep original cleaned string
                    pos_nr_31_value = str(num_val) # if 1 <= num_val <= 9 else beschreibung_pos_nr_cleaned
                except ValueError:
                    pos_nr_31_value = beschreibung_pos_nr_cleaned # Keep as string if not integer
            pos_mapped["PosNr_31"] = pos_nr_31_value

            # Initialize the set for this position number if not already present
            if pos_nr_31_value not in hinweis_conditions_met:
                hinweis_conditions_met[pos_nr_31_value] = set()

            # --- Geschoss remains "0" (default) ---

            # --- Other Mappings (Konstruktion, BehangTyp, Fenstertyp, Fensterbankart) ---
            antrieb_raw = pos_raw.get("Antrieb")
            pos_mapped["Antrieb"] = ANTRIEB_MAP.get(str(antrieb_raw), str(antrieb_raw)) if antrieb_raw else POS_DEFAULTS["Antrieb"]
            zeich_raw = pos_raw.get("Zeichnung")
            if zeich_raw and isinstance(zeich_raw, str):
                 match_konstr = re.search(r'R\d+/(\d+)', zeich_raw)
                 pos_mapped["Konstruktion"] = f"k{match_konstr.group(1)}" if match_konstr else ""
            else: pos_mapped["Konstruktion"] = ""
            desc_raw = pos_raw.get("InitialBeschreibung", "")
            is_fuehrungsschiene_item = ("Führungsschiene Alu Paarweise" in desc_raw) if isinstance(desc_raw, str) else False
            if isinstance(desc_raw, str) and "Rollladensystem Fehro_AR DM40" in desc_raw: pos_mapped["BehangTyp"] = "Rollladen Alu DM40"
            elif is_fuehrungsschiene_item: pos_mapped["BehangTyp"] = "Führungsschiene Paar"
            else: pos_mapped["BehangTyp"] = POS_DEFAULTS["BehangTyp"]
            fensterbank_text = pos_raw.get("Fensterbank", "")
            mapped_fenstertyp = POS_DEFAULTS["Fenstertyp_9"]
            if isinstance(fensterbank_text, str):
                 for pattern, ftype in FENSTERTYP_MAP.items():
                      if pattern.search(fensterbank_text): mapped_fenstertyp = ftype; break
            pos_mapped["Fenstertyp_9"] = mapped_fenstertyp
            winkel_fs = pos_raw.get("WinkelFS")
            fensterbank_text_safe = str(fensterbank_text) if fensterbank_text else ""
            if winkel_fs == "0": pos_mapped["Fensterbankart"] = "Komfortschwelle"
            elif winkel_fs == "5": pos_mapped["Fensterbankart"] = "Steinbank" if "Steinfensterbank" in fensterbank_text_safe else "Alubank"
            else: pos_mapped["Fensterbankart"] = "Steinbank" if "Steinfensterbank" in fensterbank_text_safe else POS_DEFAULTS["Fensterbankart"]

            # --- Links/Rechts and Fensteraufteilung ---
            antrieb_seite_raw = pos_raw.get("Antriebsseite")
            antrieb_seite = str(antrieb_seite_raw).strip().lower() if antrieb_seite_raw else ""
            pos_mapped['Antriebsseite'] = antrieb_seite # Store for TXT writer
            notkurbel = pos_raw.get("Notkurbel")
            notkurbel_str = str(notkurbel).strip().lower() if notkurbel is not None else ""
            antrieb_raw_for_check = pos_raw.get("Antrieb")
            antrieb_desc = str(antrieb_raw_for_check).strip().lower() if antrieb_raw_for_check else ""
            links, rechts = "0", "0"
            use_seite = antrieb_seite
            if "nhk-kit" in antrieb_desc and notkurbel_str in ["links", "rechts"]: use_seite = notkurbel_str
            if "links" in use_seite: links = "1"
            if "rechts" in use_seite: rechts = "1"
            is_beidseitig = "beidseitig" in use_seite # Check if Antriebsseite is 'beidseitig'
            if is_beidseitig: links = "1"; rechts = "1"; pos_mapped["Fensteraufteilung"] = "Komb (Fe-Ka-An) 1-1-2"
            pos_mapped["Anzahl_Links_13"] = links
            pos_mapped["Anzahl_Rechts_14"] = rechts

            # --- Direct Mapping & Raw Values ---
            pos_mapped["FeBreite_11"] = pos_raw.get("Breite", "")
            pos_mapped["FeHoehe_12"] = pos_raw.get("LaengeFS", "") # This is the Fuhrung/Kasten Length
            pos_mapped["WinkelFS_raw"] = pos_raw.get("WinkelFS", "")

            # --- Hinweistext Generation Logic ---

            # Condition i: Length >= 2396mm
            length_fs_str = pos_mapped["FeHoehe_12"]
            length_fs = 0
            try: length_fs = int(length_fs_str) if length_fs_str else 0
            except ValueError: pass # Ignore if conversion fails
            if length_fs >= 2396:
                hinweis_conditions_met[pos_nr_31_value].add("Kombi ändern")
                # print(f"DEBUG Hinweistext: Pos {pos_nr_31_value} met Length condition ({length_fs} >= 2396)")

            # Condition ii: Antriebsseite is Beidseitig
            if is_beidseitig:
                hinweis_conditions_met[pos_nr_31_value].add("Kombi ändern")
                # print(f"DEBUG Hinweistext: Pos {pos_nr_31_value} met Beidseitig condition")

            # Condition iii: Irregular color codes
            irregular_color_found = False
            standard_colors = ['hwf9006', 'hwf9016'] # Define the standard/allowed codes
            color_fields_to_check = {
                "FehroFS": pos_raw.get("FehroFS"),         # Führungsschiene
                "Endschiene": pos_raw.get("Endschiene"),   # Rollladen/Jalousie Endschiene
                "Revision": pos_raw.get("Revision")        # Revisionblende
            }
            for field_name, raw_color_text in color_fields_to_check.items():
                extracted_code = _extract_color_code(raw_color_text)
                # Check if code exists and is NOT one of the standard codes
                if extracted_code and extracted_code not in standard_colors:
                    irregular_color_found = True
                    # print(f"DEBUG Hinweistext: Pos {pos_nr_31_value} met Irregular Color condition. Field: {field_name}, Raw: '{raw_color_text}', Code: '{extracted_code}'")
                    break # No need to check other colors for this position if one is irregular
            if irregular_color_found:
                hinweis_conditions_met[pos_nr_31_value].add("mehrpreisRAL")

            # --- Add the fully mapped position ---
            positionen_data.append(pos_mapped)

        except Exception as e:
            print(f"Error mapping position index {i} (lfdNr {lfd_nr}, Formatted PosNr: {pos_mapped.get('PosNr_31', 'N/A')}): {e}", file=sys.stderr)
            print(f"Problematic Raw Data: {pos_raw}", file=sys.stderr)

    # --- Finalize Kopf Hinweistext ---
    final_hinweis_parts = []
    # Sort by position number (handle potential non-numeric PosNr_31 values)
    def sort_key(pos_nr):
        try: return int(pos_nr)
        except ValueError: return float('inf') # Put non-numeric ones at the end

    for pos_nr in sorted(hinweis_conditions_met.keys(), key=sort_key):
        conditions = hinweis_conditions_met[pos_nr]
        if not conditions: continue # Skip if no conditions met for this pos_nr

        # Add messages based on the conditions found
        if "Kombi ändern" in conditions:
            final_hinweis_parts.append(f"Pos.{pos_nr} Kombi ändern")
        if "mehrpreisRAL" in conditions:
            final_hinweis_parts.append(f"Pos.{pos_nr} mehrpreisRAL")

    # Join the parts with spaces, default to empty string if no parts
    #kopf_data["Hinweistext"] = " ".join(final_hinweis_parts) if final_hinweis_parts else ""
        kopf_data["Hinweistext"] = " "
    # print(f"DEBUG: Final Kopf Hinweistext: '{kopf_data['Hinweistext']}'")
    # print(f"INFO: Data mapping complete. {len(positionen_data)} positions processed.")
    return { "kopf": kopf_data, "positionen": positionen_data }
