
# pdf_parser.py
import fitz
import re
import logging
from config import PDF_MARKERS

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

def extract_data_from_pdf(pdf_path):
    doc = None
    try:
        doc = fitz.open(pdf_path)
        logging.info(f"Opened PDF: {pdf_path} with {len(doc)} pages using PyMuPDF.")

        header_data = {}
        all_text_blocks_with_lines = []
        full_page_text_for_search = ""

        # --- Header Extraction (Same) ---
        if len(doc) > 0:
            page1 = doc[0]
            page1_text = page1.get_text("text")
            full_page_text_for_search += page1_text + "\n"
            header_data['KdAuftrag'] = _find_first(PDF_MARKERS["KD_AUFTRAG"], page1_text)
            header_data['Bestnr'] = _find_first(PDF_MARKERS["BESTNR"], page1_text)
            header_data['VomDate'] = _find_first(PDF_MARKERS["VOM_DATE"], page1_text)
            header_data['Liefertermin'] = _find_first(PDF_MARKERS["LIEFERTERMIN"], page1_text)
            logging.info(f"Extracted Header Data: {header_data}")

            # --- Block/Line Extraction (Same) ---
            logging.info("Extracting text blocks from all pages...")
            for page_num in range(len(doc)):
                page = doc[page_num]
                blocks = page.get_text("blocks", sort=True)
                for block in blocks:
                    if block[6] == 0:
                        block_text_content = block[4].strip()
                        if block_text_content:
                            block_lines_list = [line for line in block_text_content.split('\n') if line.strip()] # Store non-empty lines
                            if block_lines_list: # Only add if block has non-empty lines
                                all_text_blocks_with_lines.append((block_text_content, block_lines_list))
                if page_num > 0:
                    full_page_text_for_search += page.get_text("text") + "\n"
            logging.info(f"Collected {len(all_text_blocks_with_lines)} text blocks with content.")
        else:
            logging.error("PDF has no pages.")
            return None

        # --- Identify Position Blocks (NEW STRATEGY) ---
        logging.info("Processing extracted text blocks to identify position blocks (Strategy 3)...")
        positions = []
        current_block_lines_accumulator = []
        in_positions_area = False

        pos_start_header_regex = re.compile(PDF_MARKERS["POS_START"], re.IGNORECASE)
        # Regex to find the Pos number at the start of a line
        pos_num_regex = re.compile(r"^(00\d{3})$") # Match ONLY Pos number

        logging.info(f"Using Pos number regex for start detection: {pos_num_regex.pattern}")

        # Find header using simple text search first
        if pos_start_header_regex.search(full_page_text_for_search):
            in_positions_area = True
            logging.info("Found Pos Header marker using text search. Starting block identification.")
        else:
            logging.warning("Could not find the 'Pos. Material Bezeichnung' header using text search. WILL ATTEMPT ANYWAY.")
            in_positions_area = True

        # --- Iterate through lines from all blocks ---
        # Flatten the list of lines for easier sequential processing
        all_lines = []
        for _, block_lines in all_text_blocks_with_lines:
            all_lines.extend(block_lines)
        logging.info(f"Processing {len(all_lines)} total non-empty lines.")

        if in_positions_area:
            for line_idx, line in enumerate(all_lines):
                line_stripped = line.strip()
                if not line_stripped: continue # Should not happen due to filter above, but safe check

                logging.debug(f"[L{line_idx+1}] Checking Line: '{line_stripped[:80]}...'")

                # Ignore Repeating Header
                if pos_start_header_regex.search(line_stripped):
                    logging.debug(f"Ignoring repeated header line: '{line_stripped}'")
                    # If header appears mid-block, process previous block? Might be complex.
                    # Let's assume headers are separate blocks for now.
                    continue

                # Check for end marker
                if "Gesamtpositionsnettowert" in line_stripped or "Lieferantenzuschlag" in line_stripped:
                    logging.info(f"Found end marker in line: '{line_stripped}'. Processing final block.")
                    if current_block_lines_accumulator:
                        pos_data = _process_position_block_pymupdf_v3(current_block_lines_accumulator)
                        if pos_data: positions.append(pos_data)
                    current_block_lines_accumulator = []
                    in_positions_area = False
                    break # Stop processing lines

                # --- NEW START DETECTION based on Pos number ---
                is_new_start_line = bool(pos_num_regex.match(line_stripped))

                if is_new_start_line:
                    logging.info(f"Potential NEW position start detected with Pos line: '{line_stripped}'")
                    # Process the PREVIOUS block
                    if current_block_lines_accumulator:
                        logging.debug(f"Processing previous block ({len(current_block_lines_accumulator)} lines)")
                        pos_data = _process_position_block_pymupdf_v3(current_block_lines_accumulator)
                        if pos_data:
                            positions.append(pos_data)
                        else:
                            logging.warning("Discarded previous block because processing failed.")
                    # Start the NEW block accumulator
                    current_block_lines_accumulator = [line_stripped]
                    logging.debug(f"Started new block accumulator with: {current_block_lines_accumulator}")

                elif current_block_lines_accumulator: # If we are inside a block, append lines
                    # Avoid appending common page footers/headers found within item details
                    if not ("SchwörerHaus KG" in line or "Bestellnummer/Datum" in line or re.match(r"\d{10}\s*/\s*\d{2}\.\d{2}\.\d{4}", line) or re.match(r"Seite\s+\d+", line)):
                         current_block_lines_accumulator.append(line_stripped)
                         logging.debug(f"Appended line '{line_stripped[:60]}...' to current block. Accumulator size: {len(current_block_lines_accumulator)}")
                    else:
                         logging.debug(f"Ignoring potential page header/footer within block: '{line_stripped}'")
                # --- End NEW START DETECTION ---

        # Process the very last block after the loop
        if current_block_lines_accumulator:
            logging.info("Processing last accumulated block after loop finish.")
            pos_data = _process_position_block_pymupdf_v3(current_block_lines_accumulator)
            if pos_data: positions.append(pos_data)
            else: logging.warning("Discarded final block after loop because processing failed.")

        logging.info(f"Identified {len(positions)} position blocks.")
        if not positions:
            logging.error(f"--- PARSING FAILURE (PyMuPDF): Failed to identify any position blocks based on Pos number start line. ---")


        # --- Post-process Positions (Extract details - Uses full BlockText) ---
        final_positions = []
        logging.info(f"Extracting details for {len(positions)} identified blocks...")
        for i, pos_data_block in enumerate(positions):
            pos_text = pos_data_block.get("BlockText", "") # Full text of the block
            if not pos_text:
                logging.warning(f"Skipping detail extraction for Pos {pos_data_block.get('Pos')} due to empty text block.")
                continue

            # Initialize with basic info found during block processing
            processed_pos_data = {
                "Pos": pos_data_block.get("Pos"),
                "Material": pos_data_block.get("Material"),
                "InitialBeschreibung": pos_data_block.get("InitialBeschreibung"), # From block processing
                "Fensternummer": pos_data_block.get("Fensternummer"), # From block processing
                "BeschreibungPosNr": pos_data_block.get("BeschreibungPosNr") # From block processing
            }

            # --- Extract ALL other details using _find_first and PDF_MARKERS ---
            # This part remains the same, relies on good BlockText
            processed_pos_data['Fensternummer'] = _find_first(PDF_MARKERS["Fensternummer"], pos_text) or processed_pos_data.get("Fensternummer")
            processed_pos_data['Breite'] = _find_first(PDF_MARKERS["Breite"], pos_text)
            processed_pos_data['LaengeFS'] = _find_first(PDF_MARKERS["LaengeFS"], pos_text)
            processed_pos_data['WinkelFS'] = _find_first(PDF_MARKERS["WinkelFS"], pos_text)
            processed_pos_data['Geschoss'] = _find_first(PDF_MARKERS["Geschoss"], pos_text)
            processed_pos_data['Antriebsseite'] = _find_first(PDF_MARKERS["Antriebsseite"], pos_text)
            processed_pos_data['Notkurbel'] = _find_first(PDF_MARKERS["Notkurbel"], pos_text, ignore_case=True)
            processed_pos_data['Antrieb'] = _find_first(PDF_MARKERS["Antrieb"], pos_text)
            processed_pos_data['Panzer'] = _find_first(PDF_MARKERS["Panzer"], pos_text)
            processed_pos_data['FehroFS'] = _find_first(PDF_MARKERS["FehroFS"], pos_text)
            processed_pos_data['Endschiene'] = _find_first(PDF_MARKERS["Endschiene"], pos_text)
            processed_pos_data['Revision'] = _find_first(PDF_MARKERS["Revision"], pos_text)
            processed_pos_data['Zeichnung'] = _find_first(PDF_MARKERS["Zeichnung"], pos_text)
            processed_pos_data['Fensterbank'] = _find_first(PDF_MARKERS["Fensterbank"], pos_text)
            processed_pos_data['FensterDesc'] = _find_first(PDF_MARKERS["FensterDesc"], pos_text)
            # --- End detail extractions ---

            final_positions.append(processed_pos_data)
            logging.debug(f"Extracted details for Pos {processed_pos_data.get('Pos')}")


        extracted_data = { "header": header_data, "positions": final_positions }
        logging.info(f"PDF parsing complete (PyMuPDF). Final positions processed: {len(final_positions)}")
        return extracted_data

    except ImportError:
         logging.error("PyMuPDF (fitz) not installed. Please run: pip install pymupdf")
         return None
    except Exception as e:
        logging.error(f"Error parsing PDF {pdf_path} using PyMuPDF: {e}", exc_info=True)
        return None
    finally:
        if doc:
             doc.close()
             logging.debug("Closed PDF document.")


# --- Helper: _process_position_block_pymupdf_v3 (NEW HELPER) ---
def _process_position_block_pymupdf_v3(accumulated_lines):
    """
    Helper function to process a block of accumulated lines for a single position.
    Attempts to find Pos, Material, and Description within the first few lines.
    Extracts the leading number from the description if present.
    """
    if not accumulated_lines or len(accumulated_lines) < 3: # Need at least Pos, Material, Desc lines ideally
        logging.warning(f"Helper _process_position_block_v3 received insufficient lines: {len(accumulated_lines)}")
        return None

    block_text = "\n".join(accumulated_lines) # Full text for detail extraction later

    pos_num_str = None
    material = None
    desc = None
    beschreibung_pos_nr = None

    # Define regexes for components
    pos_num_regex = re.compile(r"^(00\d{3})$")
    material_num_regex = re.compile(r"^(\d{8})$")
    desc_start_regex = re.compile(r"^(\d+_)")

    # Search within the first ~5 lines for the components
    search_limit = min(len(accumulated_lines), 5)
    found_pos_idx = -1
    found_mat_idx = -1
    found_desc_idx = -1

    for idx in range(search_limit):
        line = accumulated_lines[idx]
        if found_pos_idx == -1 and pos_num_regex.match(line):
            pos_num_str = line
            found_pos_idx = idx
            logging.debug(f"  HelperV3: Found Pos '{pos_num_str}' at index {idx}")
            continue # Move to next line to find Material

        if found_pos_idx != -1 and found_mat_idx == -1 and material_num_regex.match(line):
            material = line
            found_mat_idx = idx
            logging.debug(f"  HelperV3: Found Material '{material}' at index {idx}")
            continue # Move to next line to find Desc

        if found_mat_idx != -1 and found_desc_idx == -1 and desc_start_regex.match(line):
            desc = line # Store the line that starts the description
            found_desc_idx = idx
            logging.debug(f"  HelperV3: Found Desc start '{desc[:30]}...' at index {idx}")
            # Extract the description number
            match_desc_num = desc_start_regex.match(desc)
            if match_desc_num:
                beschreibung_pos_nr = match_desc_num.group(1)
                logging.debug(f"  HelperV3: Found BeschreibungPosNr '{beschreibung_pos_nr}'")
            break # Found all three key parts

    # Fallback / Validation
    if not pos_num_str or not material or not desc:
        logging.warning(f"HelperV3 failed to find all key parts (Pos/Mat/Desc) in first {search_limit} lines: {accumulated_lines[:search_limit]}")
        # Attempt basic fallback using first line as Pos if possible
        if pos_num_regex.match(accumulated_lines[0]):
            pos_num_str = accumulated_lines[0]
        else:
            pos_num_str = f"UNKNOWN_{accumulated_lines[0][:10]}" # Fallback Pos
        material = material or "UNKNOWN"
        desc = desc or (accumulated_lines[2] if len(accumulated_lines) > 2 else accumulated_lines[0]) # Guess description

    # Pre-extract Fensternummer from the full block_text
    fensternummer = _find_first(PDF_MARKERS["Fensternummer"], block_text)

    pos_data = {
        "Pos": pos_num_str,
        "Material": material,
        "InitialBeschreibung": desc, # Store the first line identified as description
        "BlockText": block_text,     # Store the full block text
        "Fensternummer": fensternummer,
        "BeschreibungPosNr": beschreibung_pos_nr
    }
    logging.info(f"Processed Block Data: Pos={pos_num_str}, Mat={material}, DescStart='{str(desc)[:30]}...', BPNr={beschreibung_pos_nr}")
    return pos_data


# --- Helper: _find_first (No change needed) ---
def _find_first(pattern, text, default=None, ignore_case=False):
    # ... (keep implementation from previous response) ...
    if not text: return default
    regex_pattern_str = pattern.pattern if hasattr(pattern, 'pattern') and isinstance(pattern.pattern, str) else str(pattern)
    if not isinstance(regex_pattern_str, str):
         logging.error(f"Invalid pattern type passed to _find_first: {type(regex_pattern_str)}")
         return default

    flags = re.IGNORECASE if ignore_case else 0
    try:
        match = re.search(regex_pattern_str, text, flags=flags)
        if match:
            try:
                 compiled_pattern = re.compile(regex_pattern_str)
                 num_groups = compiled_pattern.groups
            except re.error:
                 num_groups = 0

            if num_groups >= 1 and len(match.groups()) >= 1:
                 result = match.group(1)
            else:
                 result = match.group(0)
            return result.strip() if result else default
    except Exception as regex_error:
        logging.error(f"Regex error searching for pattern '{regex_pattern_str}' in text snippet '{text[:100]}...': {regex_error}")
    return default