import sys
import os
import PySimpleGUI as sg
import pandas as pd
import traceback
from pypdf import PdfReader, PdfWriter, PdfMerger
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from io import BytesIO
import pdfplumber
import csv
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import logging
from datetime import datetime, timedelta
import subprocess

# Configure logging with rotation
from logging.handlers import RotatingFileHandler

def setup_logging():
    """Set up logging with rotation to prevent log files from growing too large"""
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    
    # Remove any existing handlers
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)
    
    # Create rotating file handler
    # maxBytes=5MB, keep 3 backup files (total ~15MB max)
    handler = RotatingFileHandler(
        'packing_slip_organizer.log',
        maxBytes=5*1024*1024,  # 5 MB
        backupCount=3
    )
    
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    
    logger.addHandler(handler)

# Initialize logging
setup_logging()

# Define the desired order of areas
AREA_SORT_ORDER = ['A13', 'D16', 'B11', 'B12', 'B13', 'B14', 'B16', 'B17', 'B18', 'B19', 'garage']

# Print Settings - Adjust these if your printer outputs pages in wrong order
PRINT_REVERSE_ORDER = True  # Set to True if printer outputs pages face-up (reverse order)
PRINT_COLLATE = True       # Keep pages in proper sequence

def size_sort_key(size):
    """
    Assigns a numerical value to sizes for sorting.
    Smaller sizes get lower numbers.
    """
    # Define the order for letter sizes
    letter_sizes_order = {
        'XS': 1,
        'S': 2,
        'M': 3,
        'L': 4,
        'XL': 5,
        'XXL': 6
    }

    size_clean = size.lower().strip()

    # Handle sizes like 'Size 1', 'Size 2', etc.
    if size_clean.startswith('size '):
        try:
            num = int(size_clean.replace('size ', '').strip())
            return num
        except ValueError:
            pass  # Fall through to other checks

    # Handle pure numerical sizes
    try:
        num = int(size_clean)
        return num
    except ValueError:
        pass

    # Handle letter sizes
    size_upper = size.upper()
    return letter_sizes_order.get(size_upper, 100)  # Unknown sizes get a high number

def resource_path(relative_path):
    """ Get absolute path to resource, prioritizing application path for standalone execution """
    base_path = "/Applications/Yourapp"  # Explicitly set to your app's folder
    return os.path.join(base_path, relative_path)

sg.theme('LightBlue')

# Path to the warehouse_map.csv file using resource_path
csv_file_path = resource_path('assets/warehouse_map.csv')

# Function to open files on macOS
def open_file_mac(file_path):
    try:
        os.system(f'open "{file_path}"')
        logging.info(f"Opened file: {file_path}")
    except Exception as e:
        logging.error(f"Failed to open file {file_path}: {str(e)}")

# Function to load the warehouse map
def load_warehouse_map(csv_file_path):
    if os.path.exists(csv_file_path):
        try:
            df = pd.read_csv(csv_file_path)
            df = df[["Product Name", "Area"]]  # Keep only necessary columns
            df = df.dropna()  # Remove rows with any NaN values
            logging.info(f"Loaded warehouse map from {csv_file_path}")
            return df
        except Exception as e:
            sg.popup_error(f'Error reading warehouse map CSV:\n{str(e)}')
            logging.error(f"Error reading warehouse map CSV: {str(e)}")
            return pd.DataFrame(columns=["Product Name", "Area"])
    else:
        sg.popup_error(f'Warehouse map file not found at {csv_file_path}. Please ensure the CSV is included.')
        logging.error(f"Warehouse map file not found at {csv_file_path}")
        return pd.DataFrame(columns=["Product Name", "Area"])

# Function to save the warehouse map
def save_warehouse_map(df, csv_file_path):
    try:
        df.to_csv(csv_file_path, index=False)
        logging.info(f"Saved warehouse map to {csv_file_path}")
    except Exception as e:
        sg.popup_error(f'Error saving warehouse map:\n{str(e)}')
        logging.error(f"Error saving warehouse map: {str(e)}")

def determine_area_identifier(product_areas):
    # If there's more than one area and one of them is 'B16'
    if 'B16' in product_areas and len(product_areas) > 1:
        # Remove 'B16' and return the other areas
        other_areas = product_areas - {'B16'}
        area_identifier = ', '.join(sorted(other_areas))
    else:
        # Return all areas (could be just 'B16' or multiple areas)
        area_identifier = ', '.join(sorted(product_areas))
    logging.debug(f"Determined area identifier: {area_identifier}")
    return area_identifier

# Function to create the Warehouse Map Update Window
def open_warehouse_map_window():
    try:
        # Load the warehouse map DataFrame
        warehouse_map_df = load_warehouse_map(csv_file_path)

        # Ensure DataFrame columns and values are converted properly to lists
        values_list = warehouse_map_df.values.tolist()  # Convert DataFrame values to a list of lists
        headings = warehouse_map_df.columns.tolist()  # Convert DataFrame columns to a list

        # Sort the warehouse map alphabetically by product name
        warehouse_map_df = warehouse_map_df.sort_values('Product Name', key=lambda x: x.str.lower()).reset_index(drop=True)
        original_df = warehouse_map_df.copy()  # Keep original for filtering
        current_display_df = warehouse_map_df.copy()  # Initialize current display
        
        # Convert to list for display
        values_list = warehouse_map_df.values.tolist()

        # Define the layout for the Warehouse Map GUI
        layout = [
            [sg.Text('Warehouse Map', font=('Helvetica', 16, 'bold'))],
            [sg.Text('Search Product:'), sg.InputText(key='-SEARCH_PRODUCT-', enable_events=True, size=(20, 1)),
             sg.Text('Search Locker:'), sg.InputText(key='-SEARCH_LOCKER-', enable_events=True, size=(10, 1)),
             sg.Button('Clear Filters')],
            [sg.Text(f'Showing {len(values_list)} entries', key='-STATUS_TEXT-', font=('Helvetica', 10))],
            [sg.Table(
                values=values_list,
                headings=headings,
                key='-TABLE-',
                auto_size_columns=True,
                justification='left',
                num_rows=min(25, len(values_list)),
                enable_events=True,
                select_mode=sg.TABLE_SELECT_MODE_BROWSE,
                size=(None, 400)
            )],
            [sg.Text('Product Name'), sg.InputText(key='-PRODUCT-')],
            [sg.Text('Area'), sg.InputText(key='-AREA-')],
            [sg.Button('Add Entry'), sg.Button('Update Selected'), sg.Button('Remove Duplicates'), sg.Button('Save Changes'), sg.Button('Exit')]
        ]

        # Create the Warehouse Map window
        warehouse_window = sg.Window('Update Warehouse Map', layout, finalize=True)

        def filter_warehouse_data(search_product="", search_locker=""):
            """Filter the warehouse data based on search criteria"""
            filtered_df = original_df.copy()
            
            # Filter by product name if search term provided
            if search_product.strip():
                filtered_df = filtered_df[filtered_df['Product Name'].str.contains(search_product.strip(), case=False, na=False)]
            
            # Filter by locker if search term provided
            if search_locker.strip():
                filtered_df = filtered_df[filtered_df['Area'].str.contains(search_locker.strip(), case=False, na=False)]
            
            return filtered_df

        # Event loop for the Warehouse Map window
        while True:
            event, values = warehouse_window.read()

            if event == sg.WIN_CLOSED or event == 'Exit':
                break

            # Handle search functionality
            if event in ['-SEARCH_PRODUCT-', '-SEARCH_LOCKER-']:
                search_product = values['-SEARCH_PRODUCT-']
                search_locker = values['-SEARCH_LOCKER-']
                current_display_df = filter_warehouse_data(search_product, search_locker)
                warehouse_window['-TABLE-'].update(values=current_display_df.values.tolist())
                warehouse_window['-STATUS_TEXT-'].update(f'Showing {len(current_display_df)} entries')
                
            if event == 'Clear Filters':
                warehouse_window['-SEARCH_PRODUCT-'].update('')
                warehouse_window['-SEARCH_LOCKER-'].update('')
                current_display_df = warehouse_map_df.copy()
                warehouse_window['-TABLE-'].update(values=current_display_df.values.tolist())
                warehouse_window['-STATUS_TEXT-'].update(f'Showing {len(current_display_df)} entries')

            if event == '-TABLE-':
                selected_row = values['-TABLE-']
                if selected_row:
                    index = selected_row[0]
                    # Get the currently displayed data (might be filtered)
                    search_product = values['-SEARCH_PRODUCT-']
                    search_locker = values['-SEARCH_LOCKER-']
                    if search_product or search_locker:
                        current_display_df = filter_warehouse_data(search_product, search_locker)
                    else:
                        current_display_df = warehouse_map_df
                    
                    if index < len(current_display_df):
                        selected_product = current_display_df.iloc[index, 0]
                        selected_area = current_display_df.iloc[index, 1]
                        warehouse_window['-PRODUCT-'].update(selected_product)
                        warehouse_window['-AREA-'].update(selected_area)

            if event == 'Add Entry':
                product = values['-PRODUCT-'].strip()
                area = values['-AREA-'].strip()
                if product and area:
                    # Check if product already exists
                    existing_mask = warehouse_map_df['Product Name'].str.strip().str.lower() == product.lower()
                    if existing_mask.any():
                        response = sg.popup_yes_no(f"Product '{product}' already exists in the warehouse map. Do you want to update it instead of adding a duplicate?")
                        if response == 'Yes':
                            # Update existing entry
                            warehouse_map_df.loc[existing_mask, 'Area'] = area
                            # Re-sort the data alphabetically
                            warehouse_map_df = warehouse_map_df.sort_values('Product Name', key=lambda x: x.str.lower()).reset_index(drop=True)
                            original_df = warehouse_map_df.copy()
                            # Apply current filters and update display
                            search_product = values['-SEARCH_PRODUCT-']
                            search_locker = values['-SEARCH_LOCKER-']
                            current_display_df = filter_warehouse_data(search_product, search_locker)
                            warehouse_window['-TABLE-'].update(values=current_display_df.values.tolist())
                            warehouse_window['-STATUS_TEXT-'].update(f'Showing {len(current_display_df)} entries')
                            logging.info(f"Updated existing entry: Product='{product}', Area='{area}'")
                        else:
                            logging.info(f"User chose not to update existing product: {product}")
                    else:
                        # Add new entry
                        new_entry = pd.DataFrame({"Product Name": [product], "Area": [area]})
                        warehouse_map_df = pd.concat([warehouse_map_df, new_entry], ignore_index=True)
                        # Re-sort the data alphabetically
                        warehouse_map_df = warehouse_map_df.sort_values('Product Name', key=lambda x: x.str.lower()).reset_index(drop=True)
                        original_df = warehouse_map_df.copy()
                        # Apply current filters
                        search_product = values['-SEARCH_PRODUCT-']
                        search_locker = values['-SEARCH_LOCKER-']
                        current_display_df = filter_warehouse_data(search_product, search_locker)
                        warehouse_window['-TABLE-'].update(values=current_display_df.values.tolist())
                        warehouse_window['-STATUS_TEXT-'].update(f'Showing {len(current_display_df)} entries')
                        logging.info(f"Added entry: Product='{product}', Area='{area}'")
                else:
                    sg.popup_error("Please enter both a product name and an area.")
                    logging.warning("Attempted to add entry with missing product or area.")

            if event == 'Update Selected':
                selected_row = values['-TABLE-']
                if selected_row:
                    index = selected_row[0]
                    product = values['-PRODUCT-'].strip()
                    area = values['-AREA-'].strip()
                    if product and area:
                        # Get the currently displayed data to find the correct record
                        search_product = values['-SEARCH_PRODUCT-']
                        search_locker = values['-SEARCH_LOCKER-']
                        if search_product or search_locker:
                            current_display_df = filter_warehouse_data(search_product, search_locker)
                        else:
                            current_display_df = warehouse_map_df
                        
                        if index < len(current_display_df):
                            # Find the original product name to update in warehouse_map_df
                            original_product = current_display_df.iloc[index, 0]
                            # Update in the main dataframe
                            update_mask = warehouse_map_df['Product Name'] == original_product
                            warehouse_map_df.loc[update_mask, 'Product Name'] = product
                            warehouse_map_df.loc[update_mask, 'Area'] = area
                            # Re-sort the data alphabetically
                            warehouse_map_df = warehouse_map_df.sort_values('Product Name', key=lambda x: x.str.lower()).reset_index(drop=True)
                            original_df = warehouse_map_df.copy()
                            # Apply current filters and update display
                            current_display_df = filter_warehouse_data(search_product, search_locker)
                            warehouse_window['-TABLE-'].update(values=current_display_df.values.tolist())
                            warehouse_window['-STATUS_TEXT-'].update(f'Showing {len(current_display_df)} entries')
                            logging.info(f"Updated entry: Product='{product}', Area='{area}'")
                    else:
                        sg.popup_error("Please enter both a product name and an area.")
                        logging.warning("Attempted to update entry with missing product or area.")
                else:
                    sg.popup_error("Please select a row to update.")
                    logging.warning("Attempted to update without selecting a row.")

            if event == 'Remove Duplicates':
                original_count = len(warehouse_map_df)
                # Remove duplicates, keeping the last occurrence (most recent)
                warehouse_map_df = warehouse_map_df.drop_duplicates(subset=['Product Name'], keep='last')
                warehouse_map_df = warehouse_map_df.reset_index(drop=True)
                # Re-sort the data alphabetically
                warehouse_map_df = warehouse_map_df.sort_values('Product Name', key=lambda x: x.str.lower()).reset_index(drop=True)
                original_df = warehouse_map_df.copy()
                new_count = len(warehouse_map_df)
                removed_count = original_count - new_count
                
                # Apply current filters and update display
                search_product = values['-SEARCH_PRODUCT-']
                search_locker = values['-SEARCH_LOCKER-']
                current_display_df = filter_warehouse_data(search_product, search_locker)
                warehouse_window['-TABLE-'].update(values=current_display_df.values.tolist())
                warehouse_window['-STATUS_TEXT-'].update(f'Showing {len(current_display_df)} entries')
                sg.popup(f"Removed {removed_count} duplicate entries. {new_count} entries remaining.")
                logging.info(f"Removed {removed_count} duplicates from warehouse map.")

            if event == 'Save Changes':
                save_warehouse_map(warehouse_map_df, csv_file_path)
                sg.popup("Changes saved successfully!")
                logging.info("Warehouse map changes saved successfully.")

        warehouse_window.close()

    except Exception as e:
        sg.popup_error(f'Error in Warehouse Map Window: {str(e)}')
        logging.error(f"Error in Warehouse Map Window: {str(e)}", exc_info=True)

# Main App Layout
layout = [
    [sg.Text('Select the Packing Slips PDF:'), sg.Input(key='-PACKING_SLIPS-', enable_events=True), sg.FileBrowse(file_types=(("PDF Files", "*.pdf"),))],
    [sg.Text('Select the Shipping Labels PDF:'), sg.Input(key='-SHIPPING_LABELS-'), sg.FileBrowse(file_types=(("PDF Files", "*.pdf"),))],
    [sg.Text('Select Save Directory:'), sg.Input(key='-SAVE_PATH-'), sg.FolderBrowse()],
    [sg.Button('Process'), sg.Button('Update Warehouse Map'), sg.Button('Exit')],
    [sg.Text('', key='-STATUS-', size=(50, 1))]
]

window = sg.Window('Packing Slip Organizer', layout, finalize=True)

# Function to extract the first product from the packing slip
def extract_first_product(text):
    lines = text.split('\n')
    product_section_started = False
    for line in lines:
        line = line.strip()
        if 'ITEMS' in line:
            product_section_started = True
        elif product_section_started and line:
            # The first non-empty line after "ITEMS" is the product name
            return line.lower().strip()
    return None  # If no product found

def load_warehouse_map_dict(csv_file_path):
    warehouse_map = {}
    try:
        # Load the CSV file
        with open(csv_file_path, mode='r', newline='', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile)
            next(reader, None)  # Skip the header row
            for row in reader:
                if len(row) >= 2:
                    product, area = row[0], row[1]
                    # Normalize product names to lowercase and strip whitespace
                    product_normalized = product.strip().lower()
                    warehouse_map[product_normalized] = area.strip()
        # Print the loaded warehouse map for debugging
        logging.info("Warehouse Map Loaded (Normalized):")
        for product, area in warehouse_map.items():
            logging.info(f"Product: '{product}', Area: '{area}'")
    except Exception as e:
        sg.popup_error(f'Error loading warehouse map:\n{str(e)}')
        logging.error(f"Error loading warehouse map: {str(e)}", exc_info=True)
    return warehouse_map

def extract_text_from_page(pdf_path, page_number):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[page_number]
            text = page.extract_text()
        return text
    except Exception as e:
        sg.popup_error(f'Error extracting text from page {page_number + 1}:\n{str(e)}')
        logging.error(f"Error extracting text from page {page_number + 1}: {str(e)}", exc_info=True)
        return ''

def is_page_blank(pdf_path, page_number):
    """Check if a PDF page is blank or has minimal content"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if page_number >= len(pdf.pages):
                return True
            page = pdf.pages[page_number]
            text = page.extract_text()
            # Consider a page blank if it has less than 10 characters of text
            return not text or len(text.strip()) < 10
    except Exception as e:
        logging.error(f"Error checking if page {page_number + 1} is blank: {str(e)}")
        return True  # Assume blank if we can't read it

def get_warehouse_area(product, warehouse_map):
    # Check if product contains "sample" (case insensitive) - if so, assign to garage
    if 'sample' in product.lower():
        logging.info(f"SAMPLE RULE: Product '{product}' contains 'sample' - assigning to garage")
        return 'garage'
    
    # Check for exact match first
    if product in warehouse_map:
        area = warehouse_map[product]
        logging.debug(f"Found exact match for '{product}' - assigning to '{area}'")
        return area
    
    # If no exact match, return Unknown Area
    logging.debug(f"No match found for '{product}' - assigning to 'Unknown Area'")
    return 'Unknown Area'

def add_area_identifier_to_page(page, area_identifier):
    try:
        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        # Position the text at the top right corner
        x_position = 500  # Adjust based on page size and required position
        y_position = 750  # Adjust based on page size and required position
        can.drawString(x_position, y_position, area_identifier)
        can.save()
        packet.seek(0)
        new_pdf = PdfReader(packet)
        overlay_page = new_pdf.pages[0]
        page.merge_page(overlay_page)
        return page
    except Exception as e:
        sg.popup_error(f'Error adding area identifier to page:\n{str(e)}')
        logging.error(f"Error adding area identifier to page: {str(e)}", exc_info=True)
        return page

def save_pdf(writer, output_path):
    try:
        with open(output_path, 'wb') as f:
            writer.write(f)
        logging.info(f"PDF saved successfully at {output_path}")
    except Exception as e:
        sg.popup_error(f'Error saving PDF:\n{str(e)}')
        logging.error(f"Error saving PDF at {output_path}: {str(e)}", exc_info=True)

def preselect_fields_from_pdf(pdf_path, window):
    try:
        # Extract text or relevant data from the PDF
        text = extract_text_from_page(pdf_path, 0)  # Example: Extract from the first page
        first_product = extract_first_product(text)
        if first_product:
            window['-PRODUCT-'].update(first_product)
            window['-STATUS-'].update('Preselected product from PDF.')
            logging.info(f"Preselected product from PDF: {first_product}")
    except Exception as e:
        sg.popup_error(f'Error preselecting fields from PDF:\n{str(e)}')
        logging.error(f"Error preselecting fields from PDF: {str(e)}", exc_info=True)

def process_packing_slips_and_labels(packing_slips_path, shipping_labels_path, warehouse_map, window):
    try:
        window['-STATUS-'].update('Reading packing slips PDF...')
        window.refresh()

        # Access 'save_path' directly from the window
        save_path = window['-SAVE_PATH-'].get()
        if not save_path:
            sg.popup_error('Please select a directory to save the output files.')
            logging.warning("No save directory selected.")
            return

        # Verify if the save_path exists; if not, create it
        if not os.path.exists(save_path):
            try:
                os.makedirs(save_path)
                logging.info(f"Created directory: {save_path}")
            except Exception as e:
                sg.popup_error(f'Failed to create directory {save_path}:\n{str(e)}')
                logging.error(f"Failed to create directory {save_path}: {str(e)}", exc_info=True)
                return

        # Construct the full file paths using the user-specified save_path
        sorted_packing_slips_path = os.path.join(save_path, 'sorted_packing_slips.pdf')
        sorted_shipping_labels_path = os.path.join(save_path, 'sorted_shipping_labels.pdf')

        logging.info(f"Saving sorted packing slips to: {sorted_packing_slips_path}")
        logging.info(f"Saving sorted shipping labels to: {sorted_shipping_labels_path}")

        # Step 1: Read the packing slips
        packing_slips_reader = PdfReader(packing_slips_path)
        packing_slips_pages = packing_slips_reader.pages

        total_pages = len(packing_slips_pages)
        if total_pages == 0:
            sg.popup_error('The packing slips PDF is empty.')
            logging.warning("Packing slips PDF is empty.")
            return

        # Step 2: Extract product details and collect areas
        window['-STATUS-'].update('Extracting product details...')
        window.refresh()
        pages_to_keep = []
        indices_to_keep = []
        page_areas_list = []
        for idx in range(total_pages):
            text = extract_text_from_page(packing_slips_path, idx)
            product_details = extract_product_details(text, warehouse_map)
            if not product_details:
                logging.debug(f"No product details found on page {idx + 1}")
                continue  # Skip pages with no valid products
            
            # Log all extracted products for this page
            logging.info(f"Page {idx + 1} extracted products: {product_details}")
            
            product_areas = set()
            sample_products_on_page = []
            for product_title, size, quantity in product_details:
                product_title_normalized = product_title.lower().strip()
                area = get_warehouse_area(product_title_normalized, warehouse_map)
                product_areas.add(area)
                
                # Track sample products specifically
                if 'sample' in product_title_normalized:
                    sample_products_on_page.append((product_title, size, quantity, area))
                    logging.info(f"SAMPLE DETECTED on page {idx + 1}: '{product_title}' → area: '{area}'")
            
            if sample_products_on_page:
                logging.info(f"Page {idx + 1} SAMPLE PRODUCTS: {sample_products_on_page}")
            
            # Skip pages with no known areas, UNLESS they contain samples
            has_samples = any('sample' in detail[0].lower() for detail in product_details if detail)
            if not product_areas and not has_samples:
                logging.debug(f"No known areas found on page {idx + 1} and no samples detected")
                continue
            elif not product_areas and has_samples:
                logging.info(f"Page {idx + 1} has no warehouse areas but contains samples - keeping for garage processing")
            elif has_samples:
                logging.info(f"Page {idx + 1} contains samples AND other products with areas: {product_areas}")
            pages_to_keep.append(packing_slips_pages[idx])
            indices_to_keep.append(idx)
            page_areas_list.append(product_areas)
            logging.info(f"Page {idx + 1} marked for keeping with areas: {product_areas}")

        if not pages_to_keep:
            sg.popup_error('No valid products found. This could mean:\n• No products in warehouse map\n• No sample products detected\n• All pages filtered out due to parsing errors')
            logging.warning("No pages to process after filtering.")
            return

        # Step 3: Generate the summary pages
        window['-STATUS-'].update('Generating summary...')
        window.refresh()
        summary_pages = generate_summary(packing_slips_pages, warehouse_map, None, packing_slips_path, indices_to_keep)

        # Step 4: Annotate pages with area identifiers
        window['-STATUS-'].update('Annotating packing slips...')
        window.refresh()
        annotated_packing_slips_pages = []
        area_identifiers = []  # Initialize area_identifiers list
        for idx, (page, product_areas) in enumerate(zip(pages_to_keep, page_areas_list)):
            window['-STATUS-'].update(f'Annotating page {idx + 1}/{len(pages_to_keep)}...')
            window.refresh()
            # Apply the logic for area identifier
            area_identifier = determine_area_identifier(product_areas)
            annotated_page = add_area_identifier_to_page(page, area_identifier)
            annotated_packing_slips_pages.append(annotated_page)
            area_identifiers.append(area_identifier)  # Collect area identifiers for sorting
            logging.debug(f"Annotated page {idx + 1} with area identifier: {area_identifier}")

        # Step 5: Sort the pages by area, then by product title within each area
        window['-STATUS-'].update('Sorting packing slips by area and product...')
        window.refresh()
        
        # Extract product titles for each page for secondary sorting
        page_product_titles = []
        for idx in indices_to_keep:
            text = extract_text_from_page(packing_slips_path, idx)
            product_details = extract_product_details(text, warehouse_map)
            
            # Get the first (primary) product title from the page
            if product_details:
                primary_product = product_details[0][0].lower().strip()  # product_title, size, quantity
                page_product_titles.append(primary_product)
                logging.debug(f"Page {idx + 1} primary product: '{primary_product}'")
            else:
                page_product_titles.append("")  # Fallback for pages without products
                logging.debug(f"Page {idx + 1} has no identifiable products")
        
        # Create enhanced info tuples with product titles
        pages_with_info = list(zip(annotated_packing_slips_pages, area_identifiers, indices_to_keep, page_product_titles))

        # Use the globally defined AREA_SORT_ORDER
        def get_primary_area(area_identifier):
            # Split the area_identifier and get the first area
            primary_area = area_identifier.split(',')[0].strip()
            return primary_area

        def sort_key(item):
            page, area_identifier, index, product_title = item
            primary_area = get_primary_area(area_identifier)
            
            # Primary sort: Area order
            area_sort_value = AREA_SORT_ORDER.index(primary_area) if primary_area in AREA_SORT_ORDER else len(AREA_SORT_ORDER)
            
            # Secondary sort: Product title (alphabetical within each area)
            product_sort_value = product_title
            
            return (area_sort_value, product_sort_value)

        pages_with_info_sorted = sorted(pages_with_info, key=sort_key)
        sorted_pages = [item[0] for item in pages_with_info_sorted]
        sorted_indices = [item[2] for item in pages_with_info_sorted]
        sorted_products = [item[3] for item in pages_with_info_sorted]
        
        # Log the sorting results
        logging.info("📦 SORTING RESULTS BY AREA AND PRODUCT:")
        current_area = None
        for i, (index, area, product) in enumerate(zip(sorted_indices, [get_primary_area(item[1]) for item in pages_with_info_sorted], sorted_products)):
            if area != current_area:
                logging.info(f"   🏢 {area}:")
                current_area = area
            logging.info(f"      • Position {i+1}: Page {index+1} - '{product}'")
        

        logging.info("Pages sorted by area (AREA_SORT_ORDER) and then by product title within each area.")

        # Step 6: Save the sorted packing slips
        window['-STATUS-'].update('Saving sorted packing slips...')
        window.refresh()
        sorted_packing_slips_writer = PdfWriter()

        # Insert the summary pages at the beginning
        if summary_pages:
            for page in summary_pages:
                sorted_packing_slips_writer.add_page(page)
            logging.info("Added summary pages to sorted packing slips.")
        else:
            logging.info("No summary pages to add.")

        for page in sorted_pages:
            sorted_packing_slips_writer.add_page(page)
        save_pdf(sorted_packing_slips_writer, sorted_packing_slips_path)

        # Step 7: Process and sort shipping labels
        window['-STATUS-'].update('Processing shipping labels...')
        window.refresh()
        shipping_labels_reader = PdfReader(shipping_labels_path)
        shipping_labels_pages = shipping_labels_reader.pages

        if len(shipping_labels_pages) != total_pages:
            sg.popup_error('The number of shipping labels does not match the number of packing slips.')
            logging.warning("Mismatch in the number of shipping labels and packing slips.")
            return

        # Skip the time-consuming blank page analysis - we'll detect issues during processing if needed
        logging.info("Processing shipping labels without pre-analysis for better performance")

        # Log detailed information about the filtering and sorting process
        logging.info(f"=== SHIPPING LABELS PROCESSING ===")
        logging.info(f"Total original pages: {total_pages}")
        logging.info(f"Pages kept after filtering: {len(indices_to_keep)}")
        logging.info(f"Indices kept: {indices_to_keep}")
        logging.info(f"Sorted indices: {sorted_indices}")
        
        # Verify all indices are valid
        invalid_indices = [i for i in sorted_indices if i >= len(shipping_labels_pages)]
        if invalid_indices:
            error_msg = f"Invalid shipping label indices found: {invalid_indices}. Max valid index: {len(shipping_labels_pages)-1}"
            logging.error(error_msg)
            sg.popup_error(f"Error processing shipping labels:\n{error_msg}")
            return

        # Extract and sort shipping labels safely
        sorted_shipping_labels_pages = []
        
        for i, original_index in enumerate(sorted_indices):
            if original_index < len(shipping_labels_pages):
                page = shipping_labels_pages[original_index]
                sorted_shipping_labels_pages.append(page)
                
                # We'll skip individual blank page checking for performance
                # If there are issues, they'll be visible in the final output
                logging.debug(f"Added shipping label page {original_index} at position {i}")
            else:
                logging.error(f"Attempted to access shipping label page {original_index}, but only {len(shipping_labels_pages)} pages available")
        
        # Simplified processing - no blank page detection for better performance
        logging.info(f"✅ Successfully processed {len(sorted_shipping_labels_pages)} shipping label pages")
        
        logging.info(f"Successfully extracted {len(sorted_shipping_labels_pages)} shipping label pages")
        
        # --- build the sorted shipping-label PDF --------------------------
        # Use PdfWriter instead of PdfMerger for individual page selection
        sorted_shipping_labels_writer = PdfWriter()
        
        # Read the shipping labels PDF
        with open(shipping_labels_path, 'rb') as labels_file:
            labels_reader = PdfReader(labels_file)
            
            logging.info(f"🔍 SHIPPING LABELS PROCESSING:")
            logging.info(f"   • Shipping labels PDF has {len(labels_reader.pages)} pages")
            logging.info(f"   • Packing slips had {total_pages} original pages") 
            logging.info(f"   • We kept {len(indices_to_keep)} packing slip pages: {indices_to_keep}")
            logging.info(f"   • We're trying to access shipping label pages: {sorted_indices}")
            
            # Add pages in the sorted order with validation
            added_count = 0
            for i, page_index in enumerate(sorted_indices):
                if page_index < len(labels_reader.pages):
                    # Quick check if this page might be blank
                    try:
                        page_text = extract_text_from_page(shipping_labels_path, page_index)
                        is_likely_blank = not page_text or len(page_text.strip()) < 20
                        
                        sorted_shipping_labels_writer.add_page(labels_reader.pages[page_index])
                        added_count += 1
                        
                        if is_likely_blank:
                            logging.warning(f"   ⚠️  Position {i}: Using shipping label page {page_index} which appears BLANK")
                        else:
                            logging.info(f"   ✅ Position {i}: Using shipping label page {page_index} (has content)")
                            
                    except Exception as e:
                        logging.error(f"   ❌ Position {i}: Error processing page {page_index}: {str(e)}")
                        sorted_shipping_labels_writer.add_page(labels_reader.pages[page_index])
                        added_count += 1
                else:
                    logging.error(f"   🚨 Position {i}: Page index {page_index} out of range (max: {len(labels_reader.pages)-1})")
        
        # Write the sorted shipping labels PDF
        with open(sorted_shipping_labels_path, 'wb') as output_file:
            sorted_shipping_labels_writer.write(output_file)
        
        logging.info(f"Successfully created sorted shipping labels PDF with {added_count} pages")
        # Open the files after saving
        open_file_mac(sorted_packing_slips_path)
        open_file_mac(sorted_shipping_labels_path)

        # Notify user processing is complete and ask what to do next
        window['-STATUS-'].update('Processing completed. Choose your next action.')
        
        # Show print dialog with options
        success_message = (
            f"Your packing slips and shipping labels have been sorted and opened in Preview.\n\n"
            f"Sorted Packing Slips: {os.path.basename(sorted_packing_slips_path)}\n"
            f"Sorted Shipping Labels: {os.path.basename(sorted_shipping_labels_path)}\n\n"
            f"Would you like to print them now?"
        )
        
        print_choice = sg.popup_yes_no(
            success_message + "\n\nClick 'Yes' to print now, 'No' to skip printing.",
            title="Processing Complete - Print Options"
        )
        
        if print_choice == 'Yes':
            window['-STATUS-'].update('Printing documents...')
            window.refresh()
            
            # Attempt printing
            try:
                print_pdf(
                    sorted_packing_slips_path,
                    printer_name="YOUR_PACKING_SLIP_PRINTER",
                    paper_size="Letter",
                    scale_to_fit=True,
                    sides="one-sided",
                    reverse_order=PRINT_REVERSE_ORDER,
                    collate=PRINT_COLLATE
                )
                print_pdf(
                    sorted_shipping_labels_path,
                    printer_name="YOUR_LABEL_PRINTER",
                    paper_size="1744907_4_in_x_6_in",
                    scale_to_fit=True,
                    reverse_order=PRINT_REVERSE_ORDER,
                    collate=PRINT_COLLATE
                )
                window['-STATUS-'].update('Printing completed.')
                sg.popup('Print Success', 'Documents have been sent to the printers!')
                logging.info("Documents printed successfully.")
            except Exception as e:
                window['-STATUS-'].update('Printing failed.')
                sg.popup_error(f'Printing failed:\n{str(e)}')
                logging.error(f"Printing failed: {str(e)}")
        else:
            window['-STATUS-'].update('Processing completed. Printing skipped.')
            sg.popup('Complete', 'Processing completed successfully!\n\nYou can print manually from Preview if needed.')
            logging.info("Processing completed successfully. User chose not to print.")
        
        logging.info("Processing workflow completed.")

    except Exception as e:
        sg.popup_error(f'An error occurred during processing:\n{str(e)}')
        logging.error(f"An error occurred during processing: {str(e)}", exc_info=True)
        window['-STATUS-'].update('Error occurred.')

def generate_summary(packing_slips_pages, warehouse_map, _, packing_slips_path, indices_to_keep):
    summary_data = {}

    # Loop through the indices of the pages to keep
    for idx in indices_to_keep:
        # Extract detailed product information (product title, size, quantity)
        text = extract_text_from_page(packing_slips_path, idx)
        product_details = extract_product_details(text, warehouse_map)

        if not product_details:
            logging.debug(f"No product details found on page {idx + 1}")
            continue
        else:
            logging.debug(f"Product details on page {idx + 1}: {product_details}")

        # For each product, get the area from the warehouse map
        for item in product_details:
            if item is None or len(item) != 3:
                logging.debug(f"Invalid item format on page {idx + 1}: {item}")
                continue
            product_title, size, quantity = item
            product_title_normalized = product_title.lower().strip()
            area = get_warehouse_area(product_title_normalized, warehouse_map)

            if area not in summary_data:
                summary_data[area] = []

            summary_data[area].append((product_title, size, quantity))
            
            # Log sample products specifically in summary
            if 'sample' in product_title_normalized:
                logging.info(f"SUMMARY: Added SAMPLE to '{area}': Product='{product_title}', Size='{size}', Quantity='{quantity}'")
            else:
                logging.debug(f"Added to summary: Product='{product_title}', Size='{size}', Quantity='{quantity}', Area='{area}'")

    # Log the summary data for debugging
    logging.info("=== FINAL SUMMARY DATA ===")
    for area, products in summary_data.items():
        if area == 'garage':
            logging.info(f"🚚 GARAGE AREA: {len(products)} products - {products}")
        else:
            logging.info(f"📦 Area {area}: {len(products)} products - {products}")

    # Create PDF pages with the summary
    summary_pages = create_summary_pdf_page(summary_data)

    return summary_pages

def extract_product_details(text, warehouse_map):
    import re
    product_details = []
    lines = text.split('\n')
    i = 0
    acceptable_sizes = {'1', '2', 'S', 'M', 'L', 'SM', 'ML', 'LG', 'Size 1', 'Size 2', 'Sample'}
    acceptable_sizes.update(str(size) for size in range(30, 51))
    stop_phrases = [
        'Please note our return window',
        'Thank you for shopping with us!',
        'If you have any questions',
        'NOTES',
        'SIGNATURE REQUIRED SHIPPING',
        'please visit our returns portal',
        'yourcompany.com'
    ]
    while i < len(lines):
        line = lines[i].strip()
        if 'ITEMS' in line:
            i += 1
            while i < len(lines):
                product_title = lines[i].strip()
                if not product_title or any(stop_phrase in product_title for stop_phrase in stop_phrases):
                    logging.debug(f"Encountered stop phrase or empty line: '{product_title}'")
                    break  # End of items section
                # Normalize product title
                product_title_normalized = product_title.lower().strip()
                # Allow products with "sample" to pass through even if not in warehouse map
                if product_title_normalized not in warehouse_map and 'sample' not in product_title_normalized:
                    # Not a valid product title and not a sample, skip
                    logging.debug(f"Skipping unrecognized product title: '{product_title}'")
                    i += 1
                    continue
                
                # Log when we find sample products
                if 'sample' in product_title_normalized:
                    logging.info(f"Found SAMPLE product: '{product_title}' - will be assigned to garage")
                
                logging.debug(f"Extracted product_title: {product_title}")
                i += 1
                if i >= len(lines):
                    break
                size_line = lines[i].strip()
                logging.debug(f"Extracted size_line: {size_line}")
                # Try to match size and quantity
                size = None
                quantity = None
                # Try pattern: Size Quantity of Total
                match = re.match(r'(\S+(?:\s+\d+)?)\s+(\d+)\s+of\s+\d+', size_line)
                if match:
                    size_candidate = match.group(1)
                    quantity_candidate = match.group(2)
                    if size_candidate in acceptable_sizes:
                        size = size_candidate
                        quantity = quantity_candidate
                else:
                    # Check if size_line is a valid size
                    if size_line in acceptable_sizes:
                        size = size_line
                        quantity = '1'
                if size is None:
                    # For sample products, try to extract size from the product title or assign default
                    if 'sample' in product_title_normalized:
                        # Try to extract size from product title (e.g., "Sample Size 39")
                        import re
                        size_match = re.search(r'size\s+(\d+)', product_title_normalized)
                        if size_match:
                            size = size_match.group(1)
                            quantity = '1'
                            logging.info(f"SAMPLE: Extracted size '{size}' from product title: '{product_title}'")
                        else:
                            # Assign default size for samples if no size found
                            size = 'Sample'
                            quantity = '1'
                            logging.info(f"SAMPLE: Assigned default size 'Sample' for: '{product_title}'")
                    else:
                        logging.debug(f"Invalid or unrecognized size line: '{size_line}', skipping product.")
                        i += 1
                        continue
                logging.debug(f"Extracted size: {size}")
                logging.debug(f"Extracted quantity: {quantity}")
                i += 1
                # Skip SKU line if present
                if i < len(lines) and ('SKU' in lines[i] or re.match(r'^LB', lines[i].strip())):
                    logging.debug(f"Skipping SKU line: '{lines[i].strip()}'")
                    i += 1
                if product_title and size and quantity:
                    product_details.append((product_title, size, quantity))
                    logging.debug(f"Added product_detail: {product_title}, {size}, {quantity}")
                else:
                    logging.debug(f"Could not extract all details for product starting at line {i}")
            break  # Exit after processing ITEMS section
        else:
            i += 1
    return product_details

def create_summary_pdf_page(summary_data):
    from reportlab.lib.pagesizes import letter
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, KeepTogether
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch

    packet = BytesIO()

    # Create a SimpleDocTemplate
    doc = SimpleDocTemplate(packet, pagesize=letter,
                            rightMargin=0.5 * inch, leftMargin=0.5 * inch,
                            topMargin=0.5 * inch, bottomMargin=0.5 * inch)

    elements = []
    styles = getSampleStyleSheet()

    # Custom styles
    heading_style = ParagraphStyle(
        'LockerHeading',
        parent=styles['Heading2'],
        fontSize=8,
        leading=10,
        spaceAfter=4,
    )

    # Iterate through AREA_SORT_ORDER to ensure the summary follows the desired order
    for area in AREA_SORT_ORDER:
        if area in summary_data:
            products = summary_data[area]
            # Sort products by product title (alphabetically) and by size (small to large)
            sorted_products = sorted(products, key=lambda x: (x[0].lower(), size_sort_key(x[1])))

            # Build a list of flowables for this area
            area_flowables = []
            # Title for the area
            area_title = "Garage" if area.lower() == "garage" else f"Locker: {area}"
            area_flowables.append(Paragraph(area_title, heading_style))
            area_flowables.append(Spacer(1, 4))

            # Create table data
            data = [['Product Title', 'Size', 'Quantity']]
            unique_products = {}
            for product_title, size, quantity in sorted_products:
                key = (product_title, size)
                try:
                    quantity_int = int(quantity)
                except ValueError:
                    logging.warning(f"Invalid quantity '{quantity}' for product '{product_title}', size '{size}'. Skipping.")
                    continue  # Skip this entry
                if key in unique_products:
                    unique_products[key] += quantity_int
                else:
                    unique_products[key] = quantity_int

            for (product_title, size), quantity in unique_products.items():
                data.append([product_title, size, str(quantity)])

            if len(data) > 1:  # Only add the table if there is data
                # Create the table
                table = Table(data, colWidths=[3.5 * inch, 0.75 * inch, 0.75 * inch])
                # Add style
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#d3d3d3')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('ALIGN', (1, 1), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 0), (-1, 0), 8),
                    ('FONTSIZE', (0, 1), (-1, -1), 8),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 0),
                    ('TOPPADDING', (0, 0), (-1, 0), 0),
                    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f0f0f0')]),
                    ('GRID', (0, 0), (-1, -1), 0.25, colors.black),
                ]))

                # Allow table to split across pages if needed
                table.hAlign = 'LEFT'
                table.repeatRows = 1  # Repeat the header row on each page
                table_style_splittable = TableStyle([
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ])
                table.setStyle(table_style_splittable)

                area_flowables.append(table)
                area_flowables.append(Spacer(1, 8))  # Add space between lockers

                # Wrap area_flowables in KeepTogether and add to elements
                elements.append(KeepTogether(area_flowables))
            else:
                logging.debug(f"No product data for area {area}, skipping table.")

    if not elements:
        logging.warning("No summary data to generate PDF.")
        return []

    # Build the PDF
    try:
        doc.build(elements)
        logging.info("Summary PDF generated successfully.")
    except Exception as e:
        sg.popup_error(f'Error generating summary PDF:\n{str(e)}')
        logging.error(f"Error generating summary PDF: {str(e)}", exc_info=True)
        return []

    packet.seek(0)
    new_pdf = PdfReader(packet)
    summary_pages = new_pdf.pages

    return summary_pages

def find_matching_shipping_label(packing_slip_path):
    """
    Find a matching shipping label PDF in the same directory as the packing slip.
    Criteria:
    - Not the same file
    - Created within 5 minutes
    - Same number of pages
    """
    directory = os.path.dirname(packing_slip_path)
    packing_slip_time = os.path.getctime(packing_slip_path)
    packing_slip_pages = len(PdfReader(packing_slip_path).pages)

    for filename in os.listdir(directory):
        if filename.lower().endswith('.pdf'):
            file_path = os.path.join(directory, filename)
            if file_path == packing_slip_path:
                continue
            try:
                file_time = os.path.getctime(file_path)
                if abs(file_time - packing_slip_time) > 300:  # 5 minutes = 300 seconds
                    continue
                file_pages = len(PdfReader(file_path).pages)
                if file_pages == packing_slip_pages:
                    return file_path
            except Exception:
                continue
    return ""

def print_pdf(file_path, printer_name, paper_size=None, scale_to_fit=True, sides=None, reverse_order=False, collate=True):
    """
    Print a PDF file using the system's lp command with specific printer settings.
    
    Args:
        file_path: Path to the PDF file
        printer_name: Name of the printer
        paper_size: Paper size (e.g., "Letter", "A4")
        scale_to_fit: Whether to scale the document to fit the page
        sides: Printing sides ("one-sided", "two-sided-long-edge", "two-sided-short-edge")
        reverse_order: If True, prints pages in reverse order
        collate: If True, collates multiple copies (keeps pages in order)
    """
    cmd = [
        "lp",
        "-d", printer_name,
    ]
    if paper_size:
        cmd += ["-o", f"media={paper_size}"]
    if scale_to_fit:
        cmd += ["-o", "fit-to-page"]
    if sides:
        cmd += ["-o", f"sides={sides}"]
    if reverse_order:
        cmd += ["-o", "outputorder=reverse"]
    if collate:
        cmd += ["-o", "Collate=True"]
    
    # Add page order control - this helps ensure correct ordering
    cmd += ["-o", "page-set=all"]
    
    cmd.append(file_path)
    try:
        subprocess.run(cmd, check=True)
        logging.info(f"Sent {file_path} to printer {printer_name} with options: {cmd}")
    except Exception as e:
        logging.error(f"Failed to print {file_path}: {str(e)}")
        sg.popup_error(f"Failed to print {file_path}:\n{str(e)}")

def main():
    # Load the warehouse map from the CSV file
    warehouse_map = load_warehouse_map_dict(csv_file_path)

    if not warehouse_map:
        sg.popup_error('Warehouse map is empty or failed to load.')
        logging.error("Warehouse map is empty or failed to load.")
        return

    # Check for command-line arguments (PDF files)
    args = sys.argv[1:]
    preselected_pdfs = [arg for arg in args if os.path.isfile(arg) and arg.lower().endswith('.pdf')]

    # Initialize variables to track preselected PDFs
    preselected_packing_slips = None
    preselected_shipping_labels = None

    if preselected_pdfs:
        if len(preselected_pdfs) >= 2:
            preselected_packing_slips = preselected_pdfs[0]
            preselected_shipping_labels = preselected_pdfs[1]
        elif len(preselected_pdfs) == 1:
            preselected_packing_slips = preselected_pdfs[0]

    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == 'Exit':
            break

        # Auto-populate fields when Packing Slips PDF is selected
        if event == '-PACKING_SLIPS-' or (event == '__TIMEOUT__' and values['-PACKING_SLIPS-']):
            packing_slip_path = values['-PACKING_SLIPS-']
            if packing_slip_path and os.path.exists(packing_slip_path):
                # Auto-fill Save Directory
                save_dir = os.path.dirname(packing_slip_path)
                window['-SAVE_PATH-'].update(save_dir)
                # Auto-find and fill Shipping Labels PDF
                match = find_matching_shipping_label(packing_slip_path)
                if match:
                    window['-SHIPPING_LABELS-'].update(match)
                else:
                    window['-SHIPPING_LABELS-'].update('')

        if event == 'Process':
            packing_slips_path = values['-PACKING_SLIPS-']
            shipping_labels_path = values['-SHIPPING_LABELS-']
            save_path = values['-SAVE_PATH-']
            # Verify that files are selected
            if not packing_slips_path or not shipping_labels_path:
                sg.popup('Please select both PDFs before proceeding.')
                logging.warning("Process initiated without selecting both PDFs.")
            else:
                window['-STATUS-'].update('Processing started...')
                window.refresh()
                try:
                    process_packing_slips_and_labels(packing_slips_path, shipping_labels_path, warehouse_map, window)
                except Exception as e:
                    sg.popup_error(f'An unexpected error occurred:\n{str(e)}')
                    logging.error(f"Unexpected error during processing: {str(e)}", exc_info=True)
                    window['-STATUS-'].update('Error occurred.')
        elif event == 'Update Warehouse Map':
            # Open the warehouse map update window
            # Hide the main window before opening the update window
            window.hide()
            open_warehouse_map_window()
            # Unhide the main window after the update window is closed
            window.un_hide()
            logging.info("Opened Warehouse Map Update Window.")
        else:
            window['-STATUS-'].update('Ready')

        # Handle preselected PDFs via drag-and-drop
        if preselected_packing_slips:
            window['-PACKING_SLIPS-'].update(preselected_packing_slips)
            preselect_fields_from_pdf(preselected_packing_slips, window)
            preselected_packing_slips = None  # Reset to avoid reprocessing
            sg.popup('Packing Slips PDF loaded via drag-and-drop.')
            logging.info(f"Packing Slips PDF loaded via drag-and-drop: {preselected_packing_slips}")

        if preselected_shipping_labels:
            window['-SHIPPING_LABELS-'].update(preselected_shipping_labels)
            # Implement similar preselection logic for shipping labels if needed
            preselected_shipping_labels = None
            sg.popup('Shipping Labels PDF loaded via drag-and-drop.')
            logging.info(f"Shipping Labels PDF loaded via drag-and-drop: {preselected_shipping_labels}")

    window.close()
    logging.info("Application closed.")

if __name__ == '__main__':
    main()
