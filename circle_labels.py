from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.lib.colors import white, transparent, black
import pandas as pd
import re

# Gem Labels
gem_df = pd.read_excel('Gem_Labels.xlsx')


add_on_labels_oct_24 = pd.read_csv("C:/Users/hank.aungkyaw/Documents/add_on_labels_oct_24.csv")
add_on_blue_sku = add_on_labels_oct_24[add_on_labels_oct_24['Material'] == 'NB']
add_on_ivory_sku = add_on_labels_oct_24[add_on_labels_oct_24['Material'] == 'SS']
add_on_orange_sku = add_on_labels_oct_24[add_on_labels_oct_24['Material'] == 'TI']
add_on_pink_sku = add_on_labels_oct_24[add_on_labels_oct_24['Material'] == 'RG']
add_on_yellow_sku = add_on_labels_oct_24[add_on_labels_oct_24['Material'] == 'YG']
add_on_white_sku = add_on_labels_oct_24[add_on_labels_oct_24['Material'] == 'WG']
add_on_black_sku = add_on_labels_oct_24[add_on_labels_oct_24['Material'] == 'RB']


### v2
# Function to extract color code
def extract_color(name):
    # Pattern to match the specific format with "-mq" suffix
    specific_pattern = r'\d+(\.\d+)?(x\d+)?([A-Za-z]+)-mq'

    match = re.search(specific_pattern, name)
    if match:
        return match.group(3)

    # Fallback pattern for the rest of the cases
    fallback_pattern = r'\d+(\.\d+)?(x\d+)?([A-Za-z]+)'
    match = re.search(fallback_pattern, name)
    return match.group(3) if match else None


# Function to extract size
def extract_size(name):
    match = re.search(r'(\d+(\.\d+)?(x\d+)?|(\d+x\d+))', name)
    return match.group(1) if match else None


def extract_type(name):
    if '-lc' in name:
        return 'Lab Created'
    elif '-ce' in name:
        return 'Ceramic'
    elif '-so' in name:
        return 'FOP'
    elif '-ptz' in name:
        return 'Topaz'
    elif '-ge' in name:
        return 'Genuine'
    else:
        return None

def extract_cut(name):
    if '-Mq' in name:
        return 'Marquise'
    elif '-sp' in name:
        return 'Princess'
    elif '-fb' in name:
        return 'Flatback'
    elif '-hrt' in name:
        return 'Heart'
    elif '-ov' in name:
        return 'Oval'
    elif '-mq' in name:
        return 'Marquise'
    elif '-tpr' in name:
        return 'Tapered'
    elif '-tr' in name:
        return 'Trillion'
    elif '-em' in name:
        return 'Emrald'
    elif '-pe' in name:
        return 'Pear'
    elif '-oct' in name:
        return 'Octagon'
    elif '-bu' in name:
        return 'Bullet'
    else:
        return None

# Apply the function to the 'Name' column and create the 'Color' column
# df['Color'] = df['Name'].apply(extract_color)
# df['Size'] = df['Name'].apply(extract_size)
# df['Type'] = df['Name'].apply(extract_type)
# df['Cut'] = df['Name'].apply(extract_cut)

def format_sku(sku, cut_off_symbol="-", cut_off=2):
    parts = sku.split(cut_off_symbol)
    if len(parts) > cut_off: # <<< len is 2 for Assembly SKU and 5 for Gem short SKU
        return cut_off_symbol.join(parts[:cut_off]) + '\n' + cut_off_symbol.join(parts[cut_off:])
    return sku

### V 2.1
def create_avery_pdf_v2_1(df, column, fontsize=15, cut_off_symbol="-", next_line_cut_off=2, filename="avery_stickers.pdf", border_enabled=0):
    # Avery 5160 label size
    label_width = 1.75 * inch
    label_height = 0.5 * inch

    # Margins and padding
    page_margin_x = 0.3 * inch
    page_margin_y = 0.45 * inch
    label_padding_x = 0.125 * inch
    label_padding_y = 0.125 * inch

    # Spacing between labels
    horizontal_spacing = 0.3 * inch
    vertical_spacing = 0 * inch

    # Number of labels per row and column
    labels_per_row = 4
    labels_per_column = 20

    # Font size
    font_size = fontsize  # Change to desired font size

    # Create canvas
    c = canvas.Canvas(filename, pagesize=letter)

    # Draw labels
    for i, sku in enumerate(df[column]):
        formatted_sku = format_sku(sku, cut_off_symbol, next_line_cut_off)

        if i % (labels_per_row * labels_per_column) == 0 and i != 0:
            c.showPage()  # Create a new page
            c.setFont("Helvetica", font_size)  # Reset font size on new page

        c.setFont("Helvetica", font_size)  # Set font size before drawing

        col = i % labels_per_row
        row = (i // labels_per_row) % labels_per_column

        # Calculate label position with spacing
        x = page_margin_x + col * (label_width + horizontal_spacing)
        y = letter[1] - page_margin_y - (row + 1) * (label_height + vertical_spacing)

        # Draw the rectangle for the label
        if border_enabled:
            c.setStrokeColor(black)  # Set border color
        else:
            c.setStrokeColor(transparent)  # Disable the border color
        c.setFillColor(white)  # Set the fill color of the rectangle
        c.rect(x, y, label_width, label_height, fill=1)  # Draw filled rectangle

        # Set text color
        c.setFillColor(black)  # Set text color (black)

        # Calculate the width and height of the text
        lines = formatted_sku.split('\n')
        text_height = font_size * len(lines)  # Use the font size for text height calculation

        # Calculate the y positions to center the text
        for j, line in enumerate(lines):
            text_width = c.stringWidth(line, "Helvetica", font_size)
            text_x = x + (label_width - text_width) / 2
            text_y = y + (label_height - text_height) / 2 + (len(lines) - j - 1) * font_size
            c.drawString(text_x, text_y, line)

    c.save()  # Save the PDF file



create_avery_pdf_v2_1(add_on_blue_sku, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="With Borders/Add_on_Blue.pdf", border_enabled=1)
create_avery_pdf_v2_1(add_on_ivory_sku, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="With Borders/Add_on_Ivory.pdf", border_enabled=1)
create_avery_pdf_v2_1(add_on_orange_sku, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="With Borders/Add_on_Orange.pdf", border_enabled=1)
create_avery_pdf_v2_1(add_on_pink_sku, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="With Borders/Add_on_Pink.pdf", border_enabled=1)
create_avery_pdf_v2_1(add_on_white_sku, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="With Borders/Add_on_White.pdf", border_enabled=1)
create_avery_pdf_v2_1(add_on_yellow_sku, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="With Borders/Add_on_Yellow.pdf", border_enabled=1)
#
create_avery_pdf_v2_1(add_on_blue_sku, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="Without Borders/Add_on_Blue.pdf", border_enabled=0)
create_avery_pdf_v2_1(add_on_ivory_sku, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="Without Borders/Add_on_Ivory.pdf", border_enabled=0)
create_avery_pdf_v2_1(add_on_orange_sku, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="Without Borders/Add_on_Orange.pdf", border_enabled=0)
create_avery_pdf_v2_1(add_on_pink_sku, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="Without Borders/Add_on_Pink.pdf", border_enabled=0)
create_avery_pdf_v2_1(add_on_white_sku, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="Without Borders/Add_on_White.pdf", border_enabled=0)
create_avery_pdf_v2_1(add_on_yellow_sku, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="Without Borders/Add_on_Yellow.pdf", border_enabled=0)

