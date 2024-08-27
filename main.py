from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.lib.colors import white, transparent, black
import pandas as pd
import re

# Dataframe for Tip Out Bin
df = pd.read_excel("C:/Users/hank.aungkyaw/Documents/Ana 3 Mo Sales.xlsx", sheet_name="Labels ")
subset_df = df[['SKU', 'Area','Label', 'Color']].copy()
subset_df = subset_df[subset_df['Area'] == "Fulfillment"]
blue_sku = subset_df[subset_df['Color'] == 'Blue']
ivory_sku = subset_df[subset_df['Color'] == 'Ivory ']
ivory1_sku = subset_df[subset_df['Color'] == 'Ivory 1']
orange_sku = subset_df[subset_df['Color'] == 'Orange']
orange1_sku = subset_df[subset_df['Color'] == 'Orange 1']
pink_sku = subset_df[subset_df['Color'] == 'Pink']
pink1_sku = subset_df[subset_df['Color'] == 'Pink 1']
white_sku = subset_df[subset_df['Color'] == 'White']
white1_sku = subset_df[subset_df['Color'] == 'White 1']
yellow_sku = subset_df[subset_df['Color'] == 'Yellow']
yellow1_sku = subset_df[subset_df['Color'] == 'Yellow 1']

# Dataframe for Gem Bins
# df = pd.read_excel("C:/Users/hank.aungkyaw/Documents/gold_subassembly.xlsx")

# Gem Labels
gem_df = pd.read_excel('Gem_Labels.xlsx')



# # Function to extract color code after .0, .5, or other variations
# ## v1
# def extract_color(name):
#     match = re.search(r'\d+(\.\d+)?(x\d+)?([A-Za-z].*)', name)
#     return match.group(3) if match else None

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

### V 1.0
# def create_avery_pdf(df, column, filename="avery_stickers.pdf"):
#     # Avery 5160 label size
#     label_width = 1.75 * inch
#     label_height = 0.5 * inch
#
#     # Margins and padding
#     page_margin_x = 0.3 * inch
#     page_margin_y = 0.5 * inch
#     label_padding_x = 0.125 * inch
#     label_padding_y = 0.125 * inch
#
#     # Spacing between labels
#     horizontal_spacing = 0.3 * inch
#     vertical_spacing = 0 * inch
#
#     # Number of labels per row and column
#     labels_per_row = 4
#     labels_per_column = 20
#
#     # Create canvas
#     c = canvas.Canvas(filename, pagesize=letter)
#
#     # Set font size
#     font_size = 15 ## default is 12 chanage back after printing gems
#     c.setFont("Helvetica", font_size)
#
#     # Draw labels
#     for i, sku in enumerate(df[column]):
#         formatted_sku = format_sku(sku)
#
#         col = i % labels_per_row
#         row = (i // labels_per_row) % labels_per_column
#
#         # Calculate label position with spacing
#         x = page_margin_x + col * (label_width + horizontal_spacing)
#         y = letter[1] - page_margin_y - (row + 1) * (label_height + vertical_spacing)
#
#         # Draw the rectangle for the label
#         c.rect(x, y, label_width, label_height)
#
#         # Calculate the width and height of the text
#         lines = formatted_sku.split('\n')
#         text_height = font_size * len(lines)
#
#         # Calculate the y positions to center the text
#         for j, line in enumerate(lines):
#             text_width = c.stringWidth(line, "Helvetica", font_size)
#             text_x = x + (label_width - text_width) / 2
#             text_y = y + (label_height - text_height) / 2 + (len(lines) - j - 1) * font_size
#             c.drawString(text_x, text_y, line)
#
#         # Start a new page if needed
#         if (i + 1) % (labels_per_row * labels_per_column) == 0:
#             c.showPage()
#
#     c.save()


### V 2.0
# def create_avery_pdf_v2(df, column, next_line_cut_off, filename="avery_stickers.pdf"):
#     # Avery 5160 label size
#     label_width = 1.75 * inch
#     label_height = 0.5 * inch
#
#     # Margins and padding
#     page_margin_x = 0.3 * inch
#     page_margin_y = 0.5 * inch
#     label_padding_x = 0.125 * inch
#     label_padding_y = 0.125 * inch
#
#     # Spacing between labels
#     horizontal_spacing = 0.3 * inch
#     vertical_spacing = 0 * inch
#
#     # Number of labels per row and column
#     labels_per_row = 4
#     labels_per_column = 20
#
#     # Font size
#     font_size = 15  # Change to desired font size
#
#     # Create canvas
#     c = canvas.Canvas(filename, pagesize=letter)
#
#     # Draw labels
#     for i, sku in enumerate(df[column]):
#         formatted_sku = format_sku(sku, next_line_cut_off)
#
#         if i % (labels_per_row * labels_per_column) == 0 and i != 0:
#             c.showPage()  # Create a new page
#             # No need to set font size again here, will set before drawing
#
#         c.setFont("Helvetica", font_size)  # Set font size before drawing
#
#         col = i % labels_per_row
#         row = (i // labels_per_row) % labels_per_column
#
#         # Calculate label position with spacing
#         x = page_margin_x + col * (label_width + horizontal_spacing)
#         y = letter[1] - page_margin_y - (row + 1) * (label_height + vertical_spacing)
#
#         # Draw the rectangle for the label
#         c.rect(x, y, label_width, label_height)
#
#         # Calculate the width and height of the text
#         lines = formatted_sku.split('\n')
#         text_height = font_size * len(lines)  # Use the font size for text height calculation
#
#         # Calculate the y positions to center the text
#         for j, line in enumerate(lines):
#             text_width = c.stringWidth(line, "Helvetica", font_size)
#             text_x = x + (label_width - text_width) / 2
#             text_y = y + (label_height - text_height) / 2 + (len(lines) - j - 1) * font_size
#             c.drawString(text_x, text_y, line)
#
#     c.save()  # Save the PDF file

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

# Generate the PDF
# print(yellow1_sku)
# create_avery_pdf_v2_1(blue_sku, "SKU", fontsize=12, next_line_cut_off=2, filename="With Borders/BORDER_Blue.pdf", border_enabled=1)
# create_avery_pdf_v2_1(ivory_sku, "SKU", fontsize=12, next_line_cut_off=2, filename="With Borders/BORDER_Ivory.pdf", border_enabled=1)
# create_avery_pdf_v2_1(ivory1_sku, "SKU", fontsize=11, cut_off_symbol="+", next_line_cut_off=1, filename="With Borders/BORDER_Ivory1.pdf", border_enabled=1)
# create_avery_pdf_v2_1(orange_sku, "SKU", fontsize=12, next_line_cut_off=2, filename="With Borders/BORDER_Orange.pdf", border_enabled=1)
# create_avery_pdf_v2_1(orange1_sku, "SKU", fontsize=9, cut_off_symbol="+", next_line_cut_off=1, filename="With Borders/BORDER_Orange1.pdf", border_enabled=1)
# create_avery_pdf_v2_1(pink_sku, "SKU", fontsize=12, next_line_cut_off=2, filename="With Borders/BORDER_Pink.pdf", border_enabled=1)
# create_avery_pdf_v2_1(pink1_sku, "SKU", fontsize=10, cut_off_symbol="+", next_line_cut_off=1, filename="With Borders/BORDER_Pink1.pdf", border_enabled=1)
# create_avery_pdf_v2_1(white_sku, "SKU", fontsize=12, next_line_cut_off=2, filename="With Borders/BORDER_White.pdf", border_enabled=1)
# create_avery_pdf_v2_1(white1_sku, "SKU", fontsize=10, cut_off_symbol="+", next_line_cut_off=1, filename="With Borders/BORDER_White1.pdf", border_enabled=1)
# create_avery_pdf_v2_1(yellow_sku, "SKU", fontsize=12, next_line_cut_off=2, filename="With Borders/BORDER_Yellow.pdf", border_enabled=1)
# create_avery_pdf_v2_1(yellow1_sku, "SKU", fontsize=9, cut_off_symbol="+", next_line_cut_off=1, filename="With Borders/BORDER_Yellow1.pdf", border_enabled=1)
#
# create_avery_pdf_v2_1(blue_sku, "SKU", fontsize=12, next_line_cut_off=2, filename="Without Borders/Blue.pdf", border_enabled=0)
# create_avery_pdf_v2_1(ivory_sku, "SKU", fontsize=12, next_line_cut_off=2, filename="Without Borders/Ivory.pdf", border_enabled=0)
# create_avery_pdf_v2_1(ivory1_sku, "SKU", fontsize=11, cut_off_symbol="+", next_line_cut_off=1, filename="Without Borders/Ivory1.pdf", border_enabled=0)
# create_avery_pdf_v2_1(orange_sku, "SKU", fontsize=12, next_line_cut_off=2, filename="Without Borders/Orange.pdf", border_enabled=0)
# create_avery_pdf_v2_1(orange1_sku, "SKU", fontsize=9, cut_off_symbol="+", next_line_cut_off=1, filename="Without Borders/Orange1.pdf", border_enabled=0)
# create_avery_pdf_v2_1(pink_sku, "SKU", fontsize=12, next_line_cut_off=2, filename="Without Borders/Pink.pdf", border_enabled=0)
# create_avery_pdf_v2_1(pink1_sku, "SKU", fontsize=10, cut_off_symbol="+", next_line_cut_off=1, filename="Without Borders/Pink1.pdf", border_enabled=0)
# create_avery_pdf_v2_1(white_sku, "SKU", fontsize=12, next_line_cut_off=2, filename="Without Borders/White.pdf", border_enabled=0)
# create_avery_pdf_v2_1(white1_sku, "SKU", fontsize=10, cut_off_symbol="+", next_line_cut_off=1, filename="Without Borders/White1.pdf", border_enabled=0)
# create_avery_pdf_v2_1(yellow_sku, "SKU", fontsize=12, next_line_cut_off=2, filename="Without Borders/Yellow.pdf", border_enabled=0)
# create_avery_pdf_v2_1(yellow1_sku, "SKU", fontsize=9, cut_off_symbol="+", next_line_cut_off=1, filename="Without Borders/Yellow1.pdf", border_enabled=0)

# pink_df = df[df["Color"] == "Pink"]
# white_df = df[df["Color"] == "White"]
# yellow_df = df[df["Color"] == "Yellow"]
# ivory_df = df[df["Color"] == "Ivory "]
# orange_df = df[df["Color"] == "Orange"]
# print(pink_df)
# create_avery_pdf_v2_1(pink_df, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="gold assembly/border_pink.pdf", border_enabled=1)
# create_avery_pdf_v2_1(white_df, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="gold assembly/border_white.pdf", border_enabled=1)
# create_avery_pdf_v2_1(yellow_df, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="gold assembly/border_yellow.pdf", border_enabled=1)
# create_avery_pdf_v2_1(ivory_df, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="gold assembly/border_ivory.pdf", border_enabled=1)
# create_avery_pdf_v2_1(orange_df, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="gold assembly/border_orange.pdf", border_enabled=1)
#
# create_avery_pdf_v2_1(pink_df, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="gold assembly/noborder_pink.pdf", border_enabled=0)
# create_avery_pdf_v2_1(white_df, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="gold assembly/noborder_white.pdf", border_enabled=0)
# create_avery_pdf_v2_1(yellow_df, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="gold assembly/noborder_yellow.pdf", border_enabled=0)
# create_avery_pdf_v2_1(ivory_df, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="gold assembly/noborder_ivory.pdf", border_enabled=0)
# create_avery_pdf_v2_1(orange_df, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="gold assembly/noborder_orange.pdf", border_enabled=0)


# create_avery_pdf_v2_1(blue_sku, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="With Borders/Blue.pdf", border_enabled=1)
# create_avery_pdf_v2_1(ivory_sku, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="With Borders/Ivory.pdf", border_enabled=1)
# create_avery_pdf_v2_1(orange_sku, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="With Borders/Orange.pdf", border_enabled=1)
# create_avery_pdf_v2_1(pink_sku, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="With Borders/Pink.pdf", border_enabled=1)
# create_avery_pdf_v2_1(white_sku, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="With Borders/White.pdf", border_enabled=1)
# create_avery_pdf_v2_1(yellow_sku, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="With Borders/Yellow.pdf", border_enabled=1)
# create_avery_pdf_v2_1(ivory1_sku, "Label", fontsize=12, cut_off_symbol="+", next_line_cut_off=1, filename="With Borders/Ivory1.pdf", border_enabled=1)
# create_avery_pdf_v2_1(orange1_sku, "Label", fontsize=12, cut_off_symbol="*", next_line_cut_off=1, filename="With Borders/Orange1.pdf", border_enabled=1)
# create_avery_pdf_v2_1(pink1_sku, "Label", fontsize=12, cut_off_symbol="+", next_line_cut_off=1, filename="With Borders/Pink1.pdf", border_enabled=1)
# create_avery_pdf_v2_1(white1_sku, "Label", fontsize=12, cut_off_symbol="+", next_line_cut_off=1, filename="With Borders/White1.pdf", border_enabled=1)
# create_avery_pdf_v2_1(yellow1_sku, "Label", fontsize=12, cut_off_symbol="+", next_line_cut_off=1, filename="With Borders/Yellow1.pdf", border_enabled=1)
# #
# create_avery_pdf_v2_1(blue_sku, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="Without Borders/Blue.pdf", border_enabled=0)
# create_avery_pdf_v2_1(ivory_sku, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="Without Borders/Ivory.pdf", border_enabled=0)
# create_avery_pdf_v2_1(orange_sku, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="Without Borders/Orange.pdf", border_enabled=0)
# create_avery_pdf_v2_1(pink_sku, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="Without Borders/Pink.pdf", border_enabled=0)
# create_avery_pdf_v2_1(white_sku, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="Without Borders/White.pdf", border_enabled=0)
# create_avery_pdf_v2_1(yellow_sku, "Label", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="Without Borders/Yellow.pdf", border_enabled=0)
# create_avery_pdf_v2_1(ivory1_sku, "Label", fontsize=12, cut_off_symbol="+", next_line_cut_off=1, filename="Without Borders/Ivory1.pdf", border_enabled=0)
# create_avery_pdf_v2_1(orange1_sku, "Label", fontsize=12, cut_off_symbol="*", next_line_cut_off=1, filename="Without Borders/Orange1.pdf", border_enabled=0)
# create_avery_pdf_v2_1(pink1_sku, "Label", fontsize=12, cut_off_symbol="+", next_line_cut_off=1, filename="Without Borders/Pink1.pdf", border_enabled=0)
# create_avery_pdf_v2_1(white1_sku, "Label", fontsize=12, cut_off_symbol="+", next_line_cut_off=1, filename="Without Borders/White1.pdf", border_enabled=0)
# create_avery_pdf_v2_1(yellow1_sku, "Label", fontsize=12, cut_off_symbol="+", next_line_cut_off=1, filename="Without Borders/Yellow1.pdf", border_enabled=0)

create_avery_pdf_v2_1(gem_df, "Name", fontsize=15, cut_off_symbol="-", next_line_cut_off=1, filename="gem_labels/genuine_noborder.pdf", border_enabled=0)

######### For Gem short SKU ##############
#
# ## parse and export faceted synthetic gems
# faceted_syn = df[(~df['Name'].str.contains('(?i)-ge')) & (df['Name'].str.startswith('faceted-'))]
# faceted_syn.loc[:,'Color'] = faceted_syn['Name'].apply(extract_color)
# faceted_syn.loc[:,'Type'] = faceted_syn['Name'].apply(extract_type)
# faceted_syn.loc[:,'Cut'] = faceted_syn['Name'].apply(extract_cut)
# faceted_syn = faceted_syn.sort_values(['Color', 'Size'], ascending=True)
# # faceted_syn.to_csv('faceted_syn.csv', index=False)
# # create_avery_pdf(faceted_syn, 'Name',filename="faceted_syn.pdf")
# faceted_syn['Name'] = faceted_syn['Name'].str.replace('faceted-','')
# create_avery_pdf_v2_1(faceted_syn, 'Name', 15m filename="faceted_syn_short.pdf", border_enabled=1)
#
# ### parse and export genuine gems
# faceted_gen = df[(df['Name'].str.contains('(?i)-ge')) & (df['Name'].str.startswith('faceted-'))]
# faceted_gen.loc[:,'Color'] = faceted_gen['Name'].apply(extract_color)
# faceted_gen.loc[:,'Type'] = faceted_gen['Name'].apply(extract_type)
# faceted_gen.loc[:,'Cut'] = faceted_gen['Name'].apply(extract_cut)
# faceted_gen = faceted_gen.sort_values(['Color', 'Size'], ascending=True)
# # faceted_gen.to_csv('faceted_gen.csv', index=False)
# # create_avery_pdf(faceted_gen, 'Name',filename="faceted_gen.pdf")
# faceted_gen['Name'] = faceted_gen['Name'].str.replace('faceted-','')
# create_avery_pdf_v2_1(faceted_gen, 'Name',filename="faceted_gen_short.pdf", border_enabled=1)
#
# ## parse and export cab synthetic gems
# cab_syn = df[(~df['Name'].str.contains('(?i)-ge')) & (df['Name'].str.startswith('cab-'))]
# cab_syn.loc[:,'Color'] = cab_syn['Name'].apply(extract_color)
# cab_syn.loc[:,'Type'] = cab_syn['Name'].apply(extract_type)
# cab_syn.loc[:,'Cut'] = cab_syn['Name'].apply(extract_cut)
# cab_syn = cab_syn.sort_values(['Color', 'Size'], ascending=True)
# # cab_syn.to_csv('cab_syn.csv', index=False)
# # create_avery_pdf(cab_syn, 'Name',filename="cab_syn.pdf")
# cab_syn['Name'] = cab_syn['Name'].str.replace('cab-','')
# create_avery_pdf_v2_1(cab_syn, 'Name',filename="cab_syn_short.pdf", border_enabled=1)
#
# ### parse and export cab genuine gems
# cab_gen = df[(df['Name'].str.contains('(?i)-ge')) & (df['Name'].str.startswith('cab-'))]
# cab_gen.loc[:,'Color'] = cab_gen['Name'].apply(extract_color)
# cab_gen.loc[:,'Type'] = cab_gen['Name'].apply(extract_type)
# cab_gen.loc[:,'Cut'] = cab_gen['Name'].apply(extract_cut)
# cab_gen = cab_gen.sort_values(['Color', 'Size'], ascending=True)
# # cab_gen.to_csv('cab_gen.csv', index=False)
# # create_avery_pdf(cab_gen, 'Name',filename="cab_gen.pdf")
# cab_gen['Name'] = cab_gen['Name'].str.replace('cab-','')
# create_avery_pdf_v2_1(cab_gen, 'Name',filename="cab_gen_short.pdf", border_enabled=1)
#
# ## parse and export orb synthetic gems
# orb_syn = df[(~df['Name'].str.contains('(?i)-ge')) & (df['Name'].str.startswith('orb-'))]
# orb_syn.loc[:,'Color'] = orb_syn['Name'].apply(extract_color)
# orb_syn.loc[:,'Type'] = orb_syn['Name'].apply(extract_type)
# orb_syn.loc[:,'Cut'] = orb_syn['Name'].apply(extract_cut)
# orb_syn = orb_syn.sort_values(['Color', 'Size'], ascending=True)
# # orb_syn.to_csv('orb_syn.csv', index=False)
# # create_avery_pdf(orb_syn, 'Name',filename="orb_syn.pdf")
# orb_syn['Name'] = orb_syn['Name'].str.replace('orb-','')
# create_avery_pdf_v2_1(orb_syn, 'Name',filename="orb_syn_short.pdf", border_enabled=1)
#
#
# ### parse and export orb genuine gems
# orb_gen = df[(df['Name'].str.contains('(?i)-ge')) & (df['Name'].str.startswith('orb-'))]
# orb_gen.loc[:,'Color'] = orb_gen['Name'].apply(extract_color)
# orb_gen.loc[:,'Type'] = orb_gen['Name'].apply(extract_type)
# orb_gen.loc[:,'Cut'] = orb_gen['Name'].apply(extract_cut)
# orb_gen = orb_gen.sort_values(['Color', 'Size'], ascending=True)
# # orb_gen.to_csv('orb_gen.csv', index=False)
# # create_avery_pdf(orb_gen, 'Name',filename="orb_gen.pdf")
# orb_gen['Name'] = orb_gen['Name'].str.replace('orb-','')
# create_avery_pdf_v2_1(orb_gen, 'Name',filename="orb_gen_short.pdf", border_enabled=1)

