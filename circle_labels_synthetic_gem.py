from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.lib.colors import white, transparent, black
import pandas as pd
import re

# Gem Labels
# df = pd.read_excel('Gem_Labels.xlsx')

# Use NS SKU
df = pd.read_excel('ALL_GEM_SKU_241021.xlsx')

# Add on SKU
# df = pd.read_excel('henry_sku_request_list.xlsx')


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


df['Color'] = df['Name'].apply(extract_color)
# cab_df = df[df['Name'].str.startswith('cab-')].sort_values(['Color', 'Size'], ascending=True)
# faceted_df = df[df['Name'].str.startswith('faceted-')].sort_values(['Color', 'Size'], ascending=True)

# orb_df = df[df['Name'].str.startswith('orb-')].sort_values(['Color', 'Size'], ascending=True)
# topaz_df = df[df['Name'].str.contains('-ptz')].sort_values(['Name'], ascending=True)
cab_df = df[df['Name'].str.startswith('cab-')].sort_values(['Name'], ascending=True)
faceted_df = df[(df['Name'].str.startswith('faceted-')) & (~df['Name'].str.contains('-ptz'))].sort_values(['Color','Name'], ascending=True)
faceted_df['Name'] = faceted_df['Name'].replace("faceted", "facet", regex=True)
print(cab_df)
print(faceted_df)


# Define a function to extract the text between 'facet-' and '-ge'
def extract_length(sku):
    ### for synthetic
    if 'facet-' in sku and '-ge' not in sku and '-ptz' not in sku and '-Ge' not in sku and 'HSIge' not in sku and 'RUge' not in sku:
        # Split and extract the relevant part
        return sku.split('facet-')[1].split('-ge')[0]

    return ''

    # ### for topaz
    # if 'facet-' in sku and '-ptz' in sku:
    #     # Split and extract the relevant part
    #     return sku.split('facet-')[1].split('-ptz')[0]
    # return ''

def wrap_text(sku, max_line_length=8):
    words = re.findall(r'[^\s-]+|[-]', sku)  # Capture words and dashes
    wrapped_text = ''
    current_line = ''

    for word in words:
        # Check if adding the word would exceed the maximum line length
        if len(current_line) + len(word) + 1 <= max_line_length:
            current_line += word + ''
        else:
            wrapped_text += current_line.strip() + '' + '\n'  # Move to the next line
            current_line = word + ''

    wrapped_text += current_line.strip()  # Add the remaining text

    return wrapped_text



### V 2.1 - Modified for Circular Labels
def create_circular_pdf(df, column, fontsize=15, max_line_length=8,
                       filename="circular_labels_syn.pdf", border_enabled=0):
    label_diameter = 0.75 * inch  # Diameter of the circle
    radius = label_diameter / 2

    page_margin_x = 0.37 * inch
    page_margin_y = 0.61 * inch
    label_padding_x = 0.1 * inch
    label_padding_y = 0.12 * inch

    horizontal_spacing = 0.25 * inch
    vertical_spacing = 0.25 * inch

    labels_per_row = 8
    labels_per_column = 10

    c = canvas.Canvas(filename, pagesize=letter)

    for i, sku in enumerate(df[column]):
        wrapped_sku = wrap_text(sku, max_line_length)

        if i % (labels_per_row * labels_per_column) == 0 and i != 0:
            c.showPage()
            c.setFont("Helvetica", fontsize)

        c.setFont("Helvetica", fontsize)

        col = i % labels_per_row
        row = (i // labels_per_row) % labels_per_column

        x_center = page_margin_x + radius + col * (label_diameter + horizontal_spacing)
        y_center = letter[1] - page_margin_y - radius - row * (label_diameter + vertical_spacing)

        if border_enabled:
            c.setStrokeColor(black)
            c.setLineWidth(1)
        else:
            c.setStrokeColor(transparent)
            c.setLineWidth(0)

        c.setFillColor(white)
        c.circle(x_center, y_center, radius, stroke=border_enabled, fill=1)

        c.setFillColor(black)

        lines = wrapped_sku.split('\n')
        current_font_size = fontsize
        max_text_width = label_diameter - 2 * label_padding_x
        max_text_height = label_diameter - 2 * label_padding_y
        text_height = current_font_size * len(lines)

        while text_height > max_text_height and current_font_size > 6:
            current_font_size -= 1
            c.setFont("Helvetica", current_font_size)
            text_height = current_font_size * len(lines)

        if current_font_size != fontsize:
            c.setFont("Helvetica", current_font_size)

        text_y_start = y_center + (max_text_height / 2) - ((len(lines) - 1) * current_font_size / 2)

        for j, line in enumerate(lines):
            text_width = c.stringWidth(line, "Helvetica", current_font_size)
            text_x = x_center - (text_width / 2)
            text_y = text_y_start - j * current_font_size
            c.drawString(text_x, text_y, line)

    c.save()


### FOR TOPAZ BELOW
create_circular_pdf(
    faceted_df,
    "Name",
    fontsize=12,
    max_line_length=9,
    filename="With Borders/circular_labels_syn_faceted.pdf",
    border_enabled=1
)
#
# create_circular_pdf(
#     topaz_df,
#     "Name",
#     fontsize=12,
#     max_line_length=9,
#     filename="Without Borders/Circle_Topaz.pdf",
#     border_enabled=0
# )


### FOR GENUINE BELOW
######## For Gem SKU ##############

# create_circular_pdf(
#     cab_df,
#     "Name",
#     fontsize=12,
#     max_line_length=7,
#     filename="With Borders/Circle_Cab_Add_on.pdf",
#     border_enabled=1
# )

### split long and short df for faceted SKUs
# Create a new column to store the extracted text
faceted_df['extracted'] = faceted_df['Name'].apply(extract_length)
print(faceted_df)

# Create two DataFrames based on the length of the extracted text
short_df = faceted_df[faceted_df['extracted'].str.len() < 7]
long_df = faceted_df[faceted_df['extracted'].str.len() >= 7]

# Drop the temporary extracted column if not needed
short_df = short_df.drop(columns=['extracted'])
long_df = long_df.drop(columns=['extracted'])
print(short_df)
print(long_df)
#
# ########### WITH BORDER ##############
create_circular_pdf(
    short_df,
    "Name",
    fontsize=12,
    max_line_length=7,
    filename="With Borders/circular_labels_syn_short.pdf",
    border_enabled=1
)
#
# create_circular_pdf(
#     short_df,
#     "Name",
#     fontsize=12,
#     max_line_length=7,
#     filename="With Borders/Circle_Topaz_Short.pdf",
#     border_enabled=1
# )
#
create_circular_pdf(
    long_df,
    "Name",
    fontsize=12,
    max_line_length=8,
    filename="With Borders/circular_labels_syn_long.pdf",
    border_enabled=1
)
#
create_circular_pdf(
    cab_df,
    "Name",
    fontsize=12,
    max_line_length=8,
    filename="With Borders/circular_labels_syn_cab.pdf",
    border_enabled=1
)
#
#
# ########### WITHOUT BORDER ##############
#


create_circular_pdf(
    cab_df,
    "Name",
    fontsize=12,
    max_line_length=7,
    filename="Without Borders/circular_labels_syn_cab.pdf",
    border_enabled=0
)

create_circular_pdf(
    short_df,
    "Name",
    fontsize=12,
    max_line_length=7,
    filename="Without Borders/circular_labels_syn_short.pdf",
    border_enabled=0
)

create_circular_pdf(
    long_df,
    "Name",
    fontsize=10,
    max_line_length=8,
    filename="Without Borders/circular_labels_syn_long.pdf",
    border_enabled=0
)
#
# create_circular_pdf(
#     orb_df,
#     "Name",
#     fontsize=12,
#     max_line_length=8,
#     filename="Without Borders/Circle_Orb.pdf",
#     border_enabled=0
# )