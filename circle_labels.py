from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.lib.colors import white, transparent, black
import pandas as pd
import re

# Gem Labels
df = pd.read_excel('Gem_Labels.xlsx')

cab_df = df[df['Name'].str.startswith('cab-')].sort_values(['Color', 'Size'], ascending=True)
faceted_df = df[df['Name'].str.startswith('faceted-')].sort_values(['Color', 'Size'], ascending=True)
orb_df = df[df['Name'].str.startswith('orb-')].sort_values(['Color', 'Size'], ascending=True)

print(cab_df)
print(faceted_df)
print(orb_df)

def format_sku(sku, cut_off_symbol="-", cut_off=2):
    parts = sku.split(cut_off_symbol)
    if len(parts) > cut_off: # <<< len is 2 for Assembly SKU and 5 for Gem short SKU
        return cut_off_symbol.join(parts[:cut_off]) + '\n' + cut_off_symbol.join(parts[cut_off:])
    return sku

### V 2.1
def create_avery_pdf_v2_1(df, column, fontsize=15, cut_off_symbol="-", next_line_cut_off=2, filename="avery_stickers.pdf", border_enabled=0):
    # Avery 5160 label size
    label_width = 0.75 * inch
    label_height = 0.75 * inch

    # Margins and padding
    page_margin_x = 0.37 * inch
    page_margin_y = 0.37 * inch
    label_padding_x = 0.1 * inch
    label_padding_y = 0.1 * inch

    # Spacing between labels
    horizontal_spacing = 0.25 * inch
    vertical_spacing = 0.25 * inch

    # Number of labels per row and column
    labels_per_row = 8
    labels_per_column = 10

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



######## For Gem SKU ##############


create_avery_pdf_v2_1(cab_df, "Name", fontsize=12, cut_off_symbol="-", next_line_cut_off=2, filename="With Borders/Circle_Cab.pdf", border_enabled=1)
