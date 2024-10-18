from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.lib.colors import white, transparent, black
import qrcode
import pandas as pd
import re
import os
import tempfile

# Gem Labels
df = pd.read_excel('Gem_Labels.xlsx')

cab_df = df[df['Name'].str.startswith('cab-')].sort_values(['Color', 'Size'], ascending=True)
faceted_df = df[df['Name'].str.startswith('faceted-')].sort_values(['Color', 'Size'], ascending=True)
orb_df = df[df['Name'].str.startswith('orb-')].sort_values(['Color', 'Size'], ascending=True)


# Define a function to extract the text between 'facet-' and '-ge'
def extract_length(sku):
    if 'facet-' in sku and '-ge' in sku:
        return sku.split('facet-')[1].split('-ge')[0]
    return ''


def wrap_text(sku, max_line_length=8):
    words = re.findall(r'[^\s-]+|[-]', sku)
    wrapped_text = ''
    current_line = ''

    for word in words:
        if len(current_line) + len(word) + 1 <= max_line_length:
            current_line += word + ''
        else:
            wrapped_text += current_line.strip() + '' + '\n'
            current_line = word + ''

    wrapped_text += current_line.strip()
    return wrapped_text


def generate_qr_code(data):
    """Generate a QR code image and return the file path."""
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(data)
    qr.make(fit=True)

    img = qr.make_image(fill_color="black", back_color="white")

    # Create a temporary file to save the QR code
    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as temp_file:
        img.save(temp_file, format='PNG')
        temp_file_path = temp_file.name

    return temp_file_path


def create_circular_qrcode(df, column, fontsize=3.5, max_line_length=8,
                           filename="circular_labels.pdf", border_enabled=0):
    label_diameter = 0.75 * inch
    radius = label_diameter / 2

    page_margin_x = 0.37 * inch
    page_margin_y = 0.61 * inch
    label_padding_x = 0.1 * inch
    label_padding_y = 0.1 * inch

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

        # Generate and draw QR code
        qr_code_data = sku
        qr_code_path = generate_qr_code(qr_code_data)

        qr_code_x = x_center - 22  # Center QR code in the circle
        qr_code_y = y_center - 24  # Center QR code in the circle
        c.drawImage(qr_code_path, qr_code_x, qr_code_y, width=45, height=45)

        # Calculate text position for the SKU name on top of the QR code
        name_text_y = qr_code_y + 40  # Position the text above the QR code
        name_text_x = qr_code_x + 22.5  # Center the name text over the QR code

        # Draw the SKU name on top of the QR code
        c.setFillColor(black)  # Set color to black for visibility
        c.drawString(name_text_x - (c.stringWidth(sku, "Helvetica", current_font_size) / 2), name_text_y, sku)

        # Clean up temporary QR code image
        os.remove(qr_code_path)

    c.save()




########### WITH BORDER ##############
create_circular_qrcode(
    faceted_df,
    "Name",
    max_line_length=7,
    filename="qr_codes_borders/Circle_Faceted.pdf",
    border_enabled=1
)

create_circular_qrcode(
    cab_df,
    "Name",
    max_line_length=7,
    filename="qr_codes_borders/Circle_Cab.pdf",
    border_enabled=1
)

create_circular_qrcode(
    orb_df,
    "Name",
    max_line_length=8,
    filename="qr_codes_borders/Circle_Orb.pdf",
    border_enabled=1
)


########### WITHOUT BORDER ##############
create_circular_qrcode(
    faceted_df,
    "Name",
    max_line_length=7,
    filename="qr_codes_no_borders/Circle_Faceted.pdf",
    border_enabled=0
)

create_circular_qrcode(
    cab_df,
    "Name",
    max_line_length=7,
    filename="qr_codes_no_borders/Circle_Cab.pdf",
    border_enabled=0
)

create_circular_qrcode(
    orb_df,
    "Name",
    max_line_length=8,
    filename="qr_codes_no_borders/Circle_Orb.pdf",
    border_enabled=0
)
