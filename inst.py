import streamlit as st
import pandas as pd
import os
import re
import datetime
from io import BytesIO
import tempfile
from PIL import Image as PILImage, ImageDraw, ImageFont
import base64

# ReportLab imports for PDF generation
from reportlab.lib.pagesizes import landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, PageBreak, Image
from reportlab.lib.units import cm, inch
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

# Define sticker dimensions
STICKER_WIDTH = 10 * cm
STICKER_HEIGHT = 15 * cm
STICKER_PAGESIZE = (STICKER_WIDTH, STICKER_HEIGHT)

# Define content box dimensions
CONTENT_BOX_WIDTH = 9.8 * cm
CONTENT_BOX_HEIGHT = 5 * cm

def normalize_column_name(col_name):
    """Normalize column names by removing all non-alphanumeric characters and converting to lowercase"""
    return re.sub(r'[^a-zA-Z0-9]', '', str(col_name)).lower()

def find_column(df, possible_names):
    """Find a column in the DataFrame that matches any of the possible names"""
    normalized_df_columns = {normalize_column_name(col): col for col in df.columns}
    normalized_possible_names = [normalize_column_name(name) for name in possible_names]

    for norm_name in normalized_possible_names:
        if norm_name in normalized_df_columns:
            return normalized_df_columns[norm_name]

    # Check for partial matches
    for norm_name in normalized_possible_names:
        for df_norm_name, original_name in normalized_df_columns.items():
            if norm_name in df_norm_name or df_norm_name in norm_name:
                return original_name

    # Check for line location keywords
    for df_norm_name, original_name in normalized_df_columns.items():
        if ('line' in df_norm_name and 'location' in df_norm_name) or 'lineloc' in df_norm_name:
            return original_name

    return None

def process_uploaded_logo(uploaded_logo, target_width_cm, target_height_cm):
    """Process uploaded logo to fit the specified dimensions"""
    try:
        # Load image from uploaded file
        logo_img = PILImage.open(uploaded_logo)

        # Convert to RGB if necessary
        if logo_img.mode in ('RGBA', 'LA', 'P'):
            # Create white background
            background = PILImage.new('RGB', logo_img.size, (255, 255, 255))
            if logo_img.mode == 'P':
                logo_img = logo_img.convert('RGBA')
            background.paste(logo_img, mask=logo_img.split()[-1] if logo_img.mode in ('RGBA', 'LA') else None)
            logo_img = background

        # Convert cm to pixels for resizing (using 300 DPI)
        dpi = 300
        box_width_px = int(target_width_cm * dpi / 2.54)
        box_height_px = int(target_height_cm * dpi / 2.54)

        # Get original dimensions
        orig_width, orig_height = logo_img.size

        # Calculate aspect ratio and resize to fit within bounds while maintaining aspect ratio
        aspect_ratio = orig_width / orig_height
        target_aspect = box_width_px / box_height_px

        if aspect_ratio > target_aspect:
            # Image is wider, fit to width
            new_width = box_width_px
            new_height = int(box_width_px / aspect_ratio)
        else:
            # Image is taller, fit to height
            new_height = box_height_px
            new_width = int(box_height_px * aspect_ratio)

        # Resize with high quality
        logo_img = logo_img.resize((new_width, new_height), PILImage.Resampling.LANCZOS)

        # Convert to bytes for ReportLab
        img_buffer = BytesIO()
        logo_img.save(img_buffer, format='PNG', quality=100, optimize=False)
        img_buffer.seek(0)

        # Convert back to cm for ReportLab
        final_width_cm = new_width * 2.54 / dpi
        final_height_cm = new_height * 2.54 / dpi

        print(f"LOGO DEBUG: Target: {target_width_cm:.2f}cm x {target_height_cm:.2f}cm")
        print(f"LOGO DEBUG: Final: {final_width_cm:.2f}cm x {final_height_cm:.2f}cm")
        print(f"LOGO DEBUG: Pixels: {new_width}px x {new_height}px")

        # Create ReportLab Image with actual dimensions
        return Image(img_buffer, width=final_width_cm*cm, height=final_height_cm*cm)

    except Exception as e:
        st.error(f"Error processing uploaded logo: {e}")
        return None

def generate_qr_code(data_string):
    """Generate a QR code from the given data string"""
    try:
        import qrcode
        from PIL import Image as PILImage

        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=8,
            border=2,
        )

        qr.add_data(data_string)
        qr.make(fit=True)

        qr_img = qr.make_image(fill_color="black", back_color="white")

        img_buffer = BytesIO()
        qr_img.save(img_buffer, format='PNG')
        img_buffer.seek(0)

        return Image(img_buffer, width=1.8*cm, height=1.8*cm)
    except Exception as e:
        st.error(f"Error generating QR code: {e}")
        return None

def parse_line_location(location_string):
    """Parse line location string and split into 4 boxes"""
    if not location_string or pd.isna(location_string):
        return ["", "", "", ""]

    parts = str(location_string).split("_")
    result = parts[:4] + [""] * (4 - len(parts))
    return result[:4]

def generate_sticker_labels(df, line_loc_header_width, line_loc_box1_width,
                          line_loc_box2_width, line_loc_box3_width, line_loc_box4_width,
                          uploaded_first_box_logo=None):
    """Generate sticker labels with QR code from DataFrame"""
    try:
        # Define column mappings - Including bin_type mapping
        column_mappings = {
            'ASSLY': ['assly', 'ASSY NAME', 'Assy Name', 'assy name', 'assyname',
                     'assy_name', 'Assy_name', 'Assembly', 'Assembly Name', 'ASSEMBLY', 'Assembly_Name'],
            'part_no': ['PARTNO', 'PARTNO.', 'Part No', 'Part Number', 'PartNo',
                       'partnumber', 'part no', 'partnum', 'PART', 'part', 'Product Code',
                       'Item Number', 'Item ID', 'Item No', 'item', 'Item'],
            'description': ['DESCRIPTION', 'Description', 'Desc', 'Part Description',
                           'ItemDescription', 'item description', 'Product Description',
                           'Item Description', 'NAME', 'Item Name', 'Product Name'],
            'Part_per_veh': ['QYT', 'QTY / VEH', 'Qty/Veh', 'Qty Bin', 'Quantity per Bin',
                            'qty bin', 'qtybin', 'quantity bin', 'BIN QTY', 'BINQTY',
                            'QTY_BIN', 'QTY_PER_BIN', 'Bin Quantity', 'BIN'],
            'Type': ['TYPE', 'type', 'Type', 'tyPe', 'Type name'],
            'line_location': ['LINE LOCATION', 'Line Location', 'line location', 'LINELOCATION',
                             'linelocation', 'Line_Location', 'line_location', 'LINE_LOCATION',
                             'LineLocation', 'line_loc', 'lineloc', 'LINELOC', 'Line Loc'],
            'part_status': ['PART STATUS', 'Part Status', 'part status', 'PARTSTATUS',
                           'partstatus', 'Part_Status', 'part_status', 'PART_STATUS',
                           'PartStatus', 'STATUS', 'Status', 'status', 'Item Status',
                           'Component Status', 'Part State', 'State'],
            'bin_type': ['BIN TYPE', 'Bin Type', 'bin type', 'BINTYPE', 'bintype',
                        'Bin_Type', 'bin_type', 'BIN_TYPE', 'BinType', 'Container Type',
                        'CONTAINER TYPE', 'Container_Type', 'container_type', 'CONTAINER_TYPE',
                        'ContainerType', 'CONTAINER', 'Container', 'container', 'BIN', 'Bin', 'bin',
                        'Package Type', 'PACKAGE TYPE', 'Package_Type', 'package_type', 'PACKAGE_TYPE',
                        'PackageType', 'Storage Type', 'STORAGE TYPE', 'Storage_Type', 'storage_type']
        }

        # Find columns
        found_columns = {}
        for key, possible_names in column_mappings.items():
            found_col = find_column(df, possible_names)
            if found_col:
                found_columns[key] = found_col

        # Check required columns
        required_columns = ['ASSLY', 'part_no', 'description']
        missing_required = [col for col in required_columns if col not in found_columns]

        if missing_required:
            st.error(f"Missing required columns: {missing_required}")
            return None, None

        # Create a temporary file for PDF
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            output_pdf_path = tmp_file.name

        # Create PDF with adjusted margins
        def draw_border(canvas, doc):
            canvas.saveState()
            x_offset = (STICKER_WIDTH - CONTENT_BOX_WIDTH) / 2
            y_offset = STICKER_HEIGHT - CONTENT_BOX_HEIGHT - 0.2*cm
            canvas.setStrokeColor(colors.black)
            canvas.setLineWidth(1.5)
            canvas.rect(
                x_offset,
                y_offset,
                CONTENT_BOX_WIDTH,
                CONTENT_BOX_HEIGHT
            )
            canvas.restoreState()

        doc = SimpleDocTemplate(output_pdf_path, pagesize=STICKER_PAGESIZE,
                              topMargin=0.2*cm,
                              bottomMargin=(STICKER_HEIGHT - CONTENT_BOX_HEIGHT - 0.2*cm),
                              leftMargin=(STICKER_WIDTH - CONTENT_BOX_WIDTH) / 2,
                              rightMargin=(STICKER_WIDTH - CONTENT_BOX_WIDTH) / 2)

        # Define styles
        header_style = ParagraphStyle(name='HEADER', fontName='Helvetica-Bold', fontSize=8, alignment=TA_CENTER, leading=9)
        ASSLY_style = ParagraphStyle(
            name='ASSLY',
            fontName='Helvetica',
            fontSize=9,
            alignment=TA_LEFT,
            leading=11,
            spaceAfter=0,
            wordWrap='CJK',
            autoLeading="max"
        )
        Part_style = ParagraphStyle(
            name='PART NO',
            fontName='Helvetica-Bold',
            fontSize=11,
            alignment=TA_LEFT,
            leading=13,
            spaceAfter=0,
            wordWrap='CJK',
            autoLeading="max"
        )
        # Style for part status box
        Part_status_style = ParagraphStyle(
            name='PART STATUS',
            fontName='Helvetica-Bold',
            fontSize=9,
            alignment=TA_CENTER,
            leading=11,
            spaceAfter=0,
            wordWrap='CJK',
            autoLeading="max"
        )
        desc_style = ParagraphStyle(name='PART DESC', fontName='Helvetica', fontSize=7, alignment=TA_LEFT, leading=8, spaceAfter=0, wordWrap='CJK', autoLeading="max")
        partper_style = ParagraphStyle(name='Quantity', fontName='Helvetica', fontSize=9, alignment=TA_LEFT, leading=12)
        # Style for bin type
        bin_type_style = ParagraphStyle(name='BinType', fontName='Helvetica', fontSize=8, alignment=TA_CENTER, leading=10)
        Type_style = ParagraphStyle(name='Quantity', fontName='Helvetica', fontSize=7, alignment=TA_LEFT, leading=12)
        date_style = ParagraphStyle(name='DATE', fontName='Helvetica', fontSize=7, alignment=TA_LEFT, leading=12)
        location_style = ParagraphStyle(name='Location', fontName='Helvetica', fontSize=8, alignment=TA_CENTER, leading=10)

        content_width = CONTENT_BOX_WIDTH
        all_elements = []
        today_date = datetime.datetime.now().strftime("%d-%m-%Y")

        # Handle uploaded logo for first box
        first_box_logo = None
        if uploaded_first_box_logo is not None:
            # Logo takes 23% of total content width
            logo_width_cm = (content_width * 0.23) / cm  # 23% of content width in cm
            logo_height_cm = 0.75  # 0.75cm height (within 0.85cm row height)

            print(f"LOGO CALCULATION:")
            print(f"Content width: {content_width/cm:.2f}cm")
            print(f"Logo width (23%): {logo_width_cm:.2f}cm")
            print(f"Logo height: {logo_height_cm:.2f}cm")

            first_box_logo = process_uploaded_logo(uploaded_first_box_logo, logo_width_cm, logo_height_cm)
            if first_box_logo:
                st.success(f"✅ Logo processed - Size: {logo_width_cm:.2f}cm x {logo_height_cm:.2f}cm (23% width)")
            else:
                st.error("❌ Failed to process uploaded logo")

        # Process each row
        total_rows = len(df)
        progress_bar = st.progress(0)

        for index, row in df.iterrows():
            progress_bar.progress((index + 1) / total_rows)

            elements = []

            # Extract data - Including bin_type extraction
            ASSLY = str(row[found_columns.get('ASSLY', '')]) if 'ASSLY' in found_columns else "N/A"
            part_no = str(row[found_columns.get('part_no', '')]) if 'part_no' in found_columns else "N/A"
            desc = str(row[found_columns.get('description', '')]) if 'description' in found_columns else "N/A"
            Part_per_veh = str(row[found_columns.get('Part_per_veh', '')]) if 'Part_per_veh' in found_columns and pd.notna(row[found_columns['Part_per_veh']]) else ""
            Type = str(row[found_columns.get('Type', '')]) if 'Type' in found_columns and pd.notna(row[found_columns['Type']]) else ""
            line_location_raw = str(row[found_columns.get('line_location', '')]) if 'line_location' in found_columns and pd.notna(row[found_columns['line_location']]) else ""
            part_status = str(row[found_columns.get('part_status', '')]) if 'part_status' in found_columns and pd.notna(row[found_columns['part_status']]) else ""
            bin_type = str(row[found_columns.get('bin_type', '')]) if 'bin_type' in found_columns and pd.notna(row[found_columns['bin_type']]) else ""
            location_boxes = parse_line_location(line_location_raw)

            # Generate QR code - Including bin_type in QR data
            qr_data = f"ASSLY: {ASSLY}\nPart No: {part_no}\nDescription: {desc}\n"
            if Part_per_veh:
                qr_data += f"QTY/VEH: {Part_per_veh}\n"
            if bin_type:
                qr_data += f"Bin Type: {bin_type}\n"
            if Type:
                qr_data += f"Type: {Type}\n"
            if part_status:
                qr_data += f"Part Status: {part_status}\n"
            if line_location_raw:
                qr_data += f"Line Location: {line_location_raw}\n"
            qr_data += f"Date: {today_date}"

            qr_image = generate_qr_code(qr_data)
            if qr_image:
                qr_cell = qr_image
            else:
                qr_cell = Paragraph("QR", ParagraphStyle(name='QRPlaceholder', fontName='Helvetica-Bold', fontSize=12, alignment=TA_CENTER))

            # Row heights
            ASSLY_row_height = 0.85*cm
            part_row_height = 0.8*cm
            desc_row_height = 0.5*cm
            qty_row_height = 0.6*cm
            type_row_height = 0.6*cm
            date_row_height = 0.6*cm
            location_row_height = 0.5*cm

            # Process line location boxes
            location_box_1 = Paragraph(location_boxes[0], location_style) if location_boxes[0] else ""
            location_box_2 = Paragraph(location_boxes[1], location_style) if location_boxes[1] else ""
            location_box_3 = Paragraph(location_boxes[2], location_style) if location_boxes[2] else ""
            location_box_4 = Paragraph(location_boxes[3], location_style) if location_boxes[3] else ""

            # Create ASSLY row content
            first_box_content = first_box_logo if first_box_logo else ""

            # Create all table rows
            if len(ASSLY) > 3:
                assly_main = ASSLY[:-3]
                assly_suffix = ASSLY[-3:]
            else:
                assly_main = ""
                assly_suffix = ASSLY
            formatted_assly = f"{assly_main}<font size='11'><b>{assly_suffix}</b></font>"
            
            assly_row = [first_box_content, "ASSLY", Paragraph(formatted_assly, ASSLY_style)]
            partno_row = ["PART NO", Paragraph(f"<b>{part_no}</b>", Part_style), Paragraph(f"<b>{part_status}</b>", Part_status_style)]
            desc_row = ["PART DESC", Paragraph(desc, desc_style)]
            
            # Create combined table for QTY, TYPE, DATE with QR code
            qty_type_date_data = [
                ["QTY/VEH", Paragraph(str(Part_per_veh), partper_style), Paragraph(bin_type, bin_type_style), qr_cell],
                ["TYPE", Paragraph(str(Type), Type_style), "", ""],
                ["DATE", Paragraph(today_date, date_style), "", ""]
            ]
            
            location_row = ["LINE LOCATION", location_box_1, location_box_2, location_box_3, location_box_4]

            # Column widths - MAINTAINING ORIGINAL STRUCTURE
            col_widths_assly = [
                content_width * 0.25,    # Logo box: 25%
                content_width * 0.15,    # Header: 15%
                content_width * 0.60     # Value: 60%
            ]

            # Column widths for 3-column PART NO row
            col_widths_partno = [
                content_width * 0.25,    # Header: 25%
                content_width * 0.50,    # Part number: 50%
                content_width * 0.25     # Part status: 25%
            ]
            
            col_widths_standard = [content_width * 0.25, content_width * 0.75]

            # Column widths for combined QTY/TYPE/DATE table (4 columns)
            col_widths_qty = [
                content_width * 0.25,    # Header: 25%
                content_width * 0.175,   # Value 1: 17.5%
                content_width * 0.175,   # Value 2/Bin type: 17.5%
                content_width * 0.40     # QR code: 40%
            ]
            
            col_widths_bottom = [
                content_width * line_loc_header_width,
                content_width * line_loc_box1_width,
                content_width * line_loc_box2_width,
                content_width * line_loc_box3_width,
                content_width * line_loc_box4_width
            ]

            # Create tables
            assly_table = Table([assly_row], colWidths=col_widths_assly, rowHeights=[ASSLY_row_height])
            partno_table = Table([partno_row], colWidths=col_widths_partno, rowHeights=[part_row_height])
            desc_table = Table([desc_row], colWidths=col_widths_standard, rowHeights=[desc_row_height])
            qty_table = Table(qty_type_date_data, colWidths=col_widths_qty, rowHeights=[qty_row_height, type_row_height, date_row_height])
            bottom_table = Table([location_row], colWidths=col_widths_bottom, rowHeights=[location_row_height])

            # Apply table styles
            assly_table_style = [
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTNAME', (1, 0), (1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('ALIGN', (0, 0), (0, 0), 'CENTER'),
                ('ALIGN', (1, 0), (1, 0), 'CENTER'),
                ('ALIGN', (2, 0), (2, 0), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('LEFTPADDING', (0, 0), (-1, -1), 2),
                ('RIGHTPADDING', (0, 0), (-1, -1), 2),
                ('TOPPADDING', (0, 0), (-1, -1), 2),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ]

            # Style for 3-column PART NO table
            partno_table_style = [
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),  # Header bold
                ('FONTNAME', (1, 0), (1, 0), 'Helvetica-Bold'),  # Part number bold
                ('FONTNAME', (2, 0), (2, 0), 'Helvetica-Bold'),  # Part status bold
                ('FONTSIZE', (0, 0), (0, 0), 8),                # Header font size
                ('FONTSIZE', (1, 0), (1, 0), 11),               # Part number font size
                ('FONTSIZE', (2, 0), (2, 0), 9),                # Part status font size
                ('ALIGN', (0, 0), (0, 0), 'CENTER'),            # Header centered
                ('ALIGN', (1, 0), (1, 0), 'LEFT'),              # Part number left
                ('ALIGN', (2, 0), (2, 0), 'CENTER'),            # Part status centered
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('LEFTPADDING', (0, 0), (-1, -1), 3),
                ('RIGHTPADDING', (0, 0), (-1, -1), 3),
                ('TOPPADDING', (0, 0), (-1, -1), 2),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ]

            desc_table_style = [
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (0, -1), 8),
                ('FONTSIZE', (1, 0), (-1, 0), 7),
                ('ALIGN', (0, 0), (0, -1), 'CENTER'),
                ('ALIGN', (1, 0), (1, 0), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('LEFTPADDING', (0, 0), (-1, -1), 3),
                ('RIGHTPADDING', (0, 0), (-1, -1), 3),
                ('TOPPADDING', (0, 0), (-1, -1), 2),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ]

            # Style for combined QTY/TYPE/DATE table with proper QR spanning
            qty_table_style = [
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),  # Headers bold
                ('FONTSIZE', (0, 0), (0, -1), 8),                # Header font size
                ('FONTSIZE', (1, 0), (2, 0), 9),                # QTY and bin type font size
                ('FONTSIZE', (1, 1), (2, 2), 10),               # TYPE and DATE font size
                ('ALIGN', (0, 0), (0, -1), 'CENTER'),           # Headers centered
                ('ALIGN', (1, 0), (1, -1), 'LEFT'),             # Values left
                ('ALIGN', (2, 0), (2, 0), 'CENTER'),            # Bin type centered
                ('ALIGN', (3, 0), (3, 0), 'CENTER'),            # QR code centered
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('LEFTPADDING', (0, 0), (-1, -1), 3),
                ('RIGHTPADDING', (0, 0), (-1, -1), 3),
                ('TOPPADDING', (0, 0), (-1, -1), 2),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
                ('SPAN', (3, 0), (3, 2)),  # QR code spans all three rows
            ]

            bottom_table_style = [
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('ALIGN', (0, 0), (0, 0), 'CENTER'),
                ('ALIGN', (1, 0), (-1, 0), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('LEFTPADDING', (0, 0), (-1, -1), 2),
                ('RIGHTPADDING', (0, 0), (-1, -1), 2),
                ('TOPPADDING', (0, 0), (-1, -1), 2),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ]

            # Apply styles to tables
            assly_table.setStyle(assly_table_style)
            partno_table.setStyle(partno_table_style)
            desc_table.setStyle(desc_table_style)
            qty_table.setStyle(qty_table_style)
            bottom_table.setStyle(bottom_table_style)

            # Add tables to elements
            elements.extend([
                assly_table,
                partno_table,
                desc_table,
                qty_table,
                bottom_table
            ])

            # Add page break after each sticker except the last one
            if index < len(df) - 1:
                elements.append(PageBreak())
            
            all_elements.extend(elements)

        # Clear progress bar
        progress_bar.empty()

        # Build PDF with border
        doc.build(all_elements, onFirstPage=draw_border, onLaterPages=draw_border)

        # Read the generated PDF
        with open(output_pdf_path, 'rb') as pdf_file:
            pdf_data = pdf_file.read()

        # Clean up temporary file
        os.unlink(output_pdf_path)

        return pdf_data, f"sticker_labels_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"

    except Exception as e:
        st.error(f"Error generating sticker labels: {str(e)}")
        import traceback
        st.error(f"Traceback: {traceback.format_exc()}")
        return None, None

def main():
    """Main Streamlit application"""
    st.set_page_config(page_title="Sticker Label Generator", layout="wide")
    
    st.title("🏷️ Sticker Label Generator")
    st.markdown("Upload your Excel/CSV file to generate professional sticker labels with QR codes")

    # Sidebar for configuration
    with st.sidebar:
        st.header("⚙️ Configuration")
        
        # File upload
        uploaded_file = st.file_uploader(
            "Choose Excel/CSV file",
            type=['xlsx', 'xls', 'csv'],
            help="Upload your data file containing product information"
        )
        
        # Logo upload
        st.subheader("📷 Logo Settings")
        uploaded_logo = st.file_uploader(
            "Upload Logo (Optional)",
            type=['png', 'jpg', 'jpeg'],
            help="Upload a logo to appear in the first box of each sticker"
        )
        
        # Line location width settings
        st.subheader("📏 Line Location Box Widths")
        st.markdown("Adjust the width percentages for line location boxes:")
        
        line_loc_header_width = st.slider(
            "Header Width (%)", 
            min_value=0.1, max_value=0.5, 
            value=0.25, step=0.05,
            help="Width of the 'LINE LOCATION' header box"
        )
        
        remaining_width = 1.0 - line_loc_header_width
        
        line_loc_box1_width = st.slider(
            "Box 1 Width (%)", 
            min_value=0.1, max_value=remaining_width, 
            value=min(0.2, remaining_width/4), step=0.05
        )
        
        remaining_width2 = remaining_width - line_loc_box1_width
        
        line_loc_box2_width = st.slider(
            "Box 2 Width (%)", 
            min_value=0.1, max_value=remaining_width2, 
            value=min(0.2, remaining_width2/3), step=0.05
        )
        
        remaining_width3 = remaining_width2 - line_loc_box2_width
        
        line_loc_box3_width = st.slider(
            "Box 3 Width (%)", 
            min_value=0.1, max_value=remaining_width3, 
            value=min(0.15, remaining_width3/2), step=0.05
        )
        
        line_loc_box4_width = remaining_width3 - line_loc_box3_width
        
        st.info(f"Box 4 Width: {line_loc_box4_width:.2%} (auto-calculated)")

    # Main content area
    if uploaded_file is not None:
        try:
            # Read the uploaded file
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            st.success(f"✅ File uploaded successfully! Found {len(df)} rows and {len(df.columns)} columns.")
            
            # Display file preview
            with st.expander("📋 Data Preview", expanded=True):
                st.dataframe(df.head(10), use_container_width=True)
            
            # Column mapping section
            st.subheader("🔗 Column Mapping")
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**Available Columns:**")
                for i, col in enumerate(df.columns):
                    st.write(f"{i+1}. {col}")
            
            with col2:
                st.write("**Expected Column Types:**")
                expected_columns = [
                    "ASSLY (Assembly Name)",
                    "PART NO (Part Number)",
                    "DESCRIPTION (Part Description)",
                    "QTY/VEH (Quantity per Vehicle)",
                    "TYPE (Part Type)",
                    "LINE LOCATION (Location Info)",
                    "PART STATUS (Status)",
                    "BIN TYPE (Container Type)"
                ]
                for col in expected_columns:
                    st.write(f"• {col}")
            
            # Generate labels button
            st.subheader("🏷️ Generate Labels")
            
            if st.button("Generate Sticker Labels", type="primary", use_container_width=True):
                with st.spinner("Generating sticker labels... Please wait."):
                    pdf_data, filename = generate_sticker_labels(
                        df, 
                        line_loc_header_width,
                        line_loc_box1_width,
                        line_loc_box2_width,
                        line_loc_box3_width,
                        line_loc_box4_width,
                        uploaded_logo
                    )
                    
                    if pdf_data:
                        st.success("✅ Sticker labels generated successfully!")
                        
                        # Download button
                        st.download_button(
                            label="📥 Download PDF",
                            data=pdf_data,
                            file_name=filename,
                            mime="application/pdf",
                            use_container_width=True
                        )
                        
                        # Display some statistics
                        st.info(f"📊 Generated {len(df)} sticker labels")
                        
                    else:
                        st.error("❌ Failed to generate sticker labels. Please check your data and try again.")
        
        except Exception as e:
            st.error(f"❌ Error reading file: {str(e)}")
            st.error("Please make sure your file is a valid Excel (.xlsx, .xls) or CSV file.")
    
    else:
        # Instructions when no file is uploaded
        st.info("👆 Please upload an Excel or CSV file to get started.")
        
        st.subheader("📋 File Format Requirements")
        st.markdown("""
        Your file should contain the following columns (column names are flexible):
        
        **Required Columns:**
        - **Assembly Name** (ASSLY, Assembly, etc.)
        - **Part Number** (PART NO, Part Number, etc.)
        - **Description** (DESCRIPTION, Part Description, etc.)
        
        **Optional Columns:**
        - **Quantity** (QTY/VEH, Quantity, etc.)
        - **Type** (TYPE, Part Type, etc.)
        - **Line Location** (LINE LOCATION, Line Location, etc.)
        - **Part Status** (PART STATUS, Status, etc.)
        - **Bin Type** (BIN TYPE, Container Type, etc.)
        """)
        
        st.subheader("🔧 Features")
        st.markdown("""
        - 🏷️ **Professional sticker labels** with bordered layout
        - 📱 **QR codes** containing all product information
        - 🖼️ **Logo support** for branding
        - 📐 **Customizable layout** with adjustable column widths
        - 📋 **Flexible column mapping** - works with various column names
        - 📄 **PDF output** ready for printing
        - 🔄 **Batch processing** of multiple items
        """)

# Run the application
if __name__ == "__main__":
    main()
