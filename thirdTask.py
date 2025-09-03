# --- PDF Page Numbering Utility ---
import io
from PyPDF2 import PdfReader, PdfWriter


# --- Robust watermark approach for page numbers ---
def create_page_number_overlay(num_pages, page_sizes):
    """
    Create a PDF (in memory) with page numbers at the bottom right for each page size in page_sizes.
    Returns a BytesIO object containing the overlay PDF.
    """
    overlay_buffer = io.BytesIO()
    from reportlab.pdfgen import canvas
    can = canvas.Canvas(overlay_buffer)
    for i, (width, height) in enumerate(page_sizes):
        can.setPageSize((width, height))
        can.setFont("Helvetica", 10)
        can.setFillColorRGB(0.53, 0.53, 0.53)
        page_number_text = f"Page {i+1} of {num_pages}"
        can.drawRightString(width - 40, 20, page_number_text)
        can.showPage()
    can.save()
    overlay_buffer.seek(0)
    return overlay_buffer

def add_page_numbers(input_pdf_path, output_pdf_path):
    """
    Adds page numbers to the bottom right of each page in a PDF using a robust watermark overlay approach.
    """
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()
    num_pages = len(reader.pages)
    # Get the size of each page
    page_sizes = [(float(page.mediabox.width), float(page.mediabox.height)) for page in reader.pages]
    # Create overlay PDF
    overlay_pdf_buffer = create_page_number_overlay(num_pages, page_sizes)
    overlay_reader = PdfReader(overlay_pdf_buffer)
    for i, page in enumerate(reader.pages):
        overlay_page = overlay_reader.pages[i]
        page.merge_page(overlay_page)
        writer.add_page(page)
    with open(output_pdf_path, "wb") as f_out:
        writer.write(f_out)
# Streamlit app to generate mail merge document
import streamlit as st
import pandas as pd
import re
import hashlib
import pathlib
import uuid
from io import BytesIO
from xhtml2pdf import pisa  # For PDF generation
import numpy as np  # For NaN handling

st.title('Mail Merge Document Generator')

# --- Download HTML Template Section ---
with open("mail_merge_template.html", "r", encoding="utf-8") as f:
    template_contents = f.read()
st.download_button(
    label="Download HTML Template",
    data=template_contents,
    file_name="mail_merge_template.html",
    mime="text/html"
)

def get_cache_paths(hash_str):
    cache_dir = pathlib.Path('.cache')
    cache_dir.mkdir(exist_ok=True)
    main_path = cache_dir / f'{hash_str}_main.csv'
    bin_path = cache_dir / f'{hash_str}_bin.csv'
    return str(main_path), str(bin_path)

def file_hash(file):
    file.seek(0)
    content = file.read()
    file.seek(0)
    return hashlib.sha256(content).hexdigest()

# --- Template HTML uploader ---
st.subheader("Step 1: Upload Mail Merge HTML Template")
template_file = st.file_uploader("Upload HTML Template", type=["html"], key="template")


# --- Data file uploader ---
st.subheader("Step 2: Upload Data File")
uploaded_file = st.file_uploader("Upload an Excel or CSV file", type=["csv", "xlsx"], key="data")

# Optionally, allow user to clear data manually
if 'main_data' in st.session_state and st.button("Remove Data File"):
    for key in ['main_data', 'rubbish_bin', 'file_hash']:
        if key in st.session_state:
            del st.session_state[key]
    st.rerun()

    
# --- Data Processing and Editing ---
if uploaded_file is not None:
    # Read data and initialize session state
    current_hash = file_hash(uploaded_file)
    main_path, bin_path = get_cache_paths(current_hash)
    import os
    # If cached files exist, load them instead of reprocessing
    def detect_date_columns(df, threshold=0.5):
        date_cols = []
        formats = ['%b %d %Y', '%m/%d/%Y %H:%M', '%Y-%m-%d', '%Y/%m/%d', '%d-%b-%Y', '%d/%m/%Y', '%m/%d/%Y']
        for col in df.columns:
            non_null_vals = df[col].dropna()
            if len(non_null_vals) == 0:
                continue
            max_parsed = 0
            str_vals = non_null_vals.astype(str)
            for fmt in formats:
                parsed = pd.to_datetime(str_vals, format=fmt, errors='coerce')
                num_parsed = parsed.notna().sum()
                if num_parsed > max_parsed:
                    max_parsed = num_parsed
            if pd.api.types.is_numeric_dtype(non_null_vals):
                try:
                    parsed = pd.to_datetime(non_null_vals, unit='d', origin='1899-12-30', errors='coerce')
                    num_parsed = parsed.notna().sum()
                    if num_parsed > max_parsed:
                        max_parsed = num_parsed
                except Exception:
                    pass
            if max_parsed / len(non_null_vals) > threshold:
                date_cols.append(col)
        return date_cols

    if os.path.exists(main_path) and os.path.exists(bin_path):
        try:
            df = pd.read_csv(main_path)
            rubbish_bin_df = pd.read_csv(bin_path)
            # Always run date detection and formatting after loading
            detected_date_columns = set(detect_date_columns(df))
            # Add columns whose name contains 'date' (case-insensitive)
            for col in df.columns:
                if 'date' in col.lower():
                    detected_date_columns.add(col)
            st.session_state['detected_date_columns'] = list(detected_date_columns)
            for col in detected_date_columns:
                try:
                    converted = pd.to_datetime(df[col], errors='coerce')
                    if converted.notna().sum() > 0:
                        try:
                            df[col] = converted.dt.strftime('%#d %B %Y')
                        except Exception:
                            df[col] = converted.dt.strftime('%-d %B %Y')
                except Exception:
                    pass
            st.session_state['main_data'] = df
            st.session_state['rubbish_bin'] = rubbish_bin_df
            st.session_state['file_hash'] = current_hash
        except Exception as e:
            st.error(f"Error loading cached data: {e}")
            st.stop()
    else:
        try:
            file_name = uploaded_file.name.lower()
            if file_name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            # Remove the first row if it contains header-like values (e.g., duplicate column names)
            df = df.reset_index(drop=True)
            # Heuristic: if any column value in row 0 matches its column name, drop row 0
            if any(str(df.columns[i]).strip().lower() == str(df.iloc[0, i]).strip().lower() for i in range(len(df.columns))):
                df = df.iloc[1:].reset_index(drop=True)

            # Remove columns B through Q (index 1 to 16) for Qualtrics survey data
            if df.shape[1] > 16:
                cols_to_drop = df.columns[1:17]
                df = df.drop(columns=cols_to_drop)
            # Detect date columns in 'd MMM yyyy' format
            def detect_date_columns(df, threshold=0.5):
                date_cols = []
                formats = ['%b %d %Y', '%m/%d/%Y %H:%M', '%Y-%m-%d', '%Y/%m/%d', '%d-%b-%Y', '%d/%m/%Y', '%m/%d/%Y']
                for col in df.columns:
                    non_null_vals = df[col].dropna()
                    if len(non_null_vals) == 0:
                        continue
                    max_parsed = 0
                    # Try as string
                    str_vals = non_null_vals.astype(str)
                    for fmt in formats:
                        parsed = pd.to_datetime(str_vals, format=fmt, errors='coerce')
                        num_parsed = parsed.notna().sum()
                        if num_parsed > max_parsed:
                            max_parsed = num_parsed
                    # Try as numeric (Excel datetimes)
                    if pd.api.types.is_numeric_dtype(non_null_vals):
                        try:
                            parsed = pd.to_datetime(non_null_vals, unit='d', origin='1899-12-30', errors='coerce')
                            num_parsed = parsed.notna().sum()
                            if num_parsed > max_parsed:
                                max_parsed = num_parsed
                        except Exception:
                            pass
                    if max_parsed / len(non_null_vals) > threshold:
                        date_cols.append(col)
                return date_cols

            detected_date_columns = detect_date_columns(df)
            st.session_state['detected_date_columns'] = detected_date_columns
            # Convert detected date columns to datetime and format as 'MMMM d, yyyy'
            for col in detected_date_columns:
                # Try all supported formats for conversion
                converted = None
                for fmt in ['%b %d %Y', '%m/%d/%Y %H:%M', '%Y-%m-%d', '%Y/%m/%d', '%d-%b-%Y', '%d/%m/%Y', '%m/%d/%Y']:
                    try:
                        converted = pd.to_datetime(df[col], format=fmt, errors='coerce')
                        if converted.notna().sum() > 0:
                            break
                    except Exception:
                        continue
                if converted is not None:
                    # Use Windows-compatible day format
                    try:
                        df[col] = converted.dt.strftime('%#d %B %Y')
                    except Exception:
                        # Fallback for non-Windows
                        df[col] = converted.dt.strftime('%-d %B %Y')
            df['ID'] = [str(uuid.uuid4()) for _ in range(len(df))]
            st.session_state['main_data'] = df
            st.session_state['rubbish_bin'] = pd.DataFrame(columns=['ID', 'Row'])
            st.session_state['file_hash'] = current_hash
            # Save to cache
            st.session_state['main_data'].to_csv(main_path, index=False)
            st.session_state['rubbish_bin'].to_csv(bin_path, index=False)
        except Exception as e:
            st.error(f"Error reading data file: {e}")
            st.stop()

# Data editing logic

if 'main_data' in st.session_state:
    main_data = st.session_state['main_data']
    rubbish_bin = st.session_state['rubbish_bin']

    # ...existing code...

    # --- Data Preview and Editing ---
    st.subheader("Data Preview and Editing")
    df_display = main_data.copy()
    if 'Delete' in df_display.columns:
        df_display = df_display.drop(columns=['Delete'])
    df_display['Delete'] = False

    # Replace NaN with empty strings for display
    df_display = df_display.fillna('')

    edited_df = st.data_editor(
        df_display,
        num_rows='fixed',
        use_container_width=True,
    )
    st.markdown(
        "<span style='color: #d9534f; font-weight: bold;'>To delete records, use the multiselect and 'Delete Selected' button below. The table's built-in delete button is disabled to ensure records are moved to the rubbish bin.</span>",
        unsafe_allow_html=True,
    )

    if not edited_df.equals(df_display):
        edited_main_data = edited_df.copy()
        if 'Delete' in edited_main_data.columns:
            edited_main_data = edited_main_data.drop(columns=['Delete'])
        # Convert empty strings back to NaN for consistency
        edited_main_data = edited_main_data.replace('', np.nan)
        st.session_state['main_data'] = edited_main_data.copy()
        main_data = st.session_state['main_data']
        # Save to cache
        main_path, bin_path = get_cache_paths(st.session_state['file_hash'])
        st.session_state['main_data'].to_csv(main_path, index=False)
        st.session_state['rubbish_bin'].to_csv(bin_path, index=False)

    # Select rows to delete by row number
    delete_indices = st.multiselect(
        'Select rows to delete',
        options=df_display.index.tolist(),
        format_func=lambda x: f"Row {x}"
    )
    if st.button("Delete Selected"):
        to_delete = main_data.loc[delete_indices].copy()
        to_delete['Row'] = delete_indices
        st.session_state['rubbish_bin'] = pd.concat([rubbish_bin, to_delete], ignore_index=True)
        st.session_state['main_data'] = main_data.drop(index=delete_indices).reset_index(drop=True)
        # Save to cache
        main_path, bin_path = get_cache_paths(st.session_state['file_hash'])
        st.session_state['main_data'].to_csv(main_path, index=False)
        st.session_state['rubbish_bin'].to_csv(bin_path, index=False)
        st.rerun()

    # --- Rubbish Bin Section ---
    st.write('---')
    st.subheader('Rubbish Bin')
    if not st.session_state['rubbish_bin'].empty:
        bin_display = st.session_state['rubbish_bin'].copy()
        if 'ID' in bin_display.columns:
            bin_display = bin_display.drop(columns=['ID'])
        # Replace NaN with empty strings for display
        bin_display = bin_display.fillna('')
        st.dataframe(bin_display.reset_index(drop=True))
        bin_df = st.session_state['rubbish_bin']
        def record_label(row):
            row_num = int(row['Row']) if 'Row' in row and pd.notnull(row['Row']) else '?'
            return f"Row {row_num}"
        restore_options = [row['ID'] for _, row in bin_df.iterrows()]
        restore_labels = {row['ID']: record_label(row) for _, row in bin_df.iterrows()}
        restore_ids = st.multiselect(
            'Select records to restore',
            options=restore_options,
            format_func=lambda x: restore_labels.get(x, x)
        )
        if st.button("Restore Selected from Bin"):
            to_restore = st.session_state['rubbish_bin'][st.session_state['rubbish_bin']['ID'].isin(restore_ids)]
            st.session_state['main_data'] = pd.concat([st.session_state['main_data'], to_restore], ignore_index=True)
            st.session_state['rubbish_bin'] = st.session_state['rubbish_bin'][~st.session_state['rubbish_bin']['ID'].isin(restore_ids)].reset_index(drop=True)
            # Save to cache
            main_path, bin_path = get_cache_paths(st.session_state['file_hash'])
            st.session_state['main_data'].to_csv(main_path, index=False)
            st.session_state['rubbish_bin'].to_csv(bin_path, index=False)
            st.rerun()
        if st.button('Clear Bin'):
            st.session_state['rubbish_bin'] = pd.DataFrame(columns=['ID', 'Row'])
            # Save to cache
            main_path, bin_path = get_cache_paths(st.session_state['file_hash'])
            st.session_state['main_data'].to_csv(main_path, index=False)
            st.session_state['rubbish_bin'].to_csv(bin_path, index=False)
            st.rerun()
    else:
        st.info('Rubbish bin is empty.')


# --- Mail Merge Generation (PDF) ---
if template_file is not None and 'main_data' in st.session_state and not st.session_state['main_data'].empty:
    try:
        # Read template
        template_file.seek(0)
        template_str = template_file.read().decode("utf-8")
        
        # Prepare data for mail merge (remove internal columns)
        mail_merge_data = st.session_state['main_data'].copy()
        for col in ['ID', 'Row', 'Delete']:
            if col in mail_merge_data.columns:
                mail_merge_data = mail_merge_data.drop(columns=[col])
        
        # Fill NaN with empty strings
        mail_merge_data = mail_merge_data.fillna('')
        
        # --- Generate Single Document ---
        st.subheader("Step 3: Generate Document")
        st.info("This will create a single PDF document with all records merged together.")
        if st.button("Generate and Download PDF"):
            with st.spinner("Generating document..."):
                # Generate merged HTML content
                merged_html = ""
                for index, record in mail_merge_data.iterrows():
                    try:
                        # Convert record to dictionary and clean keys
                        record_dict = record.to_dict()
                        # Handle special characters in column names and preserve line breaks
                        cleaned_dict = {}
                        for key, value in record_dict.items():
                            # Remove any non-alphanumeric characters from keys except underscores
                            clean_key = re.sub(r'[^\w]', '', key)
                            # Preserve line breaks for string values
                            if isinstance(value, str):
                                value = value.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '<br>')
                            cleaned_dict[clean_key] = value

                        # Automatically detect all *_Id columns for attachments
                        attachment_urls = []
                        for k in sorted(cleaned_dict.keys()):
                            if k.endswith('_Id') and cleaned_dict[k].strip():
                                url = f'https://hkbuchtl.qualtrics.com/Q/File.php?F={cleaned_dict[k].strip()}'
                                attachment_urls.append(url)

                        # Build the HTML for the attachment links
                        if attachment_urls:
                            attachment_links_html = '<br>'.join(f'<a href="{url}">{url}</a>' for url in attachment_urls)
                        else:
                            attachment_links_html = ''
                        # Replace the two hardcoded attachment links block in the template
                        rendered_html = re.sub(
                            r'<a href="https://hkbuchtl\.qualtrics\.com/Q/File\.php\?F=\{Q23_Id\}">https://hkbuchtl\.qualtrics\.com/Q/File\.php\?F=\{Q23_Id\}</a><br>\s*<a href="https://hkbuchtl\.qualtrics\.com/Q/File\.php\?F=\{Q24_Id\}">https://hkbuchtl\.qualtrics\.com/Q/File\.php\?F=\{Q24_Id\}</a>',
                            attachment_links_html,
                            template_str
                        )
                        # Now format the rest of the template
                        rendered_html = rendered_html.format(**cleaned_dict)

                        # Hide any label (static text) that is immediately followed by an empty placeholder value
                        # This works for all templates and label styles
                        placeholder_pattern = re.compile(r'\{([^\}]+)\}')
                        placeholders = set(placeholder_pattern.findall(template_str))
                        for ph in placeholders:
                            value = cleaned_dict.get(ph, '')
                            if not str(value).strip():
                                # Match patterns like: Label: {Placeholder} or Label: <strong>{Placeholder}</strong>
                                # Remove: Label: {Placeholder}
                                rendered_html = re.sub(r'[\w\s]+:\s*\{\s*' + re.escape(ph) + r'\s*\}', '', rendered_html)
                                # Remove: Label: <strong>{Placeholder}</strong>
                                rendered_html = re.sub(r'[\w\s]+:\s*<strong>\s*\{\s*' + re.escape(ph) + r'\s*\}\s*</strong>', '', rendered_html)
                                # Remove: (Label: {Placeholder})
                                rendered_html = re.sub(r'\([^)]+\{\s*' + re.escape(ph) + r'\s*\}[^)]*\)', '', rendered_html)
                                # Remove: (Label: <strong>{Placeholder}</strong>)
                                rendered_html = re.sub(r'\([^)]+<strong>\s*\{\s*' + re.escape(ph) + r'\s*\}\s*</strong>[^)]*\)', '', rendered_html)
                                # Remove: <p>Label</p> <p>{Placeholder}</p> if placeholder is empty
                                rendered_html = re.sub(r'<p>[\w\s]+<\/p>\s*<p>\s*\{\s*' + re.escape(ph) + r'\s*\}<\/p>', '', rendered_html)
                                # Remove: <p>Label</p> <p></p> if placeholder is empty after formatting
                                rendered_html = re.sub(r'<p>[\w\s]+<\/p>\s*<p>\s*<\/p>', '', rendered_html)
                                # Remove: (URL: <strong><a href="...">...</a></strong>) if URL is empty
                                rendered_html = re.sub(r'\(URL:\s*<strong><a [^>]*>\s*<\/a><\/strong>\)', '', rendered_html)
                                # Remove: (URL: <strong></strong>) if URL is empty
                                rendered_html = re.sub(r'\(URL:\s*<strong>\s*<\/strong>\)', '', rendered_html)

                        # Add a page break before every record except the first
                        if index > 0:
                            merged_html += '<div style="page-break-before: always;"></div>'
                        merged_html += rendered_html
                    except KeyError as e:
                        st.error(f"Missing placeholder in template: {e}")
                        st.stop()
                # --- TOC Generation ---
                # Find all <h2>...</h2> headings in merged_html
                headings = re.findall(r'<h2[^>]*>(.*?)</h2>', merged_html, re.DOTALL)
                toc_html = (
                    "<div style='display: flex; flex-direction: column; justify-content: center; align-items: center; height: 80vh;'>"
                    "<h1 style='text-align:center; margin-bottom: 40px;'>Table of Contents</h1>"
                    "<table style='font-size:1.2em; width: 90%; margin: 0 auto; border-collapse: collapse;'>"
                )
                for idx, heading in enumerate(headings):
                    toc_html += f"<tr><td style='padding: 8px 0 8px 10px; text-align: left; border: none;'>{idx+1}. {heading}</td><td style='padding: 8px 10px 8px 0; text-align: right; border: none; width: 60px;'>{idx+2}</td></tr>"
                toc_html += (
                    "</table>"
                    "</div>"
                    "<div style='page-break-after: always;'></div>"
                )
                # Prepend TOC to merged_html
                merged_html_with_toc = toc_html + merged_html
                # Create PDF
                pdf_buffer = BytesIO()
                full_html = f"""<!DOCTYPE html>
                <html>
                <head>
                    <meta charset='utf-8'>
                    <title>Mail Merge Result</title>
                    <style>
                        body {{ font-family: Arial, sans-serif; line-height: 1.7; font-size: 1.25em; }}
                        h1 {{ color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 5px; }}
                        h2 {{ color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 5px; }}
                        strong {{ color: #2980b9; }}
                        ol {{ padding-left: 1.2em; }}
                    </style>
                </head>
                <body>
                    {merged_html_with_toc}
                </body>
                </html>"""
                import tempfile
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_in, \
                     tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_out:
                    # Generate PDF to temp file
                    pisa_status = pisa.CreatePDF(
                        BytesIO(full_html.encode("UTF-8")), 
                        dest=tmp_in
                    )
                    tmp_in.flush()
                    if pisa_status.err:
                        st.error("PDF generation failed. Please check your HTML template.")
                    else:
                        # Add page numbers to the temp PDF
                        add_page_numbers(tmp_in.name, tmp_out.name)
                        with open(tmp_out.name, "rb") as f_final:
                            st.success("PDF generated successfully!")
                            st.download_button(
                                label="Download Merged Document",
                                data=f_final.read(),
                                file_name="mail_merge_result.pdf",
                                mime="application/pdf"
                            )
    except Exception as e:
        st.error(f"Mail merge failed: {str(e)}")
        st.error("Please check your template and data format")
elif template_file is not None and uploaded_file is None:
    st.info("Please upload a data file to perform mail merge.")
elif template_file is None and uploaded_file is None:
    st.info("Please upload an Excel/CSV file and/or HTML template to get started.")