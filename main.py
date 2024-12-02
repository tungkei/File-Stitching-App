import streamlit as st
import img2pdf
from pypdf import PdfReader, PdfWriter, Transformation
import tempfile
import subprocess
from io import BytesIO
import os
import base64

def process_files(uploaded_files):
    merged_pdf = PdfWriter()
    for uploaded_file in uploaded_files:
        file_name, file_extension = os.path.splitext(uploaded_file.name)
        file_bytes = uploaded_file.getvalue()
        if file_extension == ".jpg" or file_extension == ".jpeg" or file_extension == ".png":
            file_pdf_bytesio = convert_img_to_pdf(file_bytes)
            merged_pdf.append(file_pdf_bytesio)
        elif file_extension == ".docx":
            file_pdf_bytesio = convert_docx_to_pdf(file_bytes)
            merged_pdf.append(file_pdf_bytesio)
        elif file_extension == ".pdf":
            file_pdf_bytesio = resize_pdf(BytesIO(file_bytes))
            merged_pdf.append(file_pdf_bytesio)
        else: 
            raise ValueError("Unsupported file type, file type: " + file_extension)
    
    merged_pdf_bytesio = BytesIO()
    merged_pdf.write(merged_pdf_bytesio)
    merged_pdf.close()
    merged_pdf_bytesio.seek(0)
    return merged_pdf_bytesio
        
# Convert jpeg and png files to PDF file type
def convert_img_to_pdf(img_bytes):
    img_pdf_bytesio = BytesIO(img2pdf.convert(img_bytes))
    img_pdf_bytesio.seek(0)  # Reset pointer to beginning of BytesIO
    resized_pdf_bytesio = resize_pdf(img_pdf_bytesio)
    resized_pdf_bytesio.seek(0)
    return resized_pdf_bytesio

# Convert docx files to PDF file type
def convert_docx_to_pdf(docx_bytes):
    # Create a temporary file in memory
    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_docx:
        temp_docx.write(docx_bytes)
        temp_docx_path = temp_docx.name
    # Create temporary file for output PDF
    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_pdf:
        temp_pdf_path = temp_pdf.name

    # Command to convert DOCX to PDF using LibreOffice
    command = [
        'libreoffice',
        '--headless',  
        '--convert-to', 'pdf',  
        '--outdir', os.path.dirname(temp_pdf_path),  
        temp_docx_path  
    ]
    
    # Run the command
    subprocess.run(command)

    # Read the content of the output PDF file
    with open(temp_docx_path.replace('.docx', '.pdf'), 'rb') as temp_pdf:
        pdf_bytesio = BytesIO(temp_pdf.read())

    pdf_bytesio.seek(0)
    resized_pdf_bytesio = resize_pdf(pdf_bytesio)
    resized_pdf_bytesio.seek(0)
    # Delete temporary files
    os.unlink(temp_docx_path)
    os.unlink(temp_docx_path.replace('.docx', '.pdf'))
    os.unlink(temp_pdf_path)

    return resized_pdf_bytesio

# Resize PDFs to A4
def resize_pdf(pdf_bytesio):
    resized_pdf_bytesio = BytesIO()
    pdf_reader = PdfReader(pdf_bytesio)
    resized_writer = PdfWriter()

    for page_number in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_number]
        og_width = page.mediabox.width
        og_height = page.mediabox.height
        a4_width = 595
        a4_height = 842
        
        # Scale and translate image to middle of page
        scale_x = a4_width / og_width
        scale_y = a4_height / og_height
        scaling_param = min(scale_x, scale_y)

        y_offset = (a4_height - og_height * scaling_param) / 2
        
        op = Transformation().scale(sx=(scaling_param), sy=(scaling_param)).translate(ty=y_offset)
        page.add_transformation(op)
        page.mediabox.lower_left = (0,0)
        page.mediabox.upper_right = (a4_width, a4_height)
        
        resized_writer.add_page(page)

    resized_writer.write(resized_pdf_bytesio)
    resized_pdf_bytesio.seek(0)
    return resized_pdf_bytesio

def create_download_link(bytesio_obj, filename, label='Download Merged PDF'):
    bytesio_obj.seek(0)
    b64 = base64.b64encode(bytesio_obj.read()).decode()
    href = f'<a href="data:application/pdf;base64,{b64}" download="{filename}">{label}</a>'
    return href

# Streamlit UI
def main():
    st.title("File Stitching App")
    if "file_uploader_key" not in st.session_state:
        st.session_state["file_uploader_key"] = 0
    if "uploaded_files" not in st.session_state:
        st.session_state["uploaded_files"] = []
    
    if "merged_file_name" not in st.session_state:
        st.session_state["merged_file_name"] = ""

    merged_file_name = st.text_input(
        "Merged File Name: ",
        value=st.session_state["merged_file_name"],
        placeholder="Enter Merged File Name"
    )
    if merged_file_name:
        st.session_state["merged_file_name"] = merged_file_name

    # Only accept files of the following types: png, jpg, jpeg, docx, pdf
    uploaded_files = st.file_uploader("Upload files", accept_multiple_files=True, key=st.session_state["file_uploader_key"], type=['png', 'jpg', 'docx', 'pdf'],)
    
    # Only show buttons if at least 1 file uploaded
    if uploaded_files:
        if st.button("Clear all uploaded files"):
            st.session_state["file_uploader_key"] += 1
            st.session_state["merged_file_name"] = ""
            st.rerun()
        files_dict = {}
        
        for uploaded_file in uploaded_files:
            if uploaded_file.name not in files_dict:
                files_dict[uploaded_file.name] = uploaded_file
            else: 
                st.error("Please do not upload duplicated files, duplicates found for: " + uploaded_file.name)
        ordered_file_names = st.multiselect("Order of files", list(files_dict.keys()))
        

        if len(ordered_file_names) == len(uploaded_files):
            ordered_files = [files_dict[i] for i in ordered_file_names]

            if st.button("Merge Files"):
                if not merged_file_name:
                    st.error("Please provide the name of the merged file.")
                else: 
                    merged_pdf_bytesio = process_files(ordered_files)
                    st.write("Files Stitched!")
                    st.markdown(create_download_link(merged_pdf_bytesio, merged_file_name + ".pdf"), unsafe_allow_html=True)

if __name__ == "__main__":
    main()