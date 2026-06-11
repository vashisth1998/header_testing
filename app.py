import streamlit as st
from PIL import Image
import fitz  # PyMuPDF (PDFs ke liye)
from docx import Document # python-docx (Word files ke liye)
from docx.shared import Inches
import io
import os

# Page setup 
st.set_page_config(page_title="Auto Letterhead Branding Tool", layout="centered")

st.title("📄 Instant Letterhead Generator")
st.write("Apna document upload karein, Central University of Haryana ka header-footer automatically lag jayega.")

st.divider()

# --- BACKGROUND SE HEADER/FOOTER LOAD KARNA ---
# Dhyaan rakhein ki aapki images ka naam exact yahi ho aur wo code wale folder mein hon
HEADER_PATH = "header.PNG"
FOOTER_PATH = "footer.PNG"

# Check karna ki background images moujood hain ya nahi
if not os.path.exists(HEADER_PATH) or not os.path.exists(FOOTER_PATH):
    st.error("⚠️ Background mein 'header.PNG' ya 'footer.PNG' nahi mili!")
    st.info("💡 Solution: Apni header aur footer images ko 'header.PNG' aur 'footer.PNG' naam se rename karke apne GitHub repo/folder mein upload kar dein.")
else:
    # --- MAIN DOCUMENT UPLOAD SECTION ---
    st.subheader("Apna Document Upload Karein")
    main_doc = st.file_uploader("Upload File (PDF, Word (.docx) ya Image)", type=['pdf', 'docx', 'PNG', 'jpg', 'jpeg'])

    if main_doc:
        file_extension = main_doc.name.split('.')[-1].lower()
        
        # Ek instant process button
        if st.button("Apply Letterhead & Generate Download Link 🚀", type="primary"):
            with st.spinner("Aapke document par header-footer lagaya ja raha hai..."):
                
                # ==========================================
                # 1. AGAR DOCUMENT IMAGE HAI (PNG/JPG)
                # ==========================================
                if file_extension in ['PNG', 'jpg', 'jpeg']:
                    header_img = Image.open(HEADER_PATH)
                    footer_img = Image.open(FOOTER_PATH)
                    main_img = Image.open(main_doc)
                    
                    target_width = main_img.width
                    
                    h_ratio = target_width / header_img.width
                    new_h_height = int(header_img.height * h_ratio)
                    header_resized = header_img.resize((target_width, new_h_height))
                    
                    f_ratio = target_width / footer_img.width
                    new_f_height = int(footer_img.height * f_ratio)
                    footer_resized = footer_img.resize((target_width, new_f_height))
                    
                    total_height = new_h_height + main_img.height + new_f_height
                    final_image = Image.new('RGB', (target_width, total_height), (255, 255, 255))
                    
                    final_image.paste(header_resized, (0, 0))
                    final_image.paste(main_img, (0, new_h_height))
                    final_image.paste(footer_resized, (0, new_h_height + main_img.height))
                    
                    img_byte_arr = io.BytesIO()
                    final_image.save(img_byte_arr, format='PNG')
                    
                    st.success("✅ Image Branded Successfully!")
                    st.image(final_image, caption="Preview", width=500)
                    
                    # DIRECT DOWNLOAD OPTION
                    st.download_button(
                        label="📥 Download Branded Image", 
                        data=img_byte_arr.getvalue(), 
                        file_name=f"branded_{main_doc.name}", 
                        mime="image/PNG"
                    )

                # ==========================================
                # 2. AGAR DOCUMENT PDF HAI
                # ==========================================
                elif file_extension == 'pdf':
                    with open(HEADER_PATH, "rb") as h_f, open(FOOTER_PATH, "rb") as f_f:
                        header_bytes = h_f.read()
                        footer_bytes = f_f.read()
                    
                    pdf_document = fitz.open(stream=main_doc.read(), filetype="pdf")
                    
                    for page_num in range(len(pdf_document)):
                        page = pdf_document.load_page(page_num)
                        page_rect = page.rect 
                        
                        # Margins layout set karna
                        header_rect = fitz.Rect(0, 0, page_rect.width, 110) 
                        footer_rect = fitz.Rect(0, page_rect.height - 90, page_rect.width, page_rect.height)
                        
                        page.insert_image(header_rect, stream=header_bytes)
                        page.insert_image(footer_rect, stream=footer_bytes)
                    
                    pdf_bytes = pdf_document.write()
                    pdf_document.close()
                    
                    st.success("✅ PDF Branded Successfully!")
                    
                    # DIRECT DOWNLOAD OPTION
                    st.download_button(
                        label="📥 Download Branded PDF", 
                        data=pdf_bytes, 
                        file_name=f"branded_{main_doc.name}", 
                        mime="application/pdf"
                    )

                # ==========================================
                # 3. AGAR DOCUMENT WORD FILE (.docx) HAI
                # ==========================================
                elif file_extension == 'docx':
                    doc = Document(main_doc)
                    
                    for section in doc.sections:
                        # Header setup
                        header = section.header
                        header.is_linked_to_previous = False
                        if header.paragraphs:
                            header.paragraphs[0].text = ""
                        else:
                            header.add_paragraph()
                        run_h = header.paragraphs[0].add_run()
                        run_h.add_picture(HEADER_PATH, width=Inches(6.2))
                        
                        # Footer setup
                        footer = section.footer
                        footer.is_linked_to_previous = False
                        if footer.paragraphs:
                            footer.paragraphs[0].text = ""
                        else:
                            footer.add_paragraph()
                        run_f = footer.paragraphs[0].add_run()
                        run_f.add_picture(FOOTER_PATH, width=Inches(6.2))
                    
                    doc_io = io.BytesIO()
                    doc.save(doc_io)
                    
                    st.success("✅ Word Document Branded Successfully!")
                    
                    # DIRECT DOWNLOAD OPTION
                    st.download_button(
                        label="📥 Download Branded DOCX", 
                        data=doc_io.getvalue(), 
                        file_name=f"branded_{main_doc.name}", 
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
