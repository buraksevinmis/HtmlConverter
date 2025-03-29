import streamlit as st
import pandas as pd
import io
import base64
from docx import Document
from bs4 import BeautifulSoup
import re
import html
import time
from docx.shared import Pt, RGBColor
import zipfile
import tempfile
import os

# Set page title and icon
st.set_page_config(
    page_title="HTML Converter",
    page_icon="📄",
    layout="centered"
)

# Function to convert table from docx to HTML
def convert_table_to_html(table):
    html_table = "<table border=\"1\">"
    
    # Process header row
    html_table += "<thead>"
    for cell in table.rows[0].cells:
        html_table += f"<th>{cell.text}</th>"
    html_table += "</thead><tbody>"
    
    # Process data rows
    for i, row in enumerate(table.rows):
        if i == 0:  # Skip header row, already processed
            continue
        html_table += "<tr>"
        for cell in row.cells:
            html_table += f"<td>{cell.text}</td>"
        html_table += "</tr>"
    
    html_table += "</tbody></table>"
    return html_table

# Function to convert docx to HTML
def convert_docx_to_html(docx_file):
    doc = Document(docx_file)
    html_content = "<div>"
    
    for element in doc.element.body:
        if element.tag.endswith('p'):
            paragraph = element  # This is a paragraph
            
            # Check if paragraph has run elements (text)
            if len(paragraph.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')) > 0:
                # Check style (heading or normal paragraph)
                style_name = paragraph.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle')
                
                # Get paragraph text
                text = ""
                for run in paragraph.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r'):
                    for text_element in run.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
                        if text_element.text:
                            text += text_element.text
                
                # Apply HTML formatting based on style
                if style_name is not None and style_name.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val').startswith('Heading'):
                    heading_level = style_name.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')[-1]
                    try:
                        heading_number = int(heading_level)
                        if 1 <= heading_number <= 6:
                            html_content += f"<h{heading_number}>{html.escape(text)}</h{heading_number}>"
                        else:
                            html_content += f"<p><strong>{html.escape(text)}</strong></p>"
                    except ValueError:
                        html_content += f"<p><strong>{html.escape(text)}</strong></p>"
                else:
                    html_content += f"<p>{html.escape(text)}</p>"
        
        elif element.tag.endswith('tbl'):
            # This is a table
            table_index = list(doc.element.body).index(element)
            html_content += convert_table_to_html(doc.tables[table_index // 2])  # Approximate index
    
    html_content += "</div>"
    return html_content

# DOCX'i analiz ederek HTML'e dönüştür - FAQ desteği ile
def convert_docx_to_html_direct(doc, use_faq_schema=False):
    """Doğrudan python-docx kullanarak DOCX'i HTML'e dönüştürür.
    Bu sürüm, dokümanı sırayla işler ve soru-cevap bölümleri için FAQ şeması oluşturur.
    
    Args:
        doc: python-docx Document nesnesi
        use_faq_schema: Eğer True ise, FAQ şeması algılanır ve işlenir. False ise, normal HTML dönüşümü yapılır.
    """
    
    # Paragraf ve tabloları orijinal sırada işleyebilmek için yeni bir yapı
    elements = []
    
    # Tüm paragraf ve tabloları bul ve sırayla dokümanı dolaş
    # Doğrudan _element kullanarak orijinal XML yapısını analiz edelim
    for child in doc._element.body:
        if child.tag.endswith('p'):  # Paragraf
            for p in doc.paragraphs:
                if p._element == child:
                    elements.append(('paragraph', p))
                    break
        elif child.tag.endswith('tbl'):  # Tablo
            for t in doc.tables:
                if t._element == child:
                    elements.append(('table', t))
                    break
    
    # FAQ şeması kullanılmıyorsa hiç FAQ analizi yapmadan normal HTML dönüşümü yap
    faq_section = None
    faq_questions_indices = []
    
    # FAQ analizi sadece use_faq_schema=True ise yapılır
    if use_faq_schema:
        # Belgedeki tüm elemanları almadan önce, soru-cevap formatındaki başlıkların analizi yapılır
        # Belgenin sonunda ard arda en az 3 başlık "?" ile bitiyorsa bunları FAQ olarak işle
        heading_elements = [(i, elem) for i, (elem_type, elem) in enumerate(elements) 
                          if elem_type == 'paragraph' and elem.style and elem.style.name.startswith('Heading')]
        
        # Olası FAQ başlıkları (soru işareti ile biten başlıklar)
        faq_candidates = []
        for i, elem in heading_elements:
            if elem.text.strip().endswith('?'):
                faq_candidates.append((i, elem))
        
        # En az 3 adet ? ile biten başlık var mı? (belgenin sonunda olmalı)
        if len(faq_candidates) >= 3:
            # En sonda bulunan FAQ adaylarını bul
            # Başlıkların indekslerini al ve küçükten büyüğe sırala (doküman sırası)
            sorted_indices = sorted([i for i, _ in faq_candidates])
        
            # Ardışık soru işaretli başlıkları bul
            # En az 3 ardışık soru işaretli başlık olmalı
            consecutive_sequences = []
            current_sequence = [sorted_indices[0]]
        
            for i in range(1, len(sorted_indices)):
                current_index = sorted_indices[i]
                prev_index = sorted_indices[i-1]
                
                # Ardışık başlıklar arasında en fazla 2 eleman olabilir
                if current_index - prev_index <= 3:
                    current_sequence.append(current_index)
                else:
                    # Ardışıklık bozuldu, yeni sekans başlat
                    if len(current_sequence) >= 3:
                        consecutive_sequences.append(current_sequence)
                    current_sequence = [current_index]
        
            # Son sekansı da kontrol et
            if len(current_sequence) >= 3:
                consecutive_sequences.append(current_sequence)
            
            # En sondaki sekansı bul (en uzun ve en sonda olan)
            if consecutive_sequences:
                # Belgenin son 1/3'ünde olan sekansları filtrele
                total_elements = len(elements)
                last_third_start = int(total_elements * 2/3)
                
                valid_sequences = []
                for seq in consecutive_sequences:
                    # Sekansın başlangıcı son 1/3'te mi?
                    if seq[0] >= last_third_start:
                        valid_sequences.append(seq)
                
                # Geçerli sekans bulunduysa en uzununu al
                if valid_sequences:
                    # En uzun sekansı bul
                    longest_sequence = max(valid_sequences, key=len)
                    faq_section = longest_sequence[0]  # İlk FAQ başlığının indeksi
                    faq_questions_indices = longest_sequence
                
                    # Sonraki soru işaretli başlıkları da ekle (eğer ardışıklık devam ediyorsa)
                    extra_indices = []
                    last_index = longest_sequence[-1]
                    
                    for idx in sorted_indices:
                        if idx > last_index and idx - last_index <= 3:
                            extra_indices.append(idx)
                            last_index = idx
                        elif idx > last_index:
                            break
                    
                    faq_questions_indices.extend(extra_indices)
    
    # HTML içeriğini oluştur
    html_parts = []
    
    # FAQ bölümünü izlemek için değişkenler
    in_faq_section = False
    current_question = None
    faq_questions = []  # Soru-cevap çiftlerini tutar
    current_answer = []  # Mevcut cevabın paragraflarını tutar
    
    # Sırayla elemanları işle
    for i, (elem_type, elem) in enumerate(elements):
        # FAQ bölümündeyiz mi kontrol et
        if faq_section is not None and i >= faq_section and not in_faq_section:
            in_faq_section = True
            # Schema.org div'i sadece use_faq_schema aktif DEĞİLSE ekle (JSON-LD modunda eklemiyoruz)
            if not use_faq_schema:
                html_parts.append('<div itemscope itemprop="mainEntity" itemtype="https://schema.org/FAQPage">')
        
        if elem_type == 'paragraph':
            para = elem
            text = para.text
            
            # Başlık mı kontrol et
            is_heading = False
            heading_level = 0
            
            if para.style and para.style.name.startswith('Heading'):
                level_str = para.style.name.replace('Heading', '')
                try:
                    level = int(level_str.strip())
                    if 1 <= level <= 6:
                        is_heading = True
                        heading_level = level
                except ValueError:
                    pass  # Geçersiz başlık numarası
            
            # FAQ bölümündeyiz ve bu bir başlık mı?
            if in_faq_section and is_heading and text.strip().endswith('?'):
                # Eğer mevcut bir soru-cevap çifti kaydedildiyse, bir önceki soruyu tamamla
                if current_question and current_answer:
                    faq_questions.append((current_question, "".join(current_answer)))
                    current_answer = []
                
                # Yeni soruyu kaydet
                current_question = text
                continue  # Soruyu doğrudan kaydet, HTML'e ekle (FAQ olarak)
            elif in_faq_section and current_question and not is_heading:
                # Bu, mevcut sorunun cevabının bir parçası
                current_answer.append(f"<p>{html.escape(text)}</p>")
                continue
            
            # Normal başlık ve paragraf işleme (FAQ bölümünde değiliz veya FAQ formatında değil)
            if is_heading:
                html_parts.append(f"<h{heading_level}>{html.escape(text)}</h{heading_level}>")
            else:
                # Normal paragraflar
                html_parts.append(f"<p>{html.escape(text)}</p>")
        
        elif elem_type == 'table':
            table = elem
            
            # Tabloyu HTML'e çevir
            table_parts = ["<table border=\"1\"><tbody>"]
            
            for i, row in enumerate(table.rows):
                table_parts.append("<tr>")
                for cell in row.cells:
                    # İlk satırı başlık olarak kabul et
                    tag = "th" if i == 0 else "td"
                    table_parts.append(f"<{tag}>{html.escape(cell.text)}</{tag}>")
                table_parts.append("</tr>")
            
            table_parts.append("</tbody></table>")
            
            # FAQ cevabı içinde bir tablo mu?
            if in_faq_section and current_question:
                current_answer.append("".join(table_parts))
            else:
                html_parts.append("".join(table_parts))
    
    # Son soruyu da ekleyelim
    if current_question and current_answer:
        faq_questions.append((current_question, "".join(current_answer)))
    
    # FAQ'ları HTML'e ekle
    if faq_questions:
        # FAQ schema.org seçeneği aktifse sadece JSON-LD ekle, normal HTML div'leri ekleme
        if use_faq_schema:
            # FAQPage JSON-LD şeması oluştur (Google için doğru format)
            faq_json = {
                "@context": "https://schema.org",
                "@type": "FAQPage",
                "mainEntity": []
            }
            
            for question, answer in faq_questions:
                # HTML öğeleri çıkar
                soup = BeautifulSoup(answer, 'html.parser')
                clean_answer = soup.get_text().strip()
                
                # JSON-LD veri yapısına ekle
                faq_json["mainEntity"].append({
                    "@type": "Question",
                    "name": question,
                    "acceptedAnswer": {
                        "@type": "Answer",
                        "text": clean_answer
                    }
                })
            
            # JSON-LD script tag'ini ekle
            import json
            json_ld_script = f'<script type="application/ld+json">\n{json.dumps(faq_json, indent=2, ensure_ascii=False)}\n</script>'
            html_parts.append(json_ld_script)
        else:
            # FAQ bölümünü normal HTML olarak ekle (FAQs görünür olmalı)
            for question, answer in faq_questions:
                html_parts.append(f'''
                <div class="faq-item">
                    <h3 class="faq-question">{html.escape(question)}</h3>
                    <div class="faq-answer">{answer}</div>
                </div>
                ''')
        
        # Sadece HTML div kullanılmışsa (Schema.org aktif değilse) kapatma div'i ekle
        if in_faq_section and not use_faq_schema:
            html_parts.append('</div>')  # FAQ bölümünü kapat
    
    # Tüm HTML parçalarını birleştir
    html_content = "<div>" + "".join(html_parts) + "</div>"
    return html_content

# Create HTML file from converted content
def create_html_file(html_content, filename):
    return f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{filename}</title>
    <style>
        body {{
            color: white;
            background-color: #1e1e1e;
            font-family: Arial, sans-serif;
            margin: 20px;
        }}
        h1, h2, h3, h4, h5, h6 {{
            color: white;
        }}
        p {{
            line-height: 1.6;
        }}
        table {{
            border-collapse: collapse;
            width: 100%;
            margin: 15px 0;
        }}
        th, td {{
            border: 1px solid #555;
            padding: 8px;
            text-align: left;
        }}
        th {{
            background-color: #333;
        }}
    </style>
</head>
<body>
{html_content}
</body>
</html>"""

# Create a DOCX document from HTML content
def create_docx_from_html(html_content, filename):
    doc = Document()
    
    # Parse the HTML content
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Process headings and paragraphs
    for tag in soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p']):
        if tag.name.startswith('h'):
            level = int(tag.name[1])
            paragraph = doc.add_paragraph(tag.text)
            paragraph.style = f'Heading {level}'
        else:
            doc.add_paragraph(tag.text)
    
    # Process tables
    for table_html in soup.find_all('table'):
        rows = table_html.find_all('tr')
        if rows:
            table = doc.add_table(rows=len(rows), cols=len(rows[0].find_all(['td', 'th'])))
            
            # Fill the table with data
            for i, row in enumerate(rows):
                cells = row.find_all(['td', 'th'])
                for j, cell in enumerate(cells):
                    table.cell(i, j).text = cell.text.strip()
    
    # Save the document to a BytesIO object
    docx_file = io.BytesIO()
    doc.save(docx_file)
    docx_file.seek(0)
    
    return docx_file

# Function to get HTML display of content with proper styling
def get_html_display(html_content):
    # Replace HTML tags with their escaped versions to display the HTML code
    escaped_html = html.escape(html_content)
    
    # Create a styled container for the HTML display
    html_display = f"""
    <div style="border: 1px solid #ddd; padding: 15px; border-radius: 5px; margin-bottom: 20px; background-color: #f9f9f9; font-family: monospace; white-space: pre-wrap;">
        {escaped_html}
    </div>
    """
    return html_display

# Create download link for a file
def get_download_link(file_bytes, download_filename, link_text):
    b64 = base64.b64encode(file_bytes).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{download_filename}">{link_text}</a>'
    return href

# Main app header
st.title("HTML Converter v2")
st.markdown("Convert DOCX files to HTML and export as DOCX or Excel - with automatic FAQ detection")

# File upload section
st.header("File Upload")
uploaded_files = st.file_uploader("Select DOCX files", type=["docx"], accept_multiple_files=True)

# FAQ schema option
use_faq_schema = st.checkbox("Use FAQ schema (JSON-LD)", value=False, help="Adds FAQ questions as JSON-LD schema")

if uploaded_files:
    st.success(f"{len(uploaded_files)} files successfully uploaded!")
    
    # Use session state to keep track of uploaded files and converted files
    if 'converted_files' not in st.session_state:
        st.session_state.converted_files = []
    
    # Check if there's a new file processing request or we're using cached results
    process_files = st.button("Process Files")
    
    # Initialize converted_files with session state data (if available)
    if process_files:
        converted_files = []
    else:
        converted_files = st.session_state.converted_files
        
    # Only process files if the button was clicked or if we have no converted files yet
    if process_files or (len(converted_files) == 0 and len(st.session_state.converted_files) > 0):
        
        # Show progress
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for i, uploaded_file in enumerate(uploaded_files):
            try:
                # Update progress and status
                status_text.text(f"Processing {uploaded_file.name}...")
                progress_bar.progress((i) / len(uploaded_files))
                
                # Process the file
                doc = Document(uploaded_file)
                html_content = convert_docx_to_html_direct(doc, use_faq_schema)
                
                converted_files.append({
                    "filename": uploaded_file.name,
                    "html_content": html_content,
                    "doc_object": doc
                })
                
            except Exception as e:
                st.error(f"Error processing {uploaded_file.name}: {str(e)}")
        
        # Complete progress bar
        progress_bar.progress(1.0)
        status_text.text("Processing complete!")
        
        # Save converted files in session state to persist across interactions
        st.session_state.converted_files = converted_files
        
        # Display results if files were processed
        if converted_files:
            st.header("Conversion Results")
            
            # Create tabs for each file
            tabs = st.tabs([file["filename"] for file in converted_files])
            
            for i, tab in enumerate(tabs):
                with tab:
                    file_data = converted_files[i]
                    st.markdown("### HTML Code View")
                    # Show the raw HTML code
                    st.code(file_data["html_content"], language="html")
                    
                    # HTML content available for reference
                    html_code = file_data["html_content"]
                    
                    # Also show rendered preview
                    st.markdown("### HTML Rendered Preview")
                    # Wrap HTML content in a full HTML document for proper rendering
                    html_file_for_render = create_html_file(file_data["html_content"], file_data["filename"])
                    st.components.v1.html(html_file_for_render, height=400, scrolling=True)
            
            # Export section
            st.header("Export Options")
            
            # Excel export section
            st.subheader("Export as Excel File")
            
            # Auto-generate Excel when files are processed
            # Create DataFrame for Excel with same format as desktop app
            data = []
            for file in converted_files:
                # Use HTML content without modifications for Excel
                # This preserves JSON-LD script tags and tables
                clean_html = file["html_content"]
                
                data.append({
                    "File Name": file["filename"],
                    "Content": clean_html
                })
            
            df = pd.DataFrame(data)
            
            # Convert to Excel
            excel_buffer = io.BytesIO()
            writer = pd.ExcelWriter(excel_buffer, engine='xlsxwriter')
            df.to_excel(writer, sheet_name='Data', index=False)
            
            # Adjust column widths
            worksheet = writer.sheets['Data']
            worksheet.set_column(0, 0, 30)   # File Name column
            worksheet.set_column(1, 1, 120)  # Content column
            
            writer.close()
            excel_buffer.seek(0)
            
            # Create a standard download button for Excel with callback
            st.download_button(
                label="Download Excel File",
                data=excel_buffer.getvalue(),
                file_name="html_conversion.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="excel_download_button",
                on_click=lambda: st.success("Excel file downloaded.")
            )

else:
    # Show instructions when no files are uploaded
    st.info("Please upload one or more DOCX files to convert to HTML.")
    
    # Display example
    st.markdown("""
    ### How to use this tool:
    1. Upload one or more DOCX files using the file uploader above
    2. Click the 'Process Files' button to convert them to HTML
    3. View HTML code and rendered preview for each file
    4. Choose from export options:
       - Export as Excel file to save all HTML content

    
    ### New Feature: Automatic FAQ Detection
    - If there are at least 3 consecutive headings ending with a question mark (?) in the **LAST PART** of the document,
      this section is automatically detected as FAQ (Frequently Asked Questions)
    - When "Use FAQ schema" option is checked, questions and answers are added in JSON-LD format in a script tag
    - This feature automatically adds structured data markup important for SEO
    - Only question headings in the last 1/3 of the document and consecutive questions are considered as FAQs
    
    ### Additional Feature: FAQ Extraction Script
    - You can use the `extract_faq.py` script to export the FAQ section of the document as a separate markdown file
    - Usage: `python extract_faq.py file.docx [output_folder]`
    - Each FAQ question-answer pair is exported in markdown format as separate headings
    """)

# Footer
st.markdown("---")
st.markdown("""
### About
Browser-based HTML editors always focus on individual editing but don't support bulk translation and conversion. So, we thought, "Why not?" and developed this. A fast, practical, and functional HTML translation/editor tool.

**Burak Sevinmiş**  
SEO Executive - Opsis Digital

---
*A tool that speeds up and automates the HTML translation process. Supports both manual and automatic input. Developed by Opsis Digital.*
""")
