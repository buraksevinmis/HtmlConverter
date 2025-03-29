import os
import sys
from docx import Document
import html
from bs4 import BeautifulSoup

def extract_faq_from_docx(docx_path, output_folder=None):
    """
    DOCX dosyasından FAQ (Sık Sorulan Sorular) bölümünü çıkarır ve ayrı bir metin dosyasına kaydeder.
    
    Args:
        docx_path (str): İşlenecek DOCX dosyasının yolu
        output_folder (str, optional): Çıktıların kaydedileceği klasör. Belirtilmezse, aynı klasöre kaydedilir.
        
    Returns:
        str: FAQ içeriği (bulunamazsa boş metin döner)
    """
    try:
        doc = Document(docx_path)
        
        # DOCX dosyasının elementlerini al ve sırala
        elements = []
        
        for paragraph in doc.paragraphs:
            elements.append(('paragraph', paragraph))
        
        # Tabloları da ekle (ve koru)
        for table in doc.tables:
            elements.append(('table', table))
        
        # Elemanları doküman içindeki sıralarına göre sırala
        elements.sort(key=lambda x: get_element_index(x[1], doc))
        
        # Belgedeki başlıkları bul
        heading_elements = [(i, elem) for i, (elem_type, elem) in enumerate(elements) 
                           if elem_type == 'paragraph' and elem.style and elem.style.name.startswith('Heading')]
        
        # Olası FAQ başlıkları (soru işareti ile biten başlıklar)
        faq_candidates = []
        for i, elem in heading_elements:
            if elem.text.strip().endswith('?'):
                faq_candidates.append((i, elem))
        
        # En az 3 adet ? ile biten başlık var mı? (belgenin sonunda olmalı)
        faq_section = None
        faq_questions_indices = []
        
        if len(faq_candidates) >= 3:
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
        
        # FAQ içeriğini oluştur
        faq_content = ""
        
        if faq_section is not None:
            # FAQ verileri
            in_faq_section = False
            current_question = None
            faq_questions = []  # Soru-cevap çiftlerini tutar
            current_answer = []  # Mevcut cevabın paragraflarını tutar
            
            # Sırayla elemanları işle
            for i, (elem_type, elem) in enumerate(elements):
                # FAQ bölümündeyiz mi kontrol et
                if faq_section is not None and i >= faq_section and not in_faq_section:
                    in_faq_section = True
                
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
                            faq_questions.append((current_question, current_answer))
                            current_answer = []
                        
                        # Yeni soruyu kaydet
                        current_question = text
                    elif in_faq_section and current_question and not is_heading:
                        # Bu, mevcut sorunun cevabının bir parçası
                        current_answer.append((elem_type, text))
                
                elif elem_type == 'table' and in_faq_section and current_question:
                    # Tablo, mevcut sorunun cevabına ait
                    table = elem
                    table_data = []
                    
                    for row in table.rows:
                        row_data = [cell.text for cell in row.cells]
                        table_data.append(row_data)
                    
                    current_answer.append((elem_type, table_data))
            
            # Son soruyu da ekleyelim
            if current_question and current_answer:
                faq_questions.append((current_question, current_answer))
            
            # FAQ içeriğini oluştur
            if faq_questions:
                faq_content = f"# FAQ - {os.path.basename(docx_path)}\n\n"
                
                for question, answer_parts in faq_questions:
                    faq_content += f"## {question}\n\n"
                    
                    for part_type, part_content in answer_parts:
                        if part_type == 'paragraph':
                            faq_content += f"{part_content}\n\n"
                        elif part_type == 'table':
                            faq_content += "### Tablo İçeriği:\n"
                            
                            for row in part_content:
                                faq_content += " | ".join(row) + "\n"
                            
                            faq_content += "\n"
                
                # Çıktıyı kaydet
                if faq_content:
                    base_name = os.path.splitext(os.path.basename(docx_path))[0]
                    
                    if output_folder:
                        if not os.path.exists(output_folder):
                            os.makedirs(output_folder)
                        
                        output_path = os.path.join(output_folder, f"{base_name}_FAQ.md")
                    else:
                        output_path = os.path.join(os.path.dirname(docx_path), f"{base_name}_FAQ.md")
                    
                    with open(output_path, 'w', encoding='utf-8') as f:
                        f.write(faq_content)
                    
                    print(f"FAQ içeriği başarıyla kaydedildi: {output_path}")
        
        return faq_content
    except Exception as e:
        print(f"Hata: {str(e)}")
        return ""

def get_element_index(element, doc):
    """Doküman içindeki bir elemanın sırasını belirle"""
    if hasattr(element, '_element'):
        # Paragraf veya tablo için
        for i, child in enumerate(doc._element.body):
            if element._element == child:
                return i
    
    # Varsayılan (bulunamazsa sona ekle)
    return 999999

def main():
    if len(sys.argv) < 2:
        print("Kullanım: python extract_faq.py <docx_dosyası> [çıktı_klasörü]")
        return
    
    docx_path = sys.argv[1]
    output_folder = sys.argv[2] if len(sys.argv) > 2 else None
    
    if not os.path.exists(docx_path):
        print(f"Hata: Dosya bulunamadı - {docx_path}")
        return
    
    extract_faq_from_docx(docx_path, output_folder)

if __name__ == "__main__":
    main()