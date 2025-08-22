import os
import fitz  # PyMuPDF
import pandas as pd
import xml.etree.ElementTree as ET
import pyreadstat
import zipfile

# ---------------- File Processing Functions ----------------

def replace_text_in_pdf(input_pdf_path, old_text, new_text):
    """Replace text in PDF file and save in same folder"""
    pdf_document = fitz.open(input_pdf_path)
    font_name = "Times-Roman"
    
    for page in pdf_document:
        text_instances = page.search_for(old_text)
        if text_instances:
            original_text_info = page.get_text("dict")['blocks']
            
            for rect in text_instances:
                page.add_redact_annot(rect)
            page.apply_redactions()
            
            for rect in text_instances:
                original_fontsize = 12
                for block in original_text_info:
                    for line in block.get("lines", []):
                        for span in line.get("spans", []):
                            if old_text in span["text"]:
                                original_fontsize = span["size"]
                                break
                        else:
                            continue
                        break
                    else:
                        continue
                    break
                
                font_params = {
                    'fontsize': original_fontsize,
                    'fontname': font_name
                }
                insert_point = fitz.Point(rect.x0, rect.y1 - 2.5)
                page.insert_text(insert_point, new_text, **font_params)
    
    base, ext = os.path.splitext(input_pdf_path)
    output_path = f"{base}_modified{ext}"
    pdf_document.save(output_path)
    pdf_document.close()
    return output_path

def replace_text_in_csv(input_csv_path, old_text, new_text):
    """Replace text in CSV file and save in same folder"""
    df = pd.read_csv(input_csv_path, dtype=str)
    df = df.applymap(lambda x: x.replace(old_text, new_text) if isinstance(x, str) else x)
    
    base, ext = os.path.splitext(input_csv_path)
    output_path = f"{base}_modified{ext}"
    df.to_csv(output_path, index=False)
    return output_path

def replace_text_in_xml(input_xml_path, old_text, new_text):
    """Replace text in XML file and save in same folder"""
    tree = ET.parse(input_xml_path)
    root = tree.getroot()
    
    def replace_in_element(elem):
        if elem.text and old_text in elem.text:
            elem.text = elem.text.replace(old_text, new_text)
        if elem.tail and old_text in elem.tail:
            elem.tail = elem.tail.replace(old_text, new_text)
        for k, v in elem.attrib.items():
            if old_text in v:
                elem.attrib[k] = v.replace(old_text, new_text)
        for child in elem:
            replace_in_element(child)
    
    replace_in_element(root)
    
    base, ext = os.path.splitext(input_xml_path)
    output_path = f"{base}_modified{ext}"
    tree.write(output_path, encoding="utf-8", xml_declaration=True)
    return output_path

def replace_text_in_xpt(input_xpt_path, old_text, new_text):
    """Replace text in XPT file and save in same folder"""
    df, meta = pyreadstat.read_xport(input_xpt_path)
    df = df.applymap(lambda x: x.replace(old_text, new_text) if isinstance(x, str) else x)
    
    base, ext = os.path.splitext(input_xpt_path)
    output_path = f"{base}_modified{ext}"
    pyreadstat.write_xport(df, output_path, file_format_version=8, table_name=meta.table_name)
    return output_path

def process_single_file(file_path, old_text, new_text):
    """Process a single file based on its extension"""
    ext = os.path.splitext(file_path)[1].lower()
    
    if ext == '.pdf':
        return replace_text_in_pdf(file_path, old_text, new_text)
    elif ext == '.csv':
        return replace_text_in_csv(file_path, old_text, new_text)
    elif ext == '.xml':
        return replace_text_in_xml(file_path, old_text, new_text)
    elif ext == '.xpt':
        return replace_text_in_xpt(file_path, old_text, new_text)
    else:
        return None

def extract_zip_and_process(zip_path, old_text, new_text):
    """Extract ZIP file and process all supported files inside, saving modified files in same folder"""
    extract_folder = os.path.splitext(zip_path)[0] + "_extracted"
    os.makedirs(extract_folder, exist_ok=True)
    
    processed_files = []
    supported_extensions = ['.pdf', '.csv', '.xml', '.xpt']
    
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_folder)
    
    for root, dirs, files in os.walk(extract_folder):
        for file in files:
            file_path = os.path.join(root, file)
            ext = os.path.splitext(file)[1].lower()
            
            if ext in supported_extensions:
                try:
                    processed_file = process_single_file(file_path, old_text, new_text)
                    if processed_file:
                        processed_files.append(processed_file)
                except Exception as e:
                    print(f"Error processing {file}: {str(e)}")
                    continue
    
    if processed_files:
        output_zip = zip_path.replace('.zip', '_modified.zip')
        with zipfile.ZipFile(output_zip, 'w') as zip_ref:
            for processed_file in processed_files:
                arcname = os.path.relpath(processed_file, extract_folder)
                zip_ref.write(processed_file, arcname)
        return output_zip
    else:
        return None

# ---------------- Example Usage ----------------
if __name__ == "__main__":
    # Example: replace text in a file
    file_path = "input.pdf"  # change this to your file path
    result = process_single_file(file_path, "CDISC", "CSIDC")
    if result:
        print(f"Modified file saved at: {result}")
