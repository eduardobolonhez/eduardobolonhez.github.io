import mammoth
import os
import sys
import shutil

def extract_html_from_docx(doc_path):
    if not os.path.exists(doc_path):
        print(f"Warning: Could not find {doc_path}. Skipping.")
        return ""
    
    # Bypass Word's file lock by copying the file first
    temp_path = doc_path + ".temp"
    try:
        shutil.copy2(doc_path, temp_path)
    except Exception as e:
        print(f"Error copying docx {doc_path} (locked?): {e}")
        return ""
    
    html_content = ""
    try:
        with open(temp_path, "rb") as docx_file:
            style_map = "p[style-name='Heading 1'] => h1:fresh\np[style-name='Heading 2'] => h2:fresh\np[style-name='Heading 3'] => h3:fresh\n"
            result = mammoth.convert_to_html(docx_file, style_map=style_map)
            html_content = result.value
            messages = result.messages
            for msg in messages:
                print(f"Mammoth warning [{os.path.basename(doc_path)}]: {msg}")
    except Exception as e:
        print(f"Error reading docx {doc_path}: {e}")
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except:
                pass
                
    return html_content

def main():
    doc_path_pt = r"F:\OneDrive\CV\Antigravity\Referencia.docx"
    doc_path_en = r"F:\OneDrive\CV\Antigravity\References.docx"
    
    template_path = os.path.join(os.path.dirname(__file__), "template.html")
    output_path = os.path.join(os.path.dirname(__file__), "index.html")
    
    print("Extracting Portuguese and English documents...")
    html_pt = extract_html_from_docx(doc_path_pt)
    html_en = extract_html_from_docx(doc_path_en)
        
    try:
        with open(template_path, "r", encoding="utf-8") as template_file:
            template = template_file.read()
    except Exception as e:
        print(f"Error reading template: {e}")
        sys.exit(1)
        
    final_html = template.replace("{{CONTENT_PT}}", html_pt).replace("{{CONTENT_EN}}", html_en)
    
    try:
        with open(output_path, "w", encoding="utf-8") as output_file:
            output_file.write(final_html)
    except Exception as e:
        print(f"Error writing output: {e}")
        sys.exit(1)
        
    print("Success! index.html generated with dual-language support!")

if __name__ == "__main__":
    main()
