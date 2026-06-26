import zipfile
import xml.etree.ElementTree as ET
import os

def extract_docx_text(docx_path):
    namespaces = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    }
    
    if not os.path.exists(docx_path):
        print(f"File not found: {docx_path}")
        return ""

    try:
        with zipfile.ZipFile(docx_path) as z:
            xml_content = z.read('word/document.xml')
            root = ET.fromstring(xml_content)
            
            paragraphs = []
            for p in root.findall('.//w:p', namespaces):
                text_runs = p.findall('.//w:t', namespaces)
                p_text = "".join(t.text for t in text_runs if t.text)
                if p_text:
                    paragraphs.append(p_text)
                    
            return "\n\n".join(paragraphs)
    except Exception as e:
        return f"Error extracting DOCX: {str(e)}"

if __name__ == "__main__":
    docx_path = r"c:\Users\ASUS\Desktop\true-soul-light-cowork\真靈光企劃書所有資訊(線上共編處)\真靈光企劃書_正式提案版.docx"
    text = extract_docx_text(docx_path)
    
    output_path = r"c:\Users\ASUS\Desktop\true-soul-light-cowork\真靈光企劃書所有資訊(線上共編處)\真靈光企劃書_正式提案版_extracted.md"
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(text)
    print(f"Extracted text saved to: {output_path}")
