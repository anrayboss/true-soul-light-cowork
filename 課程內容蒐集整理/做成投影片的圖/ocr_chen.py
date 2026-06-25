import os
import re
from PIL import Image
import numpy as np
import easyocr

def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]

def run_ocr():
    folder = r"d:\Git\true-soul-light\不用版控的\做成投影片的圖"
    # Filter only files matching "陳老師行銷內容 (x).jpg"
    files = [f for f in os.listdir(folder) if f.startswith("陳老師行銷內容") and f.lower().endswith((".jpg", ".jpeg", ".png"))]
    files.sort(key=natural_sort_key)
    
    print(f"Loading EasyOCR reader...")
    reader = easyocr.Reader(['ch_tra', 'en'])
    
    output_txt_path = os.path.join(folder, "ocr_chen_raw_results.txt")
    
    with open(output_txt_path, "w", encoding="utf-8") as out:
        for idx, img_name in enumerate(files):
            img_path = os.path.join(folder, img_name)
            print(f"Processing [{idx+1}/{len(files)}]: {img_name}")
            
            with Image.open(img_path) as img:
                img_np = np.array(img.convert('RGB'))
                
            # Run OCR
            results = reader.readtext(img_np)
            
            out.write(f"=== {img_name} ===\n")
            # Sort OCR text regions by top-left y coordinate (approximate row order)
            # then x coordinate
            results_sorted = sorted(results, key=lambda x: (x[0][0][1] // 20, x[0][0][0]))
            
            for res in results_sorted:
                bbox, text, prob = res
                out.write(f"{text} ({prob:.2f}) [box: {bbox}]\n")
            out.write("\n\n")
            
    print(f"OCR results saved to: {output_txt_path}")

if __name__ == '__main__':
    run_ocr()
