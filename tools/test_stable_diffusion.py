import os
import requests
import io
import sys
import importlib
from PIL import Image, ImageDraw, ImageFont, ImageFilter

def load_env_token():
    """
    從 node_modules/.env 檔案讀取 hf_token 密鑰
    """
    current_dir = os.path.dirname(os.path.abspath(__file__))
    env_path = os.path.join(current_dir, "..", "node_modules", ".env")
    if os.path.exists(env_path):
        try:
            with open(env_path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith("#"):
                        continue
                    if "=" in line:
                        k, v = line.split("=", 1)
                        if k.strip().lower() == "hf_token":
                            return v.strip().strip('"').strip("'")
        except Exception as e:
            print(f"讀取 .env 時發生錯誤: {e}")
    return None

def generate_via_huggingface(prompt, token=None):
    """
    方案 A：使用 Hugging Face 免費的 Inference API 呼叫 Stable Diffusion XL
    """
    print("\n--- 正在使用 Hugging Face API 生成圖像 ---")
    API_URL = "https://api-inference.huggingface.co/models/stabilityai/stable-diffusion-xl-base-1.0"
    
    headers = {}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    elif os.environ.get("HF_TOKEN"):
        headers["Authorization"] = f"Bearer {os.environ.get('HF_TOKEN')}"
    else:
        print("提示：未提供 Hugging Face Token，將使用匿名請求（可能會有限流限制）。")
        
    payload = {
        "inputs": prompt,
        "parameters": {
            "negative_prompt": "blurry, low quality, deformed, distorted, bad anatomy",
            "width": 1024,
            "height": 1024
        }
    }
    
    try:
        response = requests.post(API_URL, headers=headers, json=payload, timeout=60)
        
        if response.status_code == 503:
            print("模型正在 Hugging Face 伺服器上初始化，請等待約 20 秒後重新嘗試...")
            return None
            
        if response.status_code != 200:
            print(f"呼叫失敗，狀態碼: {response.status_code}")
            print(f"錯誤訊息: {response.text}")
            return None
            
        image_bytes = response.content
        image = Image.open(io.BytesIO(image_bytes))
        return image
        
    except Exception as e:
        print(f"發送請求時發生異常: {e}")
        return None

def generate_via_local_gpu(prompt):
    """
    方案 B：使用本地 GPU 執行 Stable Diffusion (需安裝 torch, diffusers, transformers)
    """
    print("\n--- 正在偵測本地 GPU 環境 ---")
    try:
        torch = importlib.import_module("torch")
        diffusers = importlib.import_module("diffusers")
        DiffusionPipeline = diffusers.DiffusionPipeline
    except ImportError:
        print("錯誤：本地未安裝 `torch` 或 `diffusers` 庫。")
        print("請執行以下命令安裝本地運行所需套件（需要較強的 NVIDIA 顯卡）：")
        print("pip install torch torchvision --index-url https://download.pytorch.org/whl/cu121")
        print("pip install diffusers transformers accelerate")
        return None

    if not torch.cuda.is_available():
        print("錯誤：本地未偵測到 CUDA (NVIDIA GPU)，無法在本地高效執行 Stable Diffusion。")
        return None
        
    print(f"偵測到 GPU: {torch.cuda.get_device_name(0)}")
    print("正在下載/載入 Stable Diffusion XL Base 模型 (檔案大小約 6.5GB，首次運行需耐候)...")
    
    try:
        pipe = DiffusionPipeline.from_pretrained(
            "stabilityai/stable-diffusion-xl-base-1.0", 
            torch_dtype=torch.float16, 
            use_safetensors=True, 
            variant="fp16"
        )
        pipe = pipe.to("cuda")
        
        print("模型載入成功，開始生成圖像（預設 30 步）...")
        image = pipe(
            prompt=prompt,
            negative_prompt="blurry, low quality, deformed, distorted, bad anatomy",
            num_inference_steps=30,
            width=1024,
            height=1024
        ).images[0]
        
        return image
        
    except Exception as e:
        print(f"本地生成時發生異常: {e}")
        return None

def wrap_text(text, font, max_width):
    """
    文字自動換行輔助函數
    """
    lines = []
    current_line = ""
    for char in text:
        test_line = current_line + char
        bbox = font.getbbox(test_line)
        w = bbox[2] - bbox[0]
        if w <= max_width:
            current_line = test_line
        else:
            lines.append(current_line)
            current_line = char
    if current_line:
        lines.append(current_line)
    return lines

def draw_card_info(draw, num_str, title_str, caption_str, font_num, font_title, font_caption, start_x, end_x):
    """
    繪製卡片內的編號、標題與小字圖說
    """
    # 1. 繪製編號
    draw.text((start_x, 24), num_str, font=font_num, fill="#C5A059")
    
    # 2. 自動換行並繪製步驟標題 (28px)
    max_w = end_x - start_x
    title_lines = wrap_text(title_str, font_title, max_w)
    
    y_pos = 80
    for line in title_lines:
        draw.text((start_x, y_pos), line, font=font_title, fill="#2A2421")
        bbox = font_title.getbbox(line)
        h = bbox[3] - bbox[1]
        y_pos += h + 6
        
    # 3. 自動換行並繪製底下小字圖說 (20px)，讓其置左對齊大標題
    caption_lines = wrap_text(caption_str, font_caption, max_w)
    caption_y = 205
    for line in caption_lines:
        draw.text((start_x, caption_y), line, font=font_caption, fill="#6A5D59")
        bbox = font_caption.getbbox(line)
        h = bbox[3] - bbox[1]
        caption_y += h + 5

def draw_bento_card(bg_canvas, img_path, card_num_str, title_str, caption_str, x, y, side="left"):
    """
    繪製 Bento Grid 卡片並貼上主畫布 (寬680px, 高280px)
    """
    # 1. 建立 680x280 圓角卡片畫布
    card = Image.new("RGBA", (680, 280), (255, 255, 255, 0))
    card_draw = ImageDraw.Draw(card)
    
    # 2. 繪製半透明白色圓角底板 (92% 不透明度) 加上淡金色邊框
    card_draw.rounded_rectangle(
        [(0, 0), (680, 280)], 
        radius=16, 
        fill=(255, 255, 255, 240), # 提高背景卡片不透明度到 94% 以提供更純淨的商業質感
        outline=(140, 94, 71, 38), 
        width=1
    )
    
    # 3. 貼上實景圖並施加單向融合羽化 (使用 ImageOps.fit 進行等比裁切縮放，保證不拉伸變形)
    if os.path.exists(img_path):
        try:
            src_img = Image.open(img_path).convert("RGBA")
            from PIL import ImageOps
            src_resized = ImageOps.fit(src_img, (380, 280), Image.Resampling.LANCZOS)
            
            # 建立 380x280 的羽化遮罩，讓靠近中央的圖片邊緣淡出融入卡片白色底色中
            img_mask = Image.new("L", (380, 280), 255)
            mask_draw = ImageDraw.Draw(img_mask)
            
            if side == "left":
                # 圖片在右 (300, 0)。左端 (圖片內側邊緣) 要向左融合淡出，所以左側 200px 漸變
                for px in range(200):
                    alpha = int((px / 200) * 255)
                    mask_draw.line([(px, 0), (px, 280)], fill=alpha)
                card.paste(src_resized, (300, 0), img_mask)
            else:
                # 圖片在左 (0, 0)。右端 (圖片內側邊緣) 要向右融合淡出，所以右側 200px 漸變
                for px in range(200):
                    alpha = int(((200 - px) / 200) * 255)
                    mask_draw.line([(380 - 200 + px, 0), (380 - 200 + px, 280)], fill=alpha)
                card.paste(src_resized, (0, 0), img_mask)
        except Exception as e:
            print(f"貼上卡片實景圖失敗 ({card_num_str}): {e}")
            
    # 4. 繪製文字 (載入 Windows 內建微軟正黑體)
    font_path_zh = "C:/Windows/Fonts/msjhbd.ttc"
    if not os.path.exists(font_path_zh):
        font_path_zh = "C:/Windows/Fonts/msjh.ttc"
        
    try:
        font_num = ImageFont.truetype(font_path_zh, 40)
        font_title = ImageFont.truetype(font_path_zh, 28) # 標題字 28px
        font_caption = ImageFont.truetype(font_path_zh, 20) # 底下圖說小字 20px
    except Exception as e:
        print(f"載入正黑體失敗: {e}")
        font_num = ImageFont.load_default()
        font_title = ImageFont.load_default()
        font_caption = ImageFont.load_default()
        
    if side == "left":
        # 資訊欄在左
        draw_card_info(card_draw, card_num_str, title_str, caption_str, font_num, font_title, font_caption, 24, 276)
    else:
        # 資訊欄在右
        draw_card_info(card_draw, card_num_str, title_str, caption_str, font_num, font_title, font_caption, 404, 656)
        
    # 5. 以 Alpha 混合將卡片貼上主畫布
    bg_canvas.paste(card, (x, y), card)

def draw_text_with_outline(draw, text, position, font, text_color, outline_color, outline_width=2):
    """
    文字白色描邊，提升在複雜背景上的對比度
    """
    x, y = position
    for dx in range(-outline_width, outline_width + 1):
        for dy in range(-outline_width, outline_width + 1):
            if dx*dx + dy*dy <= outline_width*outline_width:
                draw.text((x + dx, y + dy), text, font=font, fill=outline_color)
    draw.text((x, y), text, font=font, fill=text_color)

def synthesize_final_image(sd_image, target_dir):
    """
    進行 Pillow 圖像合成 (乾淨米白背景 + 左右六格卡片 + 順序調整 + 中央三件套去背 Logo)
    """
    print("\n--- 開始使用 Pillow 進行卡片拼貼圖像合成與文字排版 (純米色背景版) ---")
    
    # 1. 建立 1920x1080 滿版畫布，底色為真靈光官方米白 #FAF7F2
    bg_canvas = Image.new("RGB", (1920, 1080), "#FAF7F2")
    print("已建立乾淨米色背景 (1920x1080 px)")
    
    # 1.5 繪製中央區域底下的圓形柔和白色/金色光芒 (不遮擋左右兩側卡片圖片，放大到超過 Logo)
    try:
        # 建立一個發光圖層 (RGBA)
        glow_layer = Image.new("RGBA", (1920, 1080), (255, 255, 255, 0))
        glow_draw = ImageDraw.Draw(glow_layer)
        # 在中央繪製淡金色圓盤 (半徑 220px，直徑 440px)
        glow_draw.ellipse([(960 - 220, 540 - 220), (960 + 220, 540 + 220)], fill=(254, 240, 208, 160))
        # 在中央繪製內層更亮的乳白色圓盤 (半徑 160px，直徑 320px)
        glow_draw.ellipse([(960 - 160, 540 - 160), (960 + 160, 540 + 160)], fill=(255, 255, 255, 210))
        # 施加高斯模糊，形成向外擴散的柔和光暈 (半徑約 30px)
        glow_blurred = glow_layer.filter(ImageFilter.GaussianBlur(30))
        # 貼上主畫布
        bg_canvas.paste(glow_blurred, (0, 0), glow_blurred)
        print("已成功在背景中央繪製圓形柔和金色光芒 (確保不遮擋兩側卡片)")
    except Exception as e:
        print(f"繪製中央光芒時發生錯誤: {e}")
    
    # 2. 載入並合成中央的去背 Logo 品牌露出 (依照草圖擺放三件套：Logo -> 中文 -> 英文)
    logo_path = r"d:\Git\true-soul-light-cowork\真靈光企劃書\線上版企劃書V1\真靈光 Logo.png"
    zh_path = r"d:\Git\true-soul-light-cowork\真靈光企劃書\線上版企劃書V1\真靈光_中文去背.png"
    en_path = r"d:\Git\true-soul-light-cowork\真靈光企劃書\線上版企劃書V1\真靈光_英文去背.png"
    
    if os.path.exists(logo_path) and os.path.exists(zh_path) and os.path.exists(en_path):
        try:
            # 1. 太陽小人 Logo (高160px)
            logo_img = Image.open(logo_path).convert("RGBA")
            logo_img.thumbnail((300, 160), Image.Resampling.LANCZOS)
            bg_canvas.paste(logo_img, (960 - logo_img.width // 2, 381), logo_img)
            
            # 2. 中文去背標題 (高75px)
            zh_img = Image.open(zh_path).convert("RGBA")
            zh_img.thumbnail((300, 75), Image.Resampling.LANCZOS)
            bg_canvas.paste(zh_img, (960 - zh_img.width // 2, 561), zh_img)
            
            # 3. 英文去背標題 (高50px)
            en_img = Image.open(en_path).convert("RGBA")
            en_img.thumbnail((300, 50), Image.Resampling.LANCZOS)
            bg_canvas.paste(en_img, (960 - en_img.width // 2, 648), en_img)
            
            print("已成功合成真靈光去背三件套 (Logo、中文、英文) 於中央正圓")
        except Exception as e:
            print(f"合成中央 Logo 時發生錯誤: {e}")
    else:
        print("警告：找不到去背 Logo 圖檔，將跳過中央 Logo 合成。")
        
    # 3. 繪製六張卡片
    # 依照使用者要求順序：左側由上到下為 01, 02, 03；右側由上到下為 04, 05, 06
    # 左側三張 (邊距40px)
    draw_bento_card(bg_canvas, r"d:\Git\true-soul-light-cowork\真靈光企劃書\招募流程示意\01-見核心團隊、高級辦公室、助理.png", 
                    "01", "見核心團隊與高級辦公室與助理", "覺得公司有規模", 40, 60, "left")
    draw_bento_card(bg_canvas, r"d:\Git\true-soul-light-cowork\真靈光企劃書\招募流程示意\02-見咖啡廳與教室全滿.png", 
                    "02", "見咖啡廳與教室全滿", "覺得行銷不錯", 40, 400, "left")
    draw_bento_card(bg_canvas, r"d:\Git\true-soul-light-cowork\真靈光企劃書\招募流程示意\03-見診所與諮詢全滿.png", 
                    "03", "診所與諮詢全滿", "被行銷實力吸引", 40, 740, "left")
                    
    # 右側三張 (邊距40px，各自的標題與圖說已上下對調)
    draw_bento_card(bg_canvas, r"d:\Git\true-soul-light-cowork\真靈光企劃書\招募流程示意\04-見專業拍攝製作團隊.png", 
                    "04", "專業製作團隊", "願意給我們製作課程與代操自媒體", 1200, 60, "right")
    draw_bento_card(bg_canvas, r"d:\Git\true-soul-light-cowork\真靈光企劃書\招募流程示意\05-見完整課程與實習系統.jpg", 
                    "05", "完整課程與實習", "願意跟系統", 1200, 400, "right")
    draw_bento_card(bg_canvas, r"d:\Git\true-soul-light-cowork\真靈光企劃書\招募流程示意\06-見 IP 爆紅與出書.png", 
                    "06", "IP爆紅+出書", "願意給我們與代操自媒體與經紀", 1200, 740, "right")

    # 4. 頂部裝飾標題撤除 (依使用者要求「左上角那行字可以不用」)
    pass

    # 5. 儲存最終合成大圖
    final_output_path = os.path.join(target_dir, "招募流程示意_視覺優化版_AI融合.png")
    bg_canvas.save(final_output_path)
    
    print("=" * 60)
    print(f"[成功] 滿版圖像合成完畢！")
    print(f"成品已儲存至：{final_output_path}")
    print("=" * 60)

def main():
    # 測試 Prompt：真靈光簡報專用淺色背景底圖
    default_prompt = "A minimalist, premium light warm background for a powerpoint slide, very soft light gold gradients, pale cream color (#FAF7F2), faint abstract fluid shapes, empty clean center for layout, 8k, high-end design"
    
    print("=" * 60)
    print(" Stable Diffusion 圖像生成測試腳本 ")
    print("=" * 60)
    
    prompt = input(f"請輸入您的畫作描述 (Prompt) \n[預設: {default_prompt}]: \n").strip()
    if not prompt:
        prompt = default_prompt
        
    print("\n請選擇執行模式：")
    print("1) Hugging Face 免費雲端 API 模式 (建議，速度快，無需顯卡)")
    print("2) 本地 GPU 運算模式 (需安裝 PyTorch & Diffusers，需要強大 NVIDIA 顯卡)")
    choice = input("您的選擇 (1 或 2，預設 1): ").strip()
    
    image = None
    if choice == "2":
        image = generate_via_local_gpu(prompt)
    else:
        hf_token = load_env_token()
        if hf_token:
            print(f"已成功從 node_modules/.env 載入 Hugging Face Token (前4碼: {hf_token[:4]}...)")
        else:
            print("\n[提示] 未偵測到有效的 Hugging Face Token。")
            
        image = generate_via_huggingface(prompt, token=hf_token if hf_token else None)
        
    if image:
        bg_dir = r"d:\Git\true-soul-light-cowork\真靈光企劃書\招募流程示意"
        os.makedirs(bg_dir, exist_ok=True)
        # 儲存原始 SD 背景圖
        sd_bg_path = os.path.join(bg_dir, "sd_background.png")
        image.save(sd_bg_path)
        print(f"已儲存 Stable Diffusion 背景圖至: {sd_bg_path}")
        
        # 進行 Pillow 圖像合成 (滿版擴展 + Logo 疊加 + 文字環繞)
        synthesize_final_image(image, bg_dir)
    else:
        print("[失敗] 圖像生成失敗。請檢查您的網絡連接或本地套件狀態。")

if __name__ == "__main__":
    main()
