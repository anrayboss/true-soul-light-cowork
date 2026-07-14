import os
import subprocess
import sys
from urllib.request import pathname2url

def find_browser():
    # 尋找 Windows 上可能的 Edge 或 Chrome 瀏覽器路徑
    paths = [
        # Microsoft Edge
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        os.path.expandvars(r"%LocalAppData%\Microsoft\Edge\Application\msedge.exe"),
        # Google Chrome
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe")
    ]
    
    for path in paths:
        if os.path.exists(path):
            return path
            
    # 如果找不到，嘗試使用系統 PATH 中的執行檔
    for cmd in ["msedge.exe", "chrome.exe", "msedge", "chrome"]:
        try:
            subprocess.run([cmd, "--version"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            return cmd
        except FileNotFoundError:
            continue
            
    return None

def main():
    html_path = r"d:\Git\true-soul-light-cowork\真靈光企劃書\招募流程示意\招募流程示意_視覺優化.html"
    output_path = r"d:\Git\true-soul-light-cowork\真靈光企劃書\招募流程示意\招募流程示意_視覺優化版.png"
    
    if len(sys.argv) > 1:
        html_path = sys.argv[1]
    if len(sys.argv) > 2:
        output_path = sys.argv[2]
        
    if not os.path.exists(html_path):
        print(f"Error: HTML file not found at {html_path}")
        sys.exit(1)
        
    browser_path = find_browser()
    if not browser_path:
        print("Error: Could not find Microsoft Edge or Google Chrome executable.")
        sys.exit(1)
        
    print(f"Using browser: {browser_path}")
    
    # 轉換成 file:// URL
    file_url = f"file://{pathname2url(html_path)}"
    print(f"Loading URL: {file_url}")
    
    # Headless 截圖指令
    # --virtual-time-budget=2000 給予瀏覽器 2 秒的模擬運行時間，確保圖片載入與 CSS 渲染完畢
    # --hide-scrollbars 避免截圖中出現捲軸
    # --force-device-scale-factor=1 確保 1:1 的畫素比例 (1920x1080)
    cmd = [
        browser_path,
        "--headless",
        "--disable-gpu",
        "--window-size=1920,1080",
        "--force-device-scale-factor=1",
        "--hide-scrollbars",
        "--virtual-time-budget=2000",
        f"--screenshot={output_path}",
        file_url
    ]
    
    print("Capturing screenshot...")
    try:
        # 執行截圖
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=15)
        if result.returncode == 0 or os.path.exists(output_path):
            print(f"Success! Screenshot saved to: {output_path}")
        else:
            print("Failed to capture screenshot.")
            print(f"STDOUT: {result.stdout}")
            print(f"STDERR: {result.stderr}")
            sys.exit(1)
    except subprocess.TimeoutExpired:
        print("Error: Screenshot process timed out.")
        sys.exit(1)
    except Exception as e:
        print(f"An error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
