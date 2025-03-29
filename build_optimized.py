import os
import sys
import subprocess

def build_exe():
    """
    使用 PyInstaller 打包程序為優化體積的 exe 檔案
    """
    print("開始打包優化版駱駝競速模擬器...")
    
    # 確保所需文件存在
    readme_file = "README.md"
    if not os.path.exists(readme_file):
        with open(readme_file, "w", encoding="utf-8") as f:
            f.write("""# 駱駝競速高級模擬器

這是一個模擬駱駝競賽的進階程式，可以預測不同駱駝獲勝的機率，
提供視覺化界面、統計分析和動態模擬功能。

## 使用方法

1. 在「配置設定」頁面設置駱駝位置和比賽參數
2. 切換到「賽道視圖」可以預覽賽道和駱駝位置
3. 執行單次模擬可以觀看動畫效果
4. 執行多次模擬進行統計分析
5. 在「統計分析」頁面查看比賽結果
6. 可以將配置和結果匯出保存

## 特色功能

- 視覺化賽道和駱駝移動
- 動態模擬單場比賽
- 大量模擬統計分析
- 可自定義駱駝數量、顏色和賽道長度
- 可保存和載入配置
- 結果可匯出為Excel檔案
""")

    # 設定 PyInstaller 參數 (優化體積)
    command = [
        "pyinstaller",
        "--noconfirm",                # 覆蓋輸出目錄
        "--onefile",                  # 生成單一 exe 檔案
        "--windowed",                 # 不顯示控制台窗口
        "--clean",                    # 清理臨時文件
        "--name=駱駝競速高級模擬器",      # 應用程序名稱
        "--add-data=README.md;.",     # 添加額外文件
        
        # 明確包含核心依賴
        "--hidden-import=tkinter",
        "--hidden-import=tkinter.ttk",
        "--hidden-import=tkinter.messagebox",
        "--hidden-import=tkinter.filedialog",
        "--hidden-import=tkinter.colorchooser",
        "--hidden-import=_tkinter",
        "--hidden-import=numpy",
        "--hidden-import=matplotlib",
        "--hidden-import=matplotlib.backends.backend_tkagg",
        "--hidden-import=openpyxl",
        
        # 排除不需要的模塊以減少體積
        "--exclude-module=_ssl",
        "--exclude-module=cryptography",
        "--exclude-module=pydoc",
        "--exclude-module=sqlite3",
        "--exclude-module=pycparser",
        "--exclude-module=curses",
        "--exclude-module=setuptools",
        
        # 優化 matplotlib 大小
        "--exclude-module=matplotlib.tests",
        "--exclude-module=matplotlib.testing",
        "--exclude-module=matplotlib.examples",
        "--exclude-module=scipy",
        
        # 排除多餘的後端
        "--exclude-module=PyQt5",
        "--exclude-module=PySide2",
        "--exclude-module=IPython",
        "--exclude-module=pandas.tests",
        "--exclude-module=numpy.testing",
        
        # 主程序檔案
        "camel_race_advanced.py"
    ]
    
    # 添加自定義引入字體
    font_dir = "fonts"
    if not os.path.exists(font_dir):
        os.makedirs(font_dir)
        
    # 複製Windows字體檔案到本地目錄
    windows_font = "C:/Windows/Fonts/msjh.ttc"
    local_font = os.path.join(font_dir, "msjh.ttc")
    if os.path.exists(windows_font) and not os.path.exists(local_font):
        import shutil
        try:
            shutil.copy2(windows_font, local_font)
            print(f"已複製字體文件: {local_font}")
            # 添加字體文件到打包中
            command.insert(-1, f"--add-data={font_dir};{font_dir}")
        except Exception as e:
            print(f"複製字體文件失敗: {str(e)}")
    
    # 更新字體路徑配置
    with open("camel_race_advanced.py", "r", encoding="utf-8") as f:
        content = f.read()
    
    # 如果複製了字體，則更新路徑
    if os.path.exists(local_font):
        content = content.replace("C:/Windows/Fonts/msjh.ttc", "./fonts/msjh.ttc")
        with open("camel_race_advanced.py", "w", encoding="utf-8") as f:
            f.write(content)
    
    # 運行命令
    subprocess.run(command, check=True)
    
    print("打包完成! 請查看 dist 資料夾中的 '駱駝競速高級模擬器.exe'")

if __name__ == "__main__":
    build_exe() 