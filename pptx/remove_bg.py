import sys
import os
from rembg import remove
from PIL import Image

def remove_background(input_path, output_path):
    """將圖片去背並存檔"""
    print(f"正在讀取圖片: {input_path} ...")
    try:
        input_image = Image.open(input_path)
    except FileNotFoundError:
        print(f"找不到檔案！請確認 {input_path} 是否在同一個資料夾。")
        return

    print("AI 正在為您移除背景中，請稍候...")
    
    output_image = remove(input_image)

    output_image.save(output_path)
    print(f"去背成功！已經為您存成: {output_path}")

if __name__ == "__main__":
    # 如果使用者有在指令裡面輸入檔名，就使用指令裡的檔名
    # 例如： python remove_bg.py in.png out.png
    if len(sys.argv) >= 3:
        input_file = sys.argv[1]
        output_file = sys.argv[2]
    else:
        # 如果沒有輸入，就預設使用這兩個檔案
        input_file = "background3.png"
        output_file = "background3_transparent.png"
        print(f"小提示: 未指定檔名，將預設處理 {input_file}")
    
    remove_background(input_file, output_file)
