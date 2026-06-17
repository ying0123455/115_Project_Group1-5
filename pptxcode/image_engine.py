import os
import sys
from typing import Literal, Dict, Optional

# 保留原作者的安全性檢查跳過設定
try:
    import transformers.modeling_utils
    transformers.modeling_utils.check_torch_load_is_safe = lambda *args, **kwargs: None
except Exception:
    pass

# 確保能讀取到 84KB 專案內的 utils
# 假設這個檔案放在專案根目錄，所以直接 import
from utils.wrapper import StreamDiffusionWrapper

class ImageEngine:
    def __init__(
        self,
        model_id_or_path: str = "KBlueLeaf/kohaku-v2.1",
        lora_dict: Optional[Dict[str, float]] = None,
        width: int = 512,
        height: int = 512,
        acceleration: Literal["none", "xformers", "tensorrt"] = "xformers",
        seed: int = 2
    ):
        """
        初始化時載入模型。這個動作比較花時間，但只會執行一次。
        """
        print("⏳ 正在初始化 StreamDiffusion 圖片生成引擎，請稍候...")
        self.stream = StreamDiffusionWrapper(
            model_id_or_path=model_id_or_path,
            lora_dict=lora_dict,
            t_index_list=[0, 16, 32, 45],
            frame_buffer_size=1,
            width=width,
            height=height,
            warmup=10,
            acceleration=acceleration,
            mode="txt2img",
            use_denoising_batch=False,
            cfg_type="none",
            seed=seed,
        )
        print("✅ 圖片生成引擎準備完畢！")

    def generate_image(self, prompt: str, output_path: str) -> str:
        """
        傳入提示詞，生成圖片並回傳檔案路徑。
        """
        print(f"🎨 正在生成圖片 (Prompt: {prompt})...")
        
        # 準備提示詞與步數
        self.stream.prepare(
            prompt=prompt,
            num_inference_steps=50,
        )

        # StreamDiffusion 的串流推進邏輯
        for _ in range(self.stream.batch_size - 1):
            self.stream()

        # 取得產出的圖片物件 (PIL Image)
        output_image = self.stream()
        
        # 安全機制：確保輸出的資料夾存在，避免存檔時噴錯
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        # 存檔
        output_image.save(output_path)
        print(f"✅ 圖片已儲存至: {output_path}")
        
        return output_path
