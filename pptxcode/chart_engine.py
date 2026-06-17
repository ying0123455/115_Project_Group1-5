import re
import json
import requests
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

class ChartEngine:
    def __init__(self, ollama_url="http://163.13.143.67:1234/api/chat", ollama_model="gpt-oss:120b"):
        self.ollama_url = ollama_url
        self.ollama_model = ollama_model
        self.supported_charts = {
            "PIE": XL_CHART_TYPE.PIE,
            "RADAR": XL_CHART_TYPE.RADAR_FILLED,
            "LINE": XL_CHART_TYPE.LINE,
            "BAR": XL_CHART_TYPE.COLUMN_CLUSTERED
        }

    def _call_ollama(self, prompt, temperature=0.2):
        """內部共用的 LLM 呼叫函數，自動相容 /api/chat 與 /api/generate"""
        is_generate = "generate" in self.ollama_url
        if is_generate:
            payload = {
                "model": self.ollama_model,
                "prompt": prompt,
                "stream": False,
                "options": {"temperature": temperature}
            }
        else:
            payload = {
                "model": self.ollama_model,
                "messages": [{"role": "user", "content": prompt}],
                "stream": False,
                "options": {"temperature": temperature}
            }
        try:
            r = requests.post(self.ollama_url, json=payload, timeout=120)
            r.raise_for_status()
            res_json = r.json()
            if "message" in res_json:
                return res_json["message"]["content"]
            elif "response" in res_json:
                return res_json["response"]
            return None
        except Exception as e:
            print(f"❌ Ollama 請求失敗: {e}")
            return None

    def extract_chart_types(self, v_hint):
        """從視覺提示中提取所有圖表類型"""
        prompt = f"""
        根據以下視覺建議文字，請列出所有應使用的圖表類型，僅回傳以下之一或多個（以逗號分隔）：PIE、RADAR、LINE、BAR，若無則回傳 NONE。文字：
        { v_hint }
        """
        result = self._call_ollama(prompt)
        if not result: return []
        
        candidates = re.split(r"[,,，]+", result.strip())
        return [c.upper() for c in candidates if c.upper() in self.supported_charts]

    def extract_chart_data(self, v_hint, chart_key, document_summary):
        """根據文本摘要與視覺提示，解析出圖表的類別與數值"""
        prompt = f"""
        根據以下摘要，請為 {chart_key} 圖表提供適合的類別與數值資料。請以 JSON 格式回傳，鍵名分別為 "categories"（字串列表）與 "values"（數值列表）。如果摘要中未包含相關資料，回傳空的 JSON 物件 {{}}。
        摘要：
        {document_summary}
        視覺提示：
        {v_hint}
        """
        result = self._call_ollama(prompt)
        if not result: return {}

        try:
            # 嘗試尋找並解析 JSON 區塊
            match = re.search(r'\{.*\}', result.replace('\n', ''), re.DOTALL)
            if match:
                data = json.loads(match.group(0))
                if "categories" in data and "values" in data:
                    return data
            return {}
        except Exception as e:
            print(f"❌ JSON 解析失敗（{chart_key}）: {e}")
            return {}

    def analyze_and_get_chart_info(self, v_hint, document_summary):
        """
        整合方法：一次性回傳需要畫的圖表類型與對應的數據
        回傳格式範例：[{"type": "BAR", "data": {"categories": ["A", "B"], "values": [10, 20]}}]
        """
        chart_keys = self.extract_chart_types(v_hint)
        chart_infos = []
        for key in chart_keys:
            data = self.extract_chart_data(v_hint, key, document_summary)
            chart_infos.append({"type": key, "data": data})
        return chart_infos

    def draw_chart_on_slide(self, slide, chart_type_str, chart_data_dict, left, top, width, height):
        """將原生圖表繪製到指定的 PPT 投影片物件上"""
        chart_type = self.supported_charts.get(chart_type_str)
        if not chart_type:
            return False

        chart_data = CategoryChartData()
        if chart_data_dict and chart_data_dict.get("categories") and chart_data_dict.get("values"):
            chart_data.categories = chart_data_dict["categories"]
            chart_data.add_series("Series 1", chart_data_dict["values"])
        else:
            print(f"⚠️ 警告: 未提供有效的數據，使用預設測試資料。")
            chart_data.categories = ["A", "B", "C", "D", "E"]
            chart_data.add_series("預設資料", (10, 20, 15, 30, 35))

        graphic_frame = slide.shapes.add_chart(chart_type, int(left), int(top), int(width), int(height), chart_data)
        chart = graphic_frame.chart
        chart.has_title = False
        chart.has_legend = False
        return True
