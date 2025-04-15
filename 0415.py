import base64
import requests
from pyautocad import Autocad, APoint
import math
import json
from pictool import calculate_x_coordinate,calculate_y_coordinate,process_dwg_within_area, annotate_dimension_and_tolerance
# OpenAI API Key
api_key = "sk-proj-NqUF2ZOTGZ9EKcfiGwDhT3BlbkFJG7n0kP0jNVvpf19kAAAE"

# Function to encode the image
def encode_image(image_path):
  with open(image_path, "rb") as image_file:
    return base64.b64encode(image_file.read()).decode('utf-8')

# Path to your image
image_path = r"C:\Users\user\Desktop\LEDA\Lian Dung\501.jpg"

# Getting the base64 string
base64_image = encode_image(image_path)

headers = {
  "Content-Type": "application/json",
  "Authorization": f"Bearer {api_key}"
}

payload = {
  "model": "gpt-4o",
  "messages": [
    {
      "role": "user",
      "content": [
        {
          "type": "text",
          "text":"這是一個插頭的以正投影第三角投影法做出的視圖圖檔，圖中包含上下前後左右視圖，"
          "視圖通常以十字排列 中間是前視圖、右邊是右視圖、下方是下視圖、左邊是左視圖， 通常左前右視圖會水平繪製、上前下視圖會垂直放置"
          "圖片中已經事先用OD框選出crop 請幫我找內容含有LT-501是哪個crop 以json格式將座標+此視圖應放置在10號位置，依範例格式{'xmin': , 'ymin': , 'xmax': , 'ymax': , 'position': } 回傳,請注意回傳時僅須回傳json內容，不需任何其餘說明"
          "<annotation><folder>ilovepdf_pages-to-jpg</folder><filename>LT-202-UL_page-0001.jpg</filename><path>D:\Downloads\ilovepdf_pages-to-jpg\LT-202-UL_page-0001.jpg</path><source><database>Unknown</database></source><size><width>1755</width><height>1240</height><depth>1</depth></size><segmented>0</segmented><object><name>crop</name><pose>Unspecified</pose><truncated>0</truncated><difficult>0</difficult><bndbox><xmin>236</xmin><ymin>93</ymin><xmax>430</xmax><ymax>220</ymax></bndbox></object><object><name>crop</name><pose>Unspecified</pose><truncated>0</truncated><difficult>0</difficult><bndbox><xmin>210</xmin><ymin>385</ymin><xmax>455</xmax><ymax>835</ymax></bndbox></object><object><name>crop</name><pose>Unspecified</pose><truncated>0</truncated><difficult>0</difficult><bndbox><xmin>770</xmin><ymin>395</ymin><xmax>925</xmax><ymax>820</ymax></bndbox></object><object><name>crop</name><pose>Unspecified</pose><truncated>0</truncated><difficult>0</difficult><bndbox><xmin>1230</xmin><ymin>315</ymin><xmax>1480</xmax><ymax>820</ymax></bndbox></object><object><name>crop</name><pose>Unspecified</pose><truncated>0</truncated><difficult>0</difficult><bndbox><xmin>245</xmin><ymin>920</ymin><xmax>450</xmax><ymax>1050</ymax></bndbox></object></annotation>"
          
            },
        {
          "type": "image_url",
          "image_url": {
            "url": f"data:image/jpeg;base64,{base64_image}"
          }
        }
      ]
    }
  ],
  "max_tokens": 300
}
response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)

print(response.json())
response_json = response.json()
content = response_json['choices'][0]['message']['content']

# 去掉包裹的 ```json 標記，並轉為字典
json_string = content.strip("```json").strip("```").strip()
bounding_box = json.loads(json_string)

source_path = r"C:\Users\user\Desktop\LEDA\Lian Dung\Lian Dung\LT-501-UL.dwg"
target_path = r"C:\Users\user\Desktop\LEDA\Lian Dung\Lian Dung\成品圖框(中文版).dwg"


process_dwg_within_area(source_path, target_path, calculate_x_coordinate(bounding_box["xmin"]),calculate_y_coordinate(bounding_box["ymin"]), calculate_x_coordinate(bounding_box["xmax"]), calculate_y_coordinate(bounding_box["ymax"]),bounding_box["position"])
annotate_dimension_and_tolerance(
    calculate_x_coordinate(bounding_box["xmin"]),
    calculate_x_coordinate(bounding_box["xmax"]),
    calculate_y_coordinate(bounding_box["ymin"])  # 以下邊為基準畫尺寸
)