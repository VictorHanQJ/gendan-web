import openpyxl
import requests
import json
from io import BytesIO
import os
import tos

# ==========环境变量，全部在Vercel后台配置，不要写死代码========
ARK_API_KEY = os.environ.get("ARK_API_KEY")
MODEL_ID = "doubao-pro-4k"

#火山TOS对象存储配置
TOS_ACCESS_KEY = os.environ.get("TOS_ACCESS_KEY")
TOS_SECRET_KEY = os.environ.get("TOS_SECRET_KEY")
TOS_BUCKET_NAME = os.environ.get("TOS_BUCKET_NAME")
TOS_ENDPOINT = os.environ.get("TOS_ENDPOINT")
TOS_PUBLIC_DOMAIN = os.environ.get("TOS_PUBLIC_DOMAIN")

SYSTEM_PROMPT = """你是服装跟单表格解析专家。输入表格文本，存在合并单元格、跨行备注。
提取字段：款号，PO号，客户名称，颜色，尺码，各尺码数量，订单总数量，交期，供应商，工艺备注，溢装短装约定。
规则：
1. 合并单元格内容向下填充到对应所有行；
2. 工艺备注完整提取印花、绣花、面料、辅料、标签、包装、特殊限制；
3. 数字原样提取，不确定内容标记null；
4. 只输出纯JSON数组，不要markdown、不要多余文字；
5. 额外增加字段"风险标签"，自动标记：交期预警、存在特殊工艺备注、数量异常风险。
"""

def upload_image_to_tos(img_bytes, style_no, color_name):
    """把图片二进制上传TOS，返回公开访问URL"""
    if not all([TOS_ACCESS_KEY,TOS_SECRET_KEY,TOS_BUCKET_NAME,TOS_ENDPOINT,TOS_PUBLIC_DOMAIN]):
        return None
    try:
        client = tos.TosClientV2(TOS_ACCESS_KEY, TOS_SECRET_KEY, TOS_ENDPOINT, TOS_BUCKET_NAME)
        obj_key = f"gendan_img/{style_no}_{color_name}.png"
        client.put_object(obj_key, content=BytesIO(img_bytes))
        img_url = f"https://{TOS_PUBLIC_DOMAIN}/{obj_key}"
        return img_url
    except Exception:
        return None

def handler(event, context):
    from http import HTTPStatus
    headers = {
        "Content-Type": "application/json",
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "POST,OPTIONS",
        "Access-Control-Allow-Headers":"Content-Type"
    }
    if event["method"] == "OPTIONS":
        return {"statusCode":HTTPStatus.NO_CONTENT,"headers":headers,"body":""}
    if event["method"] != "POST":
        return {"statusCode":405,"headers":headers,"body":json.dumps({"error":"仅支持POST请求"})}
    try:
        body_bin = bytes.fromhex(event["body"])
        file_stream = BytesIO(body_bin)
        wb = openpyxl.load_workbook(file_stream, read_only=False, data_only=True)
        all_text = ""
        image_map = {} #key:(sheet,row),value:图片二进制

        #读取文本 + 提取单元格内嵌图片
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            all_text += f"\n====== SHEET:{sheet_name} ======\n"
            for row in ws.iter_rows(values_only=True):
                row_str = "\t".join([str(cell) if cell is not None else "" for cell in row])
                all_text += row_str + "\n"
            #提取本sheet所有内嵌图片
            for img in ws._images:
                row_idx = img.anchor._from.row
                image_map[(sheet_name,row_idx)] = img._data()
        wb.close()

        #截断文本，防止大表格超出模型上下文
        if len(all_text) >12000:
            all_text = all_text[:12000]

        #调用火山方舟豆包API解析
        url = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"
        payload={
            "model":MODEL_ID,
            "messages":[
                {"role":"system","content":SYSTEM_PROMPT},
                {"role":"user","content":all_text}
            ],
            "temperature":0.1
        }
        resp = requests.post(url,headers={"Authorization":f"Bearer {ARK_API_KEY}","Content-Type":"application/json"},json=payload,timeout=55)
        resp.raise_for_status()
        llm_out = resp.json()["choices"][0]["message"]["content"]
        result_json = json.loads(llm_out)

        #匹配图片上传TOS（简易匹配：款号+颜色作为图片key；复杂浮动图片无法精准绑定，原型能力）
        for idx,row_data in enumerate(result_json):
            st_no = str(row_data.get("款号","unk"))
            col = str(row_data.get("颜色","unk"))
            img_url = upload_image_to_tos((body_bin), st_no, col)
            row_data["图片url"] = img_url

        return {
            "statusCode":200,
            "headers":headers,
            "body":json.dumps({"code":0,"data":result_json})
        }
    except Exception as e:
        return {
            "statusCode":500,
            "headers":headers,
            "body":json.dumps({"code":-1,"error":str(e)})
        }
