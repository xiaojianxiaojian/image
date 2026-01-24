"""
GLM-4V-Flash 图片OCR识别脚本
使用智谱AI GLM-4V模型识别图片中的医学参数，并将结果输出到Excel
"""

import base64
import io
import json
import os
import re
from zai import ZhipuAiClient
from pathlib import Path

import pandas as pd
from PIL import Image

# 智谱AI API 配置
ZHIPU_API_KEY = os.environ.get("ZHIPU_API_KEY", "api-key-here")
ZHIPU_BASE_URL = "https://open.bigmodel.cn/api/paas/v4"

# 需要提取的参数列表
PARAMETERS = [
    "IVS;d",
    "LVID;d",
    "LVPW;d",
    "IVS;s",
    "LVID;s",
    "LVPW;s",
    "LV Vol;d",
    "LV Vol;s",
    "EF",
    "FS",
    "LV Mass",
    "LV Mass Cor"
]


def encode_image_to_base64(image_path):
    """
    将图片文件编码为base64格式

    Args:
        image_path: 图片文件路径

    Returns:
        base64编码的图片数据和MIME类型
    """
    try:
        with open(image_path, "rb") as image_file:
            # 如果是tif文件，先转换为JPEG以减少大小
            if image_path.lower().endswith(('.tif', '.tiff')):
                with Image.open(image_path) as img:
                    # 转换为RGB模式（如果是RGBA或其他模式）
                    if img.mode in ('RGBA', 'P'):
                        img = img.convert('RGB')
                    # 保存为JPEG格式到内存中
                    buffered = io.BytesIO()
                    img.save(buffered, format="JPEG", quality=85)
                    image_data = buffered.getvalue()
                    return base64.b64encode(image_data).decode('utf-8'), "image/jpeg"
            else:
                return base64.b64encode(image_file.read()).decode('utf-8'), "image/jpeg"
    except Exception as e:
        print(f"编码图片失败 {image_path}: {e}")
        return None, None


def call_glm4v_api(base64_image, mime_type="image/jpeg"):
    """
    调用智谱GLM-4V-Flash API进行图片OCR识别

    Args:
        base64_image: base64编码的图片数据
        mime_type: 图片MIME类型

    Returns:
        API返回的JSON结果
    """
    try:
        client = ZhipuAiClient(
            api_key=ZHIPU_API_KEY,
            base_url=ZHIPU_BASE_URL
        )

        # 构建提示词
        prompt = f"""请仔细识别这张图片中的文字内容，并提取以下医学参数的值：
{', '.join(PARAMETERS)}

请以JSON格式返回结果，格式如下：
{{
    "IVS;d": "值",
    "LVID;d": "值",
    "LVPW;d": "值",
    "IVS;s": "值",
    "LVID;s": "值",
    "LVPW;s": "值",
    "LV Vol;d": "值",
    "LV Vol;s": "值",
    "EF": "值",
    "FS": "值",
    "LV Mass": "值",
    "LV Mass Cor": "值"
}}

如果某个参数在图片中没有找到，请将其值设为"N/A"。只返回JSON格式，不要有其他文本。"""

        # 调用GLM-4V-Flash API
        response = client.chat.completions.create(
            model="glm-4v-flash",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": base64_image
                            }
                        },
                        {
                            "type": "text",
                            "text": prompt
                        }
                    ]
                }
            ],
            thinking={
                "type": "enabled"
            }
        )

        result_text = response.choices[0].message.content

        # 尝试解析JSON
        try:
            # 提取JSON部分
            json_match = re.search(r'\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\}', result_text, re.DOTALL)
            if json_match:
                json_str = json_match.group(0)
                return json.loads(json_str)
            else:
                # 如果没有找到JSON，尝试直接解析
                return json.loads(result_text)
        except json.JSONDecodeError:
            print(f"JSON解析失败，原始返回: {result_text[:200]}")
            return None

    except Exception as e:
        print(f"调用智谱API失败: {e}")
        return None


def parse_api_response(response_data):
    """
    解析API返回的结果，提取参数值

    Args:
        response_data: API返回的JSON数据或文本

    Returns:
        参数字典
    """
    result = {param: "N/A" for param in PARAMETERS}

    if response_data is None:
        return result

    try:
        # 如果是字典，直接提取
        if isinstance(response_data, dict):
            for param in PARAMETERS:
                if param in response_data:
                    result[param] = response_data[param]
        # 如果是文本，使用正则表达式提取
        elif isinstance(response_data, str):
            for param in PARAMETERS:
                pattern = rf"{re.escape(param)}[：:\s]*([0-9.,]+(?:\s*[a-zA-Z%]*)?)"
                match = re.search(pattern, response_data, re.IGNORECASE)
                if match:
                    result[param] = match.group(1).strip()

    except Exception as e:
        print(f"解析API响应失败: {e}")

    return result


def process_image_directory(image_dir, output_excel):
    """
    处理图片目录中的所有图片

    Args:
        image_dir: 图片目录路径
        output_excel: 输出Excel文件路径
    """
    # 支持的图片扩展名
    image_extensions = ('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tif', '.tiff')

    # 收集所有图片文件
    image_files = []
    for root, dirs, files in os.walk(image_dir):
        for file in files:
            if file.lower().endswith(image_extensions):
                image_files.append(os.path.join(root, file))

    if not image_files:
        print(f"在目录 {image_dir} 中未找到图片文件")
        return

    print(f"找到 {len(image_files)} 个图片文件")

    # 结果列表
    results = []

    # 处理每个图片文件
    for idx, image_path in enumerate(image_files, 1):
        print(f"\n处理 [{idx}/{len(image_files)}]: {os.path.basename(image_path)}")

        # 编码图片
        base64_image, mime_type = encode_image_to_base64(image_path)
        if base64_image is None:
            print(f"跳过图片: {image_path}")
            continue

        # 调用API
        print(f"正在调用GLM-4V-Flash API识别...")
        response = call_glm4v_api(base64_image, mime_type)

        # 解析结果
        params = parse_api_response(response)

        # 添加文件名
        row_data = {
            "文件名": os.path.basename(image_path),
            "文件路径": image_path
        }
        row_data.update(params)

        results.append(row_data)
        print(f"识别结果: {params}")

    # 保存到Excel
    if results:
        df = pd.DataFrame(results)

        # 调整列的顺序，将文件名放在前面
        columns = ["文件名", "文件路径"] + PARAMETERS
        df = df[columns]

        df.to_excel(output_excel, index=False, engine='openpyxl')
        print(f"\n结果已保存到: {output_excel}")
        print(f"共处理 {len(results)} 个图片文件")
    else:
        print("\n没有成功处理任何图片文件")


def process_single_image(image_path, output_excel):
    """
    处理单个图片文件

    Args:
        image_path: 图片文件路径
        output_excel: 输出Excel文件路径
    """
    print(f"处理图片: {image_path}")

    # 编码图片
    base64_image, mime_type = encode_image_to_base64(image_path)
    if base64_image is None:
        print(f"无法编码图片: {image_path}")
        return

    # 调用API
    print("正在调用GLM-4V-Flash API识别...")
    response = call_glm4v_api(base64_image, mime_type)

    print(f"\nAPI返回结果:\n{response}")

    # 解析结果
    params = parse_api_response(response)

    # 添加文件名
    row_data = {
        "文件名": os.path.basename(image_path),
        "文件路径": image_path
    }
    row_data.update(params)

    # 保存到Excel
    df = pd.DataFrame([row_data])

    # 调整列的顺序
    columns = ["文件名", "文件路径"] + PARAMETERS
    df = df[columns]

    df.to_excel(output_excel, index=False, engine='openpyxl')
    print(f"\n结果已保存到: {output_excel}")


def main():
    """
    主函数
    """
    import argparse

    parser = argparse.ArgumentParser(description='使用智谱GLM-4V-Flash API进行图片OCR识别')
    parser.add_argument('--input', '-i', type=str, help='图片文件或目录路径')
    parser.add_argument('--output', '-o', type=str, default='ocr_result.xlsx',
                        help='输出Excel文件路径 (默认: ocr_result.xlsx)')
    parser.add_argument('--api-key', '-k', type=str, help='智谱AI API密钥 (也可以通过环境变量ZHIPU_API_KEY设置)')

    args = parser.parse_args()

    # 设置API密钥
    global ZHIPU_API_KEY
    if args.api_key:
        ZHIPU_API_KEY = args.api_key
        os.environ["ZHIPU_API_KEY"] = ZHIPU_API_KEY

    if not ZHIPU_API_KEY or ZHIPU_API_KEY == "your-api-key-here":
        print("错误: 请设置智谱AI API密钥")
        print("可以通过以下方式设置:")
        print("  1. 环境变量: export ZHIPU_API_KEY=your-key")
        print("  2. 命令行参数: --api-key your-key")
        return

    # 检查输入路径
    if args.input is None:
        print("错误: 请指定输入图片或目录路径")
        print("使用方法: python python_ocr_claude.py --input 图片路径 [--output 输出路径]")
        return

    input_path = Path(args.input)

    if not input_path.exists():
        print(f"错误: 路径不存在: {args.input}")
        return

    # 处理图片
    if input_path.is_file():
        process_single_image(str(input_path), args.output)
    else:
        process_image_directory(str(input_path), args.output)


if __name__ == "__main__":
    # 示例用法
    # 处理单个图片: python python_ocr_claude.py --input MI/83051-MI/83051-MI-1.tif --output result.xlsx
    # 处理整个目录: python python_ocr_claude.py --input MI --output result.xlsx
    # 使用API密钥: python python_ocr_claude.py --input MI --output result.xlsx --api-key your-api-key

    # 如果没有提供命令行参数，使用默认配置
    if len(__import__('sys').argv) == 1:
        print("使用默认配置处理...")
        print("提示: 使用 --help 查看完整用法说明\n")

        # 设置默认输入输出
        default_input = "83001-CKO"
        default_output = "result/glm4v_ocr_result.xlsx"

        # 确保输出目录存在
        os.makedirs(os.path.dirname(default_output) or ".", exist_ok=True)

        print(f"输入目录: {default_input}")
        print(f"输出文件: {default_output}\n")

        process_image_directory(default_input, default_output)
    else:
        main()
