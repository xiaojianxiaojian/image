# image
python开发的脚本&&工具

## 项目说明

本项目包含两个核心脚本，用于医学图片的批量处理和OCR识别。

## 文件说明

### image_ui.py - 图片处理工具

**功能**：提供图形化界面，批量处理医学图片，计算特定区域面积占比。

**实现原理**：
- 使用tkinter构建GUI界面，支持多文件夹选择
- 基于OpenCV的HSV颜色空间进行背景去除
- 使用ImageJ进行图像分析：提取蓝色通道、阈值分割、计算面积占比
- 结果自动保存为Excel文件，包含各文件夹数据及平均值

### python_ocr_claude.py - OCR识别脚本

**功能**：调用智谱AI GLM-4V-Flash模型，识别图片中的医学参数。

**实现原理**：
- 将TIF图片转换为JPEG格式并编码为base64
- 通过HTTP请求调用智谱AI API，发送图片和提示词
- 模型返回JSON格式的识别结果
- 提取以下参数：IVS;d、LVID;d、LVPW;d、IVS;s、LVID;s、LVPW;s、LV Vol;d、LV Vol;s、EF、FS、LV Mass、LV Mass Cor
- 结果输出到Excel文件

## 安装依赖

```bash
pip install openai pandas pillow openpyxl zai opencv-python imagej
```

## 使用方法

### image_ui.py
```bash
python image_ui.py
```
启动后通过界面选择输入文件夹和输出路径，点击"开始处理"。

### python_ocr_claude.py
```bash
# 设置API密钥
export ZHIPU_API_KEY=your-api-key

# 运行脚本
python python_ocr_claude.py --input 图片目录 --output 结果.xlsx
```
