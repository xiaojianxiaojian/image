import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import os
import time
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
import imagej
import ctypes
from ctypes import wintypes
import sys
import cv2
import numpy as np
import shutil


# Windows API 多选文件夹对话框
class MultiFolderDialog:
    def __init__(self):
        if sys.platform != 'win32':
            raise RuntimeError("多选文件夹对话框仅支持 Windows")

        # 定义 COM 接口
        self.IID_IFileOpenDialog = ctypes.c_void_p.in_dll(
            ctypes.oledll.ole32, 'IID_IFileOpenDialog'
        )

        # 定义方法
        self.CoCreateInstance = ctypes.windll.ole32.CoCreateInstance
        self.CoInitialize = ctypes.windll.ole32.CoInitialize
        self.CoUninitialize = ctypes.windll.ole32.CoUninitialize
        self.CoTaskMemFree = ctypes.windll.ole32.CoTaskMemFree

    def select_folders(self, title="选择文件夹"):
        self.CoInitialize(None)

        try:
            # 使用标准方法，只能单选，所以使用递归方式
            root = tk.Tk()
            root.withdraw()
            folder = filedialog.askdirectory(title=title)
            root.destroy()

            if folder:
                # 检查是否包含子文件夹
                subfolders = []
                for item in os.listdir(folder):
                    item_path = os.path.join(folder, item)
                    if os.path.isdir(item_path):
                        subfolders.append(item_path)

                if subfolders:
                    # 询问是否添加所有子文件夹
                    msg = f"发现 {len(subfolders)} 个子文件夹，是否全部添加？\n\n"
                    msg += "是 - 添加所有子文件夹\n"
                    msg += "否 - 只添加选中的文件夹\n"
                    msg += "取消 - 不添加"
                    choice = messagebox.askyesnocancel("确认", msg)

                    if choice is None:
                        return []
                    elif choice:
                        return subfolders
                    else:
                        return [folder]
                else:
                    return [folder]
            return []

        finally:
            self.CoUninitialize()


def select_folders_multi(title="选择文件夹"):
    """使用递归方式模拟多选"""
    root = tk.Tk()
    root.withdraw()
    folders = []

    while True:
        folder = filedialog.askdirectory(title=title + " (可多次选择)")
        if not folder:
            break

        # 检查是否包含子文件夹
        subfolders = []
        for item in os.listdir(folder):
            item_path = os.path.join(folder, item)
            if os.path.isdir(item_path):
                subfolders.append(item_path)

        if subfolders:
            # 询问是否添加所有子文件夹
            msg = f"发现 {len(subfolders)} 个子文件夹，是否全部添加？\n\n"
            msg += "是 - 添加所有子文件夹\n"
            msg += "否 - 只添加选中的文件夹\n"
            msg += "取消 - 停止选择"
            choice = messagebox.askyesnocancel("确认", msg)

            if choice is None:
                break
            elif choice:
                folders.extend(subfolders)
            else:
                if folder not in folders:
                    folders.append(folder)
        else:
            if folder not in folders:
                folders.append(folder)

        # 询问是否继续选择
        if not messagebox.askyesno("继续", "是否继续选择其他文件夹？"):
            break

    root.destroy()
    return folders


# ============ 去除背景相关方法 ============

def remove_background(image_path, output_path, lower_color, upper_color):
    """去除图片背景并替换为白色"""
    # 读取图像
    image = cv2.imread(image_path)
    if image is None:
        raise ValueError(f"无法读取图片: {image_path}")

    # 转换为HSV颜色空间
    hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)

    # 定义颜色范围以创建掩码
    lower_bound = np.array(lower_color)
    upper_bound = np.array(upper_color)

    # 创建掩码
    mask = cv2.inRange(hsv, lower_bound, upper_bound)

    # 反转掩码
    mask = cv2.bitwise_not(mask)

    # 应用掩码
    result = cv2.bitwise_and(image, image, mask=mask)

    # 创建白色背景图像
    white_background = np.ones_like(image, dtype=np.uint8) * 255

    # 将前景图像覆盖在白色背景上
    final_result = cv2.add(result, cv2.bitwise_and(white_background, white_background, mask=cv2.bitwise_not(mask)))

    # 保存结果图像
    cv2.imwrite(output_path, final_result)


def process_folder_remove_background(input_folder, output_folder, lower_color, upper_color):
    """处理文件夹中所有图片，去除背景"""
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    processed_files = []
    for filename in os.listdir(input_folder):
        if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.tif', '.bmp')):
            input_path = os.path.join(input_folder, filename)
            output_path = os.path.join(output_folder, filename)
            remove_background(input_path, output_path, lower_color, upper_color)
            processed_files.append(output_path)

    return processed_files


# ============ UI 类 ============

class ImageProcessorUI:
    def __init__(self, root):
        self.root = root
        self.root.title("图片处理工具")
        self.root.geometry("750x700")

        # 输出文件夹路径
        self.output_folder = tk.StringVar(value="D:/code/image/result/")

        # 去除背景配置
        self.enable_remove_bg = tk.BooleanVar(value=True)
        self.lower_color_h = tk.IntVar(value=0)
        self.lower_color_s = tk.IntVar(value=0)
        self.lower_color_v = tk.IntVar(value=200)
        self.upper_color_h = tk.IntVar(value=180)
        self.upper_color_s = tk.IntVar(value=180)
        self.upper_color_v = tk.IntVar(value=255)

        # 文件夹路径列表
        self.folder_paths_list = []

        # 初始化 ImageJ
        self.ij = imagej.init(mode='headless')

        self.create_widgets()

    def create_widgets(self):
        # 标题
        title_label = tk.Label(self.root, text="图片处理工具", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)

        # 文件夹选择区域
        folder_frame = tk.LabelFrame(self.root, text="输入文件夹 (支持多选)", font=("Arial", 10))
        folder_frame.pack(pady=10, padx=20, fill="both", expand=True)

        # 列表框显示已选文件夹
        listbox_frame = tk.Frame(folder_frame)
        listbox_frame.pack(side="top", fill="both", expand=True, padx=5, pady=5)

        scrollbar = tk.Scrollbar(listbox_frame)
        scrollbar.pack(side="right", fill="y")

        self.folder_listbox = tk.Listbox(
            listbox_frame,
            font=("Courier New", 9),
            yscrollcommand=scrollbar.set,
            selectmode=tk.MULTIPLE,
            height=10
        )
        self.folder_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.folder_listbox.yview)

        # 按钮区域
        btn_frame = tk.Frame(folder_frame)
        btn_frame.pack(fill="x", padx=5, pady=5)

        btn_add = tk.Button(
            btn_frame,
            text="添加文件夹",
            command=self.add_folder,
            bg="#2196F3",
            fg="white",
            font=("Arial", 10),
            width=12
        )
        btn_add.pack(side="left", padx=5)

        btn_remove = tk.Button(
            btn_frame,
            text="移除选中",
            command=self.remove_folder,
            bg="#f44336",
            fg="white",
            font=("Arial", 10),
            width=12
        )
        btn_remove.pack(side="left", padx=5)

        btn_clear = tk.Button(
            btn_frame,
            text="清空全部",
            command=self.clear_folders,
            bg="#FF9800",
            fg="white",
            font=("Arial", 10),
            width=12
        )
        btn_clear.pack(side="left", padx=5)

        # 去除背景配置区域
        bg_frame = tk.LabelFrame(self.root, text="去除背景配置 (HSV颜色空间)", font=("Arial", 10))
        bg_frame.pack(pady=10, padx=20, fill="x")

        # 启用去除背景复选框
        enable_cb = tk.Checkbutton(
            bg_frame,
            text="启用去除背景",
            variable=self.enable_remove_bg,
            font=("Arial", 10)
        )
        enable_cb.pack(anchor="w", padx=5, pady=5)

        # HSV 颜色范围配置
        hsv_frame = tk.Frame(bg_frame)
        hsv_frame.pack(fill="x", padx=5, pady=5)

        # Lower Color
        lower_frame = tk.Frame(hsv_frame)
        lower_frame.pack(side="left", padx=10)

        tk.Label(lower_frame, text="Lower Color:", font=("Arial", 9, "bold")).pack()
        tk.Label(lower_frame, text="H:").pack()
        tk.Entry(lower_frame, textvariable=self.lower_color_h, width=5).pack()
        tk.Label(lower_frame, text="S:").pack()
        tk.Entry(lower_frame, textvariable=self.lower_color_s, width=5).pack()
        tk.Label(lower_frame, text="V:").pack()
        tk.Entry(lower_frame, textvariable=self.lower_color_v, width=5).pack()

        # Upper Color
        upper_frame = tk.Frame(hsv_frame)
        upper_frame.pack(side="left", padx=10)

        tk.Label(upper_frame, text="Upper Color:", font=("Arial", 9, "bold")).pack()
        tk.Label(upper_frame, text="H:").pack()
        tk.Entry(upper_frame, textvariable=self.upper_color_h, width=5).pack()
        tk.Label(upper_frame, text="S:").pack()
        tk.Entry(upper_frame, textvariable=self.upper_color_s, width=5).pack()
        tk.Label(upper_frame, text="V:").pack()
        tk.Entry(upper_frame, textvariable=self.upper_color_v, width=5).pack()

        # 输出文件夹配置和开始处理按钮
        output_config_frame = tk.Frame(self.root)
        output_config_frame.pack(pady=10, padx=20, fill="x")

        output_label = tk.Label(output_config_frame, text="输出文件夹:", font=("Arial", 10))
        output_label.pack(side="left")

        output_entry = tk.Entry(output_config_frame, textvariable=self.output_folder, width=50, font=("Courier New", 9))
        output_entry.pack(side="left", padx=5)

        btn_browse = tk.Button(
            output_config_frame,
            text="浏览...",
            command=self.browse_output_folder,
            width=8
        )
        btn_browse.pack(side="left", padx=5)

        self.btn_process = tk.Button(
            output_config_frame,
            text="开始处理",
            command=self.process_images,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 11, "bold"),
            width=12
        )
        self.btn_process.pack(side="left", padx=5)

        # 输出日志区域
        log_label = tk.Label(self.root, text="输出日志:", font=("Arial", 10))
        log_label.pack(pady=(5, 2), anchor="w", padx=20)

        self.text_output = scrolledtext.ScrolledText(self.root, width=85, height=12, font=("Courier New", 9))
        self.text_output.pack(pady=5, padx=20)

    def add_folder(self):
        folders = select_folders_multi(title="选择文件夹")
        for folder in folders:
            if folder not in self.folder_paths_list:
                self.folder_paths_list.append(folder)
                self.folder_listbox.insert(tk.END, folder)

    def remove_folder(self):
        selected = self.folder_listbox.curselection()
        if selected:
            # 倒序删除，避免索引变化
            for index in reversed(selected):
                self.folder_listbox.delete(index)
                del self.folder_paths_list[index]

    def clear_folders(self):
        self.folder_listbox.delete(0, tk.END)
        self.folder_paths_list.clear()

    def browse_output_folder(self):
        folder = filedialog.askdirectory(title="选择输出文件夹")
        if folder:
            # 确保路径以/结尾
            if not folder.endswith('/') and not folder.endswith('\\'):
                folder = folder + '/'
            self.output_folder.set(folder)

    def remove_background_folder(self, folder_path, temp_output_folder):
        """去除文件夹中所有图片的背景"""
        lower_color = [
            self.lower_color_h.get(),
            self.lower_color_s.get(),
            self.lower_color_v.get()
        ]
        upper_color = [
            self.upper_color_h.get(),
            self.upper_color_s.get(),
            self.upper_color_v.get()
        ]

        #self.text_output.insert(tk.END, f"  去除背景中...\n")
        self.root.update()

        # 创建临时输出文件夹
        if not os.path.exists(temp_output_folder):
            os.makedirs(temp_output_folder)

        # 处理文件夹中的所有图片
        for filename in os.listdir(folder_path):
            if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.tif', '.bmp')):
                input_path = os.path.join(folder_path, filename)
                output_path = os.path.join(temp_output_folder, filename)
                try:
                    remove_background(input_path, output_path, lower_color, upper_color)
                except Exception as e:
                    self.text_output.insert(tk.END, f"  错误: {filename} - {str(e)}\n")
                    self.root.update()

        return temp_output_folder

    def process_images_inner(self, folder_path):
        """处理单个文件夹的图片（调用原始50-106行逻辑）"""
        self.text_output.insert(tk.END, f"计算区域处理文件夹: {folder_path}\n")
        self.root.update()

        results = []

        for file in os.listdir(folder_path):
            file_path_full = os.path.join(folder_path, file)
            if os.path.isfile(file_path_full):
                try:
                    image = self.ij.io().open(file_path_full)

                    # 将 ImageJ 图像转换为 NumPy 数组
                    image_array = self.ij.py.from_java(image)

                    # 提取蓝色通道
                    blue_channel = image_array[:, :, 2]

                    # 创建仅包含蓝色通道的图像
                    processed_image = self.ij.py.to_java(blue_channel)

                    output_path = 'output.tif'
                    self.ij.io().save(processed_image, output_path)

                    image_out = self.ij.io().open(output_path)

                    imp_default = self.ij.py.to_imageplus(image_out)
                    imp_all = self.ij.py.to_imageplus(image_out)

                    self.ij.IJ.run("Set Measurements...", "area_fraction")
                    self.ij.IJ.setAutoThreshold(imp_default, "Default")
                    self.ij.IJ.setRawThreshold(imp_all, 0, 254)

                    output_default = self.ij.IJ.getValue(imp_default, "%Area")
                    output_all = self.ij.IJ.getValue(imp_all, "%Area")

                    formatted_number = "{:.2f}".format((float(output_all) - float(output_default)) / float(output_all))

                    # self.text_output.insert(tk.END, file + "  " + formatted_number + "\n")
                    # self.root.update()

                    results.append({
                        "Folder": folder_path,
                        "File": file,
                        "Value": formatted_number
                    })

                    os.remove(output_path)

                except Exception as e:
                    self.text_output.insert(tk.END, f"错误: {file} - {str(e)}\n")
                    self.root.update()

        return results

    def process_images(self):
        # 从列表框获取文件夹路径
        folder_paths = self.folder_paths_list.copy()

        if not folder_paths:
            messagebox.showwarning("警告", "请至少选择一个文件夹")
            return

        # 清空输出区域
        self.text_output.delete("1.0", tk.END)
        self.text_output.insert("1.0", "开始处理...\n\n")
        self.root.update()

        start_time = time.time()
        self.text_output.insert(tk.END, "start >>>>>\n\n")
        self.root.update()

        # 收集结果，按文件夹分组
        folder_results = {}

        # 创建临时文件夹用于存储去除背景后的图片
        temp_base = os.path.join(os.path.dirname(self.output_folder.get()), "temp_processed")
        if os.path.exists(temp_base):
            shutil.rmtree(temp_base)
        os.makedirs(temp_base)

        try:
            for folder_path in folder_paths:
                # 获取文件夹名称
                folder_name = os.path.basename(folder_path.rstrip('/\\'))

                # 如果启用去除背景，先处理背景
                if self.enable_remove_bg.get():
                    temp_output_folder = os.path.join(temp_base, folder_name)
                    processed_folder = self.remove_background_folder(folder_path, temp_output_folder)
                    self.text_output.insert(tk.END, f"去除背景处理文件夹: {processed_folder}\n")
                    self.root.update()
                    results = self.process_images_inner(processed_folder)
                else:
                    results = self.process_images_inner(folder_path)

                folder_results[folder_path] = results

                end_time = time.time()
                elapsed_time = "{:.1f}".format(end_time - start_time)
                self.text_output.insert(tk.END, f"\nstop <<<<<<<<<<<< costs {elapsed_time}s\n\n")
                self.root.update()

            # 如果启用去除背景，保存处理后的图片到输出文件夹
            if self.enable_remove_bg.get():
                self.text_output.insert(tk.END, "\n保存去除背景后的图片...\n")
                self.root.update()
                for folder_path in folder_paths:
                    folder_name = os.path.basename(folder_path.rstrip('/\\'))
                    temp_folder = os.path.join(temp_base, folder_name)
                    if os.path.exists(temp_folder):
                        output_folder = self.output_folder.get() + folder_name + "_no_bg/"
                        if os.path.exists(output_folder):
                            shutil.rmtree(output_folder)
                        shutil.copytree(temp_folder, output_folder)
                        self.text_output.insert(tk.END, f"已保存到: {output_folder}\n")
                        self.root.update()

        finally:
            # 清理临时文件夹
            if os.path.exists(temp_base):
                shutil.rmtree(temp_base)

        # 写入Excel
        if folder_results:
            output_excel = self.output_folder.get() + "result.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Result"

            # 创建粗体字体
            bold_font = Font(bold=True)

            headers = ["Folder", "File", "Value"]
            current_row = 1

            for folder_path, results in folder_results.items():
                # 写入列名（表头加粗）
                for col_idx, header in enumerate(headers, 1):
                    cell = ws.cell(row=current_row, column=col_idx, value=header)
                    cell.font = bold_font
                current_row += 1

                # 计算平均值
                values = [float(r["Value"]) for r in results]
                avg_value = sum(values) / len(values) if values else 0
                folder_name = os.path.basename(folder_path)

                # 写入数据
                for result in results:
                    ws.cell(row=current_row, column=1, value=folder_name)
                    ws.cell(row=current_row, column=2, value=result["File"])
                    ws.cell(row=current_row, column=3, value=float(result["Value"]))
                    current_row += 1

                # 写入平均值行（平均值加粗）
                avg_cell1 = ws.cell(row=current_row, column=1, value="平均值")
                avg_cell1.font = bold_font
                avg_cell3 = ws.cell(row=current_row, column=3, value=round(avg_value, 2))
                avg_cell3.font = bold_font
                current_row += 1

                # 空白行
                current_row += 1

            wb.save(output_excel)
            self.text_output.insert(tk.END, f"\nExcel已保存到: {output_excel}\n")

        messagebox.showinfo("完成", "处理完成！")


if __name__ == "__main__":
    root = tk.Tk()
    app = ImageProcessorUI(root)
    root.mainloop()
