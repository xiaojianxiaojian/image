
##============cal Measure===========


#pyimagej版本只能用1.3.0 ,升级到1.7.0后，api调用方式修改了，计算出结果是0了
import os
import imagej
import time
import numpy as np
import matplotlib.pyplot as plt



start_time = time.time()
# 初始化 ImageJ
ij = imagej.init(mode='headless')
print("start >>>>>")




# def get_file_names(folder_path):
#     file_names = []
#     for file in os.listdir(folder_path):
#         if os.path.isfile(os.path.join(folder_path, file)):
#             file_names.append(file)
#     return file_names
#
# folder_path = "input"  # 替换为您的文件夹路径
# files = get_file_names(folder_path)
# print(files)


# 示例：要处理的多个文件夹路径
folder_paths = [
    "input_bk/81983-IR",
    "input_bk/81984-IR+PSQ",
    #"flox/83818-flox+mi"
    # "flox+psq/83011-flox+psq+mi",
    # "flox+psq/83805-flox+psq+mi",
    # "flox+psq/83812-flox+psq+mi",
    # "flox+psq/83815-flox+psq+mi",
    # "flox+psq/83820-flox+psq+mi"
    # "cko/83066-MI+PSQ",
    # "cko/83068-MI+PSQ",
    # "cko/83072-MI+PSQ"
]

# 遍历文件夹列表
for folder_path in folder_paths:
    file_path = "D:/code/image/result/" + folder_path + ".txt"

# folder_path = "MI_res/83063-MI"
# #file_path = "D:/code/image/Mrp8cre sig cko/82354.txt"
# file_path = "D:/code/image/result/" + folder_path + ".txt"
# file_names = []

    with open(file_path, 'w') as res:
        for file in os.listdir(folder_path):
            if os.path.isfile(os.path.join(folder_path, file)):
                image_url = '.\\' + folder_path + '\\' + file
                image = ij.io().open(image_url)

                # 将 ImageJ 图像转换为 NumPy 数组
                image_array = ij.py.from_java(image)

                # 提取蓝色通道
                blue_channel = image_array[:, :, 2]

                # 创建仅包含蓝色通道的图像
                processed_image = ij.py.to_java(blue_channel)

                # 显示蓝色通道的图像
                #ij.py.show(processed_image, cmap='gray')

                output_path = 'output.tif'
                ij.io().save(processed_image, output_path)

                image_out = ij.io().open(output_path)

                imp_default = ij.py.to_imageplus(image_out)
                imp_all = ij.py.to_imageplus(image_out)

                ij.IJ.run("Set Measurements...", "area_fraction")

                ij.IJ.setAutoThreshold(imp_default, "Default")

                ij.IJ.setRawThreshold(imp_all, 0, 254)

                # ij.IJ.run(imp_default, "Measure", "")

                output_default = ij.IJ.getValue(imp_default, "%Area")
                output_all = ij.IJ.getValue(imp_all, "%Area")

                formatted_number = "{:.2f}".format((float(output_all) - float(output_default)) / float(output_all))

                print(file + "  " + formatted_number)

                res.write(formatted_number + "<<<<<<<<<<<<<<<" + file + "\n")

                os.remove(output_path)

    end_time = time.time()

    elapsed_time = "{:.1f}".format(end_time - start_time)
    print("stop <<<<<<<<<<<< costs ", elapsed_time, "s")



"""


#=========replace background===============


from PIL import Image
import cv2
import numpy as np
import os



# 调整图片伽马值
def adjust_gamma_image(input_path, output_path, gamma=1.0):
    # 加载图像并转换为 numpy 数组
    image = Image.open(input_path).convert("RGB")
    img_array = np.array(image).astype(np.float32) / 255.0

    # 应用伽马变换
    img_gamma = np.power(img_array, 1.0 / gamma)

    # 转换回 0-255 并为 uint8 类型
    img_gamma = np.clip(img_gamma * 255.0, 0, 255).astype(np.uint8)

    # 转换回图像
    adjusted_image = Image.fromarray(img_gamma)

    # 自动生成输出路径（如果未指定）
    if output_path is None:
        base, ext = os.path.splitext(input_path)
        output_path = f"{base}_gamma_{gamma:.1f}{ext}"

    # 保存图像
    adjusted_image.save(output_path)


def remove_background(image_path, output_path, lower_color, upper_color):
    # 读取图像
    image = cv2.imread(image_path)

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


def process_folder(input_folder, output_folder, lower_color, upper_color):
    # if not os.path.exists(output_folder):
    #     os.makedirs(output_folder)

    for filename in os.listdir(input_folder):
        if filename.endswith(('.png', '.jpg', '.jpeg', '.tif')):
            input_path = os.path.join(input_folder, filename)
            output_path = os.path.join(output_folder, filename)
            #adjust_gamma_image(input_path, output_path, gamma=0.4)
            remove_background(input_path, output_path, lower_color, upper_color)


def copy_folder_structure(src, dst, lower_color, upper_color):
    # 如果目标路径不存在，则创建目标路径
    if not os.path.exists(dst):
        os.makedirs(dst)

    # 遍历源文件夹中的所有子文件夹和文件
    for root, dirs, files in os.walk(src):
        current_folder = root

        # 计算目标文件夹中的相应路径
        target_folder = current_folder.replace(src, dst, 1)

        # 如果当前目录不存在，创建它
        if not os.path.exists(target_folder):
            os.makedirs(target_folder)
            process_folder(current_folder, target_folder, lower_color, upper_color)


# 示例使用
input_folder = "D:/code/image/flox/new/"
output_folder = "D:/code/image/flox/new_res/"

# 根据具体图像的背景颜色调整lower_color和upper_color
lower_color = [0, 0, 200]

#如果发现抠图不干净，保留了粉红色背景，增加G的值
#upper_color = [180, 80, 255]
upper_color = [180, 180, 255]

copy_folder_structure(input_folder, output_folder, lower_color, upper_color)

process_folder(input_folder, output_folder, lower_color, upper_color)



#=========replace rgb===============


"""





