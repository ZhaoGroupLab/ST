import cv2
import numpy as np
import pandas as pd
from collections import Counter

# 定义7种颜色的RGB值和对应的编码（0-6）
color_mapping = {
    (221, 101, 94): 0,   # #DD655E 皮层
    (100, 158, 185): 1,  # #649EB9 纤维
    (107, 187, 174): 2,  # #6BBBAE 基本组织区域
    (210, 104, 52): 3,   # #D26834 髓
    (149, 33, 36): 4,    # #952124 后生木质部 导管
    (227, 132, 156): 5,  # #E3849C 筛管+伴胞
    (154, 74, 44): 6     # #9A4A2C 非髓区组织
}

# 设置颜色容忍度：允许的最大颜色差异（欧氏距离）
color_tolerance = 50

# 计算两个颜色之间的欧式距离
def color_distance(c1, c2):
    return np.linalg.norm(np.array(c1) - np.array(c2))

# 读取图像（请替换为你的图像文件路径）
image_path = 'bin100.png'  # 图像路径
img = cv2.imread(image_path)

# 将图像转换为RGB格式（OpenCV默认读取为BGR）
img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)

# 获取图像的尺寸
height, width, _ = img_rgb.shape

# 假设图像的尺寸为 107 行 * 20 列，每个格子的大小
block_height = height // 107
block_width = width // 20

# 创建一个107x20的矩阵，用于存储颜色编码
color_matrix = np.zeros((107, 20), dtype=int)

# 对每个小块进行颜色分类
for i in range(107):
    for j in range(20):
        # 提取当前小块的区域
        block = img_rgb[i*block_height:(i+1)*block_height, j*block_width:(j+1)*block_width]
        
        # 跳过空白块（全为黑色或者没有有效像素的块）
        if np.all(block == 0):  # 如果块内所有像素都是黑色或无效像素
            color_matrix[i, j] = -1  # 用 -1 表示空白块
            continue
        
        # 将小块的颜色展平成一维数组（每个像素的颜色）
        pixels = block.reshape(-1, 3)
        
        # 统计每种颜色的频率
        pixel_counts = Counter(tuple(pixel) for pixel in pixels)
        
        # 找到出现频率最多的颜色
        if pixel_counts:  # 确保有有效的颜色数据
            most_common_color = pixel_counts.most_common(1)[0][0]
            
            # 匹配最接近的颜色（容忍度匹配）
            closest_color = min(color_mapping.keys(), key=lambda x: color_distance(x, most_common_color))
            
            # 仅当颜色距离足够近时才匹配
            if color_distance(closest_color, most_common_color) < color_tolerance:
                color_matrix[i, j] = color_mapping[closest_color]
            else:
                color_matrix[i, j] = -1  # 如果颜色匹配不成功，标记为 -1
        else:
            color_matrix[i, j] = -1  # 如果没有有效颜色数据，标记为 -1

# 将矩阵转换为DataFrame并保存为Excel文件
df = pd.DataFrame(color_matrix)
excel_path = 'bin100.xlsx'
df.to_excel(excel_path, index=False, header=False)

print(f"矩阵已保存为 {excel_path}")
