import cv2
import numpy as np
import pandas as pd
from collections import Counter

# 定义12种颜色的RGB值和对应的编码（0-11）
color_mapping = {
    (234, 234, 85): 0,   # 髓
    (251, 193, 114): 1,  # 髓环
    (85, 175, 216): 2,   # 基本组织
    (57, 83, 161): 3,    # 原生木质部导管
    (40, 188, 185): 4,   # 后生木质部导管
    (163, 153, 179): 5,  # 维管束
    (134, 139, 192): 6,  # 筛管
    (0, 107, 189): 7,    # 纤维细胞
    (203, 68, 63): 8,    # 伴胞
    (0, 161, 154): 9,    # 皮层细胞
    (117, 188, 55): 10,  # 皮下层细胞
    (0, 150, 61): 11     # 表皮细胞
}

# 设置颜色容忍度（允许的最大颜色差异）
color_tolerance = 29    #测试29为最好的参数

# 计算两个颜色之间的欧式距离
def color_distance(c1, c2):
    return np.linalg.norm(np.array(c1) - np.array(c2))

# 读取图像
image_path = 'jp-bin100-3.0.png'  # 图像路径
img = cv2.imread(image_path)

# 将图像转换为RGB格式（OpenCV默认读取为BGR）
img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)

# 获取图像尺寸
height, width, _ = img_rgb.shape

# 假设图像的尺寸为 107 行 * 20 列，每个格子的大小
block_height = height // 107
block_width = width // 20

# 创建一个 107x20 的列表矩阵（存储颜色编号的列表）
color_matrix = [[""] * 20 for _ in range(107)]

# 对每个小块进行颜色分类
for i in range(107):
    for j in range(20):
        # 提取当前小块的区域
        block = img_rgb[i*block_height:(i+1)*block_height, j*block_width:(j+1)*block_width]
        
        # 将小块的颜色展平成一维数组
        pixels = block.reshape(-1, 3)
        
        # 统计颜色频率
        pixel_counts = Counter(tuple(pixel) for pixel in pixels)
        
        # 找到所有主要颜色
        detected_colors = set()
        for pixel_color, count in pixel_counts.items():
            # 匹配最接近的颜色
            closest_color = min(color_mapping.keys(), key=lambda x: color_distance(x, pixel_color))
            
            # 仅当颜色距离足够近时才匹配
            if color_distance(closest_color, pixel_color) < color_tolerance:
                detected_colors.add(color_mapping[closest_color])
        
        # 将颜色编号存入矩阵（用逗号分隔）
        if detected_colors:
            color_matrix[i][j] = ",".join(map(str, sorted(detected_colors)))
        else:
            color_matrix[i][j] = "-1"  # 如果没有匹配到颜色，标记为 -1

# 转换为DataFrame并保存为Excel文件
df = pd.DataFrame(color_matrix)
excel_path = 'jp-bin100-3.06.xlsx'
df.to_excel(excel_path, index=False, header=False)

print(f"矩阵已保存为 {excel_path}")
