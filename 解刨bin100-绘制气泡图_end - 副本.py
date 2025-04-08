import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import Wedge  # 使用Wedge绘制扇形

# 定义新的颜色映射
color_mapping = {
    0: (225, 225, 89),   # 髓
    1: (47, 178, 175),  # 髓环
}

#e1e159 → RGB(225, 225, 89)
#7eb55a → RGB(126, 181, 90)
#2fb2af → RGB(47, 178, 175)



# 读取Excel文件
file_path = 'match_result.xlsx'  # 替换成你的文件路径
df = pd.read_excel(file_path, header=None)  # 假设没有表头

# 获取矩阵数据
matrix = df.to_numpy()

# 设置绘图
fig, ax = plt.subplots(figsize=(10, 10))

# 自定义圆圈的大小和轴间隔
circle_size_factor = 20  # 圆圈大小

# 绘制每个单元格的数据
for i in range(matrix.shape[0]):  # 行数
    for j in range(matrix.shape[1]):  # 列数
        cell_value = matrix[i, j]
        
        if cell_value == -1:
            continue  # 跳过不需要绘制的单元格

        # 处理每个单元格的数据
        # 解析逗号分隔的数字
        cell_values = [int(x) for x in str(cell_value).split(',')]

        # 根据数字获取对应的颜色
        colors_for_cell = [color_mapping[num] for num in cell_values if num in color_mapping]

        # 如果颜色数大于0，绘制相应的扇形
        if len(colors_for_cell) > 0:
            # 计算每个圆点的角度间隔，假设圆点总是被分成等分
            num_colors = len(colors_for_cell)
            angle_per_color = 360 / num_colors

            center_x, center_y = j * 2 * circle_size_factor, i * 2 * circle_size_factor  # 每个点之间的间隔

            # 绘制每个圆点的不同颜色扇形
            for idx, color in enumerate(colors_for_cell):
                # 计算对应的RGB颜色
                color_tuple = color
                start_angle = angle_per_color * idx
                end_angle = angle_per_color * (idx + 1)

                # 绘制扇形
                wedge = Wedge((center_x, center_y), circle_size_factor, 
                              start_angle, end_angle, 
                              color=np.array(color_tuple) / 255.0, ec="black", lw=0.5)
                ax.add_patch(wedge)

# 设置坐标轴范围和显示设置
ax.set_xlim(-circle_size_factor, matrix.shape[1] * 2 * circle_size_factor)
ax.set_ylim(-circle_size_factor, matrix.shape[0] * 2 * circle_size_factor)
ax.set_aspect('equal')  # 保持图形比例
ax.axis('off')  # 隐藏坐标轴



# 保存为可编辑的PDF文件
output_pdf_path = 'match_figure.pdf'  # 设定保存路径
plt.savefig(output_pdf_path, format='pdf', dpi=600)  # 保存为PDF格式
plt.close()  # 关闭图形，防止显示

print(f"PDF文件已保存至: {output_pdf_path}")
plt.show()
