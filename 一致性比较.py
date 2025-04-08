import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# 1. 读取数据
bin100_path = "bin100.xlsx"  # bin100: 0-6
jp_bin100_path = "jp-bin100.xlsx"  # jp-bin100: 0-11

bin100_df = pd.read_excel(bin100_path, header=None)
jp_bin100_df = pd.read_excel(jp_bin100_path, header=None)

# 2. 定义对应关系
correspondence = {
    0: {9, 10, 11},
    1: {7},
    2: {2},
    3: {0, 1},
    4: {3, 4},
    5: {6, 8},
    6: {2, 3, 4, 5, 6, 7, 8, 9, 10, 11}
}

# 3. 创建匹配结果矩阵
match_matrix = pd.DataFrame(0, index=bin100_df.index, columns=bin100_df.columns)

# 4. 遍历矩阵，检查匹配情况
for i in range(bin100_df.shape[0]):
    for j in range(bin100_df.shape[1]):
        bin_value = bin100_df.iloc[i, j]  # bin100 里的值（0-6）
        jp_values = str(jp_bin100_df.iloc[i, j]).split(',')  # jp-bin100 里的值（0-11）

        # 转换为整数集合
        jp_values = {int(val) for val in jp_values if val.isdigit()}
        
        # 检查是否符合对应关系
        if bin_value in correspondence and jp_values.intersection(correspondence[bin_value]):
            match_matrix.iloc[i, j] = 1

# 5. 计算一致性指数
total_cells = match_matrix.size  # 总单元格数
matched_cells = match_matrix.sum().sum()  # 符合对应关系的单元格数
consistency_index = matched_cells / total_cells  # 计算一致性指数
print(f"一致性指数: {consistency_index:.3f}")  # 输出结果

# 6. 保存匹配结果
match_matrix.to_excel("match_result.xlsx", index=False, header=False)

# 7. 生成可视化热图
plt.figure(figsize=(12, 6))
plt.imshow(match_matrix, cmap="coolwarm", aspect="auto", interpolation="nearest")
plt.colorbar(label="Matching Status (0 = Not Matched, 1 = Matched)")
plt.title("Matching Matrix Visualization")
plt.xlabel("Columns")
plt.ylabel("Rows")
plt.show()
