import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# 1. Load input matrices
bin100_path = "bin100.xlsx"      # Cluster-based classification matrix (0–6)
jp_bin100_path = "jp-bin100.xlsx"  # Manually annotated classification matrix (0–11)

bin100_df = pd.read_excel(bin100_path, header=None)
jp_bin100_df = pd.read_excel(jp_bin100_path, header=None)

# 2. Define correspondence relationships between bin100 categories and annotated types
# Each bin100 cluster (0–6) corresponds to one or more anatomical cell types (0–11)
correspondence = {
    0: {9, 10, 11},        # Cortex-related
    1: {7},                # Fiber
    2: {2},                # Ground tissue
    3: {0, 1},             # Pith and pith ring
    4: {3, 4},             # Xylem (proto- and metaxylem)
    5: {6, 8},             # Phloem + companion cells
    6: {2, 3, 4, 5, 6, 7, 8, 9, 10, 11}  # Non-pith mixed region
}

# 3. Initialize an empty matching matrix (same dimensions as input)
match_matrix = pd.DataFrame(0, index=bin100_df.index, columns=bin100_df.columns)

# 4. Iterate over each cell and check whether the classification pairs match
for i in range(bin100_df.shape[0]):
    for j in range(bin100_df.shape[1]):
        bin_value = bin100_df.iloc[i, j]  # Value from bin100 matrix (0–6)
        jp_values = str(jp_bin100_df.iloc[i, j]).split(',')  # Values from annotation matrix (0–11)

        # Convert annotation entries to a set of integers
        jp_values = {int(val) for val in jp_values if val.isdigit()}
        
        # Check if there is an intersection between corresponding categories
        if bin_value in correspondence and jp_values.intersection(correspondence[bin_value]):
            match_matrix.iloc[i, j] = 1  # Mark as matched

# 5. Compute the overall consistency (Jaccard-style index)
total_cells = match_matrix.size                   # Total number of grid cells
matched_cells = match_matrix.sum().sum()          # Number of matched cells
consistency_index = matched_cells / total_cells   # Consistency index
print(f"Consistency index: {consistency_index:.3f}")

# 6. Save the matching results matrix to Excel
match_matrix.to_excel("match_result.xlsx", index=False, header=False)

# 7. Generate a visualization heatmap
plt.figure(figsize=(12, 6))
plt.imshow(match_matrix, cmap="coolwarm", aspect="auto", interpolation="nearest")
plt.colorbar(label="Matching Status (0 = Not Matched, 1 = Matched)")
plt.title("Spatial Matching Matrix Visualization")
plt.xlabel("Columns")
plt.ylabel("Rows")
plt.tight_layout()
plt.show()
