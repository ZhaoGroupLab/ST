import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import Wedge  # Use Wedge to draw circular sectors

# Define RGB color mapping for each cluster (0–6)
color_mapping = {
    0: (221, 101, 94),   # Cortex (#DD655E)
    1: (100, 158, 185),  # Fiber (#649EB9)
    2: (107, 187, 174),  # Ground tissue (#6BBBAE)
    3: (210, 104, 52),   # Pith (#D26834)
    4: (149, 33, 36),    # Metaxylem vessel (#952124)
    5: (227, 132, 156),  # Phloem + companion cells (#E3849C)
    6: (154, 74, 44)     # Non‑pith region (#9A4A2C)
}

# Read the matrix file exported from previous step
file_path = 'bin100.xlsx'  # Input Excel file (no header)
df = pd.read_excel(file_path, header=None)

# Convert DataFrame to NumPy array
matrix = df.to_numpy()

# Initialize figure
fig, ax = plt.subplots(figsize=(10, 10))

# Define circle size and spacing
circle_size_factor = 20  # Adjust for desired spacing and scale

# Iterate through every cell to plot its color representation
for i in range(matrix.shape[0]):  # Rows
    for j in range(matrix.shape[1]):  # Columns
        cell_value = matrix[i, j]
        if cell_value == -1:
            continue  # Skip empty cells

        # Parse comma‑separated codes (in case multiple values exist)
        cell_values = [int(x) for x in str(cell_value).split(',')]

        # Map each numeric code to its RGB color
        colors_for_cell = [color_mapping[num] for num in cell_values if num in color_mapping]

        # Draw one or more wedges (colored sectors) per cell
        if len(colors_for_cell) > 0:
            num_colors = len(colors_for_cell)
            angle_per_color = 360 / num_colors  # Equal angular division
            center_x = j * 2 * circle_size_factor
            center_y = i * 2 * circle_size_factor

            for idx, color in enumerate(colors_for_cell):
                start_angle = angle_per_color * idx
                end_angle = angle_per_color * (idx + 1)
                wedge = Wedge(
                    (center_x, center_y),
                    circle_size_factor,
                    start_angle,
                    end_angle,
                    color=np.array(color) / 255.0,  # Normalize RGB
                    ec="black",
                    lw=0.5
                )
                ax.add_patch(wedge)

# Configure axes and appearance
ax.set_xlim(-circle_size_factor, matrix.shape[1] * 2 * circle_size_factor)
ax.set_ylim(-circle_size_factor, matrix.shape[0] * 2 * circle_size_factor)
ax.set_aspect('equal')
ax.axis('off')  # Hide axes for clean visualization

# Save figure as an editable PDF
output_pdf_path = 'bin100_figure.pdf'
plt.savefig(output_pdf_path, format='pdf', dpi=600)
plt.close()

print(f"PDF file saved to: {output_pdf_path}")
plt.show()
