import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import Wedge  # Use Wedge to draw circular sectors

# Define color mapping for matched types (you can expand this for more categories)
color_mapping = {
    0: (225, 225, 89),   # Match type A (e.g., pith)
    1: (47, 178, 175),   # Match type B (e.g., pith ring)
}

# RGB reference:
# e1e159 → (225, 225, 89)
# 7eb55a → (126, 181, 90)
# 2fb2af → (47, 178, 175)

# Load matching result matrix from Excel
file_path = 'match_result.xlsx'
df = pd.read_excel(file_path, header=None)  # No header expected

# Convert DataFrame to NumPy matrix
matrix = df.to_numpy()

# Setup figure
fig, ax = plt.subplots(figsize=(10, 10))

# Circle layout parameters
circle_size_factor = 20  # Size and spacing scale factor

# Iterate over each grid cell to draw its colored representation
for i in range(matrix.shape[0]):
    for j in range(matrix.shape[1]):
        cell_value = matrix[i, j]
        if cell_value == -1:
            continue  # Skip empty cells

        # Support for multiple values (comma-separated, if any)
        cell_values = [int(x) for x in str(cell_value).split(',')]

        # Retrieve corresponding colors for each value
        colors_for_cell = [color_mapping[num] for num in cell_values if num in color_mapping]

        # Draw one or more wedges (circular sectors) per cell
        if len(colors_for_cell) > 0:
            num_colors = len(colors_for_cell)
            angle_per_color = 360 / num_colors
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
                    color=np.array(color) / 255.0,
                    ec="black",
                    lw=0.5
                )
                ax.add_patch(wedge)

# Final layout and axis adjustments
ax.set_xlim(-circle_size_factor, matrix.shape[1] * 2 * circle_size_factor)
ax.set_ylim(-circle_size_factor, matrix.shape[0] * 2 * circle_size_factor)
ax.set_aspect('equal')
ax.axis('off')  # Hide axes

# Save as high-resolution editable PDF
output_pdf_path = 'match_figure.pdf'
plt.savefig(output_pdf_path, format='pdf', dpi=600)
plt.close()

print(f"PDF file saved to: {output_pdf_path}")
plt.show()
