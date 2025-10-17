import cv2
import numpy as np
import pandas as pd
from collections import Counter

# Define 12 anatomical region colors and their corresponding numeric codes (0–11)
color_mapping = {
    (234, 234, 85): 0,   # Pith
    (251, 193, 114): 1,  # Pith ring
    (85, 175, 216): 2,   # Ground tissue
    (57, 83, 161): 3,    # Protoxylem vessel
    (40, 188, 185): 4,   # Metaxylem vessel
    (163, 153, 179): 5,  # Vascular bundle
    (134, 139, 192): 6,  # Phloem
    (0, 107, 189): 7,    # Fiber cells
    (203, 68, 63): 8,    # Companion cells
    (0, 161, 154): 9,    # Cortex cells
    (117, 188, 55): 10,  # Hypodermal cells
    (0, 150, 61): 11     # Epidermal cells
}

# Define maximum color distance tolerance (Euclidean distance)
# The optimal tolerance was empirically determined as 29
color_tolerance = 29

# Function to compute Euclidean distance between two RGB colors
def color_distance(c1, c2):
    return np.linalg.norm(np.array(c1) - np.array(c2))

# Load the annotated image (replace with your own path if needed)
image_path = 'jp-bin100.png'
img = cv2.imread(image_path)

# Convert from BGR (OpenCV default) to RGB color format
img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)

# Get image dimensions
height, width, _ = img_rgb.shape

# The image is divided into a 107 × 20 grid
block_height = height // 107
block_width = width // 20

# Initialize a 107×20 list to store detected color codes as strings
color_matrix = [[""] * 20 for _ in range(107)]

# Iterate through each block to classify its color(s)
for i in range(107):
    for j in range(20):
        # Extract the current block region
        block = img_rgb[i*block_height:(i+1)*block_height, j*block_width:(j+1)*block_width]
        
        # Flatten all pixel colors into a single list
        pixels = block.reshape(-1, 3)
        
        # Count frequency of each RGB color
        pixel_counts = Counter(tuple(pixel) for pixel in pixels)
        
        # Detect all colors that fall within the defined tolerance
        detected_colors = set()
        for pixel_color, count in pixel_counts.items():
            # Find the closest predefined color
            closest_color = min(color_mapping.keys(), key=lambda x: color_distance(x, pixel_color))
            
            # Accept only if the distance is smaller than the tolerance
            if color_distance(closest_color, pixel_color) < color_tolerance:
                detected_colors.add(color_mapping[closest_color])
        
        # Store detected color codes (comma-separated) in the matrix
        if detected_colors:
            color_matrix[i][j] = ",".join(map(str, sorted(detected_colors)))
        else:
            color_matrix[i][j] = "-1"  # Mark as unmatched if no valid color found

# Convert matrix to a DataFrame and export to Excel
df = pd.DataFrame(color_matrix)
excel_path = 'jp-bin100.xlsx'
df.to_excel(excel_path, index=False, header=False)

print(f"Matrix saved to {excel_path}")
