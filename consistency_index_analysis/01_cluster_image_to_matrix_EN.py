import cv2
import numpy as np
import pandas as pd
from collections import Counter

# Define RGB values of the 7 anatomical regions and their corresponding codes (0–6)
color_mapping = {
    (221, 101, 94): 0,   # Cortex (#DD655E)
    (100, 158, 185): 1,  # Fiber (#649EB9)
    (107, 187, 174): 2,  # Ground tissue (#6BBBAE)
    (210, 104, 52): 3,   # Pith (#D26834)
    (149, 33, 36): 4,    # Metaxylem (#952124)
    (227, 132, 156): 5,  # Phloem + Companion cells (#E3849C)
    (154, 74, 44): 6     # Non-pith tissue (#9A4A2C)
}

# Define maximum color distance tolerance (Euclidean distance)
color_tolerance = 50

# Function to calculate Euclidean distance between two RGB colors
def color_distance(c1, c2):
    return np.linalg.norm(np.array(c1) - np.array(c2))

# Load image file (replace path if needed)
image_path = 'bin100.png'
img = cv2.imread(image_path)

# Convert image from BGR (OpenCV default) to RGB
img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)

# Get image dimensions
height, width, _ = img_rgb.shape

# The image is divided into a 107 × 20 grid
block_height = height // 107
block_width = width // 20

# Initialize an empty matrix to store color category codes
color_matrix = np.zeros((107, 20), dtype=int)

# Iterate through each grid cell to classify dominant color
for i in range(107):
    for j in range(20):
        # Extract the current block (cell)
        block = img_rgb[i*block_height:(i+1)*block_height, j*block_width:(j+1)*block_width]
        
        # Skip empty (black or invalid) blocks
        if np.all(block == 0):
            color_matrix[i, j] = -1  # Mark as empty
            continue
        
        # Flatten all pixel RGB values in the block
        pixels = block.reshape(-1, 3)
        
        # Count occurrences of each color
        pixel_counts = Counter(tuple(pixel) for pixel in pixels)
        
        # Identify the most frequent color in the block
        if pixel_counts:
            most_common_color = pixel_counts.most_common(1)[0][0]
            
            # Find the closest reference color using Euclidean distance
            closest_color = min(color_mapping.keys(), key=lambda x: color_distance(x, most_common_color))
            
            # If the color is within acceptable tolerance, assign its code
            if color_distance(closest_color, most_common_color) < color_tolerance:
                color_matrix[i, j] = color_mapping[closest_color]
            else:
                color_matrix[i, j] = -1  # Mark as unmatched
        else:
            color_matrix[i, j] = -1  # No valid pixels found

# Convert the matrix to a pandas DataFrame and save as Excel
df = pd.DataFrame(color_matrix)
excel_path = 'bin100.xlsx'
df.to_excel(excel_path, index=False, header=False)

print(f"Matrix saved to {excel_path}")
