# Spatial Consistency Evaluation Between Clustered and Annotated Cell Types at Bin100 Resolution

This project provides a reproducible pipeline to evaluate the spatial consistency between clustering results (bin100 resolution) from spatial transcriptomics and anatomically annotated cell types.  
By computing the **Jaccard index**, the pipeline quantifies the degree of spatial overlap between algorithmically defined clusters and manually segmented anatomical regions.

---

## 📁 Project Overview

This repository includes the following analysis modules:

- Extraction and formatting of classification matrices from cluster and annotation images;
- Spatial matrix encoding and transformation;
- Calculation and visualization of the Jaccard consistency index;
- Visualization of spatial distribution for both clustering and annotation.

---

## 🛠️ Script Descriptions (in execution order)

| Script | Description |
|--------|-------------|
| `01_cluster_image_to_matrix.py` | Reads the clustering image `bin100.png` and generates the spatial classification matrix `bin100.xlsx` |
| `02_annotation_image_to_matrix.py` | Reads the manually annotated image `jp-bin100.png` and generates `jp-bin100.xlsx` |
| `03_consistency_jaccard_analysis.py` | Compares the two matrices, calculates the Jaccard index, and outputs `match_result.xlsx` and a heatmap |
| `04_plot_cluster_bubbles.py` | Visualizes the spatial distribution of `bin100.xlsx` as a sector-based schematic figure |
| `05_plot_annotation_bubbles.py` | Visualizes the matching result (`match_result.xlsx`) using sector diagrams |

---

## 📊 Output Files

| File Name | Description |
|-----------|-------------|
| `bin100.xlsx` | Spatial classification matrix derived from the clustering image |
| `jp-bin100.xlsx` | Spatial classification matrix derived from the annotation image |
| `match_result.xlsx` | Matching matrix based on Jaccard comparisons (binary 0/1) |
| `match_figure.pdf` | Visualization of the spatial consistency results |
| `bin100.png`, `jp-bin100.png` | Input images (clustering and annotation) |
| `bin100_figure.pdf` | Visualization of the bin100 clustering spatial distribution |

---

## 🔧 How to Use

1. Prepare the clustering result image `bin100.png` and the annotation image `jp-bin100.png`;  
2. Run the five scripts in order: `01 → 02 → 03 → 04 → 05`;  
3. Use the generated matrices and figures for evaluation of spatial consistency or supplementary figure production.

---

## 👤 Authors

- Data processing and scripting: qntriam@163.com  
- Method design and implementation: qntriam@163.com  

Please contact the authors for any questions or suggestions.