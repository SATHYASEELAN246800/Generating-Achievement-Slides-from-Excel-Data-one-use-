# Generating-Achievement-Slides-from-Excel-Data-one-use-

# Project Title: Generating Achievement Slides from Excel Data

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/drive/1his5J4IsVmb5vrKUkmDVr86fNOhu7h48#scrollTo=0d57c2a4)

## Project Description

This project automates the process of extracting student achievement data from an Excel file, transforming it, and generating a formatted PowerPoint presentation. The presentation is based on a reference slide design and populates tables on each slide with the extracted data.

## 📁 Files

- **Input Excel File:**  
  `/content/drive/My Drive/Colab Notebooks/demo hr work excel file.xlsx`  
  Original Excel file containing raw student achievement records.

- **Intermediate Excel File:**  
  `/content/drive/My Drive/Colab Notebooks/new_excel.xlsx`  
  Contains the transformed data with combined event and institution names, and student names.

- **Reference PowerPoint File:**  
  `/content/drive/My Drive/Colab Notebooks/students outside achievement.pptx`  
  Used as a design template for layout and table formatting.

- **Output PowerPoint File:**  
  `/content/drive/My Drive/Colab Notebooks/generated_achievement_slides.pptx`  
  Final PowerPoint file with formatted tables filled with student data.

## 🔧 Project Steps

1. **Mount Google Drive**  
   Access files stored in Google Drive for reading and writing.

2. **Install Necessary Libraries**  
   Install `python-pptx` to work with PowerPoint files in Python.

3. **Read Excel Data**  
   Load the intermediate Excel file (`new_excel.xlsx`) into a pandas DataFrame.

4. **Load Reference PPTX File**  
   Use the reference slide design to replicate the look and feel in the output.

5. **Define Table and Slide Layout**  
   Configure layout for consistent formatting based on the reference design.

6. **Create New Presentation**  
   Initialize a new presentation and add slides dynamically based on data volume.

7. **Populate Table in Slides**  
   Each slide contains a table with:
   - Name of the Event/Conducting college
   - Name of the Students
   - Awards (if any) — left blank for manual entry

8. **Apply Styling**  
   Adjust table formatting (fonts, padding, widths) to match the reference style (e.g., Calibri font, table alignment).

9. **Save Presentation**  
   The final `.pptx` is saved back to your Google Drive path.

## 🧩 Dependencies

- `pandas`
- `python-pptx`
- `google.colab` (for accessing Google Drive)

Install them in Colab using:
```python
!pip install python-pptx
