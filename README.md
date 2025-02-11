# Dynamic PowerPoint Automation

## Overview
This project automates PowerPoint deck creation and modification, integrating Power BI exports into structured templates. It ensures consistency in presentation formatting and content updates.

## Features
- **Automated Renaming**: Cleans and renames PowerPoint files based on structured naming conventions.
- **Power BI Export Integration**: Maps Power BI export slides into corresponding PowerPoint sections.
- **Slide Processing**: Handles slide modifications such as text box removal, AI-generated paragraph insertions, and transparency adjustments.
- **Template Mapping**: Uses predefined mappings to place content correctly in PowerPoint decks.

## Setup & Requirements
### Prerequisites
- Python (>= 3.8)
- Required Python packages:
  ```sh
  pip install pandas pptx easyocr pywin32 pillow



---

### **Executables.ipynb**
This notebook provides multiple utility scripts for processing PowerPoint presentations, including template modifications, AI text insertions, and image transparency adjustments.

#### **README for Executables**
```markdown
# PowerPoint Executable Utilities

## Overview
This project provides a suite of scripts to automate PowerPoint modifications, including template application, text box handling, and image transparency adjustments.

## Features
- **Apply PowerPoint Templates**: Merges content from source slides into standardized templates.
- **Text Box Removal & Insertion**: Cleans unnecessary text boxes and inserts AI-generated text.
- **Image Processing**: Makes white backgrounds transparent in images inside slides.
- **Hyperlink & Notes Cleanup**: Removes embedded links and slide notes for a cleaner presentation.

## Setup & Requirements
### Prerequisites
- Python (>= 3.8)
- Required Python packages:
  ```sh
  pip install python-pptx pillow tkinter
