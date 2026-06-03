# Automated Market Intelligence Pipeline: OCR-Based Pricelist Detection

[![Python](https://img.shields.io/badge/python-3670A0?style=for-the-badge&logo=python&logoColor=ffdd54)](https://www.python.org/)
[![Pandas](https://img.shields.io/badge/pandas-%23150458.svg?style=for-the-badge&logo=pandas&logoColor=white)](https://pandas.pydata.org/)
[![OpenCV](https://img.shields.io/badge/opencv-%23white.svg?style=for-the-badge&logo=opencv&logoColor=white)](https://opencv.org/)
[![Google Colab](https://img.shields.io/badge/Colab-F9AB00?style=for-the-badge&logo=googlecolab&color=525252)](https://colab.research.google.com/)

## Project Overview
This project was developed to solve a real-world operational bottleneck encountered during my internship in the Market Intelligence and Research division at one of the largest oil and gas companies in Indonesia. It processes field survey data (Excel reports), dynamically fetches undocumented images, and utilizes Computer Vision (OpenCV) alongside Optical Character Recognition (PyTesseract) to autonomously scan, classify, and isolate authentic competitor pricelist photos from unstructured field data.

---

## Business Context & Impact
In market research, field teams often submit thousands of raw photos containing mixed displays, banners, and price tags. Sorting through this data manually is highly inefficient and prone to human error. This automated pipeline addresses the bottleneck by:

* **Eliminating Manual Sorting:** Automatically filters raw survey data and discarded irrelevant promotional banners.
* **Increasing Data Accuracy:** Employs heuristic validation (detecting numerical price chains and technical viscosity terms like "10W") to ensure only genuine pricelists are captured.
* **Providing Structured Storage:** Renames and organizes validated images based on Outlet ID, Frontliner ID, and Timestamp, creating a clean database ready for further market and pricing analysis.

---

## Architecture & Workflow

1. **Data Unnesting & Cleaning:** Reads raw Excel files, unpivots (melts) static columns, and explodes nested URL arrays into a clean, iterable DataFrame.
2. **Image Pre-processing:** Fetches images dynamically and applies Grayscale, Scaling, and Otsu Thresholding via OpenCV to enhance text legibility and mitigate background noise/shadows.
3. **Heuristic OCR Validation:** Extracts text using PyTesseract configured with a custom character whitelist. The system validates the image by scanning for explicit keywords or identifying explicit pricing patterns (e.g., recurring "000" chains).
4. **Automated Storage Pipeline:** Validated images are saved into a centralized directory with a standardized naming convention, systematically discarding invalid or blank photos.

---

## Repository Structure

* `Pricelist_READ_OCR.py` : The primary Python script containing the data wrangling pipeline and OCR logic.
* `dummy_survey_report.xlsx` : A simulated, anonymized dataset mimicking the original survey report structure for safe testing and replication.
* `README.md` : Project documentation.

---

## Getting Started

### Prerequisites
Ensure you are running this within a Google Colab environment or a local Jupyter setup with Google Drive access. The following system dependency is required for OCR:

```bash
!apt-get install tesseract-ocr
