# Python Contract Generator 📄✨

A Python automation script that dynamically generates contract documents from an Excel template. This tool was built to streamline internal workflows, reducing the time spent on manual contract creation by an estimated 50-80%. 
It leverages pandas for data handling, python-docx for document templating, and pythainlp for Thai language localization. 

## 🚀 Overview

Manually creating contracts is a repetitive, time-consuming, and error-prone process. This script automates the entire workflow by reading data from a structured Excel file and populating a pre-defined contract template. It was specifically designed for a company's real estate contracts in Thailand, with built-in logic to handle Thai language specifics. 


### Key Features

*   **Data-Driven:** Populates contracts using data from a simple `.xlsx` file.
*   **Template-Based:** Generates documents from a standardized `.docx` template.
*   **Dual Output:** Creates two distinct versions of each contract (e.g., one for the landlord, one for the tenant).
*   **Thai Language Support:** Integrates `pythainlp` for correct Thai language processing and formatting.
*   **Simple to Use:** Designed for non-technical team members with a one-click executable (`.exe`).


## ⚙️ How It Works

The script follows a simple three-step process:
1.  **Read Data**: It uses `pandas` to open `data_input.xlsx` and read the contract details row by row.
2.  **Process Template**: For each row of data, it opens `.docx` templates and dynamically replaces placeholder text with the corresponding data from the Excel file.
3.  **Generate Output**: It saves the populated documents into the `output/` folder, ready for review and printing.

## 🏁 Getting Started

There are two ways to use this application: for end-users and for developers.

### For End-Users (The Easy Way)

Follow these steps to set up the application without needing any developer tools.

**Step 1: Download the Project Files**
1.  Go to the main page of this repository on GitHub.
2.  Click the green `<> Code` button.
3.  Select **Download ZIP**.
4.  Unzip the downloaded file (e.g., `python-pdf-contract-generator-main.zip`) on your computer. This folder contains the `data_input.xlsx` file you will need to edit.

**Step 2: Download the Application Executable (.exe)**
1.  **[Click here to download `contract_maker.exe` from Google Drive](https://drive.google.com/file/d/1GVAsSgw9n7JNU0VFs6RpfVaz7MFmlksw/view?usp=sharing)**.
2.  Your browser or OS may show a security warning. This is expected behavior for `.exe` files.

**Step 3: Combine and Run**
1.  Move the `contract_maker.exe` file you just downloaded into the project folder you unzipped in Step 1.
2.  The `contract_maker.exe` file should now be in the **same folder** as `data_input.xlsx`.
3.  Open `data_input.xlsx` and fill in your contract details. Save and close the file.
4.  Double-click `contract_maker.exe` to run it.
5.  Check the `output/` folder for your generated documents!

### For Developers

If you want to modify or improve the script, you'll need to run it from the source code.


**Prerequisites:**
*   Python 3.8 or newer
*   pip (Python package installer)


## ⚠️ Limitations & Known Issues

*   **Fixed Template:** The script is currently hard-coded to work with a specific company's contract format. Modifying it for other templates requires changing the source code.
*   **Code Architecture:** As one of my first projects, the script was written in a procedural style. It consists of a long list of functions and lacks an object-oriented design, which can make it difficult to read, maintain, and extend.
*   **Performance:** The current implementation is not optimized for speed and can be slow, especially if you were to adapt it to generate a large number of contracts at once.
