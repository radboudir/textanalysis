# Topic Modeling Tool

This tool is designed to perform topic modeling and other various text analysis on textual data using R for core analysis and a Python-based user interface (UI) built with the Tkinter library. It is particularly effective with long textual responses and provides visual aids through word clouds for shorter surveys.

# Very quick steps
```
python setup_env.py
venv\Scripts\activate
python topic_modeling_app.py
```

# Quick steps
  1. First, run the following script: setup_env.py
  2. Second, activate your virtual environment with the following command: venv\Scripts\activate
  3. Third, run the last script: topic_modeling_app.py

# Execution Policy Issues?
Are you running to any issues regarding Execution Policy? You can temporarily bypass the restriction for the current PowerShell session by running the following command in your terminal: 

```
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

Afterwards, try activating the virtual environment again: 

```
venv\Scripts\activate
```

## Features

- **Topic Modeling:** Ideal for analyzing extensive text data.
- **Word Cloud:** Visualizes the most frequent terms in datasets, best suited for shorter surveys.
- **More to follow**

## Further prerequisites

Before using this tool, please ensure the following steps are completed to set up your environment:

### Install and Set Up Required Libraries

1. **R and Python:** Ensure both R and Python are installed on your computer. Download them from their official websites if necessary.
2. **Library Installation:**
   - **R Libraries:** Open your R console, navigate to the directory containing `requirements.R`, and execute `source('requirements.R')`.
   - **Python Libraries:** Open a command prompt or terminal, navigate to the directory containing `requirements.txt`, and execute `pip install -r requirements.txt`.

### Update Script Paths

- Verify that the TreeTagger tool is correctly installed and its path is appropriately set in **both** R scripts for text processing. Search for the term "teamIR" in the scripts to identify and update these paths.

## How to Run the Tool

### Starting the Application

- Open `topic_modeling_app.py` in your Python IDE (like IDLE or PyCharm) using the file browser.

### Using the Application

- The UI is designed to be user-friendly:
  - Use the **Word Cloud** option for shorter surveys to visualize key terms.
  - Use the **Topic Modeling** option for detailed analysis of more complex text data.
- Once the analysis is complete, the tool automatically saves the results in an Excel file in the same directory as the script.

## Getting Help

If you encounter any issues or need further assistance, please feel free to contact amir.khodaie@ru.nl

