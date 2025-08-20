# Excel Processor

📊 **Python script for Excel automation**: calculates averages for numeric columns, sorts data, and generates a bar chart directly in Excel.

---

## Features

- Reads an Excel file (`input.xlsx`)
- Calculates the **Average** for each row
- Sorts the table by the first numeric column
- Generates a **BarChart** for visual representation
- Saves results to `output.xlsx`

---

## Requirements

Python 3.9+ and the following packages:

pandas
openpyxl

Install dependencies:


```bash
pip install -r requirements.txt
```

Usage:

Place your Excel file in the project folder and name it input.xlsx.

Run the script

The result will be saved in output.xlsx with the new Average column and a bar chart.
