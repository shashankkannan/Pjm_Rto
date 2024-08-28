
# PJM Data Analysis Tool

This project is a GUI-based application built with Python's Tkinter library for filtering and processing PJM-related Excel data. It allows users to load Excel files, filter data based on specified dates, states, counties, and transmission owners, and provides step-by-step navigation through the filtering process.

## Table of Contents

- [Features](#features)
- [Tech Stack](#tech-stack)
- [Installation](#installation)
- [Usage](#usage)
- [Contributing](#contributing)
- [Contact](#contact)

## Features

- **Excel File Selection**: Automatically loads PJM Excel data files.
- **Date Filtering**: Filter based on a start date and end date within the "Commercial Operation Milestone" column.
- **State and County Filtering**: Select states and their corresponding counties, and filter data accordingly.
- **Transmission Owner Filtering**: Filter data based on selected transmission owners.
- **Step-by-Step Navigation**: Simple and intuitive navigation with "Next" and "Back" buttons to guide users through the process.
- **Data Visualization**: Displays filtered data in a table format within the GUI.

## Tech Stack

- **Programming Language**: Python
- **GUI Framework**: Tkinter
- **Data Processing**: Pandas
- **Excel Handling**: openpyxl, xlrd (depending on the Excel file format)

## Installation

1. **Clone the Repository**

   ```bash
   git clone https://github.com/shashankkannan/Pjm_Rto.git
   ```

2. **Create a Virtual Environment**

   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Run the Application**

   ```bash
   python Base.py
   ```

## Usage

1. **Load an Excel File**: Click on the "Load Excel File" button and choose a PJM-related Excel file.
2. **Select Date Range**: Input the start and end dates to filter the data based on the "Commercial Operation Milestone" column.
3. **Filter by State and County or Transmission Owner**: Choose to filter by state and county, or by transmission owner. Select the desired options, and the data will be filtered accordingly.
4. **View Filtered Data**: The filtered data is displayed in a table within the application.

## Contributing

Contributions are welcome! Please fork this repository, make your changes, and submit a pull request.

## Contact

For any questions or suggestions, please reach out to me at:

- **Email**: shashank.kannan.cs@gmail.com
- **GitHub**: [shashankkannan](https://github.com/shashankkannan/Pjm_Rto)
