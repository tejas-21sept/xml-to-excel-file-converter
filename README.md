# Tally XML to Excel Converter API

This project is a Flask-based API that converts Tally XML files into Excel (XLSX) format. It simplifies handling Tally data, providing an efficient and easy-to-use service for transforming financial data into a spreadsheet.

---

## Features

- Accepts Tally XML files as input.
- Validates file format before processing.
- Parses XML data and generates a structured Excel file.
- Returns the Excel file as a downloadable response.

---

## Installation

Follow these steps to set up and run the project locally:

### Prerequisites

- Python 3.8+ installed on your system.
- `pip` package manager.

### Steps

1. Clone the repository:

   ```bash
   git clone https://github.com/tejas-21sept/xml-to-excel-file-converter.git
   cd xml-to-excel-file-converter

   ```

2. Create a virtual environment:

   ```bash
   python -m venv venv
   source venv/bin/activate  # For Linux/Mac
   venv\Scripts\activate     # For Windows

   ```

3. Install dependencies:

   ```bash
   pip install -r requirements.txt

   ```

4. Set Flask environment variables:

   ```bash
   export FLASK_APP=app   # For Linux/Mac
   set FLASK_APP=app      # For Windows

   ```

5. Run the application:
   ```bash
   flask run
   ```

## Usage

    ### API Endpoint
    URL: /api/parse-xml
    Method: POST
    Description: Accepts an tally exported XML file and returns an Excel file.
    Example Request
    Use a tool like Postman or curl to send a POST request.

    Using curl:
        ```bash
        curl -X POST -F "file=@example.xml" http://127.0.0.1:5000/api/parse-xml --output transactions.xlsx
        ```
    Using Postman:
    1) Set the method to POST.
    2) Enter the endpoint URL: http://127.0.0.1:5000/api/parse-xml.
    3) Under the Body tab, choose form-data.
    4) Add a key file with the XML file as the value.
    5) Send the request and download the Excel response.

## Folder Structure

    ```bash
    tally-xml-to-excel/

    ├── app/
    │ ├── blueprints/
    │ │ ├── excel_converter/
    │ │ │ ├── **init**.py
    │ │ │ ├── routes.py
    │ │ │ └── views.py
    │ └── **init**.py
    ├── .gitignore
    ├── requirements.txt
    ├── README.md
    └── run.py
    ```

## Future Enhancements

    - Add database integration to store parsed data.
    - Implement advanced logging and error tracking.
    - Add support for more input formats (e.g., JSON).
    - Host the API for public or team-wide use.

## Contribution

This project is created and maintained by Tejas Dodal.
