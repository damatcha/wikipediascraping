## Overview
This application is designed to process an Excel file containing names and search for their corresponding Wikipedia pages. It utilizes Flask to create a web server that triggers the processing of the Excel file and retrieves URLs of relevant Wikipedia articles.

## Features
- Reads names from an Excel file.
- Searches Wikipedia for each name in batches.
- Saves the results (name and corresponding Wikipedia URL) to a new Excel file.
- Runs a Flask web server to initiate processing.

## Requirements
- Python 3.x
- Flask
- pandas
- openpyxl
- wikipedia-api

You can install the required packages via pip:

```bash
pip install Flask pandas openpyxl wikipedia-api
```

## Usage
1. **Clone the repository** or download the script.
2. **Run the script** using Python:

   ```bash
   python your_script_name.py
   ```

3. **Enter the path to your Excel file** when prompted. Ensure that the Excel file has a column titled `NAMES` containing the names you wish to search.
4. The Flask server will start automatically and open in your default web browser at `http://127.0.0.1:XXXX/`.
5. The application will process names in batches, searching Wikipedia and saving results to `wikipedia_results.xlsx`.

## Configuration
- **REQUEST_DELAY**: Adjust this variable in the script to change the delay between requests (default is 30 seconds).
- **Region**: The language setting for Wikipedia can be modified by changing the `region` variable in `process_excel_file()` function.

## Error Handling
The application includes basic error handling for:
- Disambiguation errors when multiple pages match a search term.
- Page errors when a page does not exist.
- Issues reading the Excel file.

## Output
The results will be saved in an Excel file named `wikipedia_results.xlsx` in the same directory as your script.

## License
This project is licensed under the MIT License.


