from flask import Flask, jsonify, abort, request, send_from_directory
from flask_cors import CORS
import pandas as pd
import datetime

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

# Path to the Excel file
EXCEL_FILE_PATH = "codec_template.xlsx"  # Replace with your actual Excel file path
MEDIA_ROOT_PATH = "E:/Media/2020/"


@app.route('/data/<path:path>')
def send_report(path):
    # Using request args for path will expose you to directory traversal attacks
    return send_from_directory(MEDIA_ROOT_PATH, path)

@app.route('/googlesheets', methods=['GET'])
def get_sheet_data():
    # Get the 'sheet' parameter from the query string
    sheet_name = request.args.get('sheet')
    if not sheet_name:
        abort(400, description="Missing 'sheet' parameter in the request.")

    # Get the 'offset' parameter from the query string, defaulting to 1 if not provided
    offset = request.args.get('offset', default=1, type=int)

    try:
        # Load the Excel file
        # Force openpyxl and read_only engine argument to avoid locking the excel file and allow excel to save
        excel_data = pd.ExcelFile(EXCEL_FILE_PATH, engine="openpyxl", engine_kwargs={"read_only": True})

        # Check if the requested sheet exists
        if sheet_name not in excel_data.sheet_names:
            abort(404, description=f"Sheet '{sheet_name}' not found in the Excel file.")

        # Read the specific sheet into a DataFrame, using the row at `offset - 1` as the header
        # Force openpyxl and read_only engine argument to avoid locking the excel file and allow excel to save
        df = pd.read_excel(EXCEL_FILE_PATH, sheet_name=sheet_name, skiprows=offset - 1, header=None, engine="openpyxl", engine_kwargs={"read_only": True})

        # Apply transformations to handle NaN, boolean, datetime, and time values
        def transform(value):
            # Handle NaN values
            if pd.isna(value):
                return ""
            # Handle boolean values
            elif isinstance(value, bool):
                return "TRUE" if value else "FALSE"
            # Handle datetime and time values
            elif isinstance(value, pd.Timestamp):
                return value.strftime('%Y-%m-%d %H:%M:%S')  # Convert datetime to string
            elif isinstance(value, datetime.time):
                return value.strftime('%H:%M:%S')
            # Return the value as is for any other type
            return value

        # Apply the transformation to the entire DataFrame using map for all columns
        df = df.map(transform)

        # Check if the DataFrame has exactly two columns
        if df.shape[1] == 2:
            # Convert the DataFrame to a dictionary with the first column as keys and the second as values
            data = dict(zip(df.iloc[:, 0], df.iloc[:, 1]))
        else:
            # Convert the DataFrame to a list of lists, including the headers as the first row
            data = df.values.tolist()

        return jsonify(data)

    except FileNotFoundError:
        abort(404, description="Excel file not found.")
    except Exception as e:
        abort(500, description=f"An error occurred: {str(e)}")        

if __name__ == '__main__':
    app.run(debug=True)
