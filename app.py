from flask import Flask, jsonify, abort, request
from flask_cors import CORS  # Import CORS from flask_cors
import pandas as pd

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

# Path to the Excel file
EXCEL_FILE_PATH = "codec_template.xlsx"  # Replace with your actual Excel file path

@app.route('/googlesheets', methods=['GET'])
def get_sheet_data():
    # Get the 'sheet_name' parameter from the query string
    sheet_name = request.args.get('sheet')
    if not sheet_name:
        abort(400, description="Missing 'sheet' parameter in the request.")

    # Get the 'offset' parameter from the query string, defaulting to 1 if not provided
    offset = request.args.get('offset', default=1, type=int)

    try:
        # Load the Excel file
        excel_data = pd.ExcelFile(EXCEL_FILE_PATH)

        # Check if the requested sheet exists
        if sheet_name not in excel_data.sheet_names:
            abort(404, description=f"Sheet '{sheet_name}' not found in the Excel file.")

        # Read the specific sheet into a DataFrame, using the row at `offset - 1` as the header
        df = pd.read_excel(EXCEL_FILE_PATH, sheet_name=sheet_name, header=offset - 1)

        # Apply transformations to handle NaN and boolean values
        def transform(value):
            if pd.isna(value):
                return ""
            elif isinstance(value, bool):
                return "TRUE" if value else "FALSE"
            return value

        # Apply the transformation to the entire DataFrame
        df = df.map(transform)

        # Check if the DataFrame has exactly two columns
        if df.shape[1] == 2:
            # Convert the DataFrame to a dictionary with the first column as keys and the second as values
            data = dict(zip(df.iloc[:, 0], df.iloc[:, 1]))
        else:
            # Convert the DataFrame to a list of lists, including the headers as the first row
            data = [df.columns.tolist()] + df.values.tolist()

        return jsonify(data)

    except FileNotFoundError:
        abort(404, description="Excel file not found.")
    except Exception as e:
        abort(500, description=f"An error occurred: {str(e)}")

if __name__ == '__main__':
    app.run(debug=True)
