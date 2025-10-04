import random
import io
import pandas as pd
import tempfile
import os
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'

# Remove the global variable - we'll use temporary files instead

# --- I. Randomized Quick Sort Algorithm (for Manual Input) ---

def randomized_quick_sort(data, low, high):
    """Main Quick Sort function with randomization."""
    if low < high:
        p = randomized_partition(data, low, high)
        randomized_quick_sort(data, low, p - 1)
        randomized_quick_sort(data, p + 1, high)

def randomized_partition(data, low, high):
    """Chooses a random pivot and partitions the data."""
    r = random.randint(low, high)
    data[r], data[high] = data[high], data[r]
    return partition(data, low, high)

def partition(data, low, high):
    """
    Standard partitioning process using case-insensitive comparison (str.lower()).
    """
    pivot = data[high]
    pivot_key = pivot.lower() 
    
    i = low - 1

    for j in range(low, high):
        current_key = data[j].lower()
        
        # Case-insensitive comparison
        if current_key <= pivot_key:
            i = i + 1
            data[i], data[j] = data[j], data[i]

    data[i + 1], data[high] = data[high], data[i + 1]
    return i + 1

# --- II. Flask Application Routes ---

@app.route('/')
def index():
    """Renders the main page with both sorting options."""
    return render_template('combined_sorter_index.html') 

# === Route for Manual Text Sorting ===
@app.route('/sort_manual', methods=['POST'])
def sort_manual():
    """Handles the POST request to sort the customer names from text input."""
    
    names_input = request.form.get('customer_names', '')
    
    # Convert the string into a list of names
    customer_names = [name.strip() for name in names_input.split(',') if name.strip()]

    if customer_names:
        n = len(customer_names)
        randomized_quick_sort(customer_names, 0, n - 1)
        sorted_output = ", ".join(customer_names)
    else:
        sorted_output = "No names provided to sort."

    return jsonify({'sorted_names': sorted_output})

# === Routes for Excel Sheet Sorting (UPDATED for Multiple Columns) ===
@app.route('/upload_and_sort', methods=['POST'])
def upload_and_sort():
    """Handles the Excel file upload and sorting using Pandas."""
    
    if 'excel_file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['excel_file']
    sort_column_names = request.form.get('sort_column_name', '').strip()
    
    if file.filename == '' or not sort_column_names:
        return jsonify({'error': 'Please select a file and specify at least one column name.'}), 400

    if file.filename.endswith(('.xlsx', '.xls')):
        try:
            df = pd.read_excel(file, engine='openpyxl')
            
            # Split multiple columns by comma and strip whitespace
            columns_to_sort = [col.strip() for col in sort_column_names.split(',') if col.strip()]
            
            # Validate all columns exist
            missing_columns = [col for col in columns_to_sort if col not in df.columns]
            if missing_columns:
                return jsonify({
                    'error': f'Columns {missing_columns} not found in the sheet. Available columns: {list(df.columns)}'
                }), 400
            
            # Sort by multiple columns with case-insensitive comparison
            if len(columns_to_sort) == 1:
                # Single column sorting (original behavior)
                df = df.sort_values(
                    by=columns_to_sort[0], 
                    key=lambda x: x.astype(str).str.lower(),
                    ascending=True,
                    ignore_index=True
                )
            else:
                # Multiple column sorting
                # Create a temporary DataFrame for sorting keys
                temp_df = df[columns_to_sort].copy()
                
                # Apply case-insensitive transformation to all string columns
                for col in columns_to_sort:
                    if df[col].dtype == 'object':  # String columns
                        temp_df[col] = temp_df[col].astype(str).str.lower()
                
                # Get the sorted indices
                sorted_indices = temp_df.sort_values(
                    by=columns_to_sort,
                    ascending=[True] * len(columns_to_sort)  # All ascending
                ).index
                
                # Reorder the original DataFrame
                df = df.loc[sorted_indices].reset_index(drop=True)
            
            # Create a temporary file to store the sorted data
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
            
            # Save the sorted DataFrame to the temporary file
            with pd.ExcelWriter(temp_file.name, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sorted Data')
            
            # Success message for single or multiple columns
            if len(columns_to_sort) == 1:
                message = f'File sorted successfully by column: "{columns_to_sort[0]}".'
            else:
                message = f'File sorted successfully by columns: {", ".join(columns_to_sort)}.'
            
            return jsonify({
                'success': True,
                'message': message,
                'download_filename': f"sorted_{secure_filename(file.filename)}",
                'temp_file_path': temp_file.name  # Send the temporary file path to frontend
            })
            
        except Exception as e:
            return jsonify({'error': f'An error occurred: {str(e)}'}), 500
    
    return jsonify({'error': 'Invalid file format. Must be .xlsx or .xls.'}), 400

@app.route('/download', methods=['GET'])
def download():
    """Handles the download of the sorted Excel file."""
    temp_file_path = request.args.get('file_path')
    filename = request.args.get('filename', 'sorted_file.xlsx')
    
    if temp_file_path and os.path.exists(temp_file_path):
        # Create response with file
        response = send_file(
            temp_file_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            download_name=filename,
            as_attachment=True
        )
        
        # Clean up the temporary file after the response is sent
        @response.call_on_close
        def cleanup_temp_file():
            try:
                os.unlink(temp_file_path)
            except Exception as e:
                print(f"Error deleting temporary file: {e}")
        
        return response
    else:
        return "File not found or session expired.", 404

# Cleanup function to remove any orphaned temporary files on startup
@app.before_request
def cleanup_old_temp_files():
    """Optional: Clean up any old temporary files that might be left behind"""
    temp_dir = tempfile.gettempdir()
    for filename in os.listdir(temp_dir):
        if filename.startswith('tmp') and filename.endswith('.xlsx'):
            filepath = os.path.join(temp_dir, filename)
            try:
                # Delete files older than 1 hour
                if os.path.getmtime(filepath) < (time.time() - 3600):
                    os.unlink(filepath)
            except:
                pass

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)