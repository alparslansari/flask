import os
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'supersecretkey'  # Needed for flashing messages

# Specify the upload folder and allowed file types (Excel files)
UPLOAD_FOLDER = 'uploads/'  # Path to save uploaded files
PROCESSED_FOLDER = 'processed/'  # Folder to save the processed CSV
ALLOWED_EXTENSIONS = {'xlsx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER

# Helper function to check allowed file extensions
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def upload_form():
    return render_template('upload.html')  # This will render the upload form

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('No file part')
        return redirect(request.url)
    
    file = request.files['file']
    
    if file.filename == '':
        flash('No selected file')
        return redirect(request.url)
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        flash(f'File {filename} uploaded successfully!')

        # Process the Excel file after upload
        output_csv = process_excel(filepath)
        
        # Provide download link for the processed CSV file
        return send_file(output_csv, as_attachment=True)

def process_excel(filepath):
    # Read the Excel file
    df = pd.read_excel(filepath)
    
    # Sort by column D (assuming it's the fourth column, which in pandas is 'df.iloc[:, 3]')
    sorted_df = df.sort_values(by=df.columns[3], ascending=True)
    
    # Get the top 10 rows
    top_10_df = sorted_df.head(10)
    
    # Create a CSV output path
    output_csv_path = os.path.join(app.config['PROCESSED_FOLDER'], 'top_10_results.csv')
    
    # Save the top 10 rows to CSV
    top_10_df.to_csv(output_csv_path, index=False)
    
    # Return the CSV file path for download
    return output_csv_path

if __name__ == "__main__":
    # Create the upload and processed directories if they don't exist
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    if not os.path.exists(PROCESSED_FOLDER):
        os.makedirs(PROCESSED_FOLDER)
    
    app.run(debug=True)
