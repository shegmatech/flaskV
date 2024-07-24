# THIRD

import os
import logging
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
import pandas as pd

app = Flask(__name__)

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Configure upload folder
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def split_excel_into_chunks(file_path, chunk_size=3500): #i will still make the chunk size variable, so that i can accept input from user
    dtypes = {
        'Warrant Number': str,
        'Bank Acct.': str,
        'Bank Code': str,
    }

    df = pd.read_excel(file_path)

    for col, dtype in dtypes.items():
        df[col] = df[col].astype(dtype)

    df['Bank Acct.'] = df['Bank Acct.'].str.zfill(10)
    df['Bank Code'] = df['Bank Code'].str.zfill(3)
    df['Warrant Number'] = df['Warrant Number'].apply(lambda x: str(x).zfill(12))

    num_chunks = (len(df) // chunk_size) + 1

    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        for i in range(num_chunks):
            start_idx = i * chunk_size
            end_idx = min((i + 1) * chunk_size, len(df))

            sheet_name = f'OTHERS_{i+1}'

            df_chunk = df.iloc[start_idx:end_idx]
            df_chunk.to_excel(writer, sheet_name=sheet_name, index=False)
            logger.info(f"Splitted chunk {i+1}")

    logger.info("Done splitting the data into worksheets!")

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    message = ''
    if request.method == 'POST':
        if 'file' not in request.files:
            message = 'No file part'
        else:
            file = request.files['file']
            if file.filename == '':
                message = 'No selected file'
            elif file and file.filename.endswith('.xlsx'):
                filename = secure_filename(file.filename)
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(file_path)
                try:
                    split_excel_into_chunks(file_path)
                    message = 'Done splitting the data into worksheets! See your download path for splitted file'
                    return send_file(file_path, as_attachment=True),message
                except Exception as e:
                    logger.error(f"Error processing file: {str(e)}")
                    message = 'Error processing file'
            else:
                message = 'Invalid file format. Please upload an Excel file.'
    return render_template('home.html', message=message)

if __name__ == "__main__":
    app.run(debug=True)
