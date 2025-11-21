from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
import os
import uuid
from pathlib import Path
import shutil
from process_invoices import process_invoices_web

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max upload
app.config['SECRET_KEY'] = 'your-secret-key-change-this'

# Folder configuration
UPLOAD_FOLDER = 'temp/uploads'
OUTPUT_FOLDER = 'temp/outputs'
TEMPLATE_PATH = 'template_files/recon_template.xlsx'

# Create necessary folders
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs('template_files', exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    # Check if files were uploaded
    if 'invoices' not in request.files:
        return jsonify({'error': 'No files uploaded'}), 400

    files = request.files.getlist('invoices')

    if not files or files[0].filename == '':
        return jsonify({'error': 'No files selected'}), 400

    client_name = request.form.get('clientName', '').strip()  # Optional client name

    # Create unique session folder for this upload
    session_id = str(uuid.uuid4())
    session_upload_folder = os.path.join(UPLOAD_FOLDER, session_id)
    os.makedirs(session_upload_folder, exist_ok=True)

    # Save uploaded files
    uploaded_count = 0
    for file in files:
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(session_upload_folder, filename))
            uploaded_count += 1

    if uploaded_count == 0:
        shutil.rmtree(session_upload_folder, ignore_errors=True)
        return jsonify({'error': 'No valid Excel files (.xlsx) uploaded'}), 400

    # Check if template exists
    if not os.path.exists(TEMPLATE_PATH):
        shutil.rmtree(session_upload_folder, ignore_errors=True)
        return jsonify({'error': 'Template reconciliation file not found. Please contact administrator.'}), 500

    # Process invoices
    output_filename = f"Reconciliation_{session_id}.xlsx"
    output_path = os.path.join(OUTPUT_FOLDER, output_filename)

    try:
        # Call the processing function with client name
        result = process_invoices_web(
            invoice_folder_path=session_upload_folder,
            template_recon_path=TEMPLATE_PATH,
            output_path=output_path,
            client_name=client_name
        )

        # Clean up uploaded files
        shutil.rmtree(session_upload_folder, ignore_errors=True)

        if result['success']:
            # Send file to user
            response = send_file(
                output_path,
                as_attachment=True,
                download_name='Reconciliation.xlsx',
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            # Clean up output file after sending
            @response.call_on_close
            def cleanup():
                try:
                    if os.path.exists(output_path):
                        os.remove(output_path)
                except:
                    pass
            return response
        else:
            if os.path.exists(output_path):
                os.remove(output_path)
            return jsonify({'errorgit rm -r --cached .
git add .
': result['message']}), 400

    except Exception as e:
        shutil.rmtree(session_upload_folder, ignore_errors=True)
        if os.path.exists(output_path):
            os.remove(output_path)
        return jsonify({'error': f'Unexpected error: {str(e)}'}), 500

@app.route('/health')
def health():
    return jsonify({'status': 'healthy', 'template_exists': os.path.exists(TEMPLATE_PATH)})

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
