from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import os
from werkzeug.utils import secure_filename
from document_formatter import DocumentFormatter
import tempfile
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'your-secret-key-change-this'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Ensure upload directory exists
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

ALLOWED_EXTENSIONS = {'docx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    try:
        # Check if files were uploaded
        if 'template_file' not in request.files or 'target_file' not in request.files:
            flash('Please select both template and target files')
            return redirect(url_for('index'))
        
        template_file = request.files['template_file']
        target_file = request.files['target_file']
        
        # Check if files are selected
        if template_file.filename == '' or target_file.filename == '':
            flash('Please select both files')
            return redirect(url_for('index'))
        
        # Check file extensions
        if not (allowed_file(template_file.filename) and allowed_file(target_file.filename)):
            flash('Only .docx files are allowed')
            return redirect(url_for('index'))
        
        # Create temporary files
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_template:
            template_file.save(temp_template.name)
            template_path = temp_template.name
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_target:
            target_file.save(temp_target.name)
            target_path = temp_target.name
        
        # Process the documents
        formatter = DocumentFormatter()
        output_path = formatter.apply_formatting(template_path, target_path)
        
        # Clean up temporary files
        os.unlink(template_path)
        os.unlink(target_path)
        
        # Generate output filename
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f'formatted_document_{timestamp}.docx'
        
        # Send file and clean up
        return send_file(output_path, as_attachment=True, download_name=output_filename)
        
    except Exception as e:
        flash(f'Error processing documents: {str(e)}')
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)