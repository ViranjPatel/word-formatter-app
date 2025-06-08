from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import os
from werkzeug.utils import secure_filename
from document_formatter import DocumentFormatter
import tempfile
from datetime import datetime
import logging

# Configure logging
logging.basicConfig(level=logging.WARNING)

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
    temp_files = []  # Track temporary files for cleanup
    
    try:
        # Validate file uploads
        if 'template_file' not in request.files or 'target_file' not in request.files:
            flash('Please select both template and target files')
            return redirect(url_for('index'))
        
        template_file = request.files['template_file']
        target_file = request.files['target_file']
        
        # Validate file selection
        if not template_file.filename or not target_file.filename:
            flash('Please select both files')
            return redirect(url_for('index'))
        
        # Validate file extensions
        if not (allowed_file(template_file.filename) and allowed_file(target_file.filename)):
            flash('Only .docx files are allowed')
            return redirect(url_for('index'))
        
        # Create temporary files with context management
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_template:
            template_file.save(temp_template.name)
            template_path = temp_template.name
            temp_files.append(template_path)
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_target:
            target_file.save(temp_target.name)
            target_path = temp_target.name
            temp_files.append(target_path)
        
        # Process documents with optimized formatter
        formatter = DocumentFormatter(debug=False)  # Disable debug for production
        output_path = formatter.apply_formatting(template_path, target_path)
        temp_files.append(output_path)
        
        # Generate descriptive output filename
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        base_name = os.path.splitext(target_file.filename)[0]
        output_filename = f'{base_name}_formatted_{timestamp}.docx'
        
        # Send file with automatic cleanup
        def cleanup_files():
            for filepath in temp_files:
                try:
                    if os.path.exists(filepath):
                        os.unlink(filepath)
                except OSError:
                    pass  # Ignore cleanup errors
        
        # Register cleanup to happen after response
        @app.after_request
        def cleanup_after_response(response):
            cleanup_files()
            return response
        
        return send_file(
            output_path, 
            as_attachment=True, 
            download_name=output_filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        
    except Exception as e:
        # Clean up temp files on error
        for filepath in temp_files:
            try:
                if os.path.exists(filepath):
                    os.unlink(filepath)
            except OSError:
                pass
        
        error_msg = f'Error processing documents: {str(e)}'
        app.logger.error(error_msg)
        flash(error_msg)
        return redirect(url_for('index'))

@app.errorhandler(413)
def too_large(e):
    flash('File is too large. Maximum size is 16MB.')
    return redirect(url_for('index'))

@app.errorhandler(500)
def internal_error(e):
    flash('An internal error occurred. Please try again.')
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)