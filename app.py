from flask import Flask, render_template, request, send_file, flash, redirect, url_for, jsonify
import os
from werkzeug.utils import secure_filename
from document_formatter import DocumentFormatter
import tempfile
from datetime import datetime
import logging
import atexit
import threading
import time

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

# Global cleanup registry for temporary files
cleanup_registry = []
cleanup_lock = threading.Lock()

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def register_for_cleanup(filepath, delay=30):
    """Register a file for cleanup after a delay"""
    with cleanup_lock:
        cleanup_registry.append({
            'path': filepath,
            'cleanup_time': time.time() + delay
        })

def cleanup_expired_files():
    """Clean up files that have expired"""
    current_time = time.time()
    with cleanup_lock:
        expired_files = [item for item in cleanup_registry if item['cleanup_time'] <= current_time]
        cleanup_registry[:] = [item for item in cleanup_registry if item['cleanup_time'] > current_time]
    
    for item in expired_files:
        try:
            if os.path.exists(item['path']):
                os.unlink(item['path'])
        except OSError:
            pass

def cleanup_temp_files_immediate(temp_files):
    """Immediate cleanup of temporary files"""
    for filepath in temp_files:
        try:
            if os.path.exists(filepath):
                os.unlink(filepath)
        except OSError:
            pass

# Background cleanup thread
def background_cleanup():
    """Background thread to clean up expired files"""
    while True:
        cleanup_expired_files()
        time.sleep(10)  # Check every 10 seconds

# Start background cleanup thread
cleanup_thread = threading.Thread(target=background_cleanup, daemon=True)
cleanup_thread.start()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    temp_files = []  # Track temporary files for cleanup
    
    try:
        operation = request.form.get('operation', 'format')
        
        if operation == 'format':
            return process_formatting()
        elif operation == 'latex':
            return process_latex_conversion()
        else:
            flash('Invalid operation selected')
            return redirect(url_for('index'))
            
    except Exception as e:
        error_msg = f'Error processing documents: {str(e)}'
        app.logger.error(error_msg)
        flash(error_msg)
        return redirect(url_for('index'))

def process_formatting():
    """Handle Word document formatting"""
    input_temp_files = []
    
    try:
        # Validate file uploads
        if 'template_file' not in request.files or 'target_file' not in request.files:
            flash('Please select both template and target files for formatting')
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
        
        # Create temporary files
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_template:
            template_file.save(temp_template.name)
            template_path = temp_template.name
            input_temp_files.append(template_path)
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_target:
            target_file.save(temp_target.name)
            target_path = temp_target.name
            input_temp_files.append(target_path)
        
        # Process documents with optimized formatter
        formatter = DocumentFormatter(debug=False)
        output_path = formatter.apply_formatting(template_path, target_path)
        
        # Clean up input files immediately
        cleanup_temp_files_immediate(input_temp_files)
        
        # Generate descriptive output filename
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        base_name = os.path.splitext(target_file.filename)[0]
        output_filename = f'{base_name}_formatted_{timestamp}.docx'
        
        # Register output file for delayed cleanup
        register_for_cleanup(output_path, delay=60)  # Clean up after 1 minute
        
        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        
    except Exception as e:
        # Clean up files on error
        cleanup_temp_files_immediate(input_temp_files)
        raise e

def process_latex_conversion():
    """Handle LaTeX conversion"""
    input_temp_files = []
    
    try:
        # Validate file upload
        if 'latex_file' not in request.files:
            flash('Please select a Word document to convert to LaTeX')
            return redirect(url_for('index'))
        
        latex_file = request.files['latex_file']
        
        # Validate file selection
        if not latex_file.filename:
            flash('Please select a file to convert')
            return redirect(url_for('index'))
        
        # Validate file extension
        if not allowed_file(latex_file.filename):
            flash('Only .docx files are allowed')
            return redirect(url_for('index'))
        
        # Create temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_latex:
            latex_file.save(temp_latex.name)
            latex_path = temp_latex.name
            input_temp_files.append(latex_path)
        
        # Convert to LaTeX
        formatter = DocumentFormatter(debug=False)
        output_path = formatter.convert_to_latex(latex_path)
        
        # Clean up input files immediately
        cleanup_temp_files_immediate(input_temp_files)
        
        # Generate output filename
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        base_name = os.path.splitext(latex_file.filename)[0]
        output_filename = f'{base_name}_converted_{timestamp}.tex'
        
        # Register output file for delayed cleanup
        register_for_cleanup(output_path, delay=60)  # Clean up after 1 minute
        
        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_filename,
            mimetype='text/plain'
        )
        
    except Exception as e:
        # Clean up files on error
        cleanup_temp_files_immediate(input_temp_files)
        raise e

@app.errorhandler(413)
def too_large(e):
    flash('File is too large. Maximum size is 16MB.')
    return redirect(url_for('index'))

@app.errorhandler(500)
def internal_error(e):
    flash('An internal error occurred. Please try again.')
    return redirect(url_for('index'))

# Clean up any remaining files on app shutdown
@atexit.register
def cleanup_on_exit():
    """Clean up all registered files on application exit"""
    with cleanup_lock:
        for item in cleanup_registry:
            try:
                if os.path.exists(item['path']):
                    os.unlink(item['path'])
            except OSError:
                pass
        cleanup_registry.clear()

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)