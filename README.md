# Word Document Formatter

A simple web application that extracts formatting rules from one Word document and applies them to another document using Python's `python-docx` library.

## Features

- 📄 Upload two Word documents (.docx format)
- ✨ Extract formatting from template document
- 🎨 Apply formatting to target document
- 💾 Download the formatted result
- 🎯 Simple, clean web interface
- 📱 Mobile-responsive design

## Quick Start

### Local Development

1. **Clone the repository:**
   ```bash
   git clone https://github.com/ViranjPatel/word-formatter-app.git
   cd word-formatter-app
   ```

2. **Create a virtual environment:**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the application:**
   ```bash
   python app.py
   ```

5. **Open your browser** and go to `http://localhost:5000`

### Using the Application

1. **Upload Template Document**: Choose a Word document that has the formatting you want to copy
2. **Upload Target Document**: Choose the Word document you want to apply formatting to
3. **Click "Apply Formatting"**: The app will process both documents
4. **Download Result**: Your newly formatted document will be downloaded automatically

## Supported Formatting

The application currently extracts and applies:

- ✅ Font family and size
- ✅ Bold, italic, and underline formatting
- ✅ Font colors
- ✅ Paragraph alignment
- ✅ Basic text formatting

**Note**: Complex layouts, images, headers/footers, and advanced formatting may not be preserved.

## File Requirements

- Only `.docx` files are supported
- Maximum file size: 16MB
- Both template and target documents are required

## Technology Stack

- **Backend**: Python Flask
- **Document Processing**: python-docx library
- **Frontend**: HTML, CSS, JavaScript
- **Styling**: Custom CSS with modern design

## Project Structure

```
word-formatter-app/
├── app.py                 # Main Flask application
├── document_formatter.py  # Core formatting logic
├── requirements.txt       # Python dependencies
├── templates/
│   └── index.html        # Web interface
├── static/
│   └── style.css         # Styling
└── README.md             # This file
```

## Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature-name`
3. Make your changes and commit: `git commit -am 'Add feature'`
4. Push to the branch: `git push origin feature-name`
5. Create a Pull Request

## Deployment

### Heroku

1. Create a `Procfile`:
   ```
   web: gunicorn app:app
   ```

2. Deploy to Heroku:
   ```bash
   heroku create your-app-name
   git push heroku main
   ```

### Other Platforms

This Flask application can be deployed on any platform that supports Python web applications, such as:
- Railway
- Render
- PythonAnywhere
- DigitalOcean App Platform

## License

MIT License - feel free to use this project for personal or commercial purposes.

## Support

If you encounter any issues or have questions:

1. Check the existing [Issues](https://github.com/ViranjPatel/word-formatter-app/issues)
2. Create a new issue with details about your problem
3. Include sample documents if possible (remove sensitive content)

---

**Built with ❤️ using Python, Flask, and python-docx**