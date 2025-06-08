from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
import tempfile
import os

class DocumentFormatter:
    def __init__(self):
        self.template_styles = {}
    
    def extract_formatting(self, template_path):
        """Extract formatting rules from the template document"""
        doc = Document(template_path)
        
        # Extract paragraph styles
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():  # Only process non-empty paragraphs
                style_info = {
                    'font_name': None,
                    'font_size': None,
                    'bold': False,
                    'italic': False,
                    'underline': False,
                    'color': None,
                    'alignment': None
                }
                
                # Get run formatting (font properties)
                if paragraph.runs:
                    first_run = paragraph.runs[0]
                    if first_run.font.name:
                        style_info['font_name'] = first_run.font.name
                    if first_run.font.size:
                        style_info['font_size'] = first_run.font.size
                    if first_run.font.bold:
                        style_info['bold'] = first_run.font.bold
                    if first_run.font.italic:
                        style_info['italic'] = first_run.font.italic
                    if first_run.font.underline:
                        style_info['underline'] = first_run.font.underline
                    if first_run.font.color.rgb:
                        style_info['color'] = first_run.font.color.rgb
                
                # Get paragraph alignment
                if paragraph.alignment:
                    style_info['alignment'] = paragraph.alignment
                
                # Store the first occurrence of each style
                text_preview = paragraph.text[:50]  # First 50 chars as identifier
                if text_preview not in self.template_styles:
                    self.template_styles[text_preview] = style_info
        
        # Extract table formatting if any
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if paragraph.text.strip():
                            # Similar extraction for table cells
                            pass
    
    def apply_formatting_to_paragraph(self, paragraph, style_info):
        """Apply formatting to a paragraph based on style info"""
        # Apply to all runs in the paragraph
        for run in paragraph.runs:
            if style_info['font_name']:
                run.font.name = style_info['font_name']
            if style_info['font_size']:
                run.font.size = style_info['font_size']
            if style_info['bold']:
                run.font.bold = style_info['bold']
            if style_info['italic']:
                run.font.italic = style_info['italic']
            if style_info['underline']:
                run.font.underline = style_info['underline']
            if style_info['color']:
                run.font.color.rgb = style_info['color']
        
        # Apply paragraph-level formatting
        if style_info['alignment']:
            paragraph.alignment = style_info['alignment']
    
    def apply_formatting(self, template_path, target_path):
        """Apply formatting from template to target document"""
        # Extract formatting from template
        self.extract_formatting(template_path)
        
        # Open target document
        target_doc = Document(target_path)
        
        # Apply formatting to target document
        for paragraph in target_doc.paragraphs:
            if paragraph.text.strip():
                # Try to find matching style based on content similarity
                best_match = self.find_best_style_match(paragraph.text)
                if best_match:
                    self.apply_formatting_to_paragraph(paragraph, best_match)
        
        # Apply formatting to tables if any
        for table in target_doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if paragraph.text.strip():
                            best_match = self.find_best_style_match(paragraph.text)
                            if best_match:
                                self.apply_formatting_to_paragraph(paragraph, best_match)
        
        # Save the formatted document
        output_path = tempfile.mktemp(suffix='.docx')
        target_doc.save(output_path)
        return output_path
    
    def find_best_style_match(self, text):
        """Find the best matching style for given text"""
        if not self.template_styles:
            return None
        
        # For now, use the first available style
        # In a more sophisticated version, you could implement
        # text similarity matching or pattern recognition
        return list(self.template_styles.values())[0]