from docx import Document
from docx.shared import RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
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
                    'bold': None,
                    'italic': None,
                    'underline': None,
                    'color': None,
                    'alignment': None
                }

                if paragraph.runs:
                    # Prefer a run that has explicit formatting
                    best_run = None
                    for r in paragraph.runs:
                        if r.font.name or r.font.size or r.bold is not None or r.italic is not None \
                                or r.underline is not None or (r.font.color and r.font.color.rgb):
                            best_run = r
                            break
                    if not best_run:
                        best_run = paragraph.runs[0]

                    if best_run.font.name:
                        style_info['font_name'] = best_run.font.name
                    if best_run.font.size:
                        style_info['font_size'] = best_run.font.size
                    if best_run.bold is not None:
                        style_info['bold'] = best_run.bold
                    if best_run.italic is not None:
                        style_info['italic'] = best_run.italic
                    if best_run.underline is not None:
                        style_info['underline'] = best_run.underline
                    if best_run.font.color.rgb:
                        style_info['color'] = best_run.font.color.rgb
                
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
            if style_info['bold'] is not None:
                run.bold = style_info['bold']
            if style_info['italic'] is not None:
                run.italic = style_info['italic']
            if style_info['underline'] is not None:
                run.underline = style_info['underline']
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
