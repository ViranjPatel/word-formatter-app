from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
import tempfile
import os
import re
from collections import defaultdict

class DocumentFormatter:
    def __init__(self):
        self.template_styles = {}
        self.style_mappings = {}
        self.content_patterns = {}
    
    def extract_styles_from_template(self, template_path):
        """Extract all paragraph and character styles from the template document"""
        doc = Document(template_path)
        
        # Extract built-in and custom styles
        for style in doc.styles:
            if style.type == WD_STYLE_TYPE.PARAGRAPH:
                self.extract_paragraph_style(style)
            elif style.type == WD_STYLE_TYPE.CHARACTER:
                self.extract_character_style(style)
        
        # Analyze content patterns to understand style usage
        self.analyze_content_patterns(doc)
        
        print(f"Extracted {len(self.template_styles)} styles from template")
        for style_name in self.template_styles.keys():
            print(f"  - {style_name}")
    
    def extract_paragraph_style(self, style):
        """Extract formatting information from a paragraph style"""
        style_info = {
            'type': 'paragraph',
            'name': style.name,
            'base_style': style.base_style.name if style.base_style else None,
            'font': {},
            'paragraph': {},
            'usage_examples': []
        }
        
        # Extract font formatting
        if style.font:
            if style.font.name:
                style_info['font']['name'] = style.font.name
            if style.font.size:
                style_info['font']['size'] = style.font.size
            if style.font.bold is not None:
                style_info['font']['bold'] = style.font.bold
            if style.font.italic is not None:
                style_info['font']['italic'] = style.font.italic
            if style.font.underline is not None:
                style_info['font']['underline'] = style.font.underline
            if style.font.color and style.font.color.rgb:
                style_info['font']['color'] = style.font.color.rgb
        
        # Extract paragraph formatting
        if style.paragraph_format:
            pf = style.paragraph_format
            if pf.alignment is not None:
                style_info['paragraph']['alignment'] = pf.alignment
            if pf.space_before is not None:
                style_info['paragraph']['space_before'] = pf.space_before
            if pf.space_after is not None:
                style_info['paragraph']['space_after'] = pf.space_after
            if pf.line_spacing is not None:
                style_info['paragraph']['line_spacing'] = pf.line_spacing
            if pf.first_line_indent is not None:
                style_info['paragraph']['first_line_indent'] = pf.first_line_indent
            if pf.left_indent is not None:
                style_info['paragraph']['left_indent'] = pf.left_indent
            if pf.right_indent is not None:
                style_info['paragraph']['right_indent'] = pf.right_indent
        
        self.template_styles[style.name] = style_info
    
    def extract_character_style(self, style):
        """Extract formatting information from a character style"""
        style_info = {
            'type': 'character',
            'name': style.name,
            'base_style': style.base_style.name if style.base_style else None,
            'font': {}
        }
        
        # Extract font formatting
        if style.font:
            if style.font.name:
                style_info['font']['name'] = style.font.name
            if style.font.size:
                style_info['font']['size'] = style.font.size
            if style.font.bold is not None:
                style_info['font']['bold'] = style.font.bold
            if style.font.italic is not None:
                style_info['font']['italic'] = style.font.italic
            if style.font.underline is not None:
                style_info['font']['underline'] = style.font.underline
            if style.font.color and style.font.color.rgb:
                style_info['font']['color'] = style.font.color.rgb
        
        self.template_styles[style.name] = style_info
    
    def analyze_content_patterns(self, doc):
        """Analyze how styles are used in the template document"""
        style_usage = defaultdict(list)
        
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                style_name = paragraph.style.name
                text_preview = paragraph.text.strip()[:100]
                
                # Categorize content types
                content_type = self.categorize_content(text_preview)
                
                style_usage[style_name].append({
                    'text': text_preview,
                    'content_type': content_type,
                    'length': len(paragraph.text.strip())
                })
        
        # Store usage patterns for each style
        for style_name, usage_list in style_usage.items():
            if style_name in self.template_styles:
                self.template_styles[style_name]['usage_examples'] = usage_list
                
                # Determine primary content type for this style
                content_types = [item['content_type'] for item in usage_list]
                most_common_type = max(set(content_types), key=content_types.count)
                self.template_styles[style_name]['primary_content_type'] = most_common_type
    
    def categorize_content(self, text):
        """Categorize text content to help with style matching"""
        text_lower = text.lower().strip()
        
        # Check for headings
        if re.match(r'^\d+\.', text) or re.match(r'^[A-Z][^.!?]*$', text) and len(text) < 100:
            return 'heading'
        
        # Check for list items
        if re.match(r'^[•\\-\\*]', text) or re.match(r'^\d+[.).]', text):
            return 'list_item'
        
        # Check for titles (short, all caps or title case)
        if text.isupper() and len(text) < 50:
            return 'title'
        
        # Check for emphasis or quotes
        if text.startswith('"') or text.startswith("'"):
            return 'quote'
        
        # Default to body text
        return 'body'
    
    def apply_styles_to_target(self, target_path):
        """Apply extracted styles to the target document"""
        target_doc = Document(target_path)
        
        # First, create/update styles in target document
        self.create_styles_in_target(target_doc)
        
        # Then apply styles to paragraphs based on content analysis
        for paragraph in target_doc.paragraphs:
            if paragraph.text.strip():
                best_style = self.find_best_style_for_content(paragraph.text.strip())
                if best_style:
                    try:
                        paragraph.style = target_doc.styles[best_style]
                        print(f"Applied style '{best_style}' to: {paragraph.text[:50]}...")
                    except KeyError:
                        print(f"Style '{best_style}' not found in target document")
        
        # Apply styles to table content
        for table in target_doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if paragraph.text.strip():
                            best_style = self.find_best_style_for_content(paragraph.text.strip())
                            if best_style:
                                try:
                                    paragraph.style = target_doc.styles[best_style]
                                except KeyError:
                                    pass
        
        return target_doc
    
    def create_styles_in_target(self, target_doc):
        """Create or update styles in the target document based on template"""
        for style_name, style_info in self.template_styles.items():
            try:
                # Try to get existing style
                existing_style = target_doc.styles[style_name]
                self.update_existing_style(existing_style, style_info)
                print(f"Updated existing style: {style_name}")
            except KeyError:
                # Create new style if it doesn't exist
                if style_info['type'] == 'paragraph':
                    self.create_paragraph_style(target_doc, style_name, style_info)
                elif style_info['type'] == 'character':
                    self.create_character_style(target_doc, style_name, style_info)
                print(f"Created new style: {style_name}")
    
    def update_existing_style(self, style, style_info):
        """Update an existing style with template formatting"""
        try:
            # Update font formatting
            if 'font' in style_info and hasattr(style, 'font') and style.font:
                font_info = style_info['font']
                if 'name' in font_info:
                    style.font.name = font_info['name']
                if 'size' in font_info:
                    style.font.size = font_info['size']
                if 'bold' in font_info:
                    style.font.bold = font_info['bold']
                if 'italic' in font_info:
                    style.font.italic = font_info['italic']
                if 'underline' in font_info:
                    style.font.underline = font_info['underline']
                if 'color' in font_info:
                    style.font.color.rgb = font_info['color']
            
            # Update paragraph formatting if it's a paragraph style
            if style_info['type'] == 'paragraph' and hasattr(style, 'paragraph_format'):
                if 'paragraph' in style_info:
                    para_info = style_info['paragraph']
                    pf = style.paragraph_format
                    if 'alignment' in para_info:
                        pf.alignment = para_info['alignment']
                    if 'space_before' in para_info:
                        pf.space_before = para_info['space_before']
                    if 'space_after' in para_info:
                        pf.space_after = para_info['space_after']
                    if 'line_spacing' in para_info:
                        pf.line_spacing = para_info['line_spacing']
                    if 'first_line_indent' in para_info:
                        pf.first_line_indent = para_info['first_line_indent']
                    if 'left_indent' in para_info:
                        pf.left_indent = para_info['left_indent']
                    if 'right_indent' in para_info:
                        pf.right_indent = para_info['right_indent']
        except Exception as e:
            print(f"Error updating style: {e}")
    
    def create_paragraph_style(self, doc, style_name, style_info):
        """Create a new paragraph style in the document"""
        try:
            # Skip creating built-in styles that might conflict
            if style_name in ['Normal', 'Heading 1', 'Heading 2', 'Heading 3', 'Heading 4', 'Heading 5', 'Heading 6', 'Title']:
                return
            
            styles = doc.styles
            style = styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
            
            # Set base style if specified
            if style_info.get('base_style') and style_info['base_style'] != style_name:
                try:
                    style.base_style = styles[style_info['base_style']]
                except KeyError:
                    pass
            
            self.update_existing_style(style, style_info)
        except Exception as e:
            print(f"Error creating paragraph style '{style_name}': {e}")
    
    def create_character_style(self, doc, style_name, style_info):
        """Create a new character style in the document"""
        try:
            styles = doc.styles
            style = styles.add_style(style_name, WD_STYLE_TYPE.CHARACTER)
            
            # Set base style if specified
            if style_info.get('base_style') and style_info['base_style'] != style_name:
                try:
                    style.base_style = styles[style_info['base_style']]
                except KeyError:
                    pass
            
            self.update_existing_style(style, style_info)
        except Exception as e:
            print(f"Error creating character style '{style_name}': {e}")
    
    def find_best_style_for_content(self, text):
        """Find the best style for given content based on content analysis"""
        content_type = self.categorize_content(text)
        text_lower = text.lower().strip()
        
        # Look for styles that match the content type
        matching_styles = []
        for style_name, style_info in self.template_styles.items():
            if style_info.get('primary_content_type') == content_type:
                matching_styles.append((style_name, style_info))
        
        # If we found matching styles, return the first one
        if matching_styles:
            return matching_styles[0][0]
        
        # Fallback: try to match based on style name patterns
        if content_type == 'heading':
            # Look for heading styles
            for style_name in self.template_styles.keys():
                if 'heading' in style_name.lower() or 'title' in style_name.lower():
                    return style_name
        
        elif content_type == 'list_item':
            # Look for list styles
            for style_name in self.template_styles.keys():
                if 'list' in style_name.lower() or 'bullet' in style_name.lower():
                    return style_name
        
        # Final fallback: use Normal style or the first available style
        if 'Normal' in self.template_styles:
            return 'Normal'
        elif self.template_styles:
            return list(self.template_styles.keys())[0]
        
        return None
    
    def apply_formatting(self, template_path, target_path):
        """Main method to apply formatting from template to target document"""
        print(f"Extracting styles from template: {template_path}")
        self.extract_styles_from_template(template_path)
        
        print(f"Applying styles to target: {target_path}")
        target_doc = self.apply_styles_to_target(target_path)
        
        # Save the formatted document
        output_path = tempfile.mktemp(suffix='.docx')
        target_doc.save(output_path)
        print(f"Formatted document saved to: {output_path}")
        
        return output_path