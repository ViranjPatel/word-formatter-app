from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import nsdecls, qn
import tempfile
import os
import re
from collections import defaultdict, Counter
from functools import lru_cache
import logging

# Configure logging for optional debug output
logging.basicConfig(level=logging.WARNING)
logger = logging.getLogger(__name__)

class LaTeXConverter:
    """Convert Word documents to LaTeX format with proper formatting"""
    
    def __init__(self, debug=False):
        self.debug = debug
        self.latex_content = []
        self.packages = set()
        self.document_class = "article"
        
        # LaTeX command mappings
        self.style_mappings = {
            'Heading 1': r'\section{{{content}}}',
            'Heading 2': r'\subsection{{{content}}}',
            'Heading 3': r'\subsubsection{{{content}}}',
            'Heading 4': r'\paragraph{{{content}}}',
            'Heading 5': r'\subparagraph{{{content}}}',
            'Heading 6': r'\subparagraph{{{content}}}',
            'Title': r'\title{{{content}}}',
            'Subtitle': r'\subtitle{{{content}}}',
        }
        
        if debug:
            logger.setLevel(logging.INFO)
    
    def convert_document(self, docx_path):
        """Convert Word document to LaTeX format"""
        doc = Document(docx_path)
        
        # Initialize LaTeX document
        self._initialize_latex_document()
        
        # Process document content
        self._process_paragraphs(doc.paragraphs)
        
        # Process tables
        self._process_tables(doc.tables)
        
        # Finalize document
        self._finalize_latex_document()
        
        # Generate output file
        output_path = tempfile.mktemp(suffix='.tex')
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(self.latex_content))
        
        if self.debug:
            logger.info(f"LaTeX document created: {output_path}")
        
        return output_path
    
    def _initialize_latex_document(self):
        """Initialize LaTeX document structure"""
        self.latex_content = [
            f'\\documentclass{{{self.document_class}}}',
            '',
            '% Packages',
            '\\usepackage[utf8]{inputenc}',
            '\\usepackage[T1]{fontenc}',
            '\\usepackage{geometry}',
            '\\usepackage{amsmath}',
            '\\usepackage{amsfonts}',
            '\\usepackage{amssymb}',
            '\\usepackage{graphicx}',
            '\\usepackage{hyperref}',
            '\\usepackage{booktabs}',
            '\\usepackage{array}',
            '\\usepackage{longtable}',
            '\\usepackage{enumerate}',
            '\\usepackage{enumitem}',
            '',
            '% Document settings',
            '\\geometry{margin=1in}',
            '\\setlength{\\parindent}{0pt}',
            '\\setlength{\\parskip}{6pt}',
            '',
            '\\begin{document}',
            ''
        ]
    
    def _process_paragraphs(self, paragraphs):
        """Process all paragraphs and convert to LaTeX"""
        for paragraph in paragraphs:
            if paragraph.text.strip():
                latex_para = self._convert_paragraph(paragraph)
                if latex_para:
                    self.latex_content.append(latex_para)
                    self.latex_content.append('')
    
    def _convert_paragraph(self, paragraph):
        """Convert a single paragraph to LaTeX"""
        text = paragraph.text.strip()
        if not text:
            return ''
        
        style_name = paragraph.style.name
        
        # Handle specific styles
        if style_name in self.style_mappings:
            content = self._escape_latex_chars(text)
            return self.style_mappings[style_name].format(content=content)
        
        # Handle lists
        if self._is_list_item(text):
            return self._convert_list_item(text)
        
        # Regular paragraph
        content = self._process_runs(paragraph.runs)
        return content if content.strip() else ''
    
    def _process_runs(self, runs):
        """Process runs within a paragraph for formatting"""
        content_parts = []
        
        for run in runs:
            text = run.text
            if not text:
                continue
            
            # Escape LaTeX special characters
            escaped_text = self._escape_latex_chars(text)
            
            # Apply formatting
            formatted_text = self._apply_run_formatting(escaped_text, run)
            content_parts.append(formatted_text)
        
        return ''.join(content_parts)
    
    def _apply_run_formatting(self, text, run):
        """Apply formatting to text based on run properties"""
        if not text.strip():
            return text
        
        formatted = text
        
        # Bold
        if run.bold:
            formatted = f'\\textbf{{{formatted}}}'
        
        # Italic
        if run.italic:
            formatted = f'\\textit{{{formatted}}}'
        
        # Underline
        if run.underline:
            formatted = f'\\underline{{{formatted}}}'
        
        # Font size changes (approximate)
        if run.font.size:
            size_pt = run.font.size.pt
            if size_pt > 14:
                formatted = f'\\Large {{{formatted}}}'
            elif size_pt > 12:
                formatted = f'\\large {{{formatted}}}'
            elif size_pt < 10:
                formatted = f'\\small {{{formatted}}}'
        
        return formatted
    
    def _is_list_item(self, text):
        """Check if text appears to be a list item"""
        return bool(re.match(r'^[•\-\*]\s+|^\d+[\.)]\s+', text))
    
    def _convert_list_item(self, text):
        """Convert list item to LaTeX format"""
        # Remove list markers
        cleaned_text = re.sub(r'^[•\-\*]\s+|^\d+[\.)]\s+', '', text)
        escaped_text = self._escape_latex_chars(cleaned_text)
        
        # Determine list type
        if re.match(r'^\d+[\.)]\s+', text):
            # Numbered list
            return f'\\begin{{enumerate}}\\item {escaped_text}\\end{{enumerate}}'
        else:
            # Bulleted list
            return f'\\begin{{itemize}}\\item {escaped_text}\\end{{itemize}}'
    
    def _process_tables(self, tables):
        """Process tables and convert to LaTeX"""
        for table in tables:
            latex_table = self._convert_table(table)
            if latex_table:
                self.latex_content.extend(latex_table)
                self.latex_content.append('')
    
    def _convert_table(self, table):
        """Convert a Word table to LaTeX format"""
        if not table.rows:
            return []
        
        # Determine number of columns
        max_cols = max(len(row.cells) for row in table.rows)
        
        # Start table
        latex_lines = [
            '\\begin{table}[h]',
            '\\centering',
            f'\\begin{{tabular}}{{{"l" * max_cols}}}',
            '\\toprule'
        ]
        
        # Process rows
        for i, row in enumerate(table.rows):
            row_content = []
            for cell in row.cells[:max_cols]:
                cell_text = self._escape_latex_chars(cell.text.strip())
                row_content.append(cell_text)
            
            # Pad row if necessary
            while len(row_content) < max_cols:
                row_content.append('')
            
            latex_lines.append(' & '.join(row_content) + ' \\\\')
            
            # Add midrule after first row (header)
            if i == 0:
                latex_lines.append('\\midrule')
        
        # End table
        latex_lines.extend([
            '\\bottomrule',
            '\\end{tabular}',
            '\\end{table}'
        ])
        
        return latex_lines
    
    def _escape_latex_chars(self, text):
        """Escape special LaTeX characters"""
        # Dictionary of characters to escape
        escape_chars = {
            '\\': r'\textbackslash{}',
            '{': r'\{',
            '}': r'\}',
            '$': r'\$',
            '&': r'\&',
            '%': r'\%',
            '#': r'\#',
            '^': r'\textasciicircum{}',
            '_': r'\_',
            '~': r'\textasciitilde{}',
        }
        
        for char, escaped in escape_chars.items():
            text = text.replace(char, escaped)
        
        return text
    
    def _finalize_latex_document(self):
        """Finalize LaTeX document"""
        self.latex_content.extend([
            '',
            '\\end{document}'
        ])

class DocumentFormatter:
    # Pre-compiled regex patterns for better performance
    HEADING_PATTERN = re.compile(r'^\d+\.|^[A-Z][^.!?]*$')
    LIST_PATTERN = re.compile(r'^[•\-\*]|^\d+[.).]')
    
    # Content type constants
    CONTENT_TYPES = {
        'heading': 'heading',
        'list_item': 'list_item', 
        'title': 'title',
        'quote': 'quote',
        'body': 'body'
    }
    
    # Built-in styles that shouldn't be recreated
    BUILTIN_STYLES = frozenset([
        'Normal', 'Heading 1', 'Heading 2', 'Heading 3', 'Heading 4', 
        'Heading 5', 'Heading 6', 'Title', 'Subtitle', 'Header', 'Footer'
    ])

    def __init__(self, debug=False):
        self.template_styles = {}
        self.style_content_map = {}
        self.content_style_cache = {}
        self.debug = debug
        if debug:
            logger.setLevel(logging.INFO)
    
    def extract_styles_from_template(self, template_path):
        """Extract styles and analyze content patterns efficiently"""
        doc = Document(template_path)
        
        # Single pass through styles for extraction
        paragraph_styles = {}
        character_styles = {}
        
        for style in doc.styles:
            style_data = self._extract_style_data(style)
            if style_data:
                if style.type == WD_STYLE_TYPE.PARAGRAPH:
                    paragraph_styles[style.name] = style_data
                elif style.type == WD_STYLE_TYPE.CHARACTER:
                    character_styles[style.name] = style_data
        
        self.template_styles = {**paragraph_styles, **character_styles}
        
        # Single pass content analysis with batch processing
        self._analyze_content_usage_batch(doc)
        
        if self.debug:
            logger.info(f"Extracted {len(self.template_styles)} styles")
    
    def _extract_style_data(self, style):
        """Efficiently extract style data with minimal attribute access"""
        style_data = {
            'type': 'paragraph' if style.type == WD_STYLE_TYPE.PARAGRAPH else 'character',
            'name': style.name
        }
        
        # Extract font data efficiently
        font_data = {}
        if hasattr(style, 'font') and style.font:
            font = style.font
            font_attrs = ['name', 'size', 'bold', 'italic', 'underline']
            for attr in font_attrs:
                value = getattr(font, attr, None)
                if value is not None:
                    font_data[attr] = value
            
            # Handle color separately due to complex access pattern
            if hasattr(font, 'color') and font.color and hasattr(font.color, 'rgb') and font.color.rgb:
                font_data['color'] = font.color.rgb
        
        if font_data:
            style_data['font'] = font_data
        
        # Extract paragraph data efficiently for paragraph styles
        if style.type == WD_STYLE_TYPE.PARAGRAPH and hasattr(style, 'paragraph_format') and style.paragraph_format:
            para_data = {}
            pf = style.paragraph_format
            para_attrs = ['alignment', 'space_before', 'space_after', 'line_spacing', 
                         'first_line_indent', 'left_indent', 'right_indent']
            
            for attr in para_attrs:
                value = getattr(pf, attr, None)
                if value is not None:
                    para_data[attr] = value
            
            if para_data:
                style_data['paragraph'] = para_data
        
        return style_data if len(style_data) > 2 else None  # Only return if has meaningful data
    
    def _analyze_content_usage_batch(self, doc):
        """Batch analyze content patterns for optimal performance"""
        style_usage = defaultdict(list)
        
        # Collect all text content in single pass
        all_paragraphs = []
        all_paragraphs.extend(doc.paragraphs)
        
        # Add table paragraphs
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    all_paragraphs.extend(cell.paragraphs)
        
        # Process all paragraphs with batch categorization
        content_types = []
        texts = []
        style_names = []
        
        for paragraph in all_paragraphs:
            text = paragraph.text.strip()
            if text and len(text) > 2:  # Skip very short text
                texts.append(text[:100])  # Limit text length for performance
                style_names.append(paragraph.style.name)
                content_types.append(self._categorize_content_fast(text))
        
        # Build style mappings efficiently
        for style_name, content_type in zip(style_names, content_types):
            if style_name in self.template_styles:
                if style_name not in self.style_content_map:
                    self.style_content_map[style_name] = []
                self.style_content_map[style_name].append(content_type)
        
        # Determine primary content type for each style
        for style_name, content_list in self.style_content_map.items():
            if content_list:
                primary_type = Counter(content_list).most_common(1)[0][0]
                self.template_styles[style_name]['primary_content_type'] = primary_type
    
    @lru_cache(maxsize=1000)
    def _categorize_content_fast(self, text):
        """Fast content categorization with caching and optimized patterns"""
        text_stripped = text.strip()
        text_len = len(text_stripped)
        
        # Quick length-based filtering
        if text_len == 0:
            return self.CONTENT_TYPES['body']
        
        # Check for titles (short uppercase)
        if text_stripped.isupper() and text_len < 50:
            return self.CONTENT_TYPES['title']
        
        # Check for quotes (starts with quote marks)
        first_char = text_stripped[0]
        if first_char in '\"\'':
            return self.CONTENT_TYPES['quote']
        
        # Use pre-compiled patterns for better performance
        if text_len < 100 and self.HEADING_PATTERN.match(text_stripped):
            return self.CONTENT_TYPES['heading']
        
        if self.LIST_PATTERN.match(text_stripped):
            return self.CONTENT_TYPES['list_item']
        
        return self.CONTENT_TYPES['body']
    
    def apply_styles_to_target(self, target_path):
        """Apply styles with optimized batch processing"""
        target_doc = Document(target_path)
        
        # Pre-create style mapping for fast lookup
        self._prepare_style_mapping()
        
        # Batch update existing styles
        self._batch_update_styles(target_doc)
        
        # Batch apply styles to content
        self._batch_apply_paragraph_styles(target_doc)
        
        return target_doc
    
    def _prepare_style_mapping(self):
        """Pre-compute style mappings for O(1) lookup"""
        content_type_styles = defaultdict(list)
        
        for style_name, style_data in self.template_styles.items():
            content_type = style_data.get('primary_content_type', 'body')
            content_type_styles[content_type].append(style_name)
        
        # Cache the best style for each content type
        for content_type, style_list in content_type_styles.items():
            if style_list:
                self.content_style_cache[content_type] = style_list[0]
        
        # Add fallback mappings
        style_name_lower_map = {name.lower(): name for name in self.template_styles.keys()}
        
        for content_type in self.CONTENT_TYPES.values():
            if content_type not in self.content_style_cache:
                # Try to find by name pattern
                for pattern in [content_type, 'heading', 'title', 'normal']:
                    if pattern in style_name_lower_map:
                        self.content_style_cache[content_type] = style_name_lower_map[pattern]
                        break
        
        # Ultimate fallback
        if self.template_styles:
            fallback_style = next(iter(self.template_styles.keys()))
            for content_type in self.CONTENT_TYPES.values():
                if content_type not in self.content_style_cache:
                    self.content_style_cache[content_type] = fallback_style
    
    def _batch_update_styles(self, target_doc):
        """Efficiently update or create styles in batch"""
        existing_styles = {style.name for style in target_doc.styles}
        
        for style_name, style_data in self.template_styles.items():
            try:
                if style_name in existing_styles:
                    self._update_style_fast(target_doc.styles[style_name], style_data)
                elif style_name not in self.BUILTIN_STYLES:
                    self._create_style_fast(target_doc, style_name, style_data)
            except Exception as e:
                if self.debug:
                    logger.warning(f"Style operation failed for {style_name}: {e}")
    
    def _update_style_fast(self, style, style_data):
        """Fast style update with minimal attribute access"""
        # Update font properties
        if 'font' in style_data and hasattr(style, 'font') and style.font:
            font_data = style_data['font']
            font = style.font
            
            # Batch update font attributes
            for attr, value in font_data.items():
                if attr == 'color':
                    if hasattr(font, 'color') and font.color:
                        font.color.rgb = value
                else:
                    setattr(font, attr, value)
        
        # Update paragraph properties
        if 'paragraph' in style_data and hasattr(style, 'paragraph_format') and style.paragraph_format:
            para_data = style_data['paragraph']
            pf = style.paragraph_format
            
            # Batch update paragraph attributes
            for attr, value in para_data.items():
                setattr(pf, attr, value)
    
    def _create_style_fast(self, doc, style_name, style_data):
        """Fast style creation with error handling"""
        try:
            style_type = WD_STYLE_TYPE.PARAGRAPH if style_data['type'] == 'paragraph' else WD_STYLE_TYPE.CHARACTER
            new_style = doc.styles.add_style(style_name, style_type)
            self._update_style_fast(new_style, style_data)
        except Exception as e:
            if self.debug:
                logger.warning(f"Failed to create style {style_name}: {e}")
    
    def _batch_apply_paragraph_styles(self, target_doc):
        """Apply styles to all paragraphs in optimized batch"""
        # Collect all paragraphs
        all_paragraphs = list(target_doc.paragraphs)
        
        # Add table paragraphs
        for table in target_doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    all_paragraphs.extend(cell.paragraphs)
        
        # Batch process paragraphs
        for paragraph in all_paragraphs:
            text = paragraph.text.strip()
            if text:
                content_type = self._categorize_content_fast(text)
                best_style = self.content_style_cache.get(content_type)
                
                if best_style:
                    try:
                        paragraph.style = target_doc.styles[best_style]
                        if self.debug:
                            logger.info(f"Applied '{best_style}' to: {text[:30]}...")
                    except KeyError:
                        if self.debug:
                            logger.warning(f"Style '{best_style}' not found")
    
    def apply_formatting(self, template_path, target_path):
        """Optimized main formatting method"""
        if self.debug:
            logger.info(f"Processing: {template_path} -> {target_path}")
        
        # Extract styles and patterns
        self.extract_styles_from_template(template_path)
        
        # Apply to target document
        target_doc = self.apply_styles_to_target(target_path)
        
        # Save efficiently
        output_path = tempfile.mktemp(suffix='.docx')
        target_doc.save(output_path)
        
        if self.debug:
            logger.info(f"Completed: {output_path}")
        
        return output_path
    
    def convert_to_latex(self, docx_path):
        """Convert Word document to LaTeX format"""
        converter = LaTeXConverter(debug=self.debug)
        return converter.convert_document(docx_path)
    
    def __del__(self):
        """Clean up cache when object is destroyed"""
        self._categorize_content_fast.cache_clear()