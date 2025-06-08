from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
import tempfile
import os
import re
from collections import defaultdict, Counter
from functools import lru_cache
import logging

# Configure logging for optional debug output
logging.basicConfig(level=logging.WARNING)
logger = logging.getLogger(__name__)

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
    
    def __del__(self):
        """Clean up cache when object is destroyed"""
        self._categorize_content_fast.cache_clear()