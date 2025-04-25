import re
from typing import List, Dict, Optional

'''
Improved outline parser with:
- Fixed beautification check to avoid unnecessary changes
- Better handling of indentation and page numbers
- Proper handling of leading numbers
- Added empty lines between sections for readability
'''

def is_beautified(outline: str) -> bool:
    """Check if the outline is already beautified"""
    lines = outline.splitlines()
    if not lines:
        return False
    
    last_level = -1
    
    # Check for consistent indentation and spacing around page numbers
    for line in lines:
        # Skip empty lines
        if not line.strip():
            continue
            
        indentation = len(line) - len(line.lstrip())
        content = line.strip()
        
        # Check proper indentation (multiple of 4 spaces)
        if indentation % 4 != 0:
            return False
            
        # Check for exactly one space before page numbers
        if re.search(r'\d+$', content):
            # Check if there's exactly one space before the number
            # Pattern: any character that's not a space, followed by exactly one space, followed by digits to end
            if not re.match(r'^.*[^\s]\s[0-9]+$', content):
                return False
                
            # Check that there aren't multiple spaces before the number
            if re.search(r'\s{2,}\d+$', content):
                return False
                
    return True

def beautify_outline(outline: str) -> str:
    """Formats the outline with consistent indentation and spacing"""
    # Check if already beautified to avoid double-beautification
    if is_beautified(outline):
        return outline
        
    lines = outline.splitlines()
    beautified = []
    
    # Track the current section level to add empty lines between major sections
    last_indent_level = -1
    
    for line in lines:
        # Skip empty lines
        if not line.strip():
            continue
            
        # Get indentation level
        indent = len(line) - len(line.lstrip())
        content = line.strip()
        indent_level = indent // 4  # Calculate level based on 4-space indentation
        
        # Add empty line before new major section (level 0)
        if indent_level == 0 and last_indent_level != -1:
            beautified.append("")
        
        # Split into title and page number parts
        # This completely strips all existing spacing to ensure consistent results
        parts = re.split(r'\s+(\d+)$', content.strip(), 1)
        
        if len(parts) == 3:  # Successfully split into [title, page_num, '']
            title = parts[0].rstrip()
            page_num = parts[1]
            # Reconstruct with exactly one space between title and page number
            content = f"{title} {page_num}"
        
        # Add the line with correct indentation
        beautified.append((' ' * indent_level * 4) + content)
        last_indent_level = indent_level
    
    return '\n'.join(beautified)

def normalize_outline(raw_outline: str) -> str:
    """
    Step 1: Convert messy outline into clean format with:
    - Each outline item on a single line
    - Proper indentation (4 spaces per level)
    - Page numbers at end after space/parentheses
    - No line breaks within a single item
    - Preserving leading numbers in titles
    """
    lines = [line.rstrip() for line in raw_outline.expandtabs(4).splitlines() if line.strip()]
    clean_lines = []
    current_item = None
    current_indent = 0
    last_line_was_terminated = True
    
    for line in lines:
        # Determine base indentation
        indent = len(line) - len(line.lstrip())
        content = line.strip()
        
        # Check if this is a continuation line
        has_page_number = bool(re.search(r'([\.\- ]{3,}|\(|\s)(\d+|[IVXLCDM]+)\s*$', content))
        starts_with_number = bool(re.match(r'^(\d+\.|[A-Z]\.|\-|\*)\s', content))
        
        is_continuation = (not last_line_was_terminated and 
                          not has_page_number and 
                          not starts_with_number and
                          current_item is not None)
        
        if is_continuation:
            # Append to current item (making sure current_item is not None)
            if current_item is not None:
                current_item = current_item + ' ' + content
        else:
            # Finalize previous item
            if current_item is not None:
                clean_lines.append(('    ' * current_indent) + current_item)
            
            # Start new item
            current_item = content
            current_indent = indent // 4  # Use 4 spaces as the indent unit
            last_line_was_terminated = has_page_number
    
    # Add the last item
    if current_item is not None:
        clean_lines.append(('    ' * current_indent) + current_item)
    
    # Standardize page numbers and clean up titles
    processed_lines = []
    for line in clean_lines:
        indent = len(line) - len(line.lstrip())
        content = line.lstrip()
        
        # Standardize page number format - ensure space before page number
        content = re.sub(r'([\.\- ]{3,}|\(|\s)(\d+|[IVXLCDM]+)\s*$', r' \2', content)
        content = re.sub(r'\s+(\d+|[IVXLCDM]+)\s*$', r' \1', content)
        
        processed_lines.append((' ' * indent) + content)
    
    return '\n'.join(processed_lines)

def parse_clean_outline(clean_outline: str) -> List[Dict]:
    """
    Step 2: Parse the clean outline into hierarchical JSON structure
    """
    lines = [line for line in clean_outline.splitlines() if line.strip()]
    root = {"title": "ROOT", "level": -1, "page_number": 0, "children": []}
    stack = [root]
    last_page_number = 0
    
    for line in lines:
        if not line.strip():
            continue  # Skip empty lines
            
        indent = len(line) - len(line.lstrip())
        level = indent // 4  # Use 4 spaces as the indent unit
        content = line.lstrip()
        
        # Split into title and page number
        page_match = re.search(r'\s+(\d+|[IVXLCDM]+)$', content)
        if page_match:
            title = content[:page_match.start()].strip()
            page_str = page_match.group(1)
            page_number = convert_to_number(page_str)
            last_page_number = page_number
        else:
            title = content
            page_number = last_page_number  # Use the last page number if not specified
        
        new_item = {
            "title": title,  # Keep leading numbers - don't clean title
            "level": level,
            "page_number": page_number,
            "children": []
        }
        
        # Find correct parent
        while level <= stack[-1]["level"]:
            stack.pop()
        
        # Add to parent
        stack[-1]["children"].append(new_item)
        stack.append(new_item)
    
    return root["children"]

def convert_to_number(page_str: str) -> int:
    """Convert Roman or Arabic numerals to integer"""
    if not page_str:
        return 0
    if page_str.isdigit():
        return int(page_str)
    
    # Handle Roman numerals
    roman_map = {'I':1, 'V':5, 'X':10, 'L':50, 'C':100, 'D':500, 'M':1000}
    result = 0
    prev_value = 0
    for char in reversed(page_str.upper()):
        value = roman_map.get(char, 0)
        result += -value if value < prev_value else value
        prev_value = value
    return max(result, 0)

def parse_outline(text: str) -> List[Dict]:
    """Main function to parse outline from raw text"""
    # Step 1: Normalize outline
    clean_outline = normalize_outline(text)
    
    # Step 2: Parse clean outline into structured data
    return parse_clean_outline(clean_outline)