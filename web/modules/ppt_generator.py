#!/usr/bin/env python3
"""
Wrapper module for the web application
Provides simplified interface for PPT generation
"""

import os
import sys
import re

# Import the full generation module
from generate_malayalam_hcs_ppt import generate_presentation

def parse_batch_file(batch_file_path):
    """
    Parse a batch file and return list of songs
    
    Format: HymnNum|Label|Title (one per line)
    Or: "Message" for message slide
    """
    song_list = []
    service_date = None
    
    with open(batch_file_path, 'r', encoding='utf-8') as f:
        lines = [line.strip() for line in f if line.strip()]
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Parse date directive
        if line.lower().startswith("# date:") or line.lower().startswith("date:"):
            service_date = line.split(":", 1)[1].strip()
            continue
        
        # Skip other comments
        if line.startswith("#"):
            continue
        
        # Check if it's just "Message"
        if line.lower() == 'message':
            song_list.append({
                'label': 'Message',
                'hymn_num': '',
                'title_hint': ''
            })
            continue
        
        # Parse line with format: HymnNum|Label|Title
        parts = line.split('|')
        if len(parts) < 2:
            continue
        
        hymn_num = parts[0].strip()
        label = parts[1].strip()
        title_hint = parts[2].strip() if len(parts) > 2 else ''
        
        song_list.append({
            'label': label,
            'hymn_num': hymn_num,
            'title_hint': title_hint
        })
    
    return song_list

def generate_presentation_from_song_list(song_list, output_path, service_date=None):
    """
    Generate PowerPoint presentation from song list
    
    Args:
        song_list: List of song dictionaries with keys: label, hymn_num, title_hint
        output_path: Path where the generated PPTX should be saved
        service_date: Optional service date string (e.g., "16 February 2026")
    
    Returns:
        tuple: (success: bool, message: str)
    """
    try:
        # Call the main PPT creation function
        generate_presentation(song_list, output_path, service_date)
        return (True, f"Presentation created successfully: {output_path}")
    except Exception as e:
        import traceback
        error_msg = f"{str(e)}\n\nTraceback:\n{traceback.format_exc()}"
        return (False, error_msg)
