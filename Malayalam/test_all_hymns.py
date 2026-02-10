#!/usr/bin/env python3
"""
Test script to verify each hymn from malayalam_hymns_report can be added to a PPT slide.
Creates individual test PPTs for each unique hymn number and title-only entry.
"""

import pandas as pd
import subprocess
import os
from datetime import datetime
from pathlib import Path

# Configuration
EXCEL_FILE = "malayalam_hymns_report.xlsx"
OUTPUT_DIR = "test_hymns_output"
TEST_SERVICE_DIR = "test_service_files"
GENERATOR_SCRIPT = "generate_hcs_ppt.py"
DATE_PREFIX = "7_Feb_2026"

# Create output directories
Path(OUTPUT_DIR).mkdir(exist_ok=True)
Path(TEST_SERVICE_DIR).mkdir(exist_ok=True)

def read_hymns_from_report():
    """Read and extract unique hymns from the Excel report."""
    df = pd.read_excel(EXCEL_FILE)
    
    # Get unique hymn numbers (non-empty)
    numbered_hymns = df[df['Hymn Number'].notna()]['Hymn Number'].unique()
    numbered_hymns = sorted([int(h) for h in numbered_hymns])
    
    # Get title-only entries (no hymn number)
    title_only = df[df['Hymn Number'].isna()][['Title', 'Slide Name']].drop_duplicates()
    
    return numbered_hymns, title_only

def create_test_service_file(identifier, is_numbered=True):
    """Create a service file for testing a single hymn."""
    if is_numbered:
        filename = f"{TEST_SERVICE_DIR}/test_hymn_{identifier}.txt"
        content = f"# Test service for Hymn {identifier}\n"
        content += f"# Date: {DATE_PREFIX}\n\n"
        content += f"{identifier}|Opening|\n"
    else:
        # Title-only entry
        safe_title = identifier.replace(" ", "_").replace("/", "_")[:30]
        filename = f"{TEST_SERVICE_DIR}/test_title_{safe_title}.txt"
        content = f"# Test service for title-only hymn\n"
        content += f"# Date: {DATE_PREFIX}\n\n"
        content += f"|Opening|{identifier}\n"
    
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(content)
    
    return filename

def test_hymn_generation(service_file, output_name):
    """Test generating a PPT for the given service file."""
    output_path = f"{OUTPUT_DIR}/{output_name}"
    
    try:
        result = subprocess.run(
            ["python3", GENERATOR_SCRIPT, "--batch", service_file, output_path],
            capture_output=True,
            text=True,
            timeout=30
        )
        
        # Check if output file was created
        if os.path.exists(output_path):
            return True, "Success"
        else:
            error_msg = result.stderr if result.stderr else result.stdout
            return False, f"File not created. Error: {error_msg[-200:]}"
    
    except subprocess.TimeoutExpired:
        return False, "Timeout (30s)"
    except Exception as e:
        return False, str(e)

def main():
    print(f"{'='*70}")
    print(f"Testing All Hymns from {EXCEL_FILE}")
    print(f"{'='*70}\n")
    
    # Read hymns from report
    numbered_hymns, title_only = read_hymns_from_report()
    
    print(f"Found {len(numbered_hymns)} unique hymn numbers")
    print(f"Found {len(title_only)} title-only entries")
    print(f"\nTotal tests to run: {len(numbered_hymns) + len(title_only)}\n")
    
    results = {
        'success': [],
        'failed': []
    }
    
    # Test numbered hymns
    print(f"\n{'─'*70}")
    print("Testing Numbered Hymns")
    print(f"{'─'*70}\n")
    
    for i, hymn_num in enumerate(numbered_hymns, 1):
        service_file = create_test_service_file(hymn_num, is_numbered=True)
        output_name = f"{DATE_PREFIX}_Hymn_{hymn_num}.pptx"
        
        print(f"[{i}/{len(numbered_hymns)}] Testing Hymn {hymn_num}...", end=" ")
        
        success, message = test_hymn_generation(service_file, output_name)
        
        if success:
            print(f"✅ {message}")
            results['success'].append(f"Hymn {hymn_num}")
        else:
            print(f"❌ {message}")
            results['failed'].append(f"Hymn {hymn_num}: {message}")
    
    # Test title-only hymns
    if len(title_only) > 0:
        print(f"\n{'─'*70}")
        print("Testing Title-Only Hymns")
        print(f"{'─'*70}\n")
        
        for i, row in enumerate(title_only.itertuples(index=False), 1):
            title = row.Title
            service_file = create_test_service_file(title, is_numbered=False)
            safe_title = title.replace(" ", "_").replace("/", "_")[:30]
            output_name = f"{DATE_PREFIX}_Title_{safe_title}.pptx"
            
            print(f"[{i}/{len(title_only)}] Testing '{title[:40]}'...", end=" ")
            
            success, message = test_hymn_generation(service_file, output_name)
            
            if success:
                print(f"✅ {message}")
                results['success'].append(f"Title: {title}")
            else:
                print(f"❌ {message}")
                results['failed'].append(f"Title: {title}: {message}")
    
    # Summary
    print(f"\n{'='*70}")
    print("SUMMARY")
    print(f"{'='*70}\n")
    print(f"✅ Successful: {len(results['success'])}")
    print(f"❌ Failed: {len(results['failed'])}")
    print(f"   Total: {len(results['success']) + len(results['failed'])}")
    
    if results['failed']:
        print(f"\n{'─'*70}")
        print("Failed Tests:")
        print(f"{'─'*70}")
        for fail in results['failed']:
            print(f"  • {fail}")
    
    print(f"\n{'='*70}")
    print(f"Output directory: {Path(OUTPUT_DIR).absolute()}")
    print(f"Test service files: {Path(TEST_SERVICE_DIR).absolute()}")
    print(f"{'='*70}\n")

if __name__ == "__main__":
    main()
