#!/usr/bin/env python3
"""
Malayalam Church Songs PPT Generator - Web Application
Allows users to generate PowerPoint presentations via web browser
"""

from flask import Flask, render_template, request, send_file, flash, redirect, url_for, jsonify, session
import os
import sys
from werkzeug.utils import secure_filename
from datetime import datetime
import tempfile
import shutil
from io import StringIO

# Add modules directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'modules'))

# Import the PPT generation module
from ppt_generator import generate_presentation_from_song_list, parse_batch_file

app = Flask(__name__)
app.secret_key = 'malayalam-church-songs-secret-key-2026'
app.config['GENERATED_FOLDER'] = os.path.join(os.path.dirname(__file__), 'generated')

# Store progress logs temporarily
progress_logs = {}

@app.route('/')
def index():
    """Main page with song input form"""
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    """Generate PPT from form data"""
    try:
        # Get language selection
        language = request.form.get('language', 'Malayalam').strip()
        
        # Parse manual input from form
        songs_text = request.form.get('songs_text', '').strip()
        if not songs_text:
            flash('Please enter songs.', 'error')
            return redirect(url_for('index'))
        
        # Create temp file from text input
        temp_file = tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.txt')
        temp_file.write(songs_text)
        temp_file.close()
        
        song_list = parse_batch_file(temp_file.name)
        os.remove(temp_file.name)
        
        if not song_list:
            flash('No valid songs found in input.', 'error')
            return redirect(url_for('index'))
        
        # Get service date if provided
        service_date = request.form.get('service_date', '').strip()
        
        # Generate output filename
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f'{language}_HCS_{timestamp}.pptx'
        output_path = os.path.join(app.config['GENERATED_FOLDER'], output_filename)
        
        # Capture stdout to return progress messages
        from io import StringIO
        import sys
        old_stdout = sys.stdout
        sys.stdout = captured_output = StringIO()
        
        # Generate a unique ID for this generation
        gen_id = timestamp
        progress_logs[gen_id] = []
        
        try:
            # Generate the presentation
            success, message = generate_presentation_from_song_list(
                song_list, 
                output_path, 
                service_date if service_date else None,
                language=language
            )
        finally:
            # Restore stdout and capture the log
            sys.stdout = old_stdout
            progress_log = captured_output.getvalue()
            progress_logs[gen_id] = progress_log.split('\n')
        
        if success:
            # Store the log for retrieval
            response = send_file(
                output_path,
                as_attachment=True,
                download_name=output_filename,
                mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
            )
            # Explicitly set Content-Disposition header to force correct filename
            response.headers['Content-Disposition'] = f'attachment; filename="{output_filename}"'
            # Add generation ID for log retrieval
            response.headers['X-Generation-ID'] = gen_id
            response.headers['X-Has-Log'] = 'true'
            return response
        else:
            flash(f'Error generating presentation: {message}', 'error')
            return redirect(url_for('index'))
            
    except Exception as e:
        flash(f'Error: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/get_log/<gen_id>')
def get_log(gen_id):
    """Retrieve generation log"""
    if gen_id in progress_logs:
        return jsonify({'log': progress_logs[gen_id]})
    return jsonify({'log': []})

@app.route('/cleanup')
def cleanup():
    """Clean up old generated files (keep last 10)"""
    try:
        files = []
        for filename in os.listdir(app.config['GENERATED_FOLDER']):
            filepath = os.path.join(app.config['GENERATED_FOLDER'], filename)
            if os.path.isfile(filepath):
                files.append((filepath, os.path.getmtime(filepath)))
        
        # Sort by modification time
        files.sort(key=lambda x: x[1], reverse=True)
        
        # Delete all but the 10 most recent
        for filepath, _ in files[10:]:
            os.remove(filepath)
        
        return jsonify({'status': 'success', 'message': f'Cleaned up {max(0, len(files) - 10)} old files'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/about')
def about():
    """About page"""
    return render_template('about.html')

@app.route('/extract_hymns/<language>')
def extract_hymns(language):
    """Extract all hymns from PPT files based on language and generate Excel report"""
    try:
        if language == 'Malayalam':
            import extract_malayalam_hymns as extract_module
            lang_name = 'malayalam'
        elif language == 'English':
            # Add English directory to path
            sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(__file__)), 'English'))
            import extract_english_hymns as extract_module
            lang_name = 'english'
        else:
            flash(f'Unsupported language: {language}', 'error')
            return redirect(url_for('index'))
        
        # Find all PPT files
        pptx_files = extract_module.find_all_pptx_files()
        
        # Extract hymn information
        all_hymns = []
        for pptx_file in pptx_files:
            hymns = extract_module.analyze_pptx_file(pptx_file)
            all_hymns.extend(hymns)
        
        # Create Excel report in generated folder
        output_filename = f'{lang_name}_hymns_report.xlsx'
        output_path = os.path.join(app.config['GENERATED_FOLDER'], output_filename)
        
        extract_module.create_excel_report(all_hymns, output_path)
        
        # Send the file for download
        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        import traceback
        flash(f'Error extracting {language} hymns: {str(e)}\n{traceback.format_exc()}', 'error')
        return redirect(url_for('index'))

@app.route('/ppt_count/<language>')
def ppt_count(language):
    """Get count of available PPT files for a language"""
    try:
        base_path = os.path.join(os.path.dirname(__file__), 'onedrive_git_local', 
                                 'Holy Communion Services - Slides')
        
        if language == 'Malayalam':
            search_path = os.path.join(base_path, 'Malayalam HCS')
        elif language == 'English':
            search_path = os.path.join(base_path, 'English HCS')
        else:
            return jsonify({'count': 0, 'language': language})
        
        # Count .pptx files
        count = 0
        if os.path.exists(search_path):
            for root, dirs, files in os.walk(search_path):
                count += sum(1 for f in files if f.endswith('.pptx'))
        
        return jsonify({'count': count, 'language': language})
    except Exception as e:
        return jsonify({'count': 0, 'language': language, 'error': str(e)})

if __name__ == '__main__':
    # Create directories if they don't exist
    os.makedirs(app.config['GENERATED_FOLDER'], exist_ok=True)
    
    # Run the Flask app
    # For local network access, use host='0.0.0.0'
    # For localhost only, use host='127.0.0.1'
    print("=" * 60)
    print("Church Songs PPT Generator - Web Server")
    print("=" * 60)
    print("\nStarting server...")
    print("\nAccess the application at:")
    print("  - From this computer: http://localhost:5000")
    print("  - From other computers on your network: http://<your-ip>:5000")
    print("\nTo find your IP address, run: ipconfig (Windows) or ifconfig (Linux/Mac)")
    print("\nPress Ctrl+C to stop the server")
    print("=" * 60)
    
    app.run(host='0.0.0.0', port=5000, debug=False)
