# Church Songs PPT Generator - Web Application

A Flask-based web application for generating PowerPoint presentations for church services. Supports both Malayalam and English hymns with automatic slide extraction and formatting.

## Features

- **Multi-Language Support**: Generate presentations in Malayalam or English
- **Browser-Based Interface**: No software installation needed for end users
- **Excel Export**: Extract all hymns from PPT archives to Excel spreadsheets

## Prerequisites

- Python 3.8 or higher
- Flask 3.0.0
- python-pptx 0.6.23
- openpyxl 3.1.2

## Installation & Setup

### Quick Start

1. Navigate to the web folder:
```bash
cd web
```

2. Run the startup script:
```bash
./start_server.sh
```

The script will automatically:
- Install all required Python packages
- Create necessary directories (generated/)
- Set up symbolic links to source files
- Kill any process using port 5000
- Start the Flask development server

### Manual Setup

If you prefer manual setup:

```bash
# Install dependencies
pip3 install -r requirements.txt

# Create symbolic links
ln -s ../images images
ln -s ../onedrive_git_local onedrive_git_local
cd modules
ln -s ../../Malayalam/generate_malayalam_hcs_ppt.py generate_malayalam_hcs_ppt.py
ln -s ../../Malayalam/kk_hymn_search.py kk_hymn_search.py
ln -s ../../Malayalam/kk_hymn_mapping.json kk_hymn_mapping.json
ln -s ../../Malayalam/extract_malayalam_hymns.py extract_malayalam_hymns.py
ln -s ../../English/extract_english_hymns.py extract_english_hymns.py
cd ..

# Start server
python3 app.py
```

## Usage

### Accessing the Web App

After starting the server, access the application at:
- **Local computer**: http://localhost:5000
- **Network access**: http://YOUR-IP-ADDRESS:5000

### Generating a PowerPoint Presentation

1. Select language (Malayalam or English)
2. Optionally enter a service date (e.g., "16 February 2026")
3. Enter songs in the text area using format: `HymnNum|Label|Title`
   - Example: `420|Offertory|`
   - Special: `Message` for message-only slides
4. Click "Generate PowerPoint" button
5. Download the generated PPTX file

### Extracting Hymns to Excel

1. Select language (Malayalam or English)
2. Click "Extract [Language] Hymns to Excel" button
3. Download the generated Excel file with all hymn information

## File Structure

```
web/
├── app.py                  # Flask application (main server)
├── start_server.sh         # Automated setup and start script
├── requirements.txt        # Python dependencies
├── .gitignore             # Git ignore rules
├── modules/               # Python modules (symbolic links)
│   ├── ppt_generator.py   # Web-specific wrapper
│   └── [symlinks to Malayalam/ and English/ scripts]
├── templates/             # HTML templates
│   ├── index.html         # Main page
│   └── about.html         # About page
├── static/                # Static assets
│   ├── style.css          # Stylesheet
│   └── script.js          # Client-side JavaScript
├── generated/             # Output folder for generated files
├── images/                # Symbolic link to ../images/
└── onedrive_git_local/    # Symbolic link to ../onedrive_git_local/
```

## Supported Labels

When entering songs, use these labels:
- `Opening` - Opening hymn
- `ThanksGiving` - Thanksgiving prayers
- `Offertory` - Offertory (includes QR code)
- `Message` - Message title slide (no hymn number)
- `Communion` - Holy Communion (includes special image)
- `Closing` - Closing hymn
- `Confession` - Confession
- `Dedication` - Dedication

## Configuration

### Port Configuration

Default port is 5000. To change, edit `app.py`:
```python
app.run(host='0.0.0.0', port=YOUR_PORT, debug=False)
```

### File Paths

The application automatically searches for PowerPoint files in:
- Malayalam: `onedrive_git_local/Holy Communion Services - Slides/Malayalam HCS/`
- English: `onedrive_git_local/Holy Communion Services - Slides/English HCS/`

## Deployment Notes

### For Home Network
The default configuration (`host='0.0.0.0'`) allows access from other computers on your network.

### For Internet Access
- Use SSH tunneling (serveo.net)
- Use ngrok or similar tunneling service
- Deploy to a cloud platform (Heroku, DigitalOcean, AWS, etc.)

### Security Considerations
- The development server is not suitable for production use
- For production, use a WSGI server like Gunicorn or uWSGI
- Add authentication if deploying publicly
- Set `debug=False` in production

## Troubleshooting

### Port Already in Use
The startup script automatically kills processes on port 5000. If manual intervention is needed:
```bash
lsof -ti:5000 | xargs kill -9
```

### Missing Images
Ensure the symbolic links are created properly:
```bash
ls -la images/
# Should show: images -> ../images
```

### Import Errors
Verify all symbolic links in modules/ folder:
```bash
ls -la modules/
# Should show links to Malayalam/ and English/ scripts
```

## Development

### Adding New Features
- Flask routes: Edit `app.py`
- UI changes: Edit `templates/index.html` and `static/style.css`
- Core logic: Edit source files in `Malayalam/` or `English/` directories

### Testing
Test locally before committing:
```bash
./start_server.sh
# Test in browser at http://localhost:5000
```

## License

Mar Thoma Syrian Church, Singapore

## Author

Joby - 2026
