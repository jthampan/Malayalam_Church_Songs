#!/bin/bash
# Quick start script for Malayalam Church Songs Web App

echo "================================================"
echo "Malayalam Church Songs PPT Generator - Web App"
echo "================================================"
echo ""

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 is not installed!"
    echo "Please install Python 3.8 or higher first."
    exit 1
fi

echo "✅ Python $(python3 --version) found"
echo ""

# Install dependencies
echo "📦 Installing dependencies..."
pip3 install -r requirements.txt --break-system-packages --quiet

if [ $? -ne 0 ]; then
    echo "❌ Failed to install dependencies"
    exit 1
fi

echo "✅ Dependencies installed"
echo ""

# Create necessary directories
mkdir -p generated

# Create symbolic links if they don't exist
echo "🔗 Setting up symbolic links..."

# Link to parent images folder
if [ ! -L "images" ] && [ ! -d "images" ]; then
    ln -s ../images images
    echo "✅ Created images/ → ../images link"
elif [ -L "images" ]; then
    echo "✅ images/ link already exists"
else
    echo "⚠️  Warning: images/ exists as directory, not a symbolic link"
fi

# Link to onedrive_git_local folder
if [ ! -L "onedrive_git_local" ] && [ ! -d "onedrive_git_local" ]; then
    ln -s ../onedrive_git_local onedrive_git_local
    echo "✅ Created onedrive_git_local/ link"
elif [ -L "onedrive_git_local" ]; then
    echo "✅ onedrive_git_local/ link already exists"
else
    echo "⚠️  Warning: onedrive_git_local/ exists as directory, not a symbolic link"
fi

# Create symbolic links in modules/ to Malayalam/ scripts
mkdir -p modules
cd modules

# Link core Python files from Malayalam directory
for file in generate_malayalam_hcs_ppt.py kk_hymn_search.py kk_hymn_mapping.json extract_malayalam_hymns.py; do
    if [ ! -L "$file" ] && [ ! -f "$file" ]; then
        ln -s ../../Malayalam/"$file" "$file"
        echo "✅ Linked modules/$file → Malayalam/$file"
    elif [ -L "$file" ]; then
        echo "✅ modules/$file link already exists"
    fi
done

# Link English scripts
for file in extract_english_hymns.py generate_english_hcs_ppt.py; do
    if [ ! -L "$file" ] && [ ! -f "$file" ]; then
        ln -s ../../English/"$file" "$file"
        echo "✅ Linked modules/$file → English/$file"
    elif [ -L "$file" ]; then
        echo "✅ modules/$file link already exists"
    fi
done

cd ..
echo ""

# Get local IP address
echo "🌐 Finding your IP address..."
if command -v ifconfig &> /dev/null; then
    LOCAL_IP=$(ifconfig | grep "inet " | grep -v 127.0.0.1 | awk '{print $2}' | head -1)
elif command -v ip &> /dev/null; then
    LOCAL_IP=$(ip addr show | grep "inet " | grep -v 127.0.0.1 | awk '{print $2}' | cut -d/ -f1 | head -1)
else
    LOCAL_IP="<your-ip-address>"
fi

echo ""

# Kill any existing process on port 5000
echo "🔍 Checking if port 5000 is already in use..."
PORT_PID=$(lsof -ti:5000 2>/dev/null)
if [ ! -z "$PORT_PID" ]; then
    echo "⚠️  Port 5000 is in use by process $PORT_PID"
    echo "🔪 Killing existing process..."
    kill -9 $PORT_PID
    sleep 1
    echo "✅ Process terminated"
else
    echo "✅ Port 5000 is available"
fi

echo ""
echo "================================================"
echo "🚀 Starting Web Server..."
echo "================================================"
echo ""
echo "Access URLs:"
echo "  • This computer:    http://localhost:5000"
echo "  • Other computers:  http://$LOCAL_IP:5000"
echo ""
echo "Press Ctrl+C to stop the server"
echo "================================================"
echo ""

# Run the Flask app
python3 app.py
