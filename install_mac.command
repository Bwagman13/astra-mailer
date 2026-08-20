#!/bin/bash

# ============================================================================
#  Astra Mailer — One-Click Installer for macOS
#
#  What this does:
#    1. Checks if Python 3 is installed (guides you to install if not)
#    2. Installs required Python packages
#    3. Asks for your API key
#    4. Creates a launchable app shortcut on your Desktop
#    5. Launches the app
#
#  To run: Double-click this file, or open Terminal and run:
#    bash install_mac.command
# ============================================================================

set -e

# Colors for pretty output
BLUE='\033[1;34m'
GREEN='\033[1;32m'
YELLOW='\033[1;33m'
RED='\033[1;31m'
BOLD='\033[1m'
NC='\033[0m' # No Color

APP_DIR="$(cd "$(dirname "$0")" && pwd)"
ENV_FILE="$APP_DIR/.env"

clear
echo ""
echo -e "${BLUE}  ============================================${NC}"
echo -e "${BLUE}     Welcome to Astra Mailer Installer${NC}"
echo -e "${BLUE}  ============================================${NC}"
echo ""
echo "  This will set up everything you need."
echo "  It should take about 2-5 minutes."
echo ""
read -p "  Press Enter to continue..."

# ── Step 1: Find or Install Python 3 ────────────────────────────────────────

echo ""
echo -e "${BOLD}  [Step 1/4] Checking for Python 3...${NC}"
echo ""

PYTHON_CMD=""

# Check for python3 first (standard on macOS)
if command -v python3 &>/dev/null; then
    PYVER=$(python3 --version 2>&1 | awk '{print $2}')
    echo -e "  ${GREEN}✓${NC} Found Python $PYVER"
    PYTHON_CMD="python3"
elif command -v python &>/dev/null; then
    # Make sure it's Python 3, not 2
    PYVER=$(python --version 2>&1 | awk '{print $2}')
    MAJOR=$(echo "$PYVER" | cut -d. -f1)
    if [ "$MAJOR" -ge 3 ]; then
        echo -e "  ${GREEN}✓${NC} Found Python $PYVER"
        PYTHON_CMD="python"
    fi
fi

if [ -z "$PYTHON_CMD" ]; then
    echo -e "  ${YELLOW}⚠${NC}  Python 3 is not installed."
    echo ""

    # Check if Homebrew is available
    if command -v brew &>/dev/null; then
        echo "  Homebrew detected. Installing Python via Homebrew..."
        echo ""
        brew install python3
        PYTHON_CMD="python3"
        echo ""
        echo -e "  ${GREEN}✓${NC} Python installed via Homebrew!"
    else
        echo "  Please install Python 3 using one of these methods:"
        echo ""
        echo "    Option A — Download from python.org:"
        echo "      1. Go to https://www.python.org/downloads/"
        echo "      2. Download Python 3.12 or newer for macOS"
        echo "      3. Run the .pkg installer"
        echo "      4. Run this installer again"
        echo ""
        echo "    Option B — Install Homebrew first, then Python:"
        echo "      1. Open Terminal"
        echo '      2. Run: /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"'
        echo "      3. Run: brew install python3"
        echo "      4. Run this installer again"
        echo ""
        read -p "  Press Enter to close..."
        exit 1
    fi
fi

PIP_CMD="$PYTHON_CMD -m pip"

echo -e "  Python is ready."
echo ""

# ── Step 2: Install Python Packages ─────────────────────────────────────────

echo -e "${BOLD}  [Step 2/4] Installing required packages...${NC}"
echo "  (This may take a minute or two)"
echo ""

$PIP_CMD install --upgrade pip --quiet 2>/dev/null || true
$PIP_CMD install PySide6 openpyxl anthropic python-dotenv --quiet 2>&1 | grep -v "already satisfied" || true

# Verify key packages
MISSING=""
$PYTHON_CMD -c "import PySide6" 2>/dev/null || MISSING="$MISSING PySide6"
$PYTHON_CMD -c "import openpyxl" 2>/dev/null || MISSING="$MISSING openpyxl"
$PYTHON_CMD -c "import anthropic" 2>/dev/null || MISSING="$MISSING anthropic"
$PYTHON_CMD -c "import dotenv" 2>/dev/null || MISSING="$MISSING python-dotenv"

if [ -n "$MISSING" ]; then
    echo -e "  ${YELLOW}⚠${NC}  Some packages failed to install:$MISSING"
    echo "  Trying again..."
    $PIP_CMD install $MISSING 2>&1
    echo ""
fi

echo -e "  ${GREEN}✓${NC} Packages installed."
echo ""

# ── Step 3: API Key Setup ───────────────────────────────────────────────────

echo -e "${BOLD}  [Step 3/4] Setting up your API key...${NC}"
echo ""

# Check if .env already has a real key
if [ -f "$ENV_FILE" ] && grep -q "sk-ant-" "$ENV_FILE" 2>/dev/null; then
    echo -e "  ${GREEN}✓${NC} API key already configured. Skipping."
else
    echo "  You should have received an API key that looks like:"
    echo "    sk-ant-api03-..."
    echo ""
    echo "  Paste it below and press Enter."
    echo "  (Cmd+V to paste)"
    echo ""
    read -p "  Your API Key: " API_KEY

    if [ -z "$API_KEY" ]; then
        echo ""
        echo "  No key entered. You can add it later inside the app."
        echo "  (Go to the Setup tab and paste it in the API Key field)"
        echo "ANTHROPIC_API_KEY=" > "$ENV_FILE"
    else
        echo "ANTHROPIC_API_KEY=$API_KEY" > "$ENV_FILE"
        echo ""
        echo -e "  ${GREEN}✓${NC} API key saved!"
    fi
fi

echo ""

# ── Step 4: Create Desktop Launcher ─────────────────────────────────────────

echo -e "${BOLD}  [Step 4/4] Creating desktop launcher...${NC}"
echo ""

DESKTOP="$HOME/Desktop"
LAUNCHER="$DESKTOP/Astra Mailer.command"

cat > "$LAUNCHER" << LAUNCHER_EOF
#!/bin/bash
cd "$APP_DIR"
$PYTHON_CMD "$APP_DIR/astra_mailer.py" &
disown
exit
LAUNCHER_EOF

chmod +x "$LAUNCHER"

# Also make sure this installer is executable
chmod +x "$APP_DIR/install_mac.command"

if [ -f "$LAUNCHER" ]; then
    echo -e "  ${GREEN}✓${NC} Desktop launcher created!"
    echo "  You'll find \"Astra Mailer\" on your Desktop."
    echo ""
    echo -e "  ${YELLOW}Note:${NC} The first time you open it, macOS may ask you to"
    echo "  confirm. Right-click the file → Open → click \"Open\" again."
else
    echo -e "  ${YELLOW}⚠${NC}  Could not create desktop launcher."
    echo "  You can run the app from Terminal with:"
    echo "    cd \"$APP_DIR\" && $PYTHON_CMD astra_mailer.py"
fi

# ── Done! ────────────────────────────────────────────────────────────────────

echo ""
echo -e "${BLUE}  ============================================${NC}"
echo -e "${BLUE}     Installation Complete!${NC}"
echo -e "${BLUE}  ============================================${NC}"
echo ""
echo "  You can now:"
echo "    • Double-click \"Astra Mailer\" on your Desktop"
echo "    • Or open Terminal and run:"
echo "        cd \"$APP_DIR\" && $PYTHON_CMD astra_mailer.py"
echo ""

read -p "  Launch Astra Mailer now? (y/n): " LAUNCH

if [[ "$LAUNCH" =~ ^[Yy]$ ]]; then
    echo ""
    echo "  Starting Astra Mailer..."
    cd "$APP_DIR"
    $PYTHON_CMD "$APP_DIR/astra_mailer.py" &
    disown
fi

echo ""
echo "  You can close this window."
echo ""
