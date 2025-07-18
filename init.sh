#!/bin/bash

# Exit on error
set -e

# Check if python is installed
if ! command -v python3 &> /dev/null; then
    echo "python is not installed. Please install it first using the instructions at https://www.python.org/downloads"
    exit 1
fi

# Check if uv is installed
if ! command -v uv &> /dev/null; then
    echo "uv is not installed. Installing via brew..."
    if ! command -v brew &> /dev/null; then
        echo "brew is not installed. Running install script..."
        /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
        exit 1
    fi
    brew install uv
fi

# Delete existing virtual environment
if [ -d ".venv" ]; then
    echo "Deleting existing virtual environment..."
    rm -rf .venv
fi

# Create virtual environment if it doesn't exist
if [ ! -d ".venv" ]; then
    echo "Creating virtual environment..."
    uv venv
fi

# Activate virtual environment
source .venv/bin/activate

# Install dependencies
echo "Installing dependencies..."
uv pip install -e .

# Check if .env file exists, if not create from example
if [ ! -f ".env" ]; then
    if [ -f ".env.example" ]; then
        echo "Creating .env file from .env.example..."
        cp .env.example .env
        echo "Please update .env with your settings"
    else
        echo "No .env.example file found"
        exit 1
    fi
fi

echo "Setup complete! You can now run the calendar extractor."
