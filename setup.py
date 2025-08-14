#!/usr/bin/env python3
"""
Setup script for Cell Analysis Database Auto-Update System

This script sets up the initial directory structure and installs dependencies.
"""

from setuptools import setup, find_packages
from pathlib import Path

# Read README for long description
readme_path = Path(__file__).parent / "README.md"
long_description = ""
if readme_path.exists():
    with open(readme_path, "r", encoding="utf-8") as f:
        long_description = f.read()

# Define required directories
REQUIRED_DIRS = ["source_data", "output"]

def create_directories():
    """Create required directories if they don't exist."""
    for dir_name in REQUIRED_DIRS:
        Path(dir_name).mkdir(exist_ok=True)
        print(f"✓ Created directory: {dir_name}")

setup(
    name="cell-analysis-db-updater",
    version="2.0.0",
    description="Automated Cell Analysis Database Update System with SharePoint Integration",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="Auto-generated",
    author_email="",
    packages=find_packages(),
    install_requires=[
        "pandas>=1.3.0",
        "openpyxl>=3.0.0",
    ],
    python_requires=">=3.7",
    entry_points={
        "console_scripts": [
            "cell-db-update=main:main",
        ],
    },
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Science/Research",
        "Topic :: Scientific/Engineering :: Bio-Informatics",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
    ],
)

# Create required directories after installation
if __name__ == "__main__":
    create_directories()
    print("\n✅ Setup complete! Required directories created.")
    print("\nNext steps:")
    print("1. Copy config.template.json to config.json")
    print("2. Edit config.json with your settings")
    print("3. Run: python main.py")

# import sys
# # Remove the sys.path modification as it's not needed in setup.py
# # sys.path.insert(0, str(Path(__file__).parent))

# Post-install directory creation (simplified)
print("\n📁 Creating required directories...")
create_directories() 