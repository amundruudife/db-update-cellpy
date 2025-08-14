#!/usr/bin/env python3
"""
Quick Setup Script for Cell Analysis Database Auto-Update System

This script helps new users set up the project for first use.
"""

import json
import shutil
import sys
from pathlib import Path

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent / "src"))

def main():
    """Main setup function"""
    print("🔧 Cell Analysis Database Auto-Update System - Quick Setup")
    print("=" * 60)
    
    # Check if config.json already exists
    if Path("config.json").exists():
        print("⚠️  config.json already exists!")
        response = input("Do you want to overwrite it? (y/N): ").lower()
        if response != 'y':
            print("Setup cancelled.")
            return
    
    # Copy template to config.json
    try:
        shutil.copy("config.template.json", "config.json")
        print("✅ Created config.json from template")
    except FileNotFoundError:
        print("❌ config.template.json not found!")
        return
    
    # Create necessary directories
    Path("source_data").mkdir(exist_ok=True)
    Path("output").mkdir(exist_ok=True)
    print("✅ Created necessary directories")
    
    # Load the config for editing
    with open("config.json", 'r') as f:
        config = json.load(f)
    
    print("\n📝 Configuration Setup")
    print("You can edit these settings later in config.json")
    print("-" * 40)
    
    # Get basic configuration from user
    projects_input = input(f"Enter your project names (comma-separated) [{', '.join(config['projects'])}]: ").strip()
    if projects_input:
        config['projects'] = [p.strip() for p in projects_input.split(',')]
    
    work_dir = input(f"Enter your work directory (current: {config['work_dir']}): ").strip()
    if work_dir:
        config['work_dir'] = work_dir
        config['source_path'] = f"{work_dir}/source_data/Cell_Log.xlsx"
    
    db_path = input(f"Enter your database path (current: {config['db_path']}): ").strip()
    if db_path:
        config['db_path'] = db_path
    
    # Save updated configuration
    with open("config.json", 'w') as f:
        json.dump(config, f, indent=2)
    
    print("\n✅ Configuration saved!")
    print("\n🚀 Next Steps:")
    print("1. Place your Cell_Log.xlsx in source_data/ directory")
    print("2. Test with: python main.py --dry-run")
    print("3. Run live update: python main.py --live")
    print("\n📖 For more help, see README.md")

if __name__ == "__main__":
    main() 