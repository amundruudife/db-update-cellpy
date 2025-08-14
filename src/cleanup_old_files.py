#!/usr/bin/env python3
"""
Cleanup utility for Cell Analysis Database Auto-Update System

This script helps clean up old output files, logs, and temporary data.
Run periodically to keep the workspace tidy.
"""

import os
import glob
from pathlib import Path
from datetime import datetime, timedelta

def cleanup_old_output_files(days_to_keep=7):
    """
    Remove old dry-run and backup files from output directory.
    
    Args:
        days_to_keep (int): Keep files newer than this many days
    """
    output_dir = Path("output")
    if not output_dir.exists():
        print("📁 Output directory doesn't exist - nothing to clean")
        return
    
    cutoff_date = datetime.now() - timedelta(days=days_to_keep)
    removed_count = 0
    total_size = 0
    
    # Find old files
    for pattern in ["dryrun_*.xlsx", "backup_*.xlsx"]:
        for file_path in output_dir.glob(pattern):
            file_age = datetime.fromtimestamp(file_path.stat().st_mtime)
            if file_age < cutoff_date:
                file_size = file_path.stat().st_size
                total_size += file_size
                file_path.unlink()
                removed_count += 1
                print(f"🗑️  Removed: {file_path.name}")
    
    if removed_count > 0:
        size_mb = total_size / (1024 * 1024)
        print(f"✅ Cleaned up {removed_count} files, freed {size_mb:.1f} MB")
    else:
        print(f"✅ No files older than {days_to_keep} days found")

def cleanup_python_cache():
    """Remove Python cache files"""
    cache_dirs = glob.glob("**/__pycache__", recursive=True)
    removed_count = 0
    
    for cache_dir in cache_dirs:
        try:
            import shutil
            shutil.rmtree(cache_dir)
            removed_count += 1
            print(f"🗑️  Removed cache: {cache_dir}")
        except Exception as e:
            print(f"⚠️  Could not remove {cache_dir}: {e}")
    
    if removed_count > 0:
        print(f"✅ Cleaned up {removed_count} Python cache directories")
    else:
        print("✅ No Python cache directories found")

def cleanup_source_data_copies():
    """Remove old versioned copies from source_data (keep last 5)"""
    source_dir = Path("source_data")
    if not source_dir.exists():
        return
    
    # Find versioned copies (YYMMDD_cellog_N.xlsx)
    versioned_files = list(source_dir.glob("*_cellog_*.xlsx"))
    if len(versioned_files) <= 5:
        print(f"📁 Keeping all {len(versioned_files)} source copies (≤5)")
        return
    
    # Sort by modification time, keep newest 5
    versioned_files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    files_to_remove = versioned_files[5:]  # Remove all but newest 5
    
    total_size = 0
    for file_path in files_to_remove:
        file_size = file_path.stat().st_size
        total_size += file_size
        file_path.unlink()
        print(f"🗑️  Removed old copy: {file_path.name}")
    
    if files_to_remove:
        size_mb = total_size / (1024 * 1024)
        print(f"✅ Cleaned up {len(files_to_remove)} old source copies, freed {size_mb:.1f} MB")

if __name__ == "__main__":
    print("🧹 Starting cleanup of Cell Analysis Database Auto-Update System")
    print("=" * 60)
    
    print("\n1. Cleaning up old output files (>7 days)...")
    cleanup_old_output_files(days_to_keep=7)
    
    print("\n2. Cleaning up Python cache...")
    cleanup_python_cache()
    
    print("\n3. Cleaning up old source data copies (keep newest 5)...")
    cleanup_source_data_copies()
    
    print("\n" + "=" * 60)
    print("🎉 Cleanup complete!")
    print("\n💡 To change retention:")
    print("   python cleanup_old_files.py  # Default: 7 days")
    print("   # Edit days_to_keep in script for different retention") 