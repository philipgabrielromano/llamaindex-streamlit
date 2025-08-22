# utils/helpers.py (Complete version with all missing functions)
import streamlit as st
from datetime import datetime, timedelta
from typing import List, Dict, Optional, Union
import pandas as pd
import json
import os
import contextlib

def format_timestamp(timestamp: Union[str, datetime], format_str: str = "%Y-%m-%d %H:%M:%S") -> str:
    """Format timestamp for display"""
    try:
        if isinstance(timestamp, str):
            dt = datetime.fromisoformat(timestamp.replace('Z', '+00:00'))
        else:
            dt = timestamp
        
        return dt.strftime(format_str)
    except Exception:
        return "Unknown"

def calculate_time_diff(start_time: Union[str, datetime], 
                       end_time: Optional[Union[str, datetime]] = None) -> str:
    """Calculate human-readable time difference"""
    try:
        if isinstance(start_time, str):
            start_dt = datetime.fromisoformat(start_time.replace('Z', '+00:00'))
        else:
            start_dt = start_time
        
        if end_time is None:
            end_dt = datetime.now()
        elif isinstance(end_time, str):
            end_dt = datetime.fromisoformat(end_time.replace('Z', '+00:00'))
        else:
            end_dt = end_time
        
        diff = end_dt - start_dt
        
        if diff.days > 0:
            return f"{diff.days}d ago"
        elif diff.seconds > 3600:
            hours = diff.seconds // 3600
            return f"{hours}h ago"
        elif diff.seconds > 60:
            minutes = diff.seconds // 60
            return f"{minutes}m ago"
        else:
            return "Just now"
            
    except Exception:
        return "Unknown"

def validate_file_type(filename: str, allowed_types: List[str] = None) -> bool:
    """Validate file type"""
    if allowed_types is None:
        allowed_types = ['.pdf', '.docx', '.txt', '.pptx', '.md', '.html', '.csv', '.json', '.xml']
    
    file_ext = f".{filename.lower().split('.')[-1]}" if '.' in filename else ''
    return file_ext in allowed_types

def format_file_size(size_bytes: int) -> str:
    """Format file size for display"""
    try:
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024**2:
            return f"{size_bytes/1024:.1f} KB"
        elif size_bytes < 1024**3:
            return f"{size_bytes/(1024**2):.1f} MB"
        else:
            return f"{size_bytes/(1024**3):.1f} GB"
    except Exception:
        return "Unknown size"

def create_processing_summary(processing_results: List[Dict]) -> Dict:
    """Create summary of processing results"""
    if not processing_results:
        return {
            'total': 0,
            'successful': 0,
            'failed': 0,
            'success_rate': 0
        }
    
    total = len(processing_results)
    successful = sum(1 for result in processing_results if result.get('status') == 'Success')
    failed = total - successful
    success_rate = (successful / total * 100) if total > 0 else 0
    
    return {
        'total': total,
        'successful': successful,
        'failed': failed,
        'success_rate': success_rate
    }

def safe_json_display(data: Dict) -> str:
    """Safely display JSON data"""
    try:
        return json.dumps(data, indent=2, default=str)
    except Exception:
        return str(data)

def create_metrics_dataframe(metrics_data: List[Dict]) -> pd.DataFrame:
    """Create DataFrame from metrics data"""
    try:
        if not metrics_data:
            return pd.DataFrame()
        
        df = pd.DataFrame(metrics_data)
        
        # Convert timestamp columns
        timestamp_cols = ['timestamp', 'processed_at', 'created_at', 'indexed_at']
        for col in timestamp_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
        
        return df
        
    except Exception as e:
        st.error(f"Error creating metrics DataFrame: {str(e)}")
        return pd.DataFrame()

def display_error_details(error: Exception, context: str = ""):
    """Display detailed error information"""
    error_msg = str(error)
    error_type = type(error).__name__
    
    with st.expander(f"Error Details {context}"):
        st.write(f"**Error Type:** {error_type}")
        st.write(f"**Error Message:** {error_msg}")
