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
        
        # Show stack trace in development
        if os.getenv("STREAMLIT_ENV") == "development":
            import traceback
            st.code(traceback.format_exc())

def sanitize_filename(filename: str) -> str:
    """Sanitize filename for safe storage"""
    import re
    # Remove or replace unsafe characters
    sanitized = re.sub(r'[<>:"/\\|?*]', '_', filename)
    return sanitized.strip()

def chunk_list(lst: List, chunk_size: int) -> List[List]:
    """Chunk a list into smaller lists"""
    return [lst[i:i + chunk_size] for i in range(0, len(lst), chunk_size)]

@st.cache_data(ttl=300)  # Cache for 5 minutes
def get_cached_stats(data_key: str) -> Optional[Dict]:
    """Get cached statistics data"""
    # This would integrate with your caching strategy
    return None

class ProgressTracker:
    """Context manager for progress tracking"""
    
    def __init__(self, items: List, description: str = "Processing"):
        self.items = items
        self.total = len(items)
        self.current = 0
        self.progress_bar = None
        self.status_text = None
        self.description = description
    
    def __enter__(self):
        self.progress_bar = st.progress(0)
        self.status_text = st.empty()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.progress_bar:
            self.progress_bar.empty()
        if self.status_text:
            self.status_text.empty()
    
    def update(self, item_name: str = ""):
        """Update progress"""
        self.current += 1
        if self.total > 0:
            progress = self.current / self.total
            if self.progress_bar:
                self.progress_bar.progress(progress)
            if self.status_text:
                self.status_text.text(f"{self.description}: {item_name} ({self.current}/{self.total})")
    
    def complete(self):
        """Mark as complete"""
        if self.progress_bar:
            self.progress_bar.empty()
        if self.status_text:
            self.status_text.empty()

def progress_tracker(items: List, description: str = "Processing"):
    """Context manager for progress tracking"""
    return ProgressTracker(items, description)

def format_duration(seconds: float) -> str:
    """Format duration in seconds to human readable format"""
    try:
        if seconds < 60:
            return f"{seconds:.1f}s"
        elif seconds < 3600:
            minutes = int(seconds // 60)
            remaining_seconds = seconds % 60
            return f"{minutes}m {remaining_seconds:.0f}s"
        else:
            hours = int(seconds // 3600)
            remaining_minutes = int((seconds % 3600) // 60)
            return f"{hours}h {remaining_minutes}m"
    except Exception:
        return "Unknown duration"

def get_file_extension(filename: str) -> str:
    """Get file extension from filename"""
    try:
        return filename.split('.')[-1].lower() if '.' in filename else ''
    except Exception:
        return ''

def is_text_file(filename: str) -> bool:
    """Check if file is a text-based file"""
    text_extensions = {'txt', 'md', 'csv', 'json', 'xml', 'html', 'htm'}
    return get_file_extension(filename) in text_extensions

def is_office_file(filename: str) -> bool:
    """Check if file is a Microsoft Office file"""
    office_extensions = {'docx', 'doc', 'pptx', 'ppt', 'xlsx', 'xls'}
    return get_file_extension(filename) in office_extensions

def is_pdf_file(filename: str) -> bool:
    """Check if file is a PDF"""
    return get_file_extension(filename) == 'pdf'

def create_file_summary(files: List[Dict]) -> Dict:
    """Create a summary of file types and counts"""
    if not files:
        return {'total': 0, 'by_type': {}}
    
    summary = {'total': len(files), 'by_type': {}}
    
    for file_info in files:
        filename = file_info.get('filename', '')
        file_ext = get_file_extension(filename)
        
        if file_ext:
            summary['by_type'][file_ext] = summary['by_type'].get(file_ext, 0) + 1
        else:
            summary['by_type']['unknown'] = summary['by_type'].get('unknown', 0) + 1
    
    return summary

def filter_recent_items(items: List[Dict], hours: int = 24, timestamp_key: str = 'timestamp') -> List[Dict]:
    """Filter items to only include those from the last N hours"""
    if not items:
        return []
    
    cutoff_time = datetime.now() - timedelta(hours=hours)
    recent_items = []
    
    for item in items:
        try:
            item_time = item.get(timestamp_key)
            if isinstance(item_time, str):
                item_time = datetime.fromisoformat(item_time.replace('Z', '+00:00'))
            elif isinstance(item_time, datetime):
                pass
            else:
                continue  # Skip items without valid timestamps
            
            if item_time >= cutoff_time:
                recent_items.append(item)
        except Exception:
            continue  # Skip items with invalid timestamps
    
    return recent_items

def get_system_info() -> Dict:
    """Get basic system information"""
    try:
        import platform
        import sys
        
        return {
            'python_version': sys.version.split()[0],
            'platform': platform.platform(),
            'streamlit_version': st.__version__,
            'timestamp': datetime.now().isoformat()
        }
    except Exception:
        return {
            'python_version': 'Unknown',
            'platform': 'Unknown', 
            'streamlit_version': 'Unknown',
            'timestamp': datetime.now().isoformat()
        }

def create_status_indicator(status: str) -> str:
    """Create a visual status indicator"""
    status_map = {
        'success': '✅',
        'error': '❌',
        'warning': '⚠️',
        'info': 'ℹ️',
        'processing': '🔄',
        'pending': '⏳'
    }
    return status_map.get(status.lower(), '❓')

def truncate_text(text: str, max_length: int = 100, suffix: str = "...") -> str:
    """Truncate text to specified length"""
    if not text or len(text) <= max_length:
        return text
    
    return text[:max_length - len(suffix)] + suffix

def calculate_statistics(values: List[Union[int, float]]) -> Dict:
    """Calculate basic statistics for a list of numerical values"""
    if not values:
        return {'count': 0, 'sum': 0, 'avg': 0, 'min': 0, 'max': 0}
    
    try:
        return {
            'count': len(values),
            'sum': sum(values),
            'avg': sum(values) / len(values),
            'min': min(values),
            'max': max(values)
        }
    except Exception:
        return {'count': len(values), 'sum': 0, 'avg': 0, 'min': 0, 'max': 0}
