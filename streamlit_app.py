# streamlit_app.py (Fixed - set_page_config first)

# MUST BE THE VERY FIRST STREAMLIT COMMAND
import streamlit as st

st.set_page_config(
    page_title="SharePoint ETL Dashboard",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Now import everything else
import os
import pandas as pd
from datetime import datetime, timedelta
import time
from typing import List, Dict, Optional
import json
import plotly.express as px
import plotly.graph_objects as go
import hashlib

# Updated LlamaIndex imports for v0.10+
try:
    from llama_index import VectorStoreIndex, Document, ServiceContext
    from llama_index.embeddings import OpenAIEmbedding
    from llama_index.llms import OpenAI
    from llama_index.vector_stores import AstraDBVectorStore
    LLAMA_INDEX_AVAILABLE = True
except ImportError:
    try:
        from llama_index.core import VectorStoreIndex, Document, Settings
        from llama_index.embeddings.openai import OpenAIEmbedding
        from llama_index.llms.openai import OpenAI
        from llama_index.vector_stores.astra_db import AstraDBVectorStore
        ServiceContext = None
        LLAMA_INDEX_AVAILABLE = True
    except ImportError:
        VectorStoreIndex = None
        Document = None
        Settings = None
        ServiceContext = None
        OpenAIEmbedding = None
        OpenAI = None
        AstraDBVectorStore = None
        LLAMA_INDEX_AVAILABLE = False

# Import utilities (these should not call any Streamlit functions during import)
from utils import DocumentProcessor, SharePointClient, AstraClient
from utils import format_timestamp, calculate_time_diff, validate_file_type
from utils import create_processing_summary, progress_tracker, format_file_size

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
        font-weight: bold;
    }
    .metric-card {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 10px;
        margin: 0.5rem 0;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .auto-sync-enabled {
        background: linear-gradient(90deg, #28a745 0%, #20c997 100%);
        color: white;
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
    }
    .auto-sync-disabled {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 8px;
        border: 2px dashed #dee2e6;
        margin: 1rem 0;
    }
    .sync-status {
        font-family: 'Courier New', monospace;
        background-color: #f1f1f1;
        padding: 10px;
        border-radius: 5px;
        margin: 10px 0;
    }
    .countdown-timer {
        font-size: 1.2em;
        font-weight: bold;
        color: #1f77b4;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state with auto-sync features
def init_session_state():
    if 'processing_status' not in st.session_state:
        st.session_state.processing_status = []
    if 'last_sync_time' not in st.session_state:
        st.session_state.last_sync_time = None
    if 'document_count' not in st.session_state:
        st.session_state.document_count = 0
    if 'search_history' not in st.session_state:
        st.session_state.search_history = []
    if 'system_stats' not in st.session_state:
        st.session_state.system_stats = {
            'total_processed': 0,
            'success_rate': 100,
            'avg_processing_time': 0
        }
    
    # Auto-sync specific session state
    if 'auto_sync_enabled' not in st.session_state:
        st.session_state.auto_sync_enabled = False
    if 'sync_interval_minutes' not in st.session_state:
        st.session_state.sync_interval_minutes = 60  # 1 hour default
    if 'last_auto_sync_result' not in st.session_state:
        st.session_state.last_auto_sync_result = None
    if 'auto_sync_history' not in st.session_state:
        st.session_state.auto_sync_history = []
    if 'known_files' not in st.session_state:
        st.session_state.known_files = {}

class ChangeDetector:
    """Simple change detection for SharePoint documents"""
    
    @staticmethod
    def create_file_fingerprint(doc: Dict) -> str:
        """Create a fingerprint for a document to detect changes"""
        fingerprint_data = {
            'filename': doc.get('filename', ''),
            'modified': doc.get('modified', ''),
            'content_length': len(doc.get('content', '')),
            'file_path': doc.get('file_path', '')
        }
        fingerprint_str = json.dumps(fingerprint_data, sort_keys=True)
        return hashlib.md5(fingerprint_str.encode()).hexdigest()
    
    @staticmethod
    def detect_changes(current_docs: List[Dict], known_files: Dict) -> Dict:
        """Compare current documents with known files to detect changes"""
        changes = {
            'new_files': [],
            'modified_files': [],
            'unchanged_files': []
        }
        
        current_fingerprints = {}
        
        for doc in current_docs:
            file_id = doc.get('id') or doc.get('filename', f"doc_{len(current_fingerprints)}")
            fingerprint = ChangeDetector.create_file_fingerprint(doc)
            current_fingerprints[file_id] = fingerprint
            
            if file_id not in known_files:
                changes['new_files'].append(doc)
            elif known_files[file_id] != fingerprint:
                changes['modified_files'].append(doc)
            else:
                changes['unchanged_files'].append(doc)
        
        return changes, current_fingerprints

def get_next_sync_time() -> Optional[datetime]:
    """Calculate when the next sync should occur"""
    if not st.session_state.auto_sync_enabled or not st.session_state.last_sync_time:
        return None
    
    interval = timedelta(minutes=st.session_state.sync_interval_minutes)
    return st.session_state.last_sync_time + interval

def should_auto_sync() -> bool:
    """Check if it's time for an auto-sync"""
    if not st.session_state.auto_sync_enabled:
        return False
    
    next_sync = get_next_sync_time()
    if not next_sync:
        return True  # First sync
    
    return datetime.now() >= next_sync

def run_auto_sync(astra_client, sharepoint_client, document_processor):
    """Run the automatic sync process"""
    sync_start_time = datetime.now()
    
    try:
        # Calculate how far back to look (sync interval + 1 hour buffer)
        hours_back = max(1, (st.session_state.sync_interval_minutes // 60) + 1)
        
        st.info(f"🔄 Auto-sync: Looking for documents modified in the last {hours_back} hours...")
        
        # Get recent documents from SharePoint
        recent_docs = sharepoint_client.get_recent_changes(hours=hours_back)
        
        if not recent_docs:
            sync_result = {
                'timestamp': sync_start_time,
                'status': 'success',
                'documents_found': 0,
                'new_files': 0,
                'modified_files': 0,
                'processed': 0,
                'errors': 0,
                'message': 'No new documents found'
            }
            st.info("ℹ️ No new documents found during auto-sync")
        else:
            # Detect changes
            changes, new_fingerprints = ChangeDetector.detect_changes(
                recent_docs, 
                st.session_state.known_files
            )
            
            changed_docs = changes['new_files'] + changes['modified_files']
            
            st.info(f"📊 Found {len(changes['new_files'])} new files and {len(changes['modified_files'])} modified files")
            
            if changed_docs:
                # Process only changed documents
                with st.spinner(f"Processing {len(changed_docs)} changed documents..."):
                    processed_docs = document_processor.process_sharepoint_documents(changed_docs)
                    result = astra_client.insert_documents(processed_docs)
                
                # Update document count
                successful = result.get('successful', 0)
                st.session_state.document_count += successful
                
                sync_result = {
                    'timestamp': sync_start_time,
                    'status': 'success',
                    'documents_found': len(recent_docs),
                    'new_files': len(changes['new_files']),
                    'modified_files': len(changes['modified_files']),
                    'processed': successful,
                    'errors': result.get('failed', 0),
                    'message': f'Successfully processed {successful} documents'
                }
                
                st.success(f"✅ Auto-sync complete! Processed {successful} documents")
            else:
                sync_result = {
                    'timestamp': sync_start_time,
                    'status': 'success',
                    'documents_found': len(recent_docs),
                    'new_files': 0,
                    'modified_files': 0,
                    'processed': 0,
                    'errors': 0,
                    'message': 'All documents are up to date'
                }
                st.info("ℹ️ All documents are up to date")
            
            # Update known files
            st.session_state.known_files.update(new_fingerprints)
        
        # Update sync time and history
        st.session_state.last_sync_time = sync_start_time
        st.session_state.last_auto_sync_result = sync_result
        st.session_state.auto_sync_history.append(sync_result)
        
        # Keep only last 50 sync records
        if len(st.session_state.auto_sync_history) > 50:
            st.session_state.auto_sync_history = st.session_state.auto_sync_history[-50:]
            
    except Exception as e:
        error_result = {
            'timestamp': sync_start_time,
            'status': 'error',
            'error': str(e),
            'message': f'Auto-sync failed: {str(e)}'
        }
        
        st.session_state.last_auto_sync_result = error_result
        st.session_state.auto_sync_history.append(error_result)
        
        st.error(f"❌ Auto-sync failed: {str(e)}")

@st.cache_resource
def initialize_services():
    """Initialize all services and return clients"""
    try:
        # Initialize clients
        astra_client = AstraClient()
        sharepoint_client = SharePointClient()
        document_processor = DocumentProcessor()
        
        return astra_client, sharepoint_client, document_processor, True
        
    except Exception as e:
        st.error(f"Failed to initialize services: {str(e)}")
        return None, None, None, False

def check_configuration():
    """Check if all required environment variables are set"""
    required_vars = [
        "OPENAI_API_KEY",
        "ASTRA_DB_TOKEN", 
        "ASTRA_DB_ENDPOINT",
        "SHAREPOINT_CLIENT_ID",
        "SHAREPOINT_CLIENT_SECRET", 
        "SHAREPOINT_TENANT_ID",
        "SHAREPOINT_SITE_NAME"
    ]
    
    missing_vars = [var for var in required_vars if not os.getenv(var)]
    configured_vars = [var for var in required_vars if os.getenv(var)]
    
    return missing_vars, configured_vars

def auto_sync_interface(astra_client, sharepoint_client, document_processor):
    """Interface for auto-sync settings and controls"""
    st.subheader("🔄 Automatic Sync")
    
    # Auto-sync toggle
    col1, col2 = st.columns([2, 1])
    
    with col1:
        auto_sync_enabled = st.checkbox(
            "Enable Automatic Sync",
            value=st.session_state.auto_sync_enabled,
            help="Automatically sync SharePoint documents at regular intervals"
        )
        st.session_state.auto_sync_enabled = auto_sync_enabled
    
    with col2:
        if auto_sync_enabled:
            st.markdown('<div class="auto-sync-enabled">✅ Auto-Sync Enabled</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="auto-sync-disabled">⏹️ Auto-Sync Disabled</div>', unsafe_allow_html=True)
    
    if auto_sync_enabled:
        # Sync interval settings
        col1, col2 = st.columns(2)
        
        with col1:
            sync_interval = st.selectbox(
                "Sync Interval",
                options=[15, 30, 60, 120, 180, 360, 720],  # minutes
                index=2,  # default to 60 minutes
                format_func=lambda x: f"{x} minutes" if x < 60 else f"{x//60} hour{'s' if x//60 > 1 else ''}",
                help="How often to check for new/modified documents"
            )
            st.session_state.sync_interval_minutes = sync_interval
        
        with col2:
            # Manual sync trigger
            if st.button("🔄 Run Sync Now", type="primary"):
                run_auto_sync(astra_client, sharepoint_client, document_processor)
                st.rerun()
        
        # Sync status and timing
        st.subheader("⏱️ Sync Status")
        
        status_col1, status_col2, status_col3 = st.columns(3)
        
        with status_col1:
            if st.session_state.last_sync_time:
                last_sync_str = st.session_state.last_sync_time.strftime('%H:%M:%S')
                st.metric("Last Sync", last_sync_str)
            else:
                st.metric("Last Sync", "Never")
        
        with status_col2:
            next_sync = get_next_sync_time()
            if next_sync:
                next_sync_str = next_sync.strftime('%H:%M:%S')
                st.metric("Next Sync", next_sync_str)
                
                # Countdown timer
                time_until_sync = next_sync - datetime.now()
                if time_until_sync.total_seconds() > 0:
                    minutes_left = int(time_until_sync.total_seconds() // 60)
                    seconds_left = int(time_until_sync.total_seconds() % 60)
                    st.markdown(f'<div class="countdown-timer">⏰ {minutes_left}m {seconds_left}s until next sync</div>', 
                               unsafe_allow_html=True)
                else:
                    st.markdown('<div class="countdown-timer">⏰ Sync overdue!</div>', unsafe_allow_html=True)
            else:
                st.metric("Next Sync", "Pending")
        
        with status_col3:
            if st.session_state.last_auto_sync_result:
                result = st.session_state.last_auto_sync_result
                if result['status'] == 'success':
                    processed = result.get('processed', 0)
                    st.metric("Last Result", f"{processed} docs processed")
                else:
                    st.metric("Last Result", "Error", delta="❌")
        
        # Check if auto-sync should run
        if should_auto_sync():
            st.warning("⏰ **Auto-sync is due!** The system will sync automatically on next page refresh.")
            
            # Auto-trigger sync
            with st.spinner("Running scheduled auto-sync..."):
                run_auto_sync(astra_client, sharepoint_client, document_processor)
        
        # Auto-refresh mechanism
        if auto_sync_enabled:
            # Refresh page every 30 seconds to check for due syncs and update countdown
            time.sleep(30)
            st.rerun()
    
    else:
        st.info("💡 Enable automatic sync to periodically check for new and modified SharePoint documents.")

def display_sidebar(astra_client, sharepoint_client):
    """Display sidebar with system status and auto-sync info"""
    st.sidebar.title("🎛️ Control Panel")
    
    # Connection status
    st.sidebar.markdown("### 🔗 Connection Status")
    st.sidebar.success("✅ Services Online")
    
    # Auto-sync status in sidebar
    st.sidebar.markdown("### 🔄 Auto-Sync Status")
    
    if st.session_state.auto_sync_enabled:
        st.sidebar.success("✅ Auto-Sync Enabled")
        
        # Quick stats
        interval_str = f"{st.session_state.sync_interval_minutes} min"
        st.sidebar.metric("Interval", interval_str)
        
        if st.session_state.last_sync_time:
            time_since = calculate_time_diff(st.session_state.last_sync_time)
            st.sidebar.metric("Last Sync", time_since)
        
        # Next sync countdown
        next_sync = get_next_sync_time()
        if next_sync:
            time_until = next_sync - datetime.now()
            if time_until.total_seconds() > 0:
                minutes_left = int(time_until.total_seconds() // 60)
                st.sidebar.metric("Next Sync", f"{minutes_left}m")
            else:
                st.sidebar.warning("⏰ Sync Due!")
    else:
        st.sidebar.info("⏹️ Auto-Sync Disabled")
    
    # Quick stats
    st.sidebar.markdown("### 📊 Quick Stats")
    st.sidebar.metric("Documents", st.session_state.document_count)

def main():
    """Main application function with auto-sync integration"""
    # Initialize session state
    init_session_state()
    
    # Header
    st.markdown('<h1 class="main-header">📚 SharePoint to Astra DB ETL Dashboard</h1>', 
                unsafe_allow_html=True)
    
    # Check configuration
    missing_vars, configured_vars = check_configuration()
    
    if missing_vars:
        st.error(f"❌ Missing environment variables: {', '.join(missing_vars)}")
        st.info("Please configure these in your Render dashboard environment variables.")
        return
    
    # Initialize services
    astra_client, sharepoint_client, document_processor, services_ok = initialize_services()
    
    if not services_ok:
        st.error("❌ Failed to initialize services. Please check your configuration.")
        return
    
    # Display sidebar with auto-sync info
    display_sidebar(astra_client, sharepoint_client)
    
    # Main tabs
    tab1, tab2, tab3, tab4 = st.tabs(["📥 Data Ingestion", "🔍 Search & Query", "📊 Monitoring", "⚙️ Settings"])
    
    with tab1:
        st.header("📥 Data Ingestion")
        st.info("Manual data ingestion functionality - use Settings tab for auto-sync configuration.")
    
    with tab2:
        st.header("🔍 Search & Query")
        st.info("Search functionality will be implemented here.")
    
    with tab3:
        st.header("📊 Monitoring")
        st.info("Monitoring dashboard will be implemented here.")
    
    with tab4:
        auto_sync_interface(astra_client, sharepoint_client, document_processor)
    
    # Footer
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: #666; padding: 20px;'>"
        f"📚 SharePoint ETL Dashboard | Auto-Sync: {'✅ Enabled' if st.session_state.auto_sync_enabled else '⏹️ Disabled'} | "
        f"Last updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        "</div>", 
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
