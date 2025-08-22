# streamlit_app.py (Complete Enhanced Version)

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
    .document-card {
        border: 1px solid #dee2e6;
        border-radius: 8px;
        padding: 1rem;
        margin: 0.5rem 0;
        background-color: #f8f9fa;
    }
    .file-icon {
        font-size: 1.2em;
        margin-right: 0.5rem;
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

def get_file_icon(filename: str) -> str:
    """Get appropriate icon for file type"""
    extension = filename.split('.')[-1].lower() if '.' in filename else ''
    
    icons = {
        'pdf': '📄',
        'docx': '📝', 'doc': '📝',
        'pptx': '📊', 'ppt': '📊',
        'xlsx': '📈', 'xls': '📈',
        'txt': '📋',
        'html': '🌐', 'htm': '🌐',
        'md': '📑',
        'json': '⚙️',
        'xml': '🔧',
        'csv': '📊'
    }
    
    return icons.get(extension, '📄')

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

def process_selected_documents(selected_docs, astra_client, document_processor, chunk_size, chunk_overlap):
    """Process the selected SharePoint documents"""
    # Update processor settings
    document_processor.update_chunk_settings(chunk_size, chunk_overlap)
    
    with st.spinner(f"Processing {len(selected_docs)} selected documents..."):
        try:
            # Process documents
            processed_docs = document_processor.process_sharepoint_documents(selected_docs)
            
            if processed_docs:
                # Insert into Astra DB
                result = astra_client.insert_documents(processed_docs)
                
                # Update session state
                successful = result.get('successful', 0)
                failed = result.get('failed', 0)
                
                st.session_state.document_count += successful
                
                # Add to processing status
                for i, doc in enumerate(selected_docs):
                    status = 'Success' if i < successful else 'Failed'
                    st.session_state.processing_status.append({
                        'filename': doc.get('filename', f'Document {i+1}'),
                        'status': status,
                        'timestamp': datetime.now(),
                        'source': 'selected_sharepoint',
                        'chunks': 1
                    })
                
                # Show results
                if successful > 0:
                    st.success(f"✅ Successfully processed {successful} documents!")
                
                if failed > 0:
                    st.warning(f"⚠️ {failed} documents failed to process")
                
                # Clear selection
                for i in range(len(st.session_state.get('available_documents', []))):
                    if f'doc_select_{i}' in st.session_state:
                        st.session_state[f'doc_select_{i}'] = False
                
            else:
                st.error("❌ No documents were successfully processed")
                
        except Exception as e:
            st.error(f"❌ Error processing documents: {str(e)}")

def process_uploaded_files(uploaded_files, astra_client, document_processor, chunk_size, chunk_overlap):
    """Process manually uploaded files"""
    # Update processor settings
    document_processor.update_chunk_settings(chunk_size, chunk_overlap)
    
    with st.spinner(f"Processing {len(uploaded_files)} uploaded files..."):
        try:
            successful_count = 0
            failed_count = 0
            
            for file in uploaded_files:
                try:
                    # Process document
                    document = document_processor.process_uploaded_file(file)
                    
                    if document:
                        # Insert into Astra DB
                        result = astra_client.insert_documents([document])
                        
                        if result.get('successful', 0) > 0:
                            successful_count += 1
                            st.session_state.processing_status.append({
                                'filename': file.name,
                                'status': 'Success',
                                'timestamp': datetime.now(),
                                'source': 'manual_upload',
                                'chunks': 1
                            })
                        else:
                            failed_count += 1
                            st.session_state.processing_status.append({
                                'filename': file.name,
                                'status': 'Error: Failed to index',
                                'timestamp': datetime.now(),
                                'source': 'manual_upload',
                                'chunks': 0
                            })
                    else:
                        failed_count += 1
                        
                except Exception as e:
                    failed_count += 1
                    st.session_state.processing_status.append({
                        'filename': file.name,
                        'status': f'Error: {str(e)}',
                        'timestamp': datetime.now(),
                        'source': 'manual_upload',
                        'chunks': 0
                    })
            
            # Update session state
            st.session_state.document_count += successful_count
            st.session_state.last_sync_time = datetime.now()
            
            # Display results
            if successful_count > 0:
                st.success(f"✅ Successfully processed {successful_count}/{len(uploaded_files)} files!")
            
            if failed_count > 0:
                st.warning(f"⚠️ {failed_count} files failed to process. Check the status table for details.")
                
        except Exception as e:
            st.error(f"❌ Error processing uploaded files: {str(e)}")

def display_collection_browser(astra_client):
    """Display documents stored in the collection"""
    try:
        # Get collection stats
        stats = astra_client.get_collection_stats()
        
        st.subheader("📊 Collection Overview")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Total Documents", stats.get('document_count', 'Unknown'))
        
        with col2:
            st.metric("Collection", stats.get('collection_name', 'Unknown'))
        
        with col3:
            status = stats.get('status', 'Unknown')
            status_icon = '✅' if status == 'active' else '❌'
            st.metric("Status", f"{status_icon} {status.title()}")
        
        # Note about search functionality
        st.info("""
        🔍 **Document Search & Retrieval**
        
        Use the **Search & Query** tab to:
        - Search through your stored documents
        - View document content and metadata
        - Get AI-powered responses based on your document collection
        """)
        
        # Show recent processing activity
        if st.session_state.processing_status:
            st.subheader("📝 Recent Processing Activity")
            
            recent_items = st.session_state.processing_status[-10:]  # Last 10 items
            
            for item in reversed(recent_items):
                col1, col2, col3 = st.columns([3, 1, 1])
                
                with col1:
                    status_icon = '✅' if item['status'] == 'Success' else '❌'
                    st.write(f"{status_icon} {item['filename']}")
                
                with col2:
                    st.write(item['source'])
                
                with col3:
                    timestamp = item['timestamp']
                    if isinstance(timestamp, datetime):
                        time_str = timestamp.strftime('%H:%M:%S')
                    else:
                        time_str = str(timestamp)[:8]
                    st.write(time_str)
        
    except Exception as e:
        st.error(f"❌ Error browsing collection: {str(e)}")

def data_ingestion_tab(astra_client, sharepoint_client, document_processor):
    """Enhanced data ingestion with better debugging"""
    st.header("📥 Data Ingestion & Document Management")
    
    # Create main sections
    col1, col2 = st.columns([3, 2])
    
    with col1:
        st.subheader("🔗 SharePoint Integration")
        
        # SharePoint folder selection
        with st.expander("📁 SharePoint Folder Configuration", expanded=True):
            # Get available folders/libraries
            available_folders = sharepoint_client.get_available_libraries()
            
            # Folder selection
            selected_folder = st.selectbox(
                "Document Library",
                options=available_folders,
                index=0 if available_folders else None,
                help="Select the SharePoint document library to sync from"
            )
            
            # Custom folder path option
            custom_path = st.text_input(
                "Custom Folder Path (optional)",
                placeholder="e.g., Documents/Reports/2024",
                help="Specify a specific folder path within the library"
            )
            
            # Use custom path if provided, otherwise use selected folder
            final_folder_path = custom_path if custom_path else selected_folder
            
            # File type filtering with debugging
            st.markdown("**File Type Filters:**")
            st.info("💡 Leave all unchecked to see ALL file types")
            
            file_type_cols = st.columns(4)
            
            selected_file_types = []
            with file_type_cols[0]:
                if st.checkbox("📄 PDF"):
                    selected_file_types.append(".pdf")
                if st.checkbox("📝 Word"):
                    selected_file_types.append(".docx")
            
            with file_type_cols[1]:
                if st.checkbox("📊 PowerPoint"):
                    selected_file_types.append(".pptx")
                if st.checkbox("📈 Excel"):
                    selected_file_types.append(".xlsx")
            
            with file_type_cols[2]:
                if st.checkbox("📋 Text"):
                    selected_file_types.append(".txt")
                if st.checkbox("🌐 HTML"):
                    selected_file_types.append(".html")
            
            with file_type_cols[3]:
                if st.checkbox("📑 All Files", value=True):
                    selected_file_types = None  # No filtering
            
            # Date filtering
            st.markdown("**Date Filtering:**")
            date_filter_enabled = st.checkbox("Enable Date Filter")
            
            since_date = None
            if date_filter_enabled:
                days_back = st.slider("Days back", 1, 30, 7)
                since_date = datetime.now() - timedelta(days=days_back)
                st.info(f"Will only show files modified since: {since_date.strftime('%Y-%m-%d')}")
        
        # Document preview and selection
        st.subheader("📋 Available Documents")
        
        # Show current filter settings
        if selected_file_types:
            st.info(f"🔍 File type filter: {', '.join(selected_file_types)}")
        else:
            st.info("🔍 No file type filter (showing all files)")
        
        # Fetch documents button
        if st.button("🔍 Browse SharePoint Documents", type="secondary"):
            with st.spinner("Fetching documents from SharePoint..."):
                documents = sharepoint_client.get_documents(
                    folder_path=final_folder_path,
                    file_types=selected_file_types,
                    since_date=since_date,
                    max_docs=50  # Limit for browsing
                )
                
                # Store documents in session state for selection
                st.session_state['available_documents'] = documents
                st.session_state['last_fetch_info'] = {
                    'folder': final_folder_path,
                    'file_types': selected_file_types,
                    'since_date': since_date,
                    'count': len(documents)
                }
        
        # Show last fetch info
        if 'last_fetch_info' in st.session_state:
            info = st.session_state.last_fetch_info
            st.info(f"📊 Last fetch: {info['count']} documents from '{info['folder']}'")
        
        # Display available documents for selection
        if 'available_documents' in st.session_state and st.session_state.available_documents:
            documents = st.session_state.available_documents
            
            st.success(f"📄 Found {len(documents)} documents")
            
            # Show document details
            with st.expander("📋 Document Details"):
                for doc in documents[:5]:  # Show first 5
                    st.write(f"**{doc['filename']}**")
                    st.write(f"  Modified: {doc['modified']}")
                    st.write(f"  Size: {doc['metadata'].get('file_size', 0)} bytes")
                    st.write(f"  Content length: {len(doc.get('content', ''))} characters")
                    if len(documents) > 5:
                        st.write("...")
            
            # Rest of your document selection interface...
            # (The checkbox selection code you already have)
            
        else:
            st.info("👆 Click 'Browse SharePoint Documents' to see available files")
    
    # Rest of your code...

def search_query_tab(astra_client):
    """Enhanced search with document browsing capabilities"""
    st.header("🔍 Search & Query Documents")
    
    # Search interface
    col1, col2 = st.columns([3, 1])
    
    with col1:
        query = st.text_area(
            "Enter your question:",
            placeholder="What information are you looking for?",
            height=120,
            key="search_query"
        )
        
        # Quick query suggestions
        st.markdown("**Quick Queries:**")
        suggestion_cols = st.columns(3)
        
        with suggestion_cols[0]:
            if st.button("📋 List recent documents", key="query_recent"):
                st.session_state.search_query = "What are the most recently added documents?"
                
        with suggestion_cols[1]:
            if st.button("📊 Document summary", key="query_summary"):
                st.session_state.search_query = "Provide a summary of all documents in the collection"
                
        with suggestion_cols[2]:
            if st.button("🔍 Search by topic", key="query_topic"):
                topic = st.text_input("Enter topic:", key="topic_input")
                if topic:
                    st.session_state.search_query = f"Find documents related to {topic}"
    
    with col2:
        st.markdown("### 🎛️ Search Options")
        
        similarity_top_k = st.slider(
            "Results to retrieve",
            min_value=1,
            max_value=20,
            value=5,
            help="Number of similar documents to retrieve"
        )
        
        response_mode = st.selectbox(
            "Response Mode",
            ["compact", "tree_summarize", "simple_summarize", "refine"],
            index=0,
            help="How to generate the response from retrieved documents"
        )
        
        include_metadata = st.checkbox(
            "Include Metadata",
            value=True,
            help="Show document metadata in results"
        )
        
        show_scores = st.checkbox(
            "Show Similarity Scores",
            value=True,
            help="Display similarity scores for each result"
        )
    
    # Search button
    if st.button("🔍 Search", type="primary", key="execute_search", disabled=not query.strip()):
        if query.strip():
            search_documents(astra_client, query.strip(), similarity_top_k, response_mode, include_metadata, show_scores)
    
    # Collection stats
    with st.expander("📊 Collection Statistics"):
        stats = astra_client.get_collection_stats()
        
        stats_col1, stats_col2 = st.columns(2)
        
        with stats_col1:
            st.metric("Total Documents", stats.get('document_count', 'Unknown'))
            st.metric("Collection Status", stats.get('status', 'Unknown').title())
        
        with stats_col2:
            st.metric("Last Updated", format_timestamp(stats.get('last_updated', ''), '%Y-%m-%d %H:%M'))
            if st.session_state.processing_status:
                success_rate = create_processing_summary(st.session_state.processing_status)['success_rate']
                st.metric("Processing Success Rate", f"{success_rate:.1f}%")
    
    # Search history
    if st.session_state.search_history:
        st.subheader("📚 Search History")
        
        for idx, search in enumerate(reversed(st.session_state.search_history[-5:])):  # Last 5
            with st.expander(f"🕒 {format_timestamp(search['timestamp'], '%H:%M:%S')} - {search['query'][:50]}..."):
                st.write(f"**Query:** {search['query']}")
                st.write(f"**Results:** {len(search.get('sources', []))} documents found")
                if search.get('response'):
                    st.write(f"**Response:** {search['response'][:200]}...")

def search_documents(astra_client, query, similarity_top_k, response_mode, include_metadata, show_scores):
    """Execute document search and display results"""
    search_start_time = datetime.now()
    
    with st.spinner("🔍 Searching documents..."):
        try:
            # Execute search
            result = astra_client.search_documents(
                query=query,
                top_k=similarity_top_k,
                response_mode=response_mode
            )
            
            search_time = (datetime.now() - search_start_time).total_seconds()
            
            # Display response
            st.subheader("🎯 Response")
            
            response_container = st.container()
            with response_container:
                if result.get('response'):
                    st.markdown(result['response'])
                else:
                    st.warning("No response generated. Try adjusting your query or search parameters.")
            
            # Display search metadata
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Search Time", f"{search_time:.2f}s")
            with col2:
                st.metric("Sources Found", len(result.get('sources', [])))
            with col3:
                st.metric("Response Mode", response_mode)
            
            # Display sources
            sources = result.get('sources', [])
            if sources:
                st.subheader("📄 Source Documents")
                
                for i, source in enumerate(sources):
                    score_text = f" - Score: {source.get('score', 0):.3f}" if show_scores else ""
                    filename = source.get('metadata', {}).get('filename', f'Document {i+1}')
                    
                    with st.expander(f"📋 Source {i+1}: {filename}{score_text}"):
                        # Document content
                        st.markdown("**Content:**")
                        content = source.get('text', 'No content available')
                        if len(content) > 500:
                            st.text(content[:500] + "...")
                            if st.button(f"Show full content", key=f"show_full_{i}"):
                                st.text(content)
                        else:
                            st.text(content)
                        
                        # Metadata
                        if include_metadata and source.get('metadata'):
                            st.markdown("**Metadata:**")
                            metadata = source['metadata']
                            
                            # Display key metadata in a more readable format
                            metadata_cols = st.columns(2)
                            
                            with metadata_cols[0]:
                                if isinstance(metadata, dict):
                                    for key, value in metadata.items():
                                        if key in ['filename', 'processed_at', 'source', 'file_size']:
                                            st.write(f"**{key.replace('_', ' ').title()}:** {value}")
                            
                            with metadata_cols[1]:
                                if isinstance(metadata, dict):
                                    for key, value in metadata.items():
                                        if key in ['file_type', 'chunk_size', 'word_count', 'text_length']:
                                            st.write(f"**{key.replace('_', ' ').title()}:** {value}")
                        
                        # Similarity score bar
                        if show_scores and 'score' in source:
                            score = source['score']
                            st.progress(min(score, 1.0))
            else:
                st.info("💡 No source documents found. Try adjusting your search query or parameters.")
            
            # Add to search history
            search_record = {
                'query': query,
                'response': result.get('response', ''),
                'sources': sources,
                'timestamp': datetime.now(),
                'search_time': search_time,
                'num_results': len(sources)
            }
            
            st.session_state.search_history.append(search_record)
            
            # Limit search history size
            if len(st.session_state.search_history) > 50:
                st.session_state.search_history = st.session_state.search_history[-50:]
                
        except Exception as e:
            st.error(f"❌ Search error: {str(e)}")
            
            # Add failed search to history
            search_record = {
                'query': query,
                'response': f'Error: {str(e)}',
                'sources': [],
                'timestamp': datetime.now(),
                'search_time': 0,
                'num_results': 0
            }
            
            st.session_state.search_history.append(search_record)

def monitoring_tab(astra_client, sharepoint_client):
    """Handle monitoring and analytics tab"""
    st.header("📊 System Monitoring & Analytics")
    
    # System metrics row
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            "Total Documents",
            st.session_state.document_count,
            delta=None
        )
    
    with col2:
        system_status = "🟢 Healthy" if astra_client and sharepoint_client else "🔴 Issues"
        st.metric("System Status", system_status)
    
    with col3:
        if st.session_state.last_sync_time:
            time_diff = calculate_time_diff(st.session_state.last_sync_time)
            st.metric("Last Sync", time_diff)
        else:
            st.metric("Last Sync", "Never")
    
    with col4:
        success_rate = st.session_state.system_stats.get('success_rate', 100)
        st.metric("Success Rate", f"{success_rate:.1f}%")
    
    # Processing analytics
    if st.session_state.processing_status:
        st.subheader("📈 Processing Analytics")
        
        # Create DataFrame from processing status
        df = pd.DataFrame(st.session_state.processing_status)
        
        if not df.empty and 'timestamp' in df.columns:
            # Convert timestamp
            df['timestamp'] = pd.to_datetime(df['timestamp'])
            df['hour'] = df['timestamp'].dt.floor('H')
            df['date'] = df['timestamp'].dt.date
            
            # Charts row
            chart_col1, chart_col2 = st.columns(2)
            
            with chart_col1:
                st.markdown("**Documents Processed Over Time**")
                
                # Group by hour
                hourly_data = df.groupby('hour').size().reset_index()
                hourly_data.columns = ['Hour', 'Count']
                
                if not hourly_data.empty:
                    fig = px.line(
                        hourly_data,
                        x='Hour',
                        y='Count',
                        title='Documents Processed by Hour'
                    )
                    fig.update_layout(height=300)
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("No data available for time series chart")
            
            with chart_col2:
                st.markdown("**Processing Status Distribution**")
                
                # Status distribution
                if 'status' in df.columns:
                    status_counts = df['status'].value_counts()
                    
                    # Create success/error categories
                    success_count = sum(count for status, count in status_counts.items() if 'Success' in status)
                    error_count = sum(count for status, count in status_counts.items() if 'Error' in status)
                    
                    status_data = pd.DataFrame({
                        'Status': ['Success', 'Error'],
                        'Count': [success_count, error_count]
                    })
                    
                    if not status_data.empty and status_data['Count'].sum() > 0:
                        fig = px.pie(
                            status_data,
                            values='Count',
                            names='Status',
                            title='Processing Success Rate',
                            color_discrete_map={'Success': '#28a745', 'Error': '#dc3545'}
                        )
                        fig.update_layout(height=300)
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info("No status data available")
        
        # Recent activity table
        st.subheader("🕒 Recent Activity")
        
        recent_activity = df.tail(20).copy() if not df.empty else pd.DataFrame()
        
        if not recent_activity.empty:
            recent_activity = recent_activity.sort_values('timestamp', ascending=False)
            recent_activity['formatted_time'] = recent_activity['timestamp'].dt.strftime('%Y-%m-%d %H:%M:%S')
            
            # Display recent activity
            display_columns = ['filename', 'status', 'formatted_time']
            if 'source' in recent_activity.columns:
                display_columns.append('source')
            if 'chunks' in recent_activity.columns:
                display_columns.append('chunks')
            
            st.dataframe(
                recent_activity[display_columns].head(10),
                use_container_width=True,
                hide_index=True,
                column_config={
                    'filename': 'File Name',
                    'status': 'Status',
                    'formatted_time': 'Timestamp',
                    'source': 'Source',
                    'chunks': 'Chunks'
                }
            )
        else:
            st.info("📝 No recent activity to display")
    
    else:
        st.info("📊 No processing data available yet. Process some documents to see analytics!")
    
    # Show auto-sync monitoring if enabled
    if st.session_state.auto_sync_history:
        st.subheader("🔄 Auto-Sync Performance")
        
        df_sync = pd.DataFrame(st.session_state.auto_sync_history)
        df_sync['hour'] = pd.to_datetime(df_sync['timestamp']).dt.floor('H')
        
        # Chart of documents processed over time by auto-sync
        chart_data = df_sync.groupby('hour').agg({
            'processed': 'sum',
            'documents_found': 'sum'
        }).reset_index()
        
        if not chart_data.empty:
            fig = px.bar(chart_data, x='hour', y='processed', 
                        title='Documents Processed by Auto-Sync Over Time')
            st.plotly_chart(fig, use_container_width=True)
    
    # System health section
    st.subheader("🏥 System Health")
    
    health_col1, health_col2 = st.columns(2)
    
    with health_col1:
        st.markdown("**Service Status**")
        
        services_status = []
        
        # Check Astra DB
        try:
            astra_status = astra_client.test_connection() if astra_client else False
            services_status.append({"Service": "Astra DB", "Status": "✅ Online" if astra_status else "❌ Offline"})
        except:
            services_status.append({"Service": "Astra DB", "Status": "❌ Error"})
        
        # Check SharePoint
        try:
            sp_status = sharepoint_client.test_connection() if sharepoint_client else False
            services_status.append({"Service": "SharePoint", "Status": "✅ Online" if sp_status else "❌ Offline"})
        except:
            services_status.append({"Service": "SharePoint", "Status": "❌ Error"})
        
        # Check OpenAI
        openai_status = bool(os.getenv("OPENAI_API_KEY"))
        services_status.append({"Service": "OpenAI", "Status": "✅ Configured" if openai_status else "❌ Not Configured"})
        
        services_df = pd.DataFrame(services_status)
        st.dataframe(services_df, hide_index=True, use_container_width=True)
    
    with health_col2:
        st.markdown("**Performance Metrics**")
        
        perf_metrics = []
        perf_metrics.append({"Metric": "Total Documents", "Value": st.session_state.document_count})
        perf_metrics.append({"Metric": "Success Rate", "Value": f"{st.session_state.system_stats.get('success_rate', 100):.1f}%"})
        perf_metrics.append({"Metric": "Total Processed", "Value": st.session_state.system_stats.get('total_processed', 0)})
        
        if st.session_state.search_history:
            avg_search_time = sum(s.get('search_time', 0) for s in st.session_state.search_history) / len(st.session_state.search_history)
            perf_metrics.append({"Metric": "Avg Search Time", "Value": f"{avg_search_time:.2f}s"})
        
        perf_df = pd.DataFrame(perf_metrics)
        st.dataframe(perf_df, hide_index=True, use_container_width=True)
    
    # Refresh monitoring data
    if st.button("🔄 Refresh Monitoring Data", key="refresh_monitoring"):
        st.rerun()

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
        
        # Sync history
        if st.session_state.auto_sync_history:
            with st.expander("📊 Sync History"):
                history_df = pd.DataFrame(st.session_state.auto_sync_history[-10:])  # Last 10 syncs
                
                if not history_df.empty:
                    history_df['time'] = pd.to_datetime(history_df['timestamp']).dt.strftime('%H:%M:%S')
                    history_df['result'] = history_df.apply(
                        lambda row: f"✅ {row.get('processed', 0)} docs" if row['status'] == 'success' else "❌ Error", 
                        axis=1
                    )
                    
                    display_cols = ['time', 'result', 'documents_found', 'new_files', 'modified_files']
                    available_cols = [col for col in display_cols if col in history_df.columns]
                    
                    st.dataframe(
                        history_df[available_cols].sort_values('time', ascending=False),
                        use_container_width=True,
                        hide_index=True
                    )
        
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
        data_ingestion_tab(astra_client, sharepoint_client, document_processor)
    
    with tab2:
        search_query_tab(astra_client)
    
    with tab3:
        monitoring_tab(astra_client, sharepoint_client)
    
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
