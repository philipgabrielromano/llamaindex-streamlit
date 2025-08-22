# streamlit_app.py
import streamlit as st
import os
import pandas as pd
from datetime import datetime, timedelta
import time
from typing import List, Dict, Optional
import json
import plotly.express as px
import plotly.graph_objects as go

from llama_index.core import VectorStoreIndex, Document, Settings
from llama_index.embeddings.openai import OpenAIEmbedding
from llama_index.llms.openai import OpenAI

# Import utilities
from utils import DocumentProcessor, SharePointClient, AstraClient
from utils import format_timestamp, calculate_time_diff, validate_file_type
from utils import create_processing_summary, progress_tracker, format_file_size

# Page configuration
st.set_page_config(
    page_title="SharePoint ETL Dashboard",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

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
    .success-message {
        background-color: #d4edda;
        color: #155724;
        padding: 1rem;
        border-radius: 8px;
        border: 1px solid #c3e6cb;
        margin: 1rem 0;
    }
    .error-message {
        background-color: #f8d7da;
        color: #721c24;
        padding: 1rem;
        border-radius: 8px;
        border: 1px solid #f5c6cb;
        margin: 1rem 0;
    }
    .info-box {
        background-color: #e7f3ff;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #1f77b4;
        margin: 1rem 0;
    }
    .sidebar .metric-container {
        background-color: #f8f9fa;
        padding: 10px;
        border-radius: 5px;
        margin: 5px 0;
    }
    .stButton > button {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.5rem 1rem;
        font-weight: bold;
    }
    .processing-status {
        font-family: 'Courier New', monospace;
        background-color: #f1f1f1;
        padding: 10px;
        border-radius: 5px;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
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

@st.cache_resource
def initialize_services():
    """Initialize all services and return clients"""
    try:
        # Initialize LlamaIndex settings
        Settings.llm = OpenAI(
            model="gpt-3.5-turbo",
            api_key=os.getenv("OPENAI_API_KEY")
        )
        Settings.embed_model = OpenAIEmbedding(
            api_key=os.getenv("OPENAI_API_KEY")
        )
        
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

def display_sidebar(astra_client, sharepoint_client):
    """Display sidebar with system status and controls"""
    st.sidebar.title("🎛️ Control Panel")
    
    # Connection status
    st.sidebar.markdown("### 🔗 Connection Status")
    
    # Test connections
    col1, col2 = st.sidebar.columns(2)
    
    with col1:
        if st.button("Test All", key="test_all"):
            test_all_connections(astra_client, sharepoint_client)
    
    with col2:
        if st.button("Refresh", key="refresh"):
            st.rerun()
    
    # Display connection status
    astra_status = "✅ Connected" if astra_client and astra_client.test_connection() else "❌ Disconnected"
    sharepoint_status = "✅ Connected" if sharepoint_client and sharepoint_client.test_connection() else "❌ Disconnected"
    openai_status = "✅ Connected" if os.getenv("OPENAI_API_KEY") else "❌ Not configured"
    
    st.sidebar.success(f"Astra DB: {astra_status}")
    st.sidebar.success(f"SharePoint: {sharepoint_status}")
    st.sidebar.success(f"OpenAI: {openai_status}")
    
    # Quick stats
    st.sidebar.markdown("### 📊 Quick Stats")
    
    stats_container = st.sidebar.container()
    with stats_container:
        col1, col2 = st.columns(2)
        
        with col1:
            st.metric(
                "Documents", 
                st.session_state.document_count,
                delta=None
            )
        
        with col2:
            if st.session_state.last_sync_time:
                time_diff = calculate_time_diff(st.session_state.last_sync_time)
                st.metric("Last Sync", time_diff)
            else:
                st.metric("Last Sync", "Never")
        
        # Success rate
        success_rate = st.session_state.system_stats.get('success_rate', 100)
        st.metric("Success Rate", f"{success_rate:.1f}%")
        
        # Processing status summary
        if st.session_state.processing_status:
            summary = create_processing_summary(st.session_state.processing_status)
            st.metric("Total Processed", summary['total'])

def test_all_connections(astra_client, sharepoint_client):
    """Test all service connections"""
    with st.spinner("Testing connections..."):
        results = {}
        
        # Test Astra DB
        try:
            results['astra'] = astra_client.test_connection() if astra_client else False
        except Exception as e:
            results['astra'] = False
            st.error(f"Astra DB test failed: {str(e)}")
        
        # Test SharePoint
        try:
            results['sharepoint'] = sharepoint_client.test_connection() if sharepoint_client else False
        except Exception as e:
            results['sharepoint'] = False
            st.error(f"SharePoint test failed: {str(e)}")
        
        # Test OpenAI
        try:
            from openai import OpenAI
            client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
            response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content": "Hello"}],
                max_tokens=5
            )
            results['openai'] = True
        except Exception as e:
            results['openai'] = False
            st.error(f"OpenAI test failed: {str(e)}")
        
        # Display results
        if all(results.values()):
            st.success("✅ All connections successful!")
        else:
            failed = [service for service, status in results.items() if not status]
            st.warning(f"⚠️ Failed connections: {', '.join(failed)}")

def data_ingestion_tab(astra_client, sharepoint_client, document_processor):
    """Handle data ingestion tab"""
    st.header("📥 Document Ingestion")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("🔗 SharePoint Integration")
        
        # SharePoint configuration
        with st.expander("📁 SharePoint Configuration", expanded=True):
            col1a, col1b = st.columns(2)
            
            with col1a:
                folder_path = st.text_input(
                    "Folder Path",
                    value="/Shared Documents",
                    help="Enter the SharePoint folder path to monitor",
                    key="sp_folder_path"
                )
                
                sync_mode = st.selectbox(
                    "Sync Mode",
                    ["All Files", "Recent Changes", "Specific Date Range"],
                    help="Choose how to sync documents"
                )
            
            with col1b:
                file_types = st.multiselect(
                    "File Types",
                    options=[".pdf", ".docx", ".txt", ".pptx", ".xlsx"],
                    default=[".pdf", ".docx"],
                    help="Select file types to process"
                )
                
                if sync_mode == "Recent Changes":
                    hours_back = st.number_input(
                        "Hours Back",
                        min_value=1,
                        max_value=168,
                        value=24,
                        help="Look for changes in the last N hours"
                    )
                elif sync_mode == "Specific Date Range":
                    date_range = st.date_input(
                        "Date Range",
                        value=[datetime.now().date() - timedelta(days=7), datetime.now().date()],
                        help="Select date range for document sync"
                    )
        
        # Processing options
        with st.expander("⚙️ Processing Options"):
            col1a, col1b, col1c = st.columns(3)
            
            with col1a:
                chunk_size = st.number_input(
                    "Chunk Size",
                    min_value=100,
                    max_value=2000,
                    value=1000,
                    help="Size of text chunks for processing"
                )
            
            with col1b:
                chunk_overlap = st.number_input(
                    "Chunk Overlap",
                    min_value=0,
                    max_value=500,
                    value=200,
                    help="Overlap between chunks"
                )
            
            with col1c:
                batch_size = st.number_input(
                    "Batch Size",
                    min_value=1,
                    max_value=50,
                    value=10,
                    help="Number of documents to process at once"
                )
        
        # Process button
        if st.button("🚀 Process SharePoint Documents", type="primary", key="process_sp"):
            process_sharepoint_documents(
                astra_client, sharepoint_client, document_processor,
                folder_path, file_types, chunk_size, chunk_overlap, batch_size, sync_mode
            )
    
    with col2:
        st.subheader("📤 Manual Upload")
        
        uploaded_files = st.file_uploader(
            "Upload Documents",
            accept_multiple_files=True,
            type=['pdf', 'docx', 'txt'],
            help="Upload individual files for processing",
            key="manual_upload"
        )
        
        if uploaded_files:
            st.write(f"**Selected Files:** {len(uploaded_files)}")
            
            # Show file details
            for file in uploaded_files[:3]:  # Show first 3 files
                st.write(f"• {file.name} ({format_file_size(file.size)})")
            
            if len(uploaded_files) > 3:
                st.write(f"• ... and {len(uploaded_files) - 3} more files")
            
            # Processing options for upload
            upload_chunk_size = st.number_input(
                "Chunk Size",
                min_value=100,
                max_value=2000,
                value=1000,
                key="upload_chunk_size"
            )
            
            upload_chunk_overlap = st.number_input(
                "Chunk Overlap",
                min_value=0,
                max_value=500,
                value=200,
                key="upload_chunk_overlap"
            )
            
            if st.button("📄 Process Uploaded Files", key="process_upload"):
                process_uploaded_files(
                    uploaded_files, astra_client, document_processor,
                    upload_chunk_size, upload_chunk_overlap
                )
    
    # Processing status display
    if st.session_state.processing_status:
        st.subheader("📋 Processing Status")
        
        # Summary metrics
        summary = create_processing_summary(st.session_state.processing_status[-20:])  # Last 20 items
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Processed", summary['total'])
        with col2:
            st.metric("Successful", summary['successful'])
        with col3:
            st.metric("Failed", summary['failed'])
        with col4:
            st.metric("Success Rate", f"{summary['success_rate']:.1f}%")
        
        # Detailed status table
        status_df = pd.DataFrame(st.session_state.processing_status[-10:])  # Last 10 items
        if not status_df.empty:
            status_df['timestamp'] = pd.to_datetime(status_df['timestamp'])
            status_df['formatted_time'] = status_df['timestamp'].dt.strftime('%H:%M:%S')
            
            # Display table
            display_columns = ['filename', 'status', 'formatted_time']
            if 'chunks' in status_df.columns:
                display_columns.append('chunks')
            
            st.dataframe(
                status_df[display_columns],
                use_container_width=True,
                hide_index=True
            )
            
            # Clear status button
            if st.button("🗑️ Clear Status", key="clear_status"):
                st.session_state.processing_status = []
                st.rerun()

def process_sharepoint_documents(astra_client, sharepoint_client, document_processor,
                               folder_path, file_types, chunk_size, chunk_overlap, 
                               batch_size, sync_mode):
    """Process documents from SharePoint"""
    start_time = datetime.now()
    
    with st.spinner("🔄 Connecting to SharePoint..."):
        try:
            # Update document processor settings
            document_processor.chunk_size = chunk_size
            document_processor.chunk_overlap = chunk_overlap
            
            # Get documents based on sync mode
            if sync_mode == "Recent Changes":
                hours_back = st.session_state.get('hours_back', 24)
                since_date = datetime.now() - timedelta(hours=hours_back)
                raw_docs = sharepoint_client.get_documents(
                    folder_path=folder_path,
                    file_types=file_types,
                    since_date=since_date
                )
            else:
                raw_docs = sharepoint_client.get_documents(
                    folder_path=folder_path,
                    file_types=file_types
                )
            
            if not raw_docs:
                st.warning("⚠️ No documents found matching the criteria.")
                return
            
            st.info(f"📊 Found {len(raw_docs)} documents to process")
            
            # Process documents
            with progress_tracker(raw_docs, "Processing SharePoint documents") as tracker:
                processed_docs = []
                processing_results = []
                
                for doc_info in raw_docs:
                    try:
                        # Create document
                        document = Document(
                            text=doc_info.get('content', ''),
                            metadata={
                                **doc_info.get('metadata', {}),
                                'filename': doc_info.get('filename', 'Unknown'),
                                'processed_at': datetime.now().isoformat(),
                                'chunk_size': chunk_size,
                                'chunk_overlap': chunk_overlap,
                                'source': 'sharepoint'
                            }
                        )
                        
                        processed_docs.append(document)
                        
                        # Process in batches
                        if len(processed_docs) >= batch_size:
                            batch_results = astra_client.insert_documents(processed_docs)
                            processing_results.append(batch_results)
                            processed_docs = []
                        
                        # Update tracker
                        tracker.update(doc_info.get('filename', 'Unknown'))
                        
                        # Add to processing status
                        st.session_state.processing_status.append({
                            'filename': doc_info.get('filename', 'Unknown'),
                            'status': 'Success',
                            'timestamp': datetime.now(),
                            'chunks': 1,
                            'source': 'sharepoint'
                        })
                        
                    except Exception as e:
                        st.session_state.processing_status.append({
                            'filename': doc_info.get('filename', 'Unknown'),
                            'status': f'Error: {str(e)}',
                            'timestamp': datetime.now(),
                            'chunks': 0,
                            'source': 'sharepoint'
                        })
                
                # Process remaining documents
                if processed_docs:
                    batch_results = astra_client.insert_documents(processed_docs)
                    processing_results.append(batch_results)
            
            # Calculate results
            total_successful = sum(result.get('successful', 0) for result in processing_results)
            total_failed = sum(result.get('failed', 0) for result in processing_results)
            total_docs = len(raw_docs)
            
            # Update session state
            st.session_state.document_count += total_successful
            st.session_state.last_sync_time = datetime.now()
            
            # Update system stats
            processing_time = (datetime.now() - start_time).total_seconds()
            st.session_state.system_stats['total_processed'] += total_docs
            st.session_state.system_stats['success_rate'] = (total_successful / total_docs * 100) if total_docs > 0 else 100
            st.session_state.system_stats['avg_processing_time'] = processing_time / total_docs if total_docs > 0 else 0
            
            # Display results
            if total_successful > 0:
                st.success(f"✅ Successfully processed {total_successful}/{total_docs} documents in {processing_time:.1f} seconds!")
            
            if total_failed > 0:
                st.warning(f"⚠️ {total_failed} documents failed to process. Check the status table below for details.")
            
        except Exception as e:
            st.error(f"❌ Error processing SharePoint documents: {str(e)}")
            st.session_state.processing_status.append({
                'filename': 'SharePoint Sync',
                'status': f'Critical Error: {str(e)}',
                'timestamp': datetime.now(),
                'chunks': 0,
                'source': 'sharepoint'
            })

def process_uploaded_files(uploaded_files, astra_client, document_processor,
                          chunk_size, chunk_overlap):
    """Process manually uploaded files"""
    start_time = datetime.now()
    
    with st.spinner("🔄 Processing uploaded files..."):
        try:
            # Update document processor settings
            document_processor.chunk_size = chunk_size
            document_processor.chunk_overlap = chunk_overlap
            
            successful_count = 0
            failed_count = 0
            
            with progress_tracker(uploaded_files, "Processing uploaded files") as tracker:
                for file in uploaded_files:
                    try:
                        # Validate file type
                        if not validate_file_type(file.name):
                            st.session_state.processing_status.append({
                                'filename': file.name,
                                'status': 'Error: Unsupported file type',
                                'timestamp': datetime.now(),
                                'chunks': 0,
                                'source': 'upload'
                            })
                            failed_count += 1
                            continue
                        
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
                                    'chunks': result.get('successful', 1),
                                    'source': 'upload'
                                })
                            else:
                                failed_count += 1
                                st.session_state.processing_status.append({
                                    'filename': file.name,
                                    'status': 'Error: Failed to index',
                                    'timestamp': datetime.now(),
                                    'chunks': 0,
                                    'source': 'upload'
                                })
                        else:
                            failed_count += 1
                        
                        # Update tracker
                        tracker.update(file.name)
                        
                    except Exception as e:
                        failed_count += 1
                        st.session_state.processing_status.append({
                            'filename': file.name,
                            'status': f'Error: {str(e)}',
                            'timestamp': datetime.now(),
                            'chunks': 0,
                            'source': 'upload'
                        })
            
            # Update session state
            st.session_state.document_count += successful_count
            st.session_state.last_sync_time = datetime.now()
            
            # Update system stats
            total_files = len(uploaded_files)
            processing_time = (datetime.now() - start_time).total_seconds()
            st.session_state.system_stats['total_processed'] += total_files
            st.session_state.system_stats['success_rate'] = (successful_count / total_files * 100) if total_files > 0 else 100
            
            # Display results
            if successful_count > 0:
                st.success(f"✅ Successfully processed {successful_count}/{total_files} files!")
            
            if failed_count > 0:
                st.warning(f"⚠️ {failed_count} files failed to process. Check the status table for details.")
                
        except Exception as e:
            st.error(f"❌ Error processing uploaded files: {str(e)}")

def search_query_tab(astra_client):
    """Handle search and query tab"""
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
            if st.button("📋 List all documents", key="query_list"):
                query = "What documents are available?"
                
        with suggestion_cols[1]:
            if st.button("📊 Summary statistics", key="query_stats"):
                query = "Provide a summary of the document collection"
                
        with suggestion_cols[2]:
            if st.button("🔍 Recent uploads", key="query_recent"):
                query = "What were the most recently uploaded documents?"
    
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
    
    # Search history
    if st.session_state.search_history:
        st.subheader("📚 Search History")
        
        history_df = pd.DataFrame(st.session_state.search_history[-10:])  # Last 10 searches
        history_df['formatted_time'] = pd.to_datetime(history_df['timestamp']).dt.strftime('%H:%M:%S')
        
        # Display search history
        for idx, search in enumerate(reversed(st.session_state.search_history[-5:])):  # Last 5
            with st.expander(f"🕒 {search['formatted_time']} - {search['query'][:50]}..."):
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
                        else:
                            st.text(content)
                        
                        # Metadata
                        if include_metadata and source.get('metadata'):
                            st.markdown("**Metadata:**")
                            metadata = source['metadata']
                            
                            # Display key metadata in a more readable format
                            metadata_cols = st.columns(2)
                            
                            with metadata_cols[0]:
                                if 'filename' in metadata:
                                    st.write(f"**Filename:** {metadata['filename']}")
                                if 'processed_at' in metadata:
                                    st.write(f"**Processed:** {format_timestamp(metadata['processed_at'])}")
                                if 'source' in metadata:
                                    st.write(f"**Source:** {metadata['source']}")
                            
                            with metadata_cols[1]:
                                if 'file_size' in metadata:
                                    st.write(f"**Size:** {format_file_size(metadata['file_size'])}")
                                if 'chunk_size' in metadata:
                                    st.write(f"**Chunk Size:** {metadata['chunk_size']}")
                                if 'file_type' in metadata:
                                    st.write(f"**Type:** {metadata['file_type']}")
                        
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
                'formatted_time': datetime.now().strftime('%H:%M:%S'),
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
                'formatted_time': datetime.now().strftime('%H:%M:%S'),
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

def settings_tab(astra_client, sharepoint_client):
    """Handle settings and configuration tab"""
    st.header("⚙️ Settings & Configuration")
    
    # Configuration overview
    missing_vars, configured_vars = check_configuration()
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("🔗 Service Configuration")
        
        # Environment variables status
        st.markdown("**Environment Variables**")
        
        config_status = []
        all_vars = [
            "OPENAI_API_KEY",
            "ASTRA_DB_TOKEN", 
            "ASTRA_DB_ENDPOINT",
            "ASTRA_COLLECTION_NAME",
            "SHAREPOINT_CLIENT_ID",
            "SHAREPOINT_CLIENT_SECRET", 
            "SHAREPOINT_TENANT_ID",
            "SHAREPOINT_SITE_NAME"
        ]
        
        for var in all_vars:
            value = os.getenv(var)
            status = "✅ Set" if value else "❌ Missing"
            masked_value = f"{value[:8]}..." if value and len(value) > 8 else "Not set"
            
            config_status.append({
                "Variable": var,
                "Status": status,
                "Value": masked_value
            })
        
        config_df = pd.DataFrame(config_status)
        st.dataframe(config_df, hide_index=True, use_container_width=True)
        
        # Configuration validation
        if missing_vars:
            st.error(f"❌ Missing configuration: {', '.join(missing_vars)}")
            st.info("💡 Add these environment variables in your Render dashboard.")
        else:
            st.success("✅ All required configuration variables are set!")
    
    with col2:
        st.subheader("🧪 Connection Testing")
        
        # Test individual connections
        st.markdown("**Service Connection Tests**")
        
        test_col1, test_col2 = st.columns(2)
        
        with test_col1:
            if st.button("Test Astra DB", key="test_astra_settings"):
                with st.spinner("Testing Astra DB..."):
                    try:
                        if astra_client and astra_client.test_connection():
                            st.success("✅ Astra DB connection successful!")
                            
                            # Get collection stats
                            stats = astra_client.get_collection_stats()
                            st.write(f"Collection: {stats.get('collection_name', 'N/A')}")
                            st.write(f"Status: {stats.get('status', 'N/A')}")
                        else:
                            st.error("❌ Astra DB connection failed!")
                    except Exception as e:
                        st.error(f"❌ Astra DB error: {str(e)}")
            
            if st.button("Test SharePoint", key="test_sp_settings"):
                with st.spinner("Testing SharePoint..."):
                    try:
                        if sharepoint_client and sharepoint_client.test_connection():
                            st.success("✅ SharePoint connection successful!")
                            
                            # Get SharePoint info
                            config = sharepoint_client.validate_configuration()
                            st.write(f"Site: {os.getenv('SHAREPOINT_SITE_NAME', 'N/A')}")
                            st.write(f"Tenant: {os.getenv('SHAREPOINT_TENANT_ID', 'N/A')[:8]}...")
                        else:
                            st.error("❌ SharePoint connection failed!")
                    except Exception as e:
                        st.error(f"❌ SharePoint error: {str(e)}")
        
        with test_col2:
            if st.button("Test OpenAI", key="test_openai_settings"):
                with st.spinner("Testing OpenAI..."):
                    try:
                        from openai import OpenAI
                        client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
                        
                        response = client.chat.completions.create(
                            model="gpt-3.5-turbo",
                            messages=[{"role": "user", "content": "Hello"}],
                            max_tokens=5
                        )
                        
                        st.success("✅ OpenAI connection successful!")
                        st.write("Model: gpt-3.5-turbo")
                        st.write("Embedding: text-embedding-ada-002")
                        
                    except Exception as e:
                        st.error(f"❌ OpenAI error: {str(e)}")
            
            if st.button("Test All Services", key="test_all_settings"):
                test_all_connections(astra_client, sharepoint_client)
    
    # Application settings
    st.subheader("📱 Application Settings")
    
    app_col1, app_col2 = st.columns(2)
    
    with app_col1:
        st.markdown("**Processing Defaults**")
        
        # These would be stored in session state or config
        default_chunk_size = st.number_input(
            "Default Chunk Size",
            min_value=100,
            max_value=2000,
            value=1000,
            help="Default chunk size for document processing"
        )
        
        default_chunk_overlap = st.number_input(
            "Default Chunk Overlap",
            min_value=0,
            max_value=500,
            value=200,
            help="Default overlap between chunks"
        )
        
        default_batch_size = st.number_input(
            "Default Batch Size",
            min_value=1,
            max_value=50,
            value=10,
            help="Default batch size for processing"
        )
    
    with app_col2:
        st.markdown("**UI Preferences**")
        
        # UI settings
        show_debug_info = st.checkbox(
            "Show Debug Information",
            value=False,
            help="Display additional debug information"
        )
        
        auto_refresh = st.checkbox(
            "Auto-refresh Dashboard",
            value=False,
            help="Automatically refresh dashboard data"
        )
        
        if auto_refresh:
            refresh_interval = st.slider(
                "Refresh Interval (seconds)",
                min_value=30,
                max_value=300,
                value=60,
                help="How often to refresh the dashboard"
            )
    
    # Data management
    st.subheader("🗂️ Data Management")
    
    data_col1, data_col2 = st.columns(2)
    
    with data_col1:
        st.markdown("**Session Data**")
        
        st.write(f"Processing Status Records: {len(st.session_state.processing_status)}")
        st.write(f"Search History Records: {len(st.session_state.search_history)}")
        st.write(f"Document Count: {st.session_state.document_count}")
        
        if st.button("🗑️ Clear Session Data", key="clear_session"):
            st.session_state.processing_status = []
            st.session_state.search_history = []
            st.session_state.document_count = 0
            st.session_state.system_stats = {
                'total_processed': 0,
                'success_rate': 100,
                'avg_processing_time': 0
            }
            st.success("✅ Session data cleared!")
            st.rerun()
    
    with data_col2:
        st.markdown("**Export Data**")
        
        if st.session_state.processing_status:
            # Export processing status
            status_df = pd.DataFrame(st.session_state.processing_status)
            csv_data = status_df.to_csv(index=False)
            
            st.download_button(
                "📥 Download Processing Status",
                data=csv_data,
                file_name=f"processing_status_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
        
        if st.session_state.search_history:
            # Export search history
            search_df = pd.DataFrame(st.session_state.search_history)
            search_csv = search_df.to_csv(index=False)
            
            st.download_button(
                "📥 Download Search History",
                data=search_csv,
                file_name=f"search_history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )

def main():
    """Main application function"""
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
        
        # Show configuration help
        with st.expander("📋 Configuration Help"):
            st.markdown("""
            **Required Environment Variables:**
            
            - `OPENAI_API_KEY`: Your OpenAI API key
            - `ASTRA_DB_TOKEN`: Your Astra DB token
            - `ASTRA_DB_ENDPOINT`: Your Astra DB endpoint URL
            - `SHAREPOINT_CLIENT_ID`: SharePoint application client ID
            - `SHAREPOINT_CLIENT_SECRET`: SharePoint application client secret
            - `SHAREPOINT_TENANT_ID`: Your SharePoint tenant ID
            - `SHAREPOINT_SITE_NAME`: Name of your SharePoint site
            
            **Optional:**
            - `ASTRA_COLLECTION_NAME`: Collection name (default: "documents")
            """)
        
        # Still show settings tab for configuration management
        settings_tab(None, None)
        return
    
    # Initialize services
    astra_client, sharepoint_client, document_processor, services_ok = initialize_services()
    
    if not services_ok:
        st.error("❌ Failed to initialize services. Please check your configuration.")
        return
    
    # Display sidebar
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
        settings_tab(astra_client, sharepoint_client)
    
    # Footer
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: #666; padding: 20px;'>"
        "📚 SharePoint ETL Dashboard | Built with Streamlit & LlamaIndex | "
        f"Last updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        "</div>", 
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
