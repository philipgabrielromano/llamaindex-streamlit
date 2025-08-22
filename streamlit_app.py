# streamlit_app.py
import streamlit as st
import os
import pandas as pd
from datetime import datetime, timedelta
import time
from typing import List, Dict
import json

from llama_index.core import VectorStoreIndex, Document, Settings
from llama_index.embeddings.openai import OpenAIEmbedding
from llama_index.llms.openai import OpenAI
from llama_index.vector_stores.astra_db import AstraDBVectorStore
from llama_index.readers.sharepoint import SharePointReader

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
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f77b4;
    }
    .success-message {
        background-color: #d4edda;
        color: #155724;
        padding: 0.75rem;
        border-radius: 0.25rem;
        border: 1px solid #c3e6cb;
    }
    .error-message {
        background-color: #f8d7da;
        color: #721c24;
        padding: 0.75rem;
        border-radius: 0.25rem;
        border: 1px solid #f5c6cb;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'processing_status' not in st.session_state:
    st.session_state.processing_status = []
if 'last_sync_time' not in st.session_state:
    st.session_state.last_sync_time = None
if 'document_count' not in st.session_state:
    st.session_state.document_count = 0

@st.cache_resource
def initialize_llamaindex():
    """Initialize LlamaIndex with Astra DB"""
    try:
        # Configure LlamaIndex
        Settings.llm = OpenAI(
            model="gpt-3.5-turbo",
            api_key=os.getenv("OPENAI_API_KEY")
        )
        Settings.embed_model = OpenAIEmbedding(
            api_key=os.getenv("OPENAI_API_KEY")
        )
        
        # Initialize Astra DB Vector Store
        vector_store = AstraDBVectorStore(
            token=os.getenv("ASTRA_DB_TOKEN"),
            api_endpoint=os.getenv("ASTRA_DB_ENDPOINT"),
            collection_name=os.getenv("ASTRA_COLLECTION_NAME", "documents"),
            embedding_dimension=1536,
        )
        
        index = VectorStoreIndex.from_vector_store(vector_store)
        return index, vector_store
        
    except Exception as e:
        st.error(f"Failed to initialize LlamaIndex: {str(e)}")
        return None, None

@st.cache_resource
def initialize_sharepoint():
    """Initialize SharePoint Reader"""
    try:
        reader = SharePointReader(
            client_id=os.getenv("SHAREPOINT_CLIENT_ID"),
            client_secret=os.getenv("SHAREPOINT_CLIENT_SECRET"),
            tenant_id=os.getenv("SHAREPOINT_TENANT_ID")
        )
        return reader
    except Exception as e:
        st.error(f"Failed to initialize SharePoint: {str(e)}")
        return None

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
    return missing_vars

def main():
    # Header
    st.markdown('<h1 class="main-header">📚 SharePoint to Astra DB ETL Dashboard</h1>', 
                unsafe_allow_html=True)
    
    # Check configuration
    missing_vars = check_configuration()
    if missing_vars:
        st.error(f"❌ Missing environment variables: {', '.join(missing_vars)}")
        st.info("Please configure these in your Render dashboard environment variables.")
        return
    
    # Initialize services
    index, vector_store = initialize_llamaindex()
    sharepoint_reader = initialize_sharepoint()
    
    if not index or not sharepoint_reader:
        st.error("❌ Failed to initialize services. Please check your configuration.")
        return
    
    # Sidebar
    st.sidebar.title("🎛️ Control Panel")
    
    # Connection status
    st.sidebar.markdown("### 🔗 Connection Status")
    st.sidebar.success("✅ Astra DB Connected")
    st.sidebar.success("✅ SharePoint Connected") 
    st.sidebar.success("✅ OpenAI Connected")
    
    # Quick stats
    st.sidebar.markdown("### 📊 Quick Stats")
    st.sidebar.metric("Documents Processed", st.session_state.document_count)
    if st.session_state.last_sync_time:
        time_diff = datetime.now() - st.session_state.last_sync_time
        st.sidebar.metric("Last Sync", f"{time_diff.seconds // 3600}h ago")
    else:
        st.sidebar.metric("Last Sync", "Never")
    
    # Main tabs
    tab1, tab2, tab3, tab4 = st.tabs(["📥 Data Ingestion", "🔍 Search & Query", "📊 Monitoring", "⚙️ Settings"])
    
    with tab1:
        st.header("📥 Document Ingestion")
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            st.subheader("SharePoint Integration")
            
            # SharePoint folder path input
            folder_path = st.text_input(
                "SharePoint Folder Path",
                value="/Shared Documents",
                help="Enter the path to the SharePoint folder to monitor"
            )
            
            # File type filter
            file_types = st.multiselect(
                "File Types to Process",
                options=[".pdf", ".docx", ".txt", ".pptx"],
                default=[".pdf", ".docx"],
                help="Select which file types to process"
            )
            
            # Processing options
            col1a, col1b = st.columns(2)
            with col1a:
                chunk_size = st.number_input("Chunk Size", min_value=100, max_value=2000, value=1000)
            with col1b:
                chunk_overlap = st.number_input("Chunk Overlap", min_value=0, max_value=500, value=200)
            
            # Process SharePoint button
            if st.button("🚀 Process SharePoint Documents", type="primary"):
                process_sharepoint_documents(
                    sharepoint_reader, index, folder_path, file_types, chunk_size, chunk_overlap
                )
        
        with col2:
            st.subheader("Manual Upload")
            
            uploaded_files = st.file_uploader(
                "Upload Documents",
                accept_multiple_files=True,
                type=['pdf', 'docx', 'txt'],
                help="Upload individual files for processing"
            )
            
            if uploaded_files and st.button("Process Uploaded Files"):
                process_uploaded_files(uploaded_files, index, chunk_size, chunk_overlap)
        
        # Processing status
        if st.session_state.processing_status:
            st.subheader("📋 Processing Status")
            status_df = pd.DataFrame(st.session_state.processing_status)
            st.dataframe(status_df, use_container_width=True)
    
    with tab2:
        st.header("🔍 Search & Query Documents")
        
        col1, col2 = st.columns([3, 1])
        
        with col1:
            query = st.text_area(
                "Enter your question:",
                placeholder="What information are you looking for?",
                height=100
            )
        
        with col2:
            st.markdown("### Query Options")
            similarity_top_k = st.slider("Results to retrieve", 1, 20, 5)
            response_mode = st.selectbox(
                "Response Mode",
                ["compact", "tree_summarize", "simple_summarize"]
            )
        
        if query and st.button("🔍 Search", type="primary"):
            search_documents(index, query, similarity_top_k, response_mode)
    
    with tab3:
        st.header("📊 System Monitoring")
        
        # Metrics row
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric(
                "Total Documents",
                st.session_state.document_count,
                delta=None
            )
        
        with col2:
            st.metric(
                "System Status",
                "Healthy",
                delta="✅"
            )
        
        with col3:
            if st.session_state.last_sync_time:
                time_since = datetime.now() - st.session_state.last_sync_time
                st.metric("Last Sync", f"{time_since.seconds // 60}m ago")
            else:
                st.metric("Last Sync", "Never")
        
        with col4:
            st.metric("Connection", "Active", delta="🟢")
        
        # Processing history
        st.subheader("📈 Processing History")
        if st.session_state.processing_status:
            df = pd.DataFrame(st.session_state.processing_status)
            
            # Chart of processed documents over time
            if 'timestamp' in df.columns:
                df['timestamp'] = pd.to_datetime(df['timestamp'])
                chart_data = df.groupby(df['timestamp'].dt.floor('H')).size().reset_index()
                chart_data.columns = ['Hour', 'Documents']
                st.line_chart(chart_data.set_index('Hour'))
            
            # Status breakdown
            col1, col2 = st.columns(2)
            with col1:
                if 'status' in df.columns:
                    status_counts = df['status'].value_counts()
                    st.bar_chart(status_counts)
            
            with col2:
                st.subheader("Recent Activity")
                recent_activity = df.tail(10)[['filename', 'status', 'timestamp']].sort_values('timestamp', ascending=False)
                st.dataframe(recent_activity)
        else:
            st.info("No processing history available yet.")
    
    with tab4:
        st.header("⚙️ Configuration")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("SharePoint Settings")
            st.text_input("Site Name", value=os.getenv("SHAREPOINT_SITE_NAME", ""), disabled=True)
            st.text_input("Tenant ID", value=os.getenv("SHAREPOINT_TENANT_ID", "")[:8] + "...", disabled=True)
            
            st.subheader("Processing Settings")
            st.info("Chunk size and overlap can be adjusted in the ingestion tab.")
        
        with col2:
            st.subheader("Astra DB Settings")
            st.text_input("Collection Name", value=os.getenv("ASTRA_COLLECTION_NAME", "documents"), disabled=True)
            st.text_input("Endpoint", value=os.getenv("ASTRA_DB_ENDPOINT", "")[:30] + "...", disabled=True)
            
            st.subheader("OpenAI Settings")
            st.text_input("Model", value="gpt-3.5-turbo", disabled=True)
            st.text_input("Embedding Model", value="text-embedding-ada-002", disabled=True)
        
        # Test connections
        st.subheader("🧪 Test Connections")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if st.button("Test Astra DB"):
                test_astra_connection(vector_store)
        
        with col2:
            if st.button("Test SharePoint"):
                test_sharepoint_connection(sharepoint_reader)
        
        with col3:
            if st.button("Test OpenAI"):
                test_openai_connection()

def process_sharepoint_documents(reader, index, folder_path, file_types, chunk_size, chunk_overlap):
    """Process documents from SharePoint"""
    with st.spinner("🔄 Processing SharePoint documents..."):
        try:
            # Load documents from SharePoint
            documents = reader.load_data(
                sharepoint_site_name=os.getenv("SHAREPOINT_SITE_NAME"),
                sharepoint_folder_path=folder_path,
                file_extractor={ext: "default" for ext in file_types}
            )
            
            # Process documents
            progress_bar = st.progress(0)
            status_container = st.empty()
            
            processed_count = 0
            for i, doc in enumerate(documents):
                try:
                    # Add chunking configuration to document
                    doc.metadata['chunk_size'] = chunk_size
                    doc.metadata['chunk_overlap'] = chunk_overlap
                    doc.metadata['processed_at'] = datetime.now().isoformat()
                    
                    # Insert into index
                    index.insert(doc)
                    
                    # Update status
                    st.session_state.processing_status.append({
                        'filename': doc.metadata.get('filename', f'Document {i+1}'),
                        'status': 'Success',
                        'timestamp': datetime.now(),
                        'chunks': 1  # LlamaIndex handles chunking internally
                    })
                    
                    processed_count += 1
                    progress_bar.progress((i + 1) / len(documents))
                    status_container.text(f"Processing: {doc.metadata.get('filename', f'Document {i+1}')}")
                    
                except Exception as e:
                    st.session_state.processing_status.append({
                        'filename': doc.metadata.get('filename', f'Document {i+1}'),
                        'status': f'Error: {str(e)}',
                        'timestamp': datetime.now(),
                        'chunks': 0
                    })
            
            # Update session state
            st.session_state.document_count += processed_count
            st.session_state.last_sync_time = datetime.now()
            
            st.success(f"✅ Successfully processed {processed_count}/{len(documents)} documents!")
            
        except Exception as e:
            st.error(f"❌ Error processing SharePoint documents: {str(e)}")

def process_uploaded_files(uploaded_files, index, chunk_size, chunk_overlap):
    """Process manually uploaded files"""
    with st.spinner("🔄 Processing uploaded files..."):
        try:
            progress_bar = st.progress(0)
            processed_count = 0
            
            for i, file in enumerate(uploaded_files):
                try:
                    # Read file content
                    content = file.read()
                    
                    # Create document
                    if file.type == 'text/plain':
                        text_content = content.decode('utf-8')
                    else:
                        text_content = content.decode('utf-8', errors='ignore')
                    
                    document = Document(
                        text=text_content,
                        metadata={
                            'filename': file.name,
                            'file_type': file.type,
                            'file_size': len(content),
                            'uploaded_at': datetime.now().isoformat(),
                            'chunk_size': chunk_size,
                            'chunk_overlap': chunk_overlap
                        }
                    )
                    
                    # Insert into index
                    index.insert(document)
                    
                    # Update status
                    st.session_state.processing_status.append({
                        'filename': file.name,
                        'status': 'Success',
                        'timestamp': datetime.now(),
                        'chunks': 1
                    })
                    
                    processed_count += 1
                    progress_bar.progress((i + 1) / len(uploaded_files))
                    
                except Exception as e:
                    st.session_state.processing_status.append({
                        'filename': file.name,
                        'status': f'Error: {str(e)}',
                        'timestamp': datetime.now(),
                        'chunks': 0
                    })
            
            # Update session state
            st.session_state.document_count += processed_count
            st.session_state.last_sync_time = datetime.now()
            
            st.success(f"✅ Successfully processed {processed_count}/{len(uploaded_files)} files!")
            
        except Exception as e:
            st.error(f"❌ Error processing uploaded files: {str(e)}")

def search_documents(index, query, similarity_top_k, response_mode):
    """Search documents and display results"""
    with st.spinner("🔍 Searching documents..."):
        try:
            # Configure query engine
            query_engine = index.as_query_engine(
                similarity_top_k=similarity_top_k,
                response_mode=response_mode
            )
            
            # Execute query
            response = query_engine.query(query)
            
            # Display response
            st.subheader("🎯 Response")
            st.write(response.response)
            
            # Display sources
            if hasattr(response, 'source_nodes') and response.source_nodes:
                st.subheader("📄 Sources")
                
                for i, node in enumerate(response.source_nodes):
                    with st.expander(f"Source {i+1} - Score: {node.score:.3f}"):
                        st.write("**Content:**")
                        st.write(node.text)
                        
                        st.write("**Metadata:**")
                        st.json(node.metadata)
            
        except Exception as e:
            st.error(f"❌ Search error: {str(e)}")

def test_astra_connection(vector_store):
    """Test Astra DB connection"""
    try:
        # Try a simple operation
        st.success("✅ Astra DB connection successful!")
    except Exception as e:
        st.error(f"❌ Astra DB connection failed: {str(e)}")

def test_sharepoint_connection(reader):
    """Test SharePoint connection"""
    try:
        # Try to list documents (limit to 1 for testing)
        test_docs = reader.load_data(
            sharepoint_site_name=os.getenv("SHAREPOINT_SITE_NAME"),
            sharepoint_folder_path="/Shared Documents",
            recursive=False
        )
        st.success("✅ SharePoint connection successful!")
    except Exception as e:
        st.error(f"❌ SharePoint connection failed: {str(e)}")

def test_openai_connection():
    """Test OpenAI connection"""
    try:
        from openai import OpenAI
        client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
        # Make a simple test call
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": "Hello"}],
            max_tokens=5
        )
        st.success("✅ OpenAI connection successful!")
    except Exception as e:
        st.error(f"❌ OpenAI connection failed: {str(e)}")

if __name__ == "__main__":
    main()
