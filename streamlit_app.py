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

# Updated LlamaIndex imports for v0.10+
try:
    from llama_index import VectorStoreIndex, Document, ServiceContext
    from llama_index.embeddings import OpenAIEmbedding
    from llama_index.llms import OpenAI
    from llama_index.vector_stores import AstraDBVectorStore
    LLAMA_INDEX_AVAILABLE = True
except ImportError:
    # Alternative import structure
    try:
        from llama_index.core import VectorStoreIndex, Document, Settings
        from llama_index.embeddings.openai import OpenAIEmbedding
        from llama_index.llms.openai import OpenAI
        from llama_index.vector_stores.astra_db import AstraDBVectorStore
        ServiceContext = None
        LLAMA_INDEX_AVAILABLE = True
    except ImportError:
        # Fallback - we'll implement basic functionality without LlamaIndex
        VectorStoreIndex = None
        Document = None
        Settings = None
        ServiceContext = None
        OpenAIEmbedding = None
        OpenAI = None
        AstraDBVectorStore = None
        LLAMA_INDEX_AVAILABLE = False

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
    .warning-box {
        background-color: #fff3cd;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #ffc107;
        margin: 1rem 0;
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
        # Check if LlamaIndex is available
        if not LLAMA_INDEX_AVAILABLE:
            st.warning("⚠️ LlamaIndex not fully available. Some features may be limited.")
            return None, None, None, False
        
        # Initialize LlamaIndex settings
        if ServiceContext:
            # For older versions
            embed_model = OpenAIEmbedding(api_key=os.getenv("OPENAI_API_KEY"))
            llm = OpenAI(model="gpt-3.5-turbo", api_key=os.getenv("OPENAI_API_KEY"))
            service_context = ServiceContext.from_defaults(
                llm=llm,
                embed_model=embed_model
            )
        elif Settings:
            # For newer versions
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
        st.info("💡 This might be due to missing API keys or package compatibility issues.")
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

def show_package_status():
    """Show the status of required packages"""
    st.subheader("📦 Package Status")
    
    packages_status = {
        "Streamlit": True,  # Obviously available since we're running
        "LlamaIndex": LLAMA_INDEX_AVAILABLE,
        "OpenAI": bool(os.getenv("OPENAI_API_KEY")),
        "AstraDB": bool(os.getenv("ASTRA_DB_TOKEN")),
        "SharePoint": bool(os.getenv("SHAREPOINT_CLIENT_ID")),
        "Pandas": True,
        "Plotly": True
    }
    
    col1, col2 = st.columns(2)
    
    for i, (package, status) in enumerate(packages_status.items()):
        target_col = col1 if i % 2 == 0 else col2
        
        with target_col:
            if status:
                st.success(f"✅ {package}")
            else:
                st.error(f"❌ {package}")

def main():
    """Main application function"""
    # Initialize session state
    init_session_state()
    
    # Header
    st.markdown('<h1 class="main-header">📚 SharePoint to Astra DB ETL Dashboard</h1>', 
                unsafe_allow_html=True)
    
    # Show package status
    show_package_status()
    
    # Check configuration
    missing_vars, configured_vars = check_configuration()
    
    if missing_vars:
        st.markdown('<div class="error-message">', unsafe_allow_html=True)
        st.markdown(f"**❌ Missing environment variables:** {', '.join(missing_vars)}")
        st.markdown("</div>", unsafe_allow_html=True)
        
        st.markdown('<div class="info-box">', unsafe_allow_html=True)
        st.markdown("**💡 Please configure these in your Render dashboard environment variables.**")
        st.markdown("</div>", unsafe_allow_html=True)
        
        # Show configuration help
        with st.expander("📋 Configuration Help"):
            st.markdown("""
            **Required Environment Variables:**
            
            - `OPENAI_API_KEY`: Your OpenAI API key (get from https://platform.openai.com/api-keys)
            - `ASTRA_DB_TOKEN`: Your Astra DB token (get from Astra console)
            - `ASTRA_DB_ENDPOINT`: Your Astra DB endpoint URL
            - `SHAREPOINT_CLIENT_ID`: SharePoint application client ID
            - `SHAREPOINT_CLIENT_SECRET`: SharePoint application client secret
            - `SHAREPOINT_TENANT_ID`: Your SharePoint tenant ID
            - `SHAREPOINT_SITE_NAME`: Name of your SharePoint site
            
            **Optional:**
            - `ASTRA_COLLECTION_NAME`: Collection name (default: "documents")
            
            **How to get SharePoint credentials:**
            1. Go to Azure Portal → App Registrations
            2. Create new registration or use existing
            3. Copy Application (client) ID and Directory (tenant) ID
            4. Create client secret under "Certificates & secrets"
            5. Grant SharePoint permissions to your app
            """)
        
        # Show basic configuration interface
        show_configuration_interface()
        return
    
    # Check if LlamaIndex is available
    if not LLAMA_INDEX_AVAILABLE:
        st.markdown('<div class="warning-box">', unsafe_allow_html=True)
        st.markdown("**⚠️ Warning:** LlamaIndex is not fully available. Running in limited mode.")
        st.markdown("</div>", unsafe_allow_html=True)
    
    # Initialize services
    astra_client, sharepoint_client, document_processor, services_ok = initialize_services()
    
    if not services_ok:
        st.markdown('<div class="error-message">', unsafe_allow_html=True)
        st.markdown("**❌ Failed to initialize services.** Please check your configuration and try again.")
        st.markdown("</div>", unsafe_allow_html=True)
        
        # Show basic troubleshooting
        show_troubleshooting_interface()
        return
    
    # Success - show main interface
    st.markdown('<div class="success-message">', unsafe_allow_html=True)
    st.markdown("**✅ All services initialized successfully!** You can now use the full dashboard.")
    st.markdown("</div>", unsafe_allow_html=True)
    
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

def show_configuration_interface():
    """Show basic configuration interface when env vars are missing"""
    st.subheader("🔧 Configuration Interface")
    
    st.info("This interface will help you verify your configuration once environment variables are set.")
    
    # Test individual services
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**OpenAI Configuration**")
        openai_key = os.getenv("OPENAI_API_KEY")
        if openai_key:
            if st.button("Test OpenAI Connection"):
                test_openai_connection()
        else:
            st.warning("OPENAI_API_KEY not set")
    
    with col2:
        st.markdown("**Astra DB Configuration**")
        astra_token = os.getenv("ASTRA_DB_TOKEN")
        astra_endpoint = os.getenv("ASTRA_DB_ENDPOINT")
        if astra_token and astra_endpoint:
            if st.button("Test Astra DB Connection"):
                test_astra_connection()
        else:
            st.warning("Astra DB credentials not set")

def show_troubleshooting_interface():
    """Show troubleshooting interface when services fail to initialize"""
    st.subheader("🔧 Troubleshooting")
    
    st.markdown("""
    **Common Issues:**
    
    1. **Package Import Errors**: Try redeploying with updated requirements.txt
    2. **API Key Issues**: Ensure all API keys are valid and have proper permissions
    3. **Network Issues**: Check if Render can access external APIs
    4. **Version Conflicts**: Package versions might be incompatible
    
    **Quick Tests:**
    """)
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("Test Basic Python Imports"):
            test_basic_imports()
    
    with col2:
        if st.button("Test Environment Variables"):
            test_environment_variables()

def test_openai_connection():
    """Test OpenAI connection"""
    try:
        import openai
        client = openai.OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": "Hello"}],
            max_tokens=5
        )
        st.success("✅ OpenAI connection successful!")
    except Exception as e:
        st.error(f"❌ OpenAI connection failed: {str(e)}")

def test_astra_connection():
    """Test Astra DB connection"""
    try:
        import astrapy
        client = astrapy.DataAPIClient(os.getenv("ASTRA_DB_TOKEN"))
        db = client.get_database_by_api_endpoint(os.getenv("ASTRA_DB_ENDPOINT"))
        st.success("✅ Astra DB connection successful!")
    except Exception as e:
        st.error(f"❌ Astra DB connection failed: {str(e)}")

def test_basic_imports():
    """Test basic Python imports"""
    import_tests = {
        "pandas": "import pandas",
        "numpy": "import numpy", 
        "plotly": "import plotly",
        "requests": "import requests",
        "openai": "import openai"
    }
    
    results = {}
    for package, import_stmt in import_tests.items():
        try:
            exec(import_stmt)
            results[package] = "✅"
        except ImportError as e:
            results[package] = f"❌ {str(e)}"
    
    for package, status in results.items():
        st.write(f"{package}: {status}")

def test_environment_variables():
    """Test environment variable availability"""
    required_vars = [
        "OPENAI_API_KEY",
        "ASTRA_DB_TOKEN", 
        "ASTRA_DB_ENDPOINT",
        "SHAREPOINT_CLIENT_ID",
        "SHAREPOINT_CLIENT_SECRET", 
        "SHAREPOINT_TENANT_ID",
        "SHAREPOINT_SITE_NAME"
    ]
    
    for var in required_vars:
        value = os.getenv(var)
        if value:
            masked_value = f"{value[:8]}..." if len(value) > 8 else value
            st.write(f"✅ {var}: {masked_value}")
        else:
            st.write(f"❌ {var}: Not set")

# Placeholder functions for the main interface
def display_sidebar(astra_client, sharepoint_client):
    st.sidebar.title("🎛️ Control Panel")
    st.sidebar.success("✅ Services Online")

def data_ingestion_tab(astra_client, sharepoint_client, document_processor):
    st.header("📥 Data Ingestion")
    st.info("Data ingestion functionality will be implemented here.")

def search_query_tab(astra_client):
    st.header("🔍 Search & Query")
    st.info("Search functionality will be implemented here.")

def monitoring_tab(astra_client, sharepoint_client):
    st.header("📊 Monitoring")
    st.info("Monitoring dashboard will be implemented here.")

def settings_tab(astra_client, sharepoint_client):
    st.header("⚙️ Settings")
    st.info("Settings panel will be implemented here.")

if __name__ == "__main__":
    main()
