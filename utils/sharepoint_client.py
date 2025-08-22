# utils/sharepoint_client.py (Fixed with better fallback)
import streamlit as st
import os
import requests
from datetime import datetime, timedelta
from typing import List, Dict, Optional
import json

# Try multiple SharePoint reader import paths
SHAREPOINT_AVAILABLE = False
SharePointReader = None

# Try different import paths
import_attempts = [
    "llama_index.readers.microsoft_sharepoint",
    "llama_index_readers_microsoft_sharepoint", 
    "llama_index.readers.sharepoint",
    "llama_index_readers_sharepoint"
]

for import_path in import_attempts:
    try:
        module = __import__(import_path, fromlist=['SharePointReader'])
        SharePointReader = getattr(module, 'SharePointReader', None)
        if SharePointReader:
            SHAREPOINT_AVAILABLE = True
            st.success(f"✅ SharePoint reader loaded from: {import_path}")
            break
    except ImportError:
        continue

if not SHAREPOINT_AVAILABLE:
    st.warning("⚠️ SharePoint reader not available. Install with: pip install llama-index-readers-microsoft-sharepoint")

class SharePointClient:
    """Handles SharePoint integration with fallback for missing reader"""
    
    def __init__(self):
        self.client_id = os.getenv("SHAREPOINT_CLIENT_ID")
        self.client_secret = os.getenv("SHAREPOINT_CLIENT_SECRET") 
        self.tenant_id = os.getenv("SHAREPOINT_TENANT_ID")
        self.site_name = os.getenv("SHAREPOINT_SITE_NAME")
        
        if not all([self.client_id, self.client_secret, self.tenant_id, self.site_name]):
            st.warning("⚠️ Missing SharePoint configuration. Please set environment variables.")
            self.reader = None
            return
        
        self.reader = None
        if SHAREPOINT_AVAILABLE and SharePointReader:
            self._initialize_reader()
        else:
            st.info("💡 SharePoint reader not available. You can still test other features.")
    
    def _initialize_reader(self):
        """Initialize SharePoint reader"""
        try:
            if not SharePointReader:
                raise Exception("SharePoint reader not available")
                
            self.reader = SharePointReader(
                client_id=self.client_id,
                client_secret=self.client_secret,
                tenant_id=self.tenant_id
            )
            st.success("✅ SharePoint reader initialized successfully!")
        except Exception as e:
            st.error(f"Failed to initialize SharePoint reader: {str(e)}")
            self.reader = None
    
    def test_connection(self) -> bool:
        """Test SharePoint connection"""
        if not self.reader:
            st.warning("⚠️ SharePoint reader not available for connection test")
            return self._test_basic_config()
        
        try:
            # Try to load a minimal set of documents to test connection
            test_docs = self.reader.load_data(
                sharepoint_site_name=self.site_name,
                sharepoint_folder_path="/Shared Documents",
                recursive=False
            )
            
            st.success(f"✅ SharePoint connection successful! Found {len(test_docs)} documents in test.")
            return True
            
        except Exception as e:
            error_msg = str(e)
            if "401" in error_msg or "Unauthorized" in error_msg:
                st.error("❌ SharePoint authentication failed. Check your client credentials.")
            elif "404" in error_msg or "Not Found" in error_msg:
                st.error(f"❌ SharePoint site '{self.site_name}' not found. Check your site name.")
            elif "403" in error_msg or "Forbidden" in error_msg:
                st.error("❌ SharePoint access denied. Check your app permissions.")
            else:
                st.error(f"❌ SharePoint connection test failed: {error_msg}")
            return False
    
    def _test_basic_config(self) -> bool:
        """Test basic configuration without SharePoint reader"""
        config_ok = all([self.client_id, self.client_secret, self.tenant_id, self.site_name])
        
        if config_ok:
            st.info("✅ SharePoint configuration appears complete (reader not available for full test)")
        else:
            st.error("❌ SharePoint configuration incomplete")
        
        return config_ok
    
    def get_documents(self, folder_path: str = "/Shared Documents", 
                     file_types: List[str] = None, 
                     since_date: Optional[datetime] = None,
                     max_docs: Optional[int] = None) -> List[Dict]:
        """Get documents from SharePoint"""
        
        if not self.reader:
            st.warning("⚠️ SharePoint reader not available. Returning mock data for testing.")
            return self._get_mock_documents()
        
        try:
            # Set up file extractor for specified types
            file_extractor = None
            if file_types:
                file_extractor = {}
                for ext in file_types:
                    if not ext.startswith('.'):
                        ext = f'.{ext}'
                    file_extractor[ext] = "default"
            
            st.info(f"📂 Loading documents from: {folder_path}")
            if file_types:
                st.info(f"🔍 Filtering for file types: {', '.join(file_types)}")
            
            # Load documents with proper parameters
            documents = self.reader.load_data(
                sharepoint_site_name=self.site_name,
                sharepoint_folder_path=folder_path,
                file_extractor=file_extractor,
                recursive=True
            )
            
            st.info(f"📄 Loaded {len(documents)} raw documents from SharePoint")
            
            # Process documents (same as before)
            return self._process_documents(documents, since_date, max_docs)
            
        except Exception as e:
            error_msg = str(e)
            if "401" in error_msg:
                st.error("❌ SharePoint authentication failed. Check your credentials.")
            elif "404" in error_msg:
                st.error(f"❌ SharePoint folder '{folder_path}' not found.")
            elif "403" in error_msg:
                st.error("❌ Access denied to SharePoint folder. Check permissions.")
            else:
                st.error(f"❌ Error retrieving SharePoint documents: {error_msg}")
            return []
    
    def _get_mock_documents(self) -> List[Dict]:
        """Return mock documents for testing when SharePoint reader isn't available"""
        mock_docs = [
            {
                'id': 'mock_1',
                'filename': 'Sample Document 1.pdf',
                'content': 'This is a sample document for testing the ETL pipeline. It contains various information about project management and best practices.',
                'modified': datetime.now().isoformat(),
                'file_path': '/Shared Documents/Sample Document 1.pdf',
                'metadata': {
                    'source': 'mock_data',
                    'created_at': datetime.now().isoformat(),
                    'file_size': 1024,
                    'author': 'System'
                }
            },
            {
                'id': 'mock_2', 
                'filename': '
