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
                'filename': 'Test Report.docx',
                'content': 'This is a test report containing quarterly analysis and recommendations for improvement. The report covers multiple areas including performance metrics.',
                'modified': (datetime.now() - timedelta(hours=2)).isoformat(),
                'file_path': '/Shared Documents/Test Report.docx',
                'metadata': {
                    'source': 'mock_data',
                    'created_at': (datetime.now() - timedelta(hours=2)).isoformat(),
                    'file_size': 2048,
                    'author': 'System'
                }
            }
        ]
        
        st.info(f"📋 Generated {len(mock_docs)} mock documents for testing")
        return mock_docs
    
    def _process_documents(self, documents, since_date, max_docs):
        """Process SharePoint documents into the expected format"""
        doc_list = []
        processed_count = 0
        
        for doc in documents:
            try:
                # Extract metadata with fallbacks
                filename = (
                    doc.metadata.get('filename') or 
                    doc.metadata.get('file_name') or 
                    doc.metadata.get('title') or 
                    doc.metadata.get('name') or
                    f'Document_{processed_count + 1}'
                )
                
                file_path = (
                    doc.metadata.get('file_path') or 
                    doc.metadata.get('source') or 
                    doc.metadata.get('url') or
                    ''
                )
                
                modified_date = (
                    doc.metadata.get('last_modified') or 
                    doc.metadata.get('modified') or 
                    doc.metadata.get('date_modified') or
                    datetime.now().isoformat()
                )
                
                doc_id = (
                    doc.metadata.get('id') or 
                    doc.metadata.get('doc_id') or 
                    doc.metadata.get('document_id') or
                    f'doc_{processed_count + 1}'
                )
                
                # Create document info
                doc_info = {
                    'content': doc.text or '',
                    'filename': filename,
                    'file_path': file_path,
                    'modified': modified_date,
                    'id': doc_id,
                    'metadata': {
                        **doc.metadata,
                        'source': 'sharepoint',
                        'processed_at': datetime.now().isoformat(),
                        'text_length': len(doc.text) if doc.text else 0,
                        'word_count': len(doc.text.split()) if doc.text else 0
                    }
                }
                
                # Apply filters
                if since_date:
                    try:
                        doc_modified_str = doc_info['modified']
                        if isinstance(doc_modified_str, str):
                            if 'T' in doc_modified_str:
                                doc_modified = datetime.fromisoformat(doc_modified_str.replace('Z', '+00:00'))
                            else:
                                doc_modified = datetime.fromisoformat(doc_modified_str)
                        else:
                            doc_modified = doc_modified_str
                        
                        if doc_modified < since_date:
                            continue
                    except Exception as date_error:
                        st.warning(f"Could not parse date for {filename}: {date_error}")
                
                # Check content
                if not doc.text or not doc.text.strip():
                    st.warning(f"⚠️ No text content found in {filename}")
                    continue
                
                doc_list.append(doc_info)
                processed_count += 1
                
                # Apply max docs limit
                if max_docs and processed_count >= max_docs:
                    st.info(f"📊 Reached maximum document limit: {max_docs}")
                    break
                    
            except Exception as doc_error:
                st.warning(f"Error processing document: {str(doc_error)}")
                continue
        
        st.success(f"✅ Successfully processed {len(doc_list)} documents")
        return doc_list
    
    def get_recent_changes(self, hours: int = 24) -> List[Dict]:
        """Get documents modified in the last N hours"""
        since_date = datetime.now() - timedelta(hours=hours)
        st.info(f"🕒 Looking for documents modified since: {since_date.strftime('%Y-%m-%d %H:%M:%S')}")
        return self.get_documents(since_date=since_date)
    
    def validate_configuration(self) -> Dict[str, bool]:
        """Validate SharePoint configuration"""
        config_status = {
            'client_id': bool(self.client_id),
            'client_secret': bool(self.client_secret),
            'tenant_id': bool(self.tenant_id),
            'site_name': bool(self.site_name),
            'reader_available': SHAREPOINT_AVAILABLE and SharePointReader is not None,
            'reader_initialized': bool(self.reader)
        }
        
        return config_status
    
    def get_site_info(self) -> Dict:
        """Get SharePoint site information"""
        try:
            site_info = {
                'site_name': self.site_name or 'Not configured',
                'tenant_id': self.tenant_id[:8] + "..." if self.tenant_id else "Not configured",
                'client_id': self.client_id[:8] + "..." if self.client_id else "Not configured",
                'reader_status': 'Available' if self.reader else 'Not available',
                'reader_package_status': 'Installed' if SHAREPOINT_AVAILABLE else 'Missing',
                'estimated_url': f"https://[tenant].sharepoint.com/sites/{self.site_name}" if self.site_name else "Not configured"
            }
            
            return site_info
            
        except Exception as e:
            st.error(f"Error getting site info: {str(e)}")
            return {}
