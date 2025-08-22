# utils/sharepoint_client.py
import streamlit as st
import os
import requests
from datetime import datetime, timedelta
from typing import List, Dict, Optional
import json

# Updated import for SharePoint
try:
    from llama_index.readers.microsoft_sharepoint import SharePointReader
except ImportError:
    # Fallback for different package structure
    try:
        from llama_index.readers.sharepoint import SharePointReader
    except ImportError:
        SharePointReader = None
        st.error("SharePoint reader not available. Please check package installation.")

class SharePointClient:
    """Handles SharePoint integration and document retrieval"""
    
    def __init__(self):
        self.client_id = os.getenv("SHAREPOINT_CLIENT_ID")
        self.client_secret = os.getenv("SHAREPOINT_CLIENT_SECRET") 
        self.tenant_id = os.getenv("SHAREPOINT_TENANT_ID")
        self.site_name = os.getenv("SHAREPOINT_SITE_NAME")
        
        if not all([self.client_id, self.client_secret, self.tenant_id, self.site_name]):
            raise ValueError("Missing required SharePoint configuration")
        
        self.reader = None
        if SharePointReader:
            self._initialize_reader()
    
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
        except Exception as e:
            st.error(f"Failed to initialize SharePoint reader: {str(e)}")
    
    def test_connection(self) -> bool:
        """Test SharePoint connection"""
        try:
            if not self.reader:
                return False
            
            # Try to load a minimal set of documents
            test_docs = self.reader.load_data(
                sharepoint_site_name=self.site_name,
                sharepoint_folder_path="/Shared Documents",
                recursive=False
            )
            return True
            
        except Exception as e:
            st.error(f"SharePoint connection test failed: {str(e)}")
            return False
    
    def get_documents(self, folder_path: str = "/Shared Documents", 
                     file_types: List[str] = None, 
                     since_date: Optional[datetime] = None) -> List[Dict]:
        """Get documents from SharePoint"""
        try:
            if not self.reader:
                raise Exception("SharePoint reader not initialized")
            
            # Set up file extractor for specified types
            file_extractor = None
            if file_types:
                file_extractor = {}
                for ext in file_types:
                    if not ext.startswith('.'):
                        ext = f'.{ext}'
                    file_extractor[ext] = "default"
            
            # Load documents with proper parameters
            documents = self.reader.load_data(
                sharepoint_site_name=self.site_name,
                sharepoint_folder_path=folder_path,
                file_extractor=file_extractor,
                recursive=True
            )
            
            # Convert to list of dicts with metadata
            doc_list = []
            for doc in documents:
                doc_info = {
                    'content': doc.text,
                    'filename': doc.metadata.get('filename', doc.metadata.get('file_name', 'Unknown')),
                    'file_path': doc.metadata.get('file_path', doc.metadata.get('source', '')),
                    'modified': doc.metadata.get('last_modified', doc.metadata.get('modified', datetime.now().isoformat())),
                    'id': doc.metadata.get('id', doc.metadata.get('doc_id', '')),
                    'metadata': doc.metadata
                }
                
                # Filter by date if specified
                if since_date:
                    try:
                        doc_modified_str = doc_info['modified']
                        if isinstance(doc_modified_str, str):
                            # Handle different date formats
                            if 'T' in doc_modified_str:
                                doc_modified = datetime.fromisoformat(doc_modified_str.replace('Z', '+00:00'))
                            else:
                                doc_modified = datetime.fromisoformat(doc_modified_str)
                        else:
                            doc_modified = doc_modified_str
                        
                        if doc_modified < since_date:
                            continue
                    except Exception:
                        # If date parsing fails, include the document
                        pass
                
                doc_list.append(doc_info)
            
            return doc_list
            
        except Exception as e:
            st.error(f"Error retrieving SharePoint documents: {str(e)}")
            return []
    
    def get_folder_structure(self, root_path: str = "/Shared Documents") -> Dict:
        """Get SharePoint folder structure"""
        try:
            # This would require direct SharePoint API calls
            # For now, return a simple structure
            return {
                "folders": ["/Shared Documents", "/Documents", "/Reports"],
                "root": root_path
            }
        except Exception as e:
            st.error(f"Error getting folder structure: {str(e)}")
            return {"folders": [], "root": root_path}
    
    def get_recent_changes(self, hours: int = 24) -> List[Dict]:
        """Get documents modified in the last N hours"""
        since_date = datetime.now() - timedelta(hours=hours)
        return self.get_documents(since_date=since_date)
    
    def validate_configuration(self) -> Dict[str, bool]:
        """Validate SharePoint configuration"""
        config_status = {
            'client_id': bool(self.client_id),
            'client_secret': bool(self.client_secret),
            'tenant_id': bool(self.tenant_id),
            'site_name': bool(self.site_name),
            'reader_available': SharePointReader is not None,
            'reader_initialized': bool(self.reader)
        }
        
        return config_status
