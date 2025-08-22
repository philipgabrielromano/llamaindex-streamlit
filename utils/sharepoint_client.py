# utils/sharepoint_client.py
import streamlit as st
import os
import requests
from datetime import datetime, timedelta
from typing import List, Dict, Optional
import json

from llama_index.readers.sharepoint import SharePointReader

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
        self._initialize_reader()
    
    def _initialize_reader(self):
        """Initialize SharePoint reader"""
        try:
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
            file_extractor = {}
            if file_types:
                for ext in file_types:
                    file_extractor[ext] = "default"
            
            # Load documents
            documents = self.reader.load_data(
                sharepoint_site_name=self.site_name,
                sharepoint_folder_path=folder_path,
                file_extractor=file_extractor if file_extractor else None,
                recursive=True
            )
            
            # Convert to list of dicts with metadata
            doc_list = []
            for doc in documents:
                doc_info = {
                    'content': doc.text,
                    'filename': doc.metadata.get('filename', 'Unknown'),
                    'file_path': doc.metadata.get('file_path', ''),
                    'modified': doc.metadata.get('last_modified', datetime.now().isoformat()),
                    'id': doc.metadata.get('id', ''),
                    'metadata': doc.metadata
                }
                
                # Filter by date if specified
                if since_date:
                    doc_modified = datetime.fromisoformat(doc_info['modified'].replace('Z', '+00:00'))
                    if doc_modified < since_date:
                        continue
                
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
            'reader_initialized': bool(self.reader)
        }
        
        return config_status
