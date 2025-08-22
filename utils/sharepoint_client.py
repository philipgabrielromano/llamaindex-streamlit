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
    SHAREPOINT_AVAILABLE = True
except ImportError:
    try:
        # Fallback import path
        from llama_index_readers_microsoft_sharepoint import SharePointReader
        SHAREPOINT_AVAILABLE = True
    except ImportError:
        try:
            # Another fallback
            from llama_index.readers.sharepoint import SharePointReader
            SHAREPOINT_AVAILABLE = True
        except ImportError:
            SharePointReader = None
            SHAREPOINT_AVAILABLE = False

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
        if SHAREPOINT_AVAILABLE and SharePointReader:
            self._initialize_reader()
        else:
            st.warning("⚠️ SharePoint reader not available. Some features may be limited.")
    
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
            self.reader = None
    
    def test_connection(self) -> bool:
        """Test SharePoint connection"""
        try:
            if not self.reader:
                st.warning("SharePoint reader not initialized")
                return False
            
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
    
    def get_documents(self, folder_path: str = "/Shared Documents", 
                     file_types: List[str] = None, 
                     since_date: Optional[datetime] = None,
                     max_docs: Optional[int] = None) -> List[Dict]:
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
            
            # Convert to list of dicts with metadata
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
                    
                    # Filter by date if specified
                    if since_date:
                        try:
                            doc_modified_str = doc_info['modified']
                            if isinstance(doc_modified_str, str):
                                # Handle different date formats
                                if 'T' in doc_modified_str:
                                    # ISO format
                                    doc_modified = datetime.fromisoformat(doc_modified_str.replace('Z', '+00:00'))
                                else:
                                    # Try parsing as ISO date
                                    doc_modified = datetime.fromisoformat(doc_modified_str)
                            else:
                                doc_modified = doc_modified_str
                            
                            if doc_modified < since_date:
                                continue
                        except Exception as date_error:
                            # If date parsing fails, include the document
                            st.warning(f"Could not parse date for {filename}: {date_error}")
                    
                    # Check if document has content
                    if not doc.text or not doc.text.strip():
                        st.warning(f"⚠️ No text content found in {filename}")
                        continue
                    
                    doc_list.append(doc_info)
                    processed_count += 1
                    
                    # Limit number of documents if specified
                    if max_docs and processed_count >= max_docs:
                        st.info(f"📊 Reached maximum document limit: {max_docs}")
                        break
                        
                except Exception as doc_error:
                    st.warning(f"Error processing document: {str(doc_error)}")
                    continue
            
            st.success(f"✅ Successfully processed {len(doc_list)} documents")
            return doc_list
            
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
    
    def get_folder_structure(self, root_path: str = "/Shared Documents") -> Dict:
        """Get SharePoint folder structure"""
        try:
            if not self.reader:
                return {"folders": ["/Shared Documents"], "root": root_path}
            
            # This would require direct SharePoint API calls for full implementation
            # For now, return common SharePoint folder structure
            common_folders = [
                "/Shared Documents",
                "/Documents", 
                "/Site Assets",
                "/Lists",
                "/Forms"
            ]
            
            return {
                "folders": common_folders,
                "root": root_path,
                "note": "Common SharePoint folders. Actual structure may vary."
            }
            
        except Exception as e:
            st.error(f"Error getting folder structure: {str(e)}")
            return {"folders": [root_path], "root": root_path}
    
    def get_recent_changes(self, hours: int = 24) -> List[Dict]:
        """Get documents modified in the last N hours"""
        since_date = datetime.now() - timedelta(hours=hours)
        st.info(f"🕒 Looking for documents modified since: {since_date.strftime('%Y-%m-%d %H:%M:%S')}")
        return self.get_documents(since_date=since_date)
    
    def get_document_by_id(self, doc_id: str) -> Optional[Dict]:
        """Get a specific document by ID"""
        try:
            # This would require direct SharePoint API implementation
            # For now, return None
            st.info(f"Document lookup by ID not yet implemented: {doc_id}")
            return None
        except Exception as e:
            st.error(f"Error getting document by ID: {str(e)}")
            return None
    
    def search_documents(self, query: str, max_results: int = 10) -> List[Dict]:
        """Search documents by query"""
        try:
            # This would require SharePoint Search API implementation
            # For now, get all documents and filter by content
            all_docs = self.get_documents(max_docs=max_results * 2)
            
            # Simple text search in content
            matching_docs = []
            query_lower = query.lower()
            
            for doc in all_docs:
                content = doc.get('content', '').lower()
                filename = doc.get('filename', '').lower()
                
                if query_lower in content or query_lower in filename:
                    matching_docs.append(doc)
                
                if len(matching_docs) >= max_results:
                    break
            
            st.info(f"🔍 Found {len(matching_docs)} documents matching '{query}'")
            return matching_docs
            
        except Exception as e:
            st.error(f"Error searching documents: {str(e)}")
            return []
    
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
                'site_name': self.site_name,
                'tenant_id': self.tenant_id[:8] + "..." if self.tenant_id else "Not configured",
                'client_id': self.client_id[:8] + "..." if self.client_id else "Not configured",
                'reader_status': 'Available' if self.reader else 'Not available',
                'estimated_url': f"https://[tenant].sharepoint.com/sites/{self.site_name}" if self.site_name else "Not configured"
            }
            
            return site_info
            
        except Exception as e:
            st.error(f"Error getting site info: {str(e)}")
            return {}
    
    def list_available_libraries(self) -> List[str]:
        """List available document libraries"""
        try:
            # Common SharePoint document library names
            common_libraries = [
                "Shared Documents",
                "Documents", 
                "Site Assets",
                "Pages",
                "Reports",
                "Templates"
            ]
            
            return common_libraries
            
        except Exception as e:
            st.error(f"Error listing libraries: {str(e)}")
            return ["Shared Documents"]
    
    def get_supported_file_types(self) -> List[str]:
        """Get list of supported file types for SharePoint reader"""
        return [
            ".pdf",
            ".docx", 
            ".doc",
            ".txt",
            ".md",
            ".html",
            ".pptx",
            ".ppt",
            ".xlsx",
            ".xls"
        ]
    
    def cleanup(self):
        """Cleanup resources"""
        try:
            if self.reader:
                # Cleanup reader resources if needed
                self.reader = None
        except Exception as e:
            st.warning(f"Error during cleanup: {str(e)}")
