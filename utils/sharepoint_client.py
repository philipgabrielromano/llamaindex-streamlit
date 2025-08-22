# utils/sharepoint_client.py (Pure Office365 REST client implementation)
import streamlit as st
import os
import requests
from datetime import datetime, timedelta
from typing import List, Dict, Optional
import json
import io

# Use Office365 REST client directly instead of LlamaIndex
try:
    from office365.runtime.auth.client_credential import ClientCredential
    from office365.sharepoint.client_context import ClientContext
    from office365.sharepoint.files.file import File
    OFFICE365_AVAILABLE = True
    st.success("✅ Office365 REST client available")
except ImportError:
    OFFICE365_AVAILABLE = False
    ClientCredential = None
    ClientContext = None
    File = None
    st.warning("⚠️ Office365 REST client not available")

class SharePointClient:
    """Direct SharePoint integration using Office365 REST client"""
    
    def __init__(self):
        self.client_id = os.getenv("SHAREPOINT_CLIENT_ID")
        self.client_secret = os.getenv("SHAREPOINT_CLIENT_SECRET") 
        self.tenant_id = os.getenv("SHAREPOINT_TENANT_ID")
        self.site_name = os.getenv("SHAREPOINT_SITE_NAME")
        
        # Extract tenant name from tenant_id or site_name
        self.tenant_name = self._extract_tenant_name()
        
        if not all([self.client_id, self.client_secret, self.tenant_id, self.site_name]):
            st.warning("⚠️ Missing SharePoint configuration. Please set environment variables.")
            self.ctx = None
            return
        
        self.site_url = f"https://{self.tenant_name}.sharepoint.com/sites/{self.site_name}"
        self.ctx = None
        
        if OFFICE365_AVAILABLE:
            self._initialize_client()
        else:
            st.info("💡 Office365 REST client not available. Using mock data for testing.")
    
    def _extract_tenant_name(self):
        """Extract tenant name from configuration"""
        # If site_name looks like a full URL, extract tenant
        if self.site_name and 'sharepoint.com' in self.site_name:
            parts = self.site_name.split('.')
            return parts[0].replace('https://', '')
        
        # Use a default or ask user to configure
        return os.getenv("SHAREPOINT_TENANT_NAME", "your-tenant")
    
    def _initialize_client(self):
        """Initialize SharePoint client context"""
        try:
            if not OFFICE365_AVAILABLE:
                return
                
            credentials = ClientCredential(self.client_id, self.client_secret)
            self.ctx = ClientContext(self.site_url).with_credentials(credentials)
            
            st.success(f"✅ SharePoint client initialized for: {self.site_url}")
        except Exception as e:
            st.error(f"Failed to initialize SharePoint client: {str(e)}")
            self.ctx = None
    
    def test_connection(self) -> bool:
        """Test SharePoint connection"""
        if not OFFICE365_AVAILABLE:
            st.warning("⚠️ Office365 client not available for connection test")
            return self._test_basic_config()
        
        if not self.ctx:
            st.error("❌ SharePoint client not initialized")
            return False
        
        try:
            # Try to access the web to test connection
            web = self.ctx.web
            self.ctx.load(web)
            self.ctx.execute_query()
            
            st.success(f"✅ SharePoint connection successful! Site: {web.title}")
            return True
            
        except Exception as e:
            error_msg = str(e)
            if "401" in error_msg or "Unauthorized" in error_msg:
                st.error("❌ SharePoint authentication failed. Check your client credentials.")
            elif "404" in error_msg or "Not Found" in error_msg:
                st.error(f"❌ SharePoint site not found. Check your site URL: {self.site_url}")
            elif "403" in error_msg or "Forbidden" in error_msg:
                st.error("❌ SharePoint access denied. Check your app permissions.")
            else:
                st.error(f"❌ SharePoint connection test failed: {error_msg}")
            return False
    
    def _test_basic_config(self) -> bool:
        """Test basic configuration without full client"""
        config_items = [
            ("Client ID", self.client_id),
            ("Client Secret", self.client_secret),
            ("Tenant ID", self.tenant_id), 
            ("Site Name", self.site_name),
            ("Tenant Name", self.tenant_name)
        ]
        
        missing = [name for name, value in config_items if not value]
        
        if missing:
            st.error(f"❌ Missing configuration: {', '.join(missing)}")
            return False
        else:
            st.info("✅ SharePoint configuration appears complete")
            st.info(f"🔗 Target URL: {self.site_url}")
            return True
    
    def get_documents(self, folder_path: str = "Shared Documents", 
                     file_types: List[str] = None, 
                     since_date: Optional[datetime] = None,
                     max_docs: Optional[int] = None) -> List[Dict]:
        """Get documents from SharePoint"""
        
        if not OFFICE365_AVAILABLE or not self.ctx:
            st.warning("⚠️ SharePoint client not available. Returning mock data for testing.")
            return self._get_mock_documents()
        
        try:
            st.info(f"📂 Loading documents from: {folder_path}")
            
            # Get document library
            lists = self.ctx.web.lists
            self.ctx.load(lists)
            self.ctx.execute_query()
            
            # Find the document library
            target_list = None
            for lst in lists:
                if lst.properties['Title'] == folder_path or lst.properties['Title'] == folder_path.replace('/', ''):
                    target_list = lst
                    break
            
            if not target_list:
                st.error(f"❌ Document library '{folder_path}' not found")
                return []
            
            # Get items from the library
            items = target_list.items
            self.ctx.load(items)
            self.ctx.execute_query()
            
            documents = []
            processed_count = 0
            
            for item in items:
                try:
                    # Extract item properties
                    props = item.properties
                    filename = props.get('FileLeafRef', f'Document_{processed_count}')
                    
                    # Filter by file type if specified
                    if file_types:
                        file_ext = f".{filename.split('.')[-1].lower()}" if '.' in filename else ''
                        if file_ext not in file_types:
                            continue
                    
                    # Extract metadata
                    modified_str = props.get('Modified', datetime.now().isoformat())
                    file_path = props.get('FileRef', '')
                    item_id = props.get('ID', f'item_{processed_count}')
                    
                    # Filter by date if specified
                    if since_date:
                        try:
                            if isinstance(modified_str, str):
                                modified_dt = datetime.fromisoformat(modified_str.replace('Z', '+00:00'))
                            else:
                                modified_dt = modified_str
                            
                            if modified_dt < since_date:
                                continue
                        except Exception:
                            pass  # Include document if date parsing fails
                    
                    # Get file content
                    content = self._get_file_content(file_path, filename)
                    
                    # Create document info
                    doc_info = {
                        'id': item_id,
                        'filename': filename,
                        'content': content,
                        'modified': modified_str,
                        'file_path': file_path,
                        'metadata': {
                            'sharepoint_id': item_id,
                            'file_size': props.get('File_x0020_Size', 0),
                            'created': props.get('Created', ''),
                            'author': props.get('Author', {}).get('Title', 'Unknown') if isinstance(props.get('Author'), dict) else str(props.get('Author', 'Unknown')),
                            'source': 'sharepoint_direct',
                            'site_url': self.site_url,
                            'library': folder_path,
                            'processed_at': datetime.now().isoformat(),
                            'text_length': len(content),
                            'word_count': len(content.split()) if content else 0
                        }
                    }
                    
                    documents.append(doc_info)
                    processed_count += 1
                    
                    # Apply max docs limit
                    if max_docs and processed_count >= max_docs:
                        st.info(f"📊 Reached maximum document limit: {max_docs}")
                        break
                        
                except Exception as item_error:
                    st.warning(f"Error processing item: {str(item_error)}")
                    continue
            
            st.success(f"✅ Successfully loaded {len(documents)} documents from SharePoint")
            return documents
            
        except Exception as e:
            error_msg = str(e)
            if "401" in error_msg:
                st.error("❌ SharePoint authentication failed. Check your credentials.")
            elif "404" in error_msg:
                st.error(f"❌ SharePoint library '{folder_path}' not found.")
            elif "403" in error_msg:
                st.error("❌ Access denied to SharePoint library. Check permissions.")
            else:
                st.error(f"❌ Error retrieving SharePoint documents: {error_msg}")
            return []
    
    def _get_file_content(self, file_path: str, filename: str) -> str:
        """Get content from a SharePoint file"""
        try:
            if not self.ctx or not file_path:
                return f"[Content not available for {filename}]"
            
            # Get file object
            file_obj = self.ctx.web.get_file_by_server_relative_url(file_path)
            file_content = file_obj.read()
            self.ctx.execute_query()
            
            # Extract text based on file type
            file_ext = filename.split('.')[-1].lower() if '.' in filename else ''
            
            if file_ext == 'txt':
                return file_content.decode('utf-8', errors='ignore')
            elif file_ext in ['pdf', 'docx']:
                # For now, return a placeholder
                # In a full implementation, you'd use PyPDF2 or python-docx here
                return f"[{file_ext.upper()} content from {filename} - text extraction would be implemented here]"
            else:
                # Try to decode as text
                try:
                    return file_content.decode('utf-8', errors='ignore')
                except:
                    return f"[Binary content from {filename} - {len(file_content)} bytes]"
                    
        except Exception as e:
            st.warning(f"Could not read content from {filename}: {str(e)}")
            return f"[Could not extract content from {filename}]"
    
    def _get_mock_documents(self) -> List[Dict]:
        """Return mock documents for testing"""
        mock_docs = [
            {
                'id': 'mock_1',
                'filename': 'Quarterly Report Q1 2024.pdf',
                'content': 'This is a sample quarterly report containing financial analysis, performance metrics, and strategic recommendations for Q1 2024. The report shows strong growth in key performance indicators.',
                'modified': datetime.now().isoformat(),
                'file_path': '/sites/yoursite/Shared Documents/Reports/Quarterly Report Q1 2024.pdf',
                'metadata': {
                    'source': 'mock_data',
                    'created_at': (datetime.now() - timedelta(days=7)).isoformat(),
                    'file_size': 1024000,
                    'author': 'Finance Team',
                    'text_length': 150,
                    'word_count': 25
                }
            },
            {
                'id': 'mock_2', 
                'filename': 'Project Status Update.docx',
                'content': 'Weekly project status update covering milestone achievements, resource allocation, risk assessment, and next steps. All major deliverables are on track for the planned timeline.',
                'modified': (datetime.now() - timedelta(hours=3)).isoformat(),
                'file_path': '/sites/yoursite/Shared Documents/Projects/Project Status Update.docx',
                'metadata': {
                    'source': 'mock_data',
                    'created_at': (datetime.now() - timedelta(hours=3)).isoformat(),
                    'file_size': 512000,
                    'author': 'Project Manager',
                    'text_length': 140,
                    'word_count': 23
                }
            },
            {
                'id': 'mock_3',
                'filename': 'Meeting Notes - Team Sync.txt',
                'content': 'Team synchronization meeting notes including action items, decisions made, and follow-up tasks. Key topics discussed: budget planning, resource requirements, and timeline adjustments.',
                'modified': (datetime.now() - timedelta(hours=8)).isoformat(),
                'file_path': '/sites/yoursite/Shared Documents/Meetings/Meeting Notes - Team Sync.txt',
                'metadata': {
                    'source': 'mock_data',
                    'created_at': (datetime.now() - timedelta(hours=8)).isoformat(),
                    'file_size': 256000,
                    'author': 'Team Lead',
                    'text_length': 130,
                    'word_count': 21
                }
            }
        ]
        
        st.info(f"📋 Generated {len(mock_docs)} mock documents for testing")
        return mock_docs
    
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
            'tenant_name': bool(self.tenant_name),
            'office365_available': OFFICE365_AVAILABLE,
            'client_initialized': bool(self.ctx),
            'site_url': bool(self.site_url)
        }
        
        return config_status
    
    def get_site_info(self) -> Dict:
        """Get SharePoint site information"""
        return {
            'site_name': self.site_name or 'Not configured',
            'tenant_name': self.tenant_name or 'Not configured',
            'tenant_id': self.tenant_id[:8] + "..." if self.tenant_id else "Not configured",
            'client_id': self.client_id[:8] + "..." if self.client_id else "Not configured",
            'site_url': self.site_url,
            'client_status': 'Available' if self.ctx else 'Not available',
            'office365_package_status': 'Installed' if OFFICE365_AVAILABLE else 'Missing'
        }
    
    def get_available_libraries(self) -> List[str]:
        """Get list of available document libraries"""
        if not OFFICE365_AVAILABLE or not self.ctx:
            return ["Shared Documents", "Documents", "Site Assets"]
        
        try:
            lists = self.ctx.web.lists
            self.ctx.load(lists)
            self.ctx.execute_query()
            
            libraries = []
            for lst in lists:
                list_props = lst.properties
                if list_props.get('BaseTemplate') == 101:  # Document library template
                    libraries.append(list_props.get('Title', 'Unknown'))
            
            return libraries if libraries else ["Shared Documents"]
            
        except Exception as e:
            st.warning(f"Could not get libraries: {str(e)}")
            return ["Shared Documents"]
