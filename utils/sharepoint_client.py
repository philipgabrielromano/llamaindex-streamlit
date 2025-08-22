# utils/sharepoint_client.py (Remove st calls during import)
import os
import requests
from datetime import datetime, timedelta
from typing import List, Dict, Optional
import json

# Try Office365 REST client import without Streamlit calls
try:
    from office365.runtime.auth.client_credential import ClientCredential
    from office365.sharepoint.client_context import ClientContext
    from office365.sharepoint.files.file import File
    OFFICE365_AVAILABLE = True
except ImportError:
    OFFICE365_AVAILABLE = False
    ClientCredential = None
    ClientContext = None
    File = None

class SharePointClient:
    """Direct SharePoint integration using Office365 REST client"""
    
    def __init__(self):
        # Only import streamlit here when methods are called, not during import
        import streamlit as st
        
        self.client_id = os.getenv("SHAREPOINT_CLIENT_ID")
        self.client_secret = os.getenv("SHAREPOINT_CLIENT_SECRET") 
        self.tenant_id = os.getenv("SHAREPOINT_TENANT_ID")
        self.site_name = os.getenv("SHAREPOINT_SITE_NAME")
        
        # Extract tenant name from configuration
        self.tenant_name = self._extract_tenant_name()
        
        if not all([self.client_id, self.client_secret, self.tenant_id, self.site_name]):
            # Only show warning when actually using the client, not during import
            pass
        
        self.site_url = f"https://{self.tenant_name}.sharepoint.com/sites/{self.site_name}"
        self.ctx = None
        
        if OFFICE365_AVAILABLE:
            self._initialize_client()
    
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
            
        except Exception as e:
            # Don't call streamlit during initialization
            self.ctx = None
    
    def test_connection(self) -> bool:
        """Test SharePoint connection"""
        import streamlit as st
        
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
        import streamlit as st
        
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
    
    # ... rest of your methods remain the same but import streamlit locally when needed
    
    def get_recent_changes(self, hours: int = 24) -> List[Dict]:
        """Get documents modified in the last N hours"""
        import streamlit as st
        
        since_date = datetime.now() - timedelta(hours=hours)
        st.info(f"🕒 Looking for documents modified since: {since_date.strftime('%Y-%m-%d %H:%M:%S')}")
        return self.get_documents(since_date=since_date)
    
    def get_documents(self, folder_path: str = "Shared Documents", 
                     file_types: List[str] = None, 
                     since_date: Optional[datetime] = None,
                     max_docs: Optional[int] = None) -> List[Dict]:
        """Get documents from SharePoint"""
        import streamlit as st
        
        if not OFFICE365_AVAILABLE or not self.ctx:
            st.warning("⚠️ SharePoint client not available. Returning mock data for testing.")
            return self._get_mock_documents()
        
        # ... rest of method
        return []
    
    def _get_mock_documents(self) -> List[Dict]:
        """Return mock documents for testing"""
        import streamlit as st
        
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
    
    def validate_configuration(self) -> Dict[str, bool]:
        """Validate SharePoint configuration"""
        return {
            'client_id': bool(self.client_id),
            'client_secret': bool(self.client_secret),
            'tenant_id': bool(self.tenant_id),
            'site_name': bool(self.site_name),
            'tenant_name': bool(self.tenant_name),
            'office365_available': OFFICE365_AVAILABLE,
            'client_initialized': bool(self.ctx),
            'site_url': bool(self.site_url)
        }
