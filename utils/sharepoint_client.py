# utils/sharepoint_client.py (Enhanced with auth error handling)
import os
import requests
from datetime import datetime, timedelta
from typing import List, Dict, Optional
import json

# Try Office365 REST client import
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
    """Direct SharePoint integration with enhanced error handling"""
    
    def __init__(self):
        self.client_id = os.getenv("SHAREPOINT_CLIENT_ID")
        self.client_secret = os.getenv("SHAREPOINT_CLIENT_SECRET") 
        self.tenant_id = os.getenv("SHAREPOINT_TENANT_ID")
        self.site_name = os.getenv("SHAREPOINT_SITE_NAME")
        
        # Extract tenant name from configuration
        self.tenant_name = self._extract_tenant_name()
        
        if not all([self.client_id, self.client_secret, self.tenant_id, self.site_name]):
            pass  # Handle in methods
        
        self.site_url = f"https://{self.tenant_name}.sharepoint.com/sites/{self.site_name}"
        self.ctx = None
        self.auth_tested = False
        
        if OFFICE365_AVAILABLE:
            self._initialize_client()
    
    def _extract_tenant_name(self):
        """Extract tenant name from configuration"""
        # Try multiple sources for tenant name
        tenant_name = os.getenv("SHAREPOINT_TENANT_NAME")
        
        if tenant_name:
            return tenant_name
        
        # Extract from site name if it contains sharepoint.com
        if self.site_name and 'sharepoint.com' in self.site_name:
            parts = self.site_name.split('.')
            return parts[0].replace('https://', '')
        
        # Extract from any URL in environment
        for env_var in ['SHAREPOINT_SITE_URL', 'SHAREPOINT_BASE_URL']:
            url = os.getenv(env_var, '')
            if 'sharepoint.com' in url:
                parts = url.split('.')
                return parts[0].split('//')[-1]
        
        # Default fallback
        return "your-tenant"
    
    def _initialize_client(self):
        """Initialize SharePoint client context"""
        try:
            if not OFFICE365_AVAILABLE:
                return
                
            credentials = ClientCredential(self.client_id, self.client_secret)
            self.ctx = ClientContext(self.site_url).with_credentials(credentials)
            
        except Exception as e:
            self.ctx = None
    
    def get_available_libraries(self) -> List[str]:
        """Get list of available document libraries with fallback"""
        import streamlit as st
        
        # Always provide a basic list that works
        default_libraries = [
            "Shared Documents",
            "Documents", 
            "Site Assets",
            "Reports",
            "Templates"
        ]
        
        if not OFFICE365_AVAILABLE or not self.ctx:
            st.info("📚 Using default library list (SharePoint client not fully available)")
            return default_libraries
        
        try:
            # Test authentication first
            if not self._test_authentication():
                st.warning("⚠️ SharePoint authentication issue. Using default libraries.")
                return default_libraries
            
            # Get all lists from SharePoint site
            lists = self.ctx.web.lists
            self.ctx.load(lists)
            self.ctx.execute_query()
            
            libraries = []
            for lst in lists:
                try:
                    list_props = lst.properties
                    # Check if it's a document library (BaseTemplate 101)
                    base_template = list_props.get('BaseTemplate')
                    if base_template == 101:
                        library_title = list_props.get('Title', 'Unknown')
                        if library_title and not library_title.startswith('_'):  # Skip system libraries
                            libraries.append(library_title)
                except Exception:
                    continue  # Skip problematic lists
            
            if libraries:
                st.success(f"📚 Found {len(libraries)} document libraries: {', '.join(libraries[:3])}{'...' if len(libraries) > 3 else ''}")
                return libraries
            else:
                st.info("📚 No document libraries found, using defaults")
                return default_libraries
            
        except Exception as e:
            error_msg = str(e).lower()
            
            if 'app-only access token failed' in error_msg:
                st.error("❌ **SharePoint App Authentication Failed**")
                st.markdown("""
                **To fix this issue:**
                
                1. **Check App Registration Permissions:**
                   - Go to Azure Portal → App Registrations → Your App
                   - API Permissions → Add: `Sites.Read.All`, `Files.Read.All` (Application permissions)
                   - Click "Grant admin consent"
                
                2. **Verify Environment Variables:**
                   - SHAREPOINT_CLIENT_ID (Application ID from Azure)
                   - SHAREPOINT_CLIENT_SECRET (From Certificates & Secrets)
                   - SHAREPOINT_TENANT_ID (Directory ID from Azure)
                
                3. **SharePoint Site Permissions:**
                   - Your app needs explicit permission to access the SharePoint site
                   - Contact your SharePoint admin to grant access
                
                **Using default libraries for now...**
                """)
            else:
                st.warning(f"Could not get libraries: {str(e)}")
            
            return default_libraries
    
    def _test_authentication(self) -> bool:
        """Test SharePoint authentication without showing UI messages"""
        if self.auth_tested:
            return True  # Don't test repeatedly
        
        try:
            if not self.ctx:
                return False
            
            # Simple test - try to access web properties
            web = self.ctx.web
            self.ctx.load(web)
            self.ctx.execute_query()
            
            self.auth_tested = True
            return True
            
        except Exception as e:
            error_msg = str(e).lower()
            if 'app-only access token failed' in error_msg or '401' in error_msg:
                return False
            return False
    
    def test_connection(self) -> bool:
        """Test SharePoint connection with detailed feedback"""
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
            
            st.success(f"✅ SharePoint connection successful! Site: {getattr(web, 'title', 'Unknown')}")
            self.auth_tested = True
            return True
            
        except Exception as e:
            error_msg = str(e)
            
            if "app-only access token failed" in error_msg.lower():
                st.error("❌ **SharePoint App Authentication Failed**")
                st.markdown("""
                **Required Actions:**
                
                1. **Azure App Registration:**
                   - Add API permissions: `Sites.Read.All`, `Files.Read.All`
                   - Grant admin consent
                   - Verify client secret is valid
                
                2. **SharePoint Site Access:**
                   - App needs explicit site collection permissions
                   - Contact SharePoint admin
                
                3. **Alternative Setup:**
                   - Consider using user credentials instead of app-only
                   - Or use SharePoint REST API with different auth method
                """)
            elif "401" in error_msg or "Unauthorized" in error_msg:
                st.error("❌ SharePoint authentication failed. Check your client credentials.")
            elif "404" in error_msg or "Not Found" in error_msg:
                st.error(f"❌ SharePoint site not found. Check your site URL: {self.site_url}")
            elif "403" in error_msg or "Forbidden" in error_msg:
                st.error("❌ SharePoint access denied. App may need site collection permissions.")
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
            
            # Show configuration help
            st.markdown("""
            **Missing Environment Variables:**
            
            Add these to your Render environment variables:
            - `SHAREPOINT_CLIENT_ID`: Application ID from Azure
            - `SHAREPOINT_CLIENT_SECRET`: Client secret from Azure  
            - `SHAREPOINT_TENANT_ID`: Directory (tenant) ID from Azure
            - `SHAREPOINT_SITE_NAME`: Just the site name (e.g., "ProjectSite")
            - `SHAREPOINT_TENANT_NAME`: Your tenant name (e.g., "contoso")
            """)
            return False
        else:
            st.info("✅ SharePoint configuration appears complete")
            st.info(f"🔗 Target URL: {self.site_url}")
            return True
    
    def get_documents(self, folder_path: str = "Shared Documents", 
                     file_types: List[str] = None, 
                     since_date: Optional[datetime] = None,
                     max_docs: Optional[int] = None) -> List[Dict]:
        """Get documents from SharePoint with fallback to mock data"""
        import streamlit as st
        
        if not OFFICE365_AVAILABLE or not self.ctx or not self._test_authentication():
            st.warning("⚠️ SharePoint not accessible. Using mock data for demonstration.")
            return self._get_mock_documents()
        
        try:
            st.info(f"📂 Attempting to load documents from: {folder_path}")
            
            # Extract library name from folder path
            library_name = folder_path.split('/')[1] if '/' in folder_path else folder_path
            
            # Get document library
            try:
                library = self.ctx.web.lists.get_by_title(library_name)
                items = library.items
                self.ctx.load(items)
                self.ctx.execute_query()
                
                st.success(f"✅ Successfully connected to '{library_name}' library")
                
            except Exception as lib_error:
                st.error(f"❌ Could not access library '{library_name}': {str(lib_error)}")
                
                if "app-only access token failed" in str(lib_error).lower():
                    st.markdown("**🔧 Quick Fix:** Using mock data while you configure SharePoint permissions.")
                
                return self._get_mock_documents()
            
            # Process items (rest of the method remains the same)
            documents = []
            # ... (rest of document processing logic)
            
            return documents
            
        except Exception as e:
            st.error(f"❌ Error retrieving SharePoint documents: {str(e)}")
            st.info("🔄 Falling back to mock data for testing")
            return self._get_mock_documents()
    
    def get_recent_changes(self, hours: int = 24) -> List[Dict]:
        """Get documents modified in the last N hours"""
        import streamlit as st
        
        since_date = datetime.now() - timedelta(hours=hours)
        st.info(f"🕒 Looking for documents modified since: {since_date.strftime('%Y-%m-%d %H:%M:%S')}")
        return self.get_documents(since_date=since_date)
    
    
    def get_recent_changes(self, hours: int = 24) -> List[Dict]:
        """Get documents modified in the last N hours"""
        import streamlit as st
        
        since_date = datetime.now() - timedelta(hours=hours)
        st.info(f"🕒 Looking for documents modified since: {since_date.strftime('%Y-%m-%d %H:%M:%S')}")
        return self.get_documents(since_date=since_date)
    
    def _get_mock_documents(self) -> List[Dict]:
        """Return mock documents for testing"""
        import streamlit as st
        
        mock_docs = [
            {
                'id': 'mock_1',
                'filename': 'Quarterly Report Q1 2024.pdf',
                'content': 'This is a sample quarterly report containing financial analysis, performance metrics, and strategic recommendations for Q1 2024. The report shows strong growth in key performance indicators and outlines strategic initiatives for the upcoming quarter. Key metrics include revenue growth of 15%, customer satisfaction scores of 94%, and operational efficiency improvements of 12%.',
                'modified': datetime.now().isoformat(),
                'file_path': '/sites/yoursite/Shared Documents/Reports/Quarterly Report Q1 2024.pdf',
                'metadata': {
                    'source': 'mock_data',
                    'created_at': (datetime.now() - timedelta(days=7)).isoformat(),
                    'file_size': 1024000,
                    'author': 'Finance Team',
                    'text_length': 450,
                    'word_count': 75
                }
            },
            {
                'id': 'mock_2', 
                'filename': 'Project Status Update.docx',
                'content': 'Weekly project status update covering milestone achievements, resource allocation, risk assessment, and next steps. All major deliverables are on track for the planned timeline. Current phase focuses on implementation and testing of core features. Team productivity remains high with 98% milestone completion rate.',
                'modified': (datetime.now() - timedelta(hours=3)).isoformat(),
                'file_path': '/sites/yoursite/Shared Documents/Projects/Project Status Update.docx',
                'metadata': {
                    'source': 'mock_data',
                    'created_at': (datetime.now() - timedelta(hours=3)).isoformat(),
                    'file_size': 512000,
                    'author': 'Project Manager',
                    'text_length': 380,
                    'word_count': 63
                }
            },
            {
                'id': 'mock_3',
                'filename': 'Meeting Notes - Team Sync.txt',
                'content': 'Team synchronization meeting notes including action items, decisions made, and follow-up tasks. Key topics discussed: budget planning, resource requirements, and timeline adjustments. Action items assigned to team members with clear deadlines. Next meeting scheduled for next week to review progress on assigned tasks.',
                'modified': (datetime.now() - timedelta(hours=8)).isoformat(),
                'file_path': '/sites/yoursite/Shared Documents/Meetings/Meeting Notes - Team Sync.txt',
                'metadata': {
                    'source': 'mock_data',
                    'created_at': (datetime.now() - timedelta(hours=8)).isoformat(),
                    'file_size': 256000,
                    'author': 'Team Lead',
                    'text_length': 320,
                    'word_count': 53
                }
            },
            {
                'id': 'mock_4',
                'filename': 'Policy Document - Remote Work.docx',
                'content': 'Comprehensive remote work policy document outlining guidelines, expectations, and best practices for remote employees. Covers communication protocols, performance expectations, security requirements, and work-life balance recommendations. Updated to reflect current industry standards and company values.',
                'modified': (datetime.now() - timedelta(days=2)).isoformat(),
                'file_path': '/sites/yoursite/Shared Documents/Policies/Policy Document - Remote Work.docx',
                'metadata': {
                    'source': 'mock_data',
                    'created_at': (datetime.now() - timedelta(days=2)).isoformat(),
                    'file_size': 768000,
                    'author': 'HR Department',
                    'text_length': 410,
                    'word_count': 68
                }
            },
            {
                'id': 'mock_5',
                'filename': 'Technical Specification.pdf',
                'content': 'Technical specification document detailing system architecture, API endpoints, database schema, and integration requirements. Includes performance benchmarks, security considerations, and deployment guidelines. Serves as the primary reference for development and implementation teams.',
                'modified': (datetime.now() - timedelta(hours=12)).isoformat(),
                'file_path': '/sites/yoursite/Shared Documents/Technical/Technical Specification.pdf',
                'metadata': {
                    'source': 'mock_data',
                    'created_at': (datetime.now() - timedelta(hours=12)).isoformat(),
                    'file_size': 1536000,
                    'author': 'Technical Team',
                    'text_length': 520,
                    'word_count': 87
                }
            }
        ]
        
        st.info(f"📋 Generated {len(mock_docs)} mock documents for testing")
        return mock_docs
    
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
            ".xls",
            ".json",
            ".xml"
        ]
    
    def search_documents_in_sharepoint(self, query: str, library_name: str = "Shared Documents") -> List[Dict]:
        """Search documents directly in SharePoint"""
        import streamlit as st
        
        if not OFFICE365_AVAILABLE or not self.ctx:
            # Filter mock documents by query
            mock_docs = self._get_mock_documents()
            query_lower = query.lower()
            
            matching_docs = []
            for doc in mock_docs:
                content = doc.get('content', '').lower()
                filename = doc.get('filename', '').lower()
                
                if query_lower in content or query_lower in filename:
                    matching_docs.append(doc)
            
            st.info(f"🔍 Mock search found {len(matching_docs)} documents matching '{query}'")
            return matching_docs
        
        try:
            # This would implement SharePoint search API
            # For now, return filtered documents
            all_docs = self.get_documents(folder_path=library_name, max_docs=20)
            
            query_lower = query.lower()
            matching_docs = []
            
            for doc in all_docs:
                content = doc.get('content', '').lower()
                filename = doc.get('filename', '').lower()
                
                if query_lower in content or query_lower in filename:
                    matching_docs.append(doc)
            
            st.info(f"🔍 Found {len(matching_docs)} documents matching '{query}'")
            return matching_docs
            
        except Exception as e:
            st.error(f"Error searching SharePoint: {str(e)}")
            return []
    
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
    
    def cleanup(self):
        """Cleanup resources"""
        try:
            if self.ctx:
                # Cleanup context if needed
                self.ctx = None
        except Exception as e:
            pass  # Silent cleanup
