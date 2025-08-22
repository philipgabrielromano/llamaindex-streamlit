# utils/sharepoint_client.py (Updated for your specific site)
import os
import requests
from datetime import datetime, timedelta
from typing import List, Dict, Optional
import json

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
    """SharePoint client configured for your specific site"""
    
    def __init__(self):
        self.client_id = os.getenv("SHAREPOINT_CLIENT_ID")
        self.client_secret = os.getenv("SHAREPOINT_CLIENT_SECRET") 
        self.tenant_id = os.getenv("SHAREPOINT_TENANT_ID")
        self.site_name = os.getenv("SHAREPOINT_SITE_NAME")
        self.tenant_name = os.getenv("SHAREPOINT_TENANT_NAME", "goodwillgoodskills")
        
        self.site_url = f"https://{self.tenant_name}.sharepoint.com/sites/{self.site_name}"
        self.ctx = None
        self.auth_tested = False
        
        if OFFICE365_AVAILABLE and all([self.client_id, self.client_secret, self.tenant_id, self.site_name]):
            self._initialize_client()
    
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
        """Get actual document libraries from your SharePoint site"""
        import streamlit as st
        
        if not OFFICE365_AVAILABLE or not self.ctx:
            # Return the actual libraries we found
            return [
                "Documents",           # ✅ This exists in your site
                "Form Templates",      # ✅ This exists in your site  
                "Site Assets",         # ✅ This exists in your site
                "Style Library",       # ✅ This exists in your site
                "Teams Wiki Data"      # ✅ This exists in your site
            ]
        
        try:
            # Get all lists from SharePoint site
            lists = self.ctx.web.lists
            self.ctx.load(lists)
            self.ctx.execute_query()
            
            libraries = []
            for lst in lists:
                try:
                    list_props = lst.properties
                    # Check if it's a document library (BaseTemplate 101)
                    if list_props.get('BaseTemplate') == 101:
                        library_title = list_props.get('Title', 'Unknown')
                        if library_title and not library_title.startswith('_'):
                            libraries.append(library_title)
                except Exception:
                    continue
            
            if libraries:
                st.success(f"📚 Found {len(libraries)} document libraries: {', '.join(libraries)}")
                return libraries
            else:
                st.info("📚 No document libraries found, using defaults")
                return ["Documents", "Site Assets"]
            
        except Exception as e:
            st.warning(f"Could not get libraries: {str(e)}")
            # Return the actual libraries we know exist
            return ["Documents", "Form Templates", "Site Assets", "Style Library", "Teams Wiki Data"]
    
    def test_connection(self) -> bool:
        """Test SharePoint connection"""
        import streamlit as st
        
        if not OFFICE365_AVAILABLE:
            st.warning("⚠️ Office365 client not available")
            return False
        
        if not self.ctx:
            st.error("❌ SharePoint client not initialized")
            return False
        
        try:
            # Test basic connection
            web = self.ctx.web
            self.ctx.load(web)
            self.ctx.execute_query()
            
            st.success(f"✅ SharePoint connection successful! Site: {web.title}")
            
            # Test accessing the main Documents library
            try:
                documents_lib = self.ctx.web.lists.get_by_title("Documents")
                items = documents_lib.items
                self.ctx.load(items)
                self.ctx.execute_query()
                
                st.success(f"✅ Can access 'Documents' library - found {len(items)} items")
                
                # Show sample files
                if len(items) > 0:
                    st.info("📄 Sample files found:")
                    for i, item in enumerate(items[:3]):
                        try:
                            filename = item.properties.get('FileLeafRef', f'Item {i+1}')
                            st.write(f"  • {filename}")
                        except:
                            st.write(f"  • Item {i+1}")
                
            except Exception as lib_error:
                st.warning(f"⚠️ Could not access 'Documents' library: {str(lib_error)}")
            
            self.auth_tested = True
            return True
            
        except Exception as e:
            st.error(f"❌ SharePoint connection test failed: {str(e)}")
            return False
    
    def get_documents
