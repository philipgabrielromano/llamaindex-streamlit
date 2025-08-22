# utils/sharepoint_client.py (Enhanced debugging version)
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
    """Enhanced SharePoint client with detailed debugging"""
    
    def __init__(self):
        self.client_id = os.getenv("SHAREPOINT_CLIENT_ID")
        self.client_secret = os.getenv("SHAREPOINT_CLIENT_SECRET") 
        self.tenant_id = os.getenv("SHAREPOINT_TENANT_ID")
        self.site_name = os.getenv("SHAREPOINT_SITE_NAME")
        self.tenant_name = os.getenv("SHAREPOINT_TENANT_NAME", "your-tenant")
        
        self.site_url = f"https://{self.tenant_name}.sharepoint.com/sites/{self.site_name}"
        self.ctx = None
        self.auth_tested = False
        self.last_error = None
        
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
            self.last_error = str(e)
            self.ctx = None
    
    def get_debug_info(self) -> Dict:
        """Get detailed debug information"""
        return {
            'client_id': self.client_id[:8] + "..." if self.client_id else "Not set",
            'tenant_id': self.tenant_id[:8] + "..." if self.tenant_id else "Not set", 
            'site_name': self.site_name or "Not set",
            'tenant_name': self.tenant_name or "Not set",
            'site_url': self.site_url,
            'office365_available': OFFICE365_AVAILABLE,
            'context_initialized': bool(self.ctx),
            'last_error': self.last_error,
            'auth_tested': self.auth_tested
        }
    
    def test_connection_detailed(self) -> Dict:
        """Detailed connection test with step-by-step feedback"""
        import streamlit as st
        
        result = {
            'overall_success': False,
            'steps': []
        }
        
        # Step 1: Check configuration
        step1 = {'name': 'Configuration Check', 'success': False, 'message': ''}
        
        if all([self.client_id, self.client_secret, self.tenant_id, self.site_name]):
            step1['success'] = True
            step1['message'] = '✅ All required environment variables are set'
        else:
            step1['message'] = '❌ Missing environment variables'
        
        result['steps'].append(step1)
        
        # Step 2: Check Office365 package
        step2 = {'name': 'Office365 Package', 'success': OFFICE365_AVAILABLE, 'message': ''}
        step2['message'] = '✅ Office365 REST client available' if OFFICE365_AVAILABLE else '❌ Office365 package not available'
        result['steps'].append(step2)
        
        # Step 3: Test authentication
        step3 = {'name': 'SharePoint Authentication', 'success': False, 'message': ''}
        
        if self.ctx and OFFICE365_AVAILABLE:
            try:
                web = self.ctx.web
                self.ctx.load(web)
                self.ctx.execute_query()
                
                step3['success'] = True
                step3['message'] = f'✅ Successfully authenticated to site: {getattr(web, "title", "Unknown")}'
                result['overall_success'] = True
                
            except Exception as e:
                step3['message'] = f'❌ Authentication failed: {str(e)}'
                
                # Provide specific guidance based on error
                error_lower = str(e).lower()
                if 'app-only access token failed' in error_lower:
                    step3['message'] += '\n\n🔧 Fix: Configure app-only permissions in Azure AD'
                elif '401' in error_lower:
                    step3['message'] += '\n\n🔧 Fix: Check client ID and secret'
                elif '403' in error_lower:
                    step3['message'] += '\n\n🔧 Fix: Grant app permission to SharePoint site'
                elif '404' in error_lower:
                    step3['message'] += '\n\n🔧 Fix: Verify site name and URL'
        else:
            step3['message'] = '❌ Client context not initialized'
        
        result['steps'].append(step3)
        
        return result
    
    def get_available_libraries(self) -> List[str]:
        """Get available libraries with enhanced error handling"""
        import streamlit as st
        
        # Always provide defaults
        default_libraries = [
            "Shared Documents",
            "Documents", 
            "Site Assets",
            "Reports",
            "Templates"
        ]
        
        if not OFFICE365_AVAILABLE:
            st.info("📚 Office365 client not available - using default libraries")
            return default_libraries
        
        if not self.ctx:
            st.warning("⚠️ SharePoint client not initialized - using default libraries")
            return default_libraries
        
        try:
            # Test auth first
            web = self.ctx.web
            self.ctx.load(web)
            self.ctx.execute_query()
            
            # Get lists
            lists = self.ctx.web.lists
            self.ctx.load(lists)
            self.ctx.execute_query()
            
            libraries = []
            for lst in lists:
                try:
                    list_props = lst.properties
                    if list_props.get('BaseTemplate') == 101:  # Document library
                        title = list_props.get('Title', '')
                        if title and not title.startswith('_'):
                            libraries.append(title)
                except:
                    continue
            
            if libraries:
                st.success(f"📚 Found {len(libraries)} libraries: {', '.join(libraries)}")
                return libraries
            else:
                st.info("📚 No custom libraries found, using defaults")
                return default_libraries
                
        except Exception as e:
            error_msg = str(e)
            
            if 'app-only access token failed' in error_msg.lower():
                st.error("❌ **App-only authentication failed**")
                st.markdown("""
                **To fix this:**
                
                1. **Azure AD App Registration:**
                   ```
                   • Go to portal.azure.com
                   • Azure Active Directory → App registrations
                   • Your app → API permissions
                   • Add: Sites.FullControl.All (Application permission)
                   • Grant admin consent ✅
                   ```
                
                2. **SharePoint Site Access:**
                   ```
                   • Go to your SharePoint site
                   • Settings → Site permissions  
                   • Advanced permissions → Grant permissions
                   • Add: [your-client-id]@[your-tenant-id]
                   ```
                
                3. **Alternative: Use SharePoint Admin PowerShell:**
                   ```powershell
                   Connect-PnPOnline -Url "https://yourtenant-admin.sharepoint.com" -Interactive
                   Grant-PnPSiteCollectionAppCatalogAccess -Site "https://yourtenant.sharepoint.com/sites/yoursite"
                   ```
                """)
            else:
                st.warning(f"Library access error: {error_msg}")
            
            st.info("📚 Using default libraries for now")
            return default_libraries
    
    def test_connection(self) -> bool:
        """Test connection with detailed feedback"""
        import streamlit as st
        
        # Run detailed test
        test_result = self.test_connection_detailed()
        
        # Display results
        st.markdown("**Connection Test Results:**")
        
        for step in test_result['steps']:
            if step['success']:
                st.success(f"✅ {step['name']}: {step['message']}")
            else:
                st.error(f"❌ {step['name']}: {step['message']}")
        
        return test_result['overall_success']
    
    # Add the rest of your existing methods here...
    def get_documents(self, folder_path: str = "Shared Documents", 
                     file_types: List[str] = None, 
                     since_date: Optional[datetime] = None,
                     max_docs: Optional[int] = None) -> List[Dict]:
        """Get documents with fallback to mock data"""
        import streamlit as st
        
        # Always fall back to mock data for now while fixing auth
        st.info("🎭 Using mock data while SharePoint authentication is being configured")
        return self._get_mock_documents()
    
    def _get_mock_documents(self) -> List[Dict]:
        """Return enhanced mock documents for testing"""
        import streamlit as st
        
        mock_docs = [
            {
                'id': 'mock_1',
                'filename': 'Quarterly Report Q1 2024.pdf',
                'content': 'Executive Summary: This quarterly report presents a comprehensive analysis of our business performance for Q1 2024. Key highlights include revenue growth of 15% year-over-year, successful launch of three new product lines, and expansion into two new markets. Customer satisfaction scores reached 94%, reflecting our commitment to quality service. Operational efficiency improved by 12% through process optimization and technology investments. Looking ahead, we anticipate continued growth in Q2 with projected revenue increase of 8-10%. Strategic initiatives for the upcoming quarter include digital transformation projects, talent acquisition in key areas, and sustainability program implementation.',
                'modified': datetime.now().isoformat(),
                'file_path': '/sites/yoursite/Shared Documents/Reports/Quarterly Report Q1 2024.pdf',
                'metadata': {
                    'source': 'mock_data',
                    'created_at': (datetime.now() - timedelta(days=7)).isoformat(),
                    'file_size': 1024000,
                    'author': 'Finance Team',
                    'text_length': 800,
                    'word_count': 133
                }
            },
            {
                'id': 'mock_2', 
                'filename': 'Project Alpha Status Update.docx',
                'content': 'Project Alpha Weekly Status Report - Week of March 15, 2024. Current phase focuses on implementation and testing of core features. Development team has completed 85% of planned deliverables for this sprint. Key accomplishments include API integration, database optimization, and user interface enhancements. Testing phase revealed minor issues that have been resolved. Resource allocation remains optimal with full team utilization. Risk assessment shows low probability of delays. Next week priorities include final testing, documentation updates, and deployment preparation. Stakeholder feedback has been overwhelmingly positive.',
                'modified': (datetime.now() - timedelta(hours=3)).isoformat(),
                'file_path': '/sites/yoursite/Shared Documents/Projects/Project Alpha Status Update.docx',
                'metadata': {
                    'source': 'mock_data',
                    'created_at': (datetime.now() - timedelta(hours=3)).isoformat(),
                    'file_size': 512000,
                    'author': 'Project Manager',
                    'text_length': 750,
                    'word_count': 125
                }
            },
            {
                'id': 'mock_3',
                'filename': 'Employee Handbook 2024.pdf',
                'content': 'Welcome to our organization! This employee handbook serves as your guide to company policies, procedures, and culture. Our mission is to deliver exceptional value to customers while fostering an inclusive and innovative workplace. Core values include integrity, collaboration, innovation, and customer focus. Employment policies cover equal opportunity, anti-discrimination, workplace safety, and professional development. Benefits package includes health insurance, retirement planning, paid time off, and professional development opportunities. Code of conduct outlines expected behaviors and ethical standards. For questions about policies or procedures, contact Human Resources.',
                'modified': (datetime.now() - timedelta(days=1)).isoformat(),
                'file_path': '/sites/yoursite/Shared Documents/HR/Employee Handbook 2024.pdf',
                'metadata': {
                    'source': 'mock_data',
                    'created_at': (datetime.now() - timedelta(days=1)).isoformat(),
                    'file_size': 2048000,
                    'author': 'HR Department',
                    'text_length': 920,
                    'word_count': 153
                }
            }
        ]
        
        st.info(f"📋 Generated {len(mock_docs)} mock documents for testing")
        return mock_docs
    
    def get_recent_changes(self, hours: int = 24) -> List[Dict]:
        """Get recent changes with mock data filtering"""
        import streamlit as st
        
        since_date = datetime.now() - timedelta(hours=hours)
        st.info(f"🕒 Looking for documents modified since: {since_date.strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Filter mock documents by date
        all_docs = self._get_mock_documents()
        recent_docs = []
        
        for doc in all_docs:
            try:
                doc_date = datetime.fromisoformat(doc['modified'].replace('Z', '+00:00'))
                if doc_date >= since_date:
                    recent_docs.append(doc)
            except:
                recent_docs.append(doc)  # Include if date parsing fails
        
        st.info(f"📄 Found {len(recent_docs)} documents modified in the last {hours} hours")
        return recent_docs
    
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
