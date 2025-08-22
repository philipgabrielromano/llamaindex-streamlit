# utils/sharepoint_client.py (Complete version with missing methods)
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
    
    def get_available_libraries(self) -> List[str]:
        """Get list of available document libraries"""
        import streamlit as st
        
        if not OFFICE365_AVAILABLE or not self.ctx:
            st.info("📚 Using default library list (Office365 client not available)")
            return [
                "Shared Documents",
                "Documents", 
                "Site Assets",
                "Reports",
                "Templates",
                "Archive"
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
                        libraries.append(library_title)
                except Exception as e:
                    continue  # Skip problematic lists
            
            if not libraries:
                # Fallback to common libraries
                libraries = ["Shared Documents", "Documents"]
            
            st.success(f"📚 Found {len(libraries)} document libraries")
            return libraries
            
        except Exception as e:
            st.warning(f"Could not get libraries: {str(e)}, using defaults")
            return [
                "Shared Documents",
                "Documents", 
                "Site Assets"
            ]
    
    def get_folder_structure(self, library_name: str = "Shared Documents") -> Dict:
        """Get folder structure within a library"""
        import streamlit as st
        
        if not OFFICE365_AVAILABLE or not self.ctx:
            return {
                "folders": [
                    f"/{library_name}",
                    f"/{library_name}/Reports",
                    f"/{library_name}/Archive",
                    f"/{library_name}/Templates"
                ],
                "library": library_name
            }
        
        try:
            # Get the document library
            library = self.ctx.web.lists.get_by_title(library_name)
            folders = library.root_folder.folders
            self.ctx.load(folders)
            self.ctx.execute_query()
            
            folder_paths = [f"/{library_name}"]  # Root folder
            
            for folder in folders:
                try:
                    folder_name = folder.properties.get('Name', '')
                    if folder_name and not folder_name.startswith('_'):  # Skip system folders
                        folder_paths.append(f"/{library_name}/{folder_name}")
                except Exception:
                    continue
            
            return {
                "folders": folder_paths,
                "library": library_name
            }
            
        except Exception as e:
            st.warning(f"Could not get folder structure for {library_name}: {str(e)}")
            return {
                "folders": [f"/{library_name}"],
                "library": library_name
            }
    
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
    
    def get_documents(self, folder_path: str = "Shared Documents", 
                     file_types: List[str] = None, 
                     since_date: Optional[datetime] = None,
                     max_docs: Optional[int] = None) -> List[Dict]:
        """Get documents from SharePoint"""
        import streamlit as st
        
        if not OFFICE365_AVAILABLE or not self.ctx:
            st.warning("⚠️ SharePoint client not available. Returning mock data for testing.")
            return self._get_mock_documents()
        
        try:
            st.info(f"📂 Loading documents from: {folder_path}")
            
            # Extract library name from folder path
            library_name = folder_path.split('/')[1] if '/' in folder_path else folder_path
            
            # Get document library
            try:
                library = self.ctx.web.lists.get_by_title(library_name)
                items = library.items
                self.ctx.load(items)
                self.ctx.execute_query()
            except Exception as e:
                st.error(f"❌ Could not access library '{library_name}': {str(e)}")
                return []
            
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
                    file_size = props.get('File_x0020_Size', 0)
                    
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
                    
                    # Get file content (simplified for now)
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
                            'file_size': file_size,
                            'created': props.get('Created', ''),
                            'author': self._extract_author(props.get('Author', {})),
                            'source': 'sharepoint_direct',
                            'site_url': self.site_url,
                            'library': library_name,
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
                    continue  # Skip problematic items
            
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
    
    def _extract_author(self, author_field) -> str:
        """Extract author name from SharePoint author field"""
        try:
            if isinstance(author_field, dict):
                return author_field.get('Title', 'Unknown')
            elif isinstance(author_field, str):
                return author_field
            else:
                return 'Unknown'
        except:
            return 'Unknown'
    
    def _get_file_content(self, file_path: str, filename: str) -> str:
        """Get content from a SharePoint file"""
        try:
            if not self.ctx or not file_path:
                return f"[Content not available for {filename}]"
            
            # For now, return placeholder content
            # In a full implementation, you'd:
            # 1. Download the file content
            # 2. Use your DocumentProcessor to extract text
            
            file_ext = filename.split('.')[-1].lower() if '.' in filename else ''
            
            if file_ext == 'txt':
                return f"Sample text content from {filename}. This would contain the actual file content in a real implementation."
            elif file_ext == 'pdf':
                return f"Sample PDF content from {filename}. This document contains important information about various topics and would be fully extracted in a real implementation."
            elif file_ext == 'docx':
                return f"Sample Word document content from {filename}. This document includes detailed analysis, recommendations, and supporting data that would be properly extracted."
            else:
                return f"Sample content from {filename}. File type: {file_ext}. Content extraction would be implemented based on file type."
                    
        except Exception as e:
            return f"[Could not extract content from {filename}: {str(e)}]"
    
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
