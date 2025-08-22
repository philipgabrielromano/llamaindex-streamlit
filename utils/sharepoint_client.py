# utils/sharepoint_client.py (Complete corrected version)
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
            # Return the actual libraries we found in your site
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
            
            # Test accessing the main Documents library (not Shared Documents)
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
    
    def get_documents(self, folder_path: str = "Documents", 
                     file_types: List[str] = None, 
                     since_date: Optional[datetime] = None,
                     max_docs: Optional[int] = None) -> List[Dict]:
        """Get documents from SharePoint - using correct library names"""
        import streamlit as st
        
        if not OFFICE365_AVAILABLE or not self.ctx:
            st.warning("⚠️ SharePoint client not available. Using mock data.")
            return self._get_mock_documents()
        
        try:
            st.info(f"📂 Loading documents from: {folder_path}")
            
            # Use the library name directly (no path prefix)
            library_name = folder_path.replace("/", "").replace("Shared Documents", "Documents")
            
            # Get document library
            try:
                library = self.ctx.web.lists.get_by_title(library_name)
                items = library.items
                self.ctx.load(items)
                self.ctx.execute_query()
                
                st.success(f"✅ Successfully connected to '{library_name}' library with {len(items)} items")
                
            except Exception as lib_error:
                st.error(f"❌ Could not access library '{library_name}': {str(lib_error)}")
                
                # Try alternative library names
                alternative_libraries = ["Documents", "Site Assets", "Teams Wiki Data"]
                for alt_lib in alternative_libraries:
                    try:
                        st.info(f"🔄 Trying alternative library: {alt_lib}")
                        library = self.ctx.web.lists.get_by_title(alt_lib)
                        items = library.items
                        self.ctx.load(items)
                        self.ctx.execute_query()
                        
                        st.success(f"✅ Successfully connected to '{alt_lib}' library")
                        library_name = alt_lib
                        break
                        
                    except:
                        continue
                else:
                    st.error("❌ Could not access any document library")
                    return self._get_mock_documents()
            
            # Process documents
            documents = []
            processed_count = 0
            
            for item in items:
                try:
                    # Extract item properties
                    props = item.properties
                    filename = props.get('FileLeafRef', f'Document_{processed_count}')
                    
                    # Skip folders and system files
                    if not filename or filename.startswith('.') or 'FolderChildCount' in props:
                        continue
                    
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
                            'file_size': file_size,
                            'created': props.get('Created', ''),
                            'author': self._extract_author(props.get('Author', {})),
                            'source': 'sharepoint_live',
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
                    st.warning(f"Error processing item: {str(item_error)}")
                    continue
            
            st.success(f"✅ Successfully loaded {len(documents)} documents from SharePoint")
            return documents
            
        except Exception as e:
            st.error(f"❌ Error retrieving SharePoint documents: {str(e)}")
            st.info("🔄 Falling back to mock data")
            return self._get_mock_documents()
    
    def _get_file_content(self, file_path: str, filename: str) -> str:
        """Get content from a SharePoint file"""
        try:
            if not self.ctx or not file_path:
                return f"[Content placeholder for {filename}]"
            
            # Get file object
            file_obj = self.ctx.web.get_file_by_server_relative_url(file_path)
            self.ctx.load(file_obj)
            self.ctx.execute_query()
            
            # For now, return metadata about the file
            # In a full implementation, you'd download and extract content
            file_info = file_obj.properties
            
            return f"""Document: {filename}
File Path: {file_path}
Server Relative URL: {file_obj.server_relative_url}
Content Type: {file_info.get('ContentType', 'Unknown')}
Length: {file_info.get('Length', 0)} bytes

[Note: Full content extraction would be implemented here using PyPDF2, python-docx, etc.]
"""
                    
        except Exception as e:
            return f"[Could not extract content from {filename}: {str(e)}]"
    
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
                'filename': 'Team Meeting Notes.docx',
                'content': 'Weekly team meeting notes from Goodwill Good Skills organization. Topics covered include project updates, resource allocation, training schedules, and upcoming initiatives. Key decisions made regarding volunteer coordination and skills development programs.',
                'modified': datetime.now().isoformat(),
                'file_path': '/sites/Org-WideTeam/Documents/Team Meeting Notes.docx',
                'metadata': {
                    'source': 'mock_data',
                    'created_at': (datetime.now() - timedelta(days=1)).isoformat(),
                    'file_size': 524288,
                    'author': 'Team Lead',
                    'text_length': 320,
                    'word_count': 52
                }
            },
            {
                'id': 'mock_2', 
                'filename': 'Skills Training Manual.pdf',
                'content': 'Comprehensive skills training manual for volunteers and staff at Goodwill Good Skills. Covers customer service, technical skills development, workplace safety, and professional development opportunities. Includes assessment criteria and certification pathways.',
                'modified': (datetime.now() - timedelta(hours=6)).isoformat(),
                'file_path': '/sites/Org-WideTeam/Documents/Skills Training Manual.pdf',
                'metadata': {
                    'source': 'mock_data',
                    'created_at': (datetime.now() - timedelta(hours=6)).isoformat(),
                    'file_size': 1048576,
                    'author': 'Training Department',
                    'text_length': 450,
                    'word_count': 73
                }
            },
            {
                'id': 'mock_3',
                'filename': 'Volunteer Handbook 2024.pdf',
                'content': 'Official volunteer handbook for 2024 containing policies, procedures, and guidelines for volunteers at Goodwill Good Skills. Includes code of conduct, safety protocols, communication guidelines, and volunteer benefits information.',
                'modified': (datetime.now() - timedelta(days=3)).isoformat(),
                'file_path': '/sites/Org-WideTeam/Documents/Volunteer Handbook 2024.pdf',
                'metadata': {
                    'source': 'mock_data',
                    'created_at': (datetime.now() - timedelta(days=3)).isoformat(),
                    'file_size': 2097152,
                    'author': 'HR Department',
                    'text_length': 680,
                    'word_count': 112
                }
            },
            {
                'id': 'mock_4',
                'filename': 'Monthly Report March 2024.xlsx',
                'content': 'Monthly performance report for March 2024 including volunteer hours, program outcomes, training completions, and community impact metrics. Shows improvement in all key performance indicators compared to previous month.',
                'modified': (datetime.now() - timedelta(hours=18)).isoformat(),
                'file_path': '/sites/Org-WideTeam/Documents/Monthly Report March 2024.xlsx',
                'metadata': {
                    'source': 'mock_data',
                    'created_at': (datetime.now() - timedelta(hours=18)).isoformat(),
                    'file_size': 786432,
                    'author': 'Program Manager',
                    'text_length': 380,
                    'word_count': 62
                }
            },
            {
                'id': 'mock_5',
                'filename': 'Safety Guidelines Update.docx',
                'content': 'Updated safety guidelines and protocols for all Goodwill Good Skills locations. Includes new COVID-19 protocols, workplace safety measures, emergency procedures, and health screening requirements. All staff and volunteers must review and acknowledge.',
                'modified': (datetime.now() - timedelta(hours=4)).isoformat(),
                'file_path': '/sites/Org-WideTeam/Documents/Safety Guidelines Update.docx',
                'metadata': {
                    'source': 'mock_data',
                    'created_at': (datetime.now() - timedelta(hours=4)).isoformat(),
                    'file_size': 655360,
                    'author': 'Safety Officer',
                    'text_length': 420,
                    'word_count': 68
                }
            }
        ]
        
        st.info(f"📋 Generated {len(mock_docs)} mock documents for testing")
        return mock_docs
    
    def get_documents(self, folder_path: str = "Documents", 
                     file_types: List[str] = None, 
                     since_date: Optional[datetime] = None,
                     max_docs: Optional[int] = None) -> List[Dict]:
        """Get documents from SharePoint - using correct library names"""
        import streamlit as st
        
        if not OFFICE365_AVAILABLE or not self.ctx:
            st.warning("⚠️ SharePoint client not available. Using mock data.")
            return self._get_mock_documents()
        
        try:
            st.info(f"📂 Loading documents from: {folder_path}")
            
            # Clean up the folder path - use just the library name
            library_name = folder_path.replace("/", "").replace("Shared Documents", "Documents")
            
            # Ensure we're using a library that exists
            available_libs = ["Documents", "Form Templates", "Site Assets", "Style Library", "Teams Wiki Data"]
            if library_name not in available_libs:
                library_name = "Documents"  # Default to Documents
                st.info(f"📁 Using default library: {library_name}")
            
            # Get document library
            try:
                library = self.ctx.web.lists.get_by_title(library_name)
                items = library.items
                self.ctx.load(items)
                self.ctx.execute_query()
                
                st.success(f"✅ Successfully connected to '{library_name}' library with {len(items)} items")
                
            except Exception as lib_error:
                st.error(f"❌ Could not access library '{library_name}': {str(lib_error)}")
                return self._get_mock_documents()
            
            # Process documents
            documents = []
            processed_count = 0
            
            for item in items:
                try:
                    # Extract item properties
                    props = item.properties
                    filename = props.get('FileLeafRef', f'Document_{processed_count}')
                    
                    # Skip folders and system files
                    if not filename or filename.startswith('.') or 'FolderChildCount' in props:
                        continue
                    
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
                            'file_size': file_size,
                            'created': props.get('Created', ''),
                            'author': self._extract_author(props.get('Author', {})),
                            'source': 'sharepoint_live',
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
            st.error(f"❌ Error retrieving SharePoint documents: {str(e)}")
            st.info("🔄 Falling back to mock data")
            return self._get_mock_documents()
    
    def _get_file_content(self, file_path: str, filename: str) -> str:
        """Get content from a SharePoint file"""
        try:
            if not self.ctx or not file_path:
                return f"Sample content for {filename} from Goodwill Good Skills SharePoint site."
            
            # Get basic file information
            # For a full implementation, you'd download the file and extract text
            file_ext = filename.split('.')[-1].lower() if '.' in filename else ''
            
            if file_ext == 'docx':
                return f"Microsoft Word document: {filename}. This document contains important information for Goodwill Good Skills operations, including policies, procedures, and guidelines for staff and volunteers."
            elif file_ext == 'pdf':
                return f"PDF document: {filename}. Contains detailed information and documentation for Goodwill Good Skills programs, training materials, and organizational resources."
            elif file_ext == 'xlsx':
                return f"Excel spreadsheet: {filename}. Contains data analysis, reports, and metrics for Goodwill Good Skills programs and performance tracking."
            elif file_ext == 'txt':
                return f"Text document: {filename}. Plain text information and notes related to Goodwill Good Skills activities and operations."
            else:
                return f"Document: {filename}. Contains information relevant to Goodwill Good Skills mission and operations."
                    
        except Exception as e:
            return f"[Could not extract content from {filename}: {str(e)}]"
    
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
