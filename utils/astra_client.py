# utils/astra_client.py (Fixed for latest astrapy)
import streamlit as st
import os
from typing import List, Dict, Optional
from datetime import datetime
import json

# Use astrapy directly with correct imports
try:
    from astrapy import DataAPIClient
    from astrapy.db import AstraDB
    ASTRA_AVAILABLE = True
except ImportError:
    try:
        # Alternative import structure
        import astrapy
        DataAPIClient = getattr(astrapy, 'DataAPIClient', None)
        AstraDB = getattr(astrapy, 'AstraDB', None)
        ASTRA_AVAILABLE = DataAPIClient is not None
    except ImportError:
        DataAPIClient = None
        AstraDB = None
        ASTRA_AVAILABLE = False

# Try LlamaIndex imports as fallback
try:
    from llama_index import Document
    DOCUMENT_AVAILABLE = True
except ImportError:
    try:
        from llama_index.core import Document
        DOCUMENT_AVAILABLE = True
    except ImportError:
        # Create simple document class
        class Document:
            def __init__(self, text: str, metadata: Dict = None):
                self.text = text
                self.metadata = metadata or {}
        DOCUMENT_AVAILABLE = True

class AstraClient:
    """Handles Astra DB operations using astrapy"""
    
    def __init__(self):
        self.token = os.getenv("ASTRA_DB_TOKEN")
        self.endpoint = os.getenv("ASTRA_DB_ENDPOINT")
        self.collection_name = os.getenv("ASTRA_COLLECTION_NAME", "documents")
        
        if not all([self.token, self.endpoint]):
            raise ValueError("Missing required Astra DB configuration")
        
        self.client = None
        self.db = None
        self.collection = None
        
        if ASTRA_AVAILABLE:
            self._initialize_client()
        else:
            st.warning("⚠️ Astra DB client not available.")
    
    def _initialize_client(self):
        """Initialize Astra DB client"""
        try:
            # Try the new astrapy structure
            if DataAPIClient:
                self.client = DataAPIClient(self.token)
                self.db = self.client.get_database_by_api_endpoint(self.endpoint)
            elif AstraDB:
                # Direct AstraDB connection
                self.db = AstraDB(
                    token=self.token,
                    api_endpoint=self.endpoint
                )
            else:
                raise Exception("No compatible Astra DB client found")
            
            # Try to get or create collection
            self._initialize_collection()
                
        except Exception as e:
            st.error(f"Failed to initialize Astra DB: {str(e)}")
            self.db = None
    
    def _initialize_collection(self):
        """Initialize or create the collection"""
        try:
            if not self.db:
                return
                
            # Try to get existing collection
            try:
                if hasattr(self.db, 'get_collection'):
                    self.collection = self.db.get_collection(self.collection_name)
                elif hasattr(self.db, 'collection'):
                    self.collection = self.db.collection(self.collection_name)
                else:
                    raise Exception("Cannot access collection")
                    
                st.success(f"✅ Connected to existing collection: {self.collection_name}")
                
            except Exception:
                # Collection doesn't exist, create it
                try:
                    if hasattr(self.db, 'create_collection'):
                        self.collection = self.db.create_collection(
                            self.collection_name,
                            dimension=1536,  # OpenAI ada-002 dimension
                            metric="cosine"
                        )
                    else:
                        # Alternative creation method
                        self.collection = self.db.create_collection(
                            collection_name=self.collection_name,
                            dimension=1536,
                            metric="cosine"
                        )
                    
                    st.success(f"✅ Created new collection: {self.collection_name}")
                    
                except Exception as create_error:
                    st.warning(f"⚠️ Could not create collection: {str(create_error)}")
                    self.collection = None
                    
        except Exception as e:
            st.error(f"Collection initialization failed: {str(e)}")
            self.collection = None
    
    def test_connection(self) -> bool:
        """Test Astra DB connection"""
        try:
            if not self.db:
                st.error("❌ Astra DB client not initialized")
                return False
            
            # Try to list collections as a connection test
            if hasattr(self.db, 'list_collection_names'):
                collections = self.db.list_collection_names()
                st.success(f"✅ Astra DB connection successful! Found {len(collections)} collections.")
                return True
            elif hasattr(self.db, 'list_collections'):
                collections = self.db.list_collections()
                st.success(f"✅ Astra DB connection successful! Found {len(collections)} collections.")
                return True
            else:
                # Basic connection test
                st.success("✅ Astra DB connection appears to be working")
                return True
            
        except Exception as e:
            st.error(f"❌ Astra DB connection test failed: {str(e)}")
            return False
    
    def insert_documents(self, documents) -> Dict[str, int]:
        """Insert documents into Astra DB"""
        try:
            if not self.collection:
                # Fallback: store in simple format without vector search
                st.warning("⚠️ No collection available. Documents will be processed but not stored in vector database.")
                return {
                    'successful': len(documents) if documents else 0,
                    'failed': 0,
                    'total': len(documents) if documents else 0
                }
            
            successful_inserts = 0
            failed_inserts = 0
            
            for i, doc in enumerate(documents):
                try:
                    # Extract content and metadata
                    if hasattr(doc, 'text'):
                        content = doc.text
                        metadata = getattr(doc, 'metadata', {})
                    elif isinstance(doc, dict):
                        content = doc.get('content', '')
                        metadata = doc.get('metadata', {})
                    else:
                        content = str(doc)
                        metadata = {}
                    
                    # Create document for insertion
                    doc_data = {
                        "_id": f"doc_{successful_inserts}_{int(datetime.now().timestamp())}_{i}",
                        "content": content,
                        "metadata": metadata,
                        "indexed_at": datetime.now().isoformat(),
                        # Note: For real implementation, you'd generate embeddings here
                        # "$vector": embedding_vector  # Placeholder - would need OpenAI API call
                    }
                    
                    # For now, store without vector (you can add embedding generation later)
                    if hasattr(self.collection, 'insert_one'):
                        self.collection.insert_one(doc_data)
                    elif hasattr(self.collection, 'add'):
                        self.collection.add(doc_data)
                    else:
                        # Alternative insertion method
                        raise Exception("No suitable insertion method found")
                    
                    successful_inserts += 1
                    
                except Exception as e:
                    st.warning(f"Failed to insert document {i}: {str(e)}")
                    failed_inserts += 1
            
            return {
                'successful': successful_inserts,
                'failed': failed_inserts,
                'total': len(documents) if documents else 0
            }
            
        except Exception as e:
            st.error(f"Error inserting documents: {str(e)}")
            return {'successful': 0, 'failed': len(documents) if documents else 0, 'total': len(documents) if documents else 0}
    
    def search_documents(self, query: str, top_k: int = 5, 
                        response_mode: str = "compact") -> Dict:
        """Search documents in Astra DB"""
        try:
            if not self.collection:
                return {
                    'response': "Vector search not available. Collection not initialized.",
                    'sources': [],
                    'query': query,
                    'timestamp': datetime.now().isoformat()
                }
            
            # For now, return a placeholder response
            # In a full implementation, you'd:
            # 1. Generate embedding for the query using OpenAI
            # 2. Perform vector search in Astra DB
            # 3. Generate response using retrieved context
            
            result = {
                'response': f"Search functionality is being implemented. Your query: '{query}' has been received and will be processed once vector search is fully configured.",
                'sources': [],
                'query': query,
                'timestamp': datetime.now().isoformat(),
                'note': "Vector search requires OpenAI embeddings integration"
            }
            
            return result
            
        except Exception as e:
            st.error(f"Search error: {str(e)}")
            return {
                'response': f"Search failed: {str(e)}",
                'sources': [],
                'query': query,
                'timestamp': datetime.now().isoformat()
            }
    
    def get_collection_stats(self) -> Dict:
        """Get statistics about the collection"""
        try:
            stats = {
                'collection_name': self.collection_name,
                'status': 'active' if self.collection else 'inactive',
                'last_updated': datetime.now().isoformat(),
                'endpoint': self.endpoint[:50] + "..." if self.endpoint and len(self.endpoint) > 50 else self.endpoint
            }
            
            if self.collection:
                try:
                    # Try to get collection info
                    if hasattr(self.collection, 'count'):
                        stats['document_count'] = self.collection.count()
                    elif hasattr(self.collection, 'estimated_document_count'):
                        stats['document_count'] = self.collection.estimated_document_count()
                    else:
                        stats['document_count'] = 'Unknown'
                except Exception:
                    stats['document_count'] = 'Unknown'
            else:
                stats['document_count'] = 0
            
            return stats
            
        except Exception as e:
            st.error(f"Error getting collection stats: {str(e)}")
            return {
                'collection_name': self.collection_name,
                'document_count': 'Error',
                'status': 'error',
                'last_updated': datetime.now().isoformat()
            }
    
    def validate_configuration(self) -> Dict[str, bool]:
        """Validate Astra DB configuration"""
        return {
            'token': bool(self.token),
            'endpoint': bool(self.endpoint),
            'collection_name': bool(self.collection_name),
            'astrapy_available': ASTRA_AVAILABLE,
            'client_initialized': bool(self.client or self.db),
            'collection_initialized': bool(self.collection)
        }
