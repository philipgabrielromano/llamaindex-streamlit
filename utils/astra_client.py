# utils/astra_client.py (Updated for astrapy 1.5.2)
import streamlit as st
import os
from typing import List, Dict, Optional
from datetime import datetime
import json

# Use astrapy with correct imports for version 1.5.2
try:
    from astrapy import DataAPIClient
    ASTRA_AVAILABLE = True
except ImportError:
    try:
        # Fallback import
        from astrapy.client import DataAPIClient
        ASTRA_AVAILABLE = True
    except ImportError:
        DataAPIClient = None
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
    """Handles Astra DB operations using astrapy 1.5.2"""
    
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
        """Initialize Astra DB client for astrapy 1.5.2"""
        try:
            # Initialize the Data API client
            self.client = DataAPIClient(self.token)
            
            # Get database using the endpoint
            self.db = self.client.get_database_by_api_endpoint(self.endpoint)
            
            st.success("✅ Astra DB client initialized successfully!")
            
            # Initialize collection
            self._initialize_collection()
                
        except Exception as e:
            st.error(f"Failed to initialize Astra DB client: {str(e)}")
            self.db = None
    
    def _initialize_collection(self):
        """Initialize or create the collection"""
        try:
            if not self.db:
                return
            
            # List existing collections
            existing_collections = self.db.list_collection_names()
            
            if self.collection_name in existing_collections:
                # Get existing collection
                self.collection = self.db.get_collection(self.collection_name)
                st.success(f"✅ Connected to existing collection: {self.collection_name}")
            else:
                # Create new collection
                try:
                    self.collection = self.db.create_collection(
                        self.collection_name,
                        dimension=1536,  # OpenAI ada-002 dimension
                        metric="cosine"
                    )
                    st.success(f"✅ Created new collection: {self.collection_name}")
                except Exception as create_error:
                    st.warning(f"⚠️ Could not create collection: {str(create_error)}")
                    st.info("💡 You may need to create the collection manually in the Astra DB console.")
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
            collections = self.db.list_collection_names()
            st.success(f"✅ Astra DB connection successful! Found {len(collections)} collections: {', '.join(collections[:3])}")
            return True
            
        except Exception as e:
            st.error(f"❌ Astra DB connection test failed: {str(e)}")
            return False
    
    def insert_documents(self, documents) -> Dict[str, int]:
        """Insert documents into Astra DB"""
        try:
            if not self.collection:
                # Provide helpful guidance
                st.warning("⚠️ No collection available for document storage.")
                st.info("💡 Please ensure your collection exists or can be created.")
                return {
                    'successful': 0,
                    'failed': len(documents) if documents else 0,
                    'total': len(documents) if documents else 0,
                    'message': 'Collection not available'
                }
            
            successful_inserts = 0
            failed_inserts = 0
            
            # Process documents in batches for better performance
            batch_size = 20
            document_batches = [documents[i:i + batch_size] for i in range(0, len(documents), batch_size)]
            
            for batch_num, batch in enumerate(document_batches):
                try:
                    batch_docs = []
                    
                    for i, doc in enumerate(batch):
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
                            doc_id = f"doc_{batch_num}_{i}_{int(datetime.now().timestamp())}"
                            doc_data = {
                                "_id": doc_id,
                                "content": content[:10000],  # Limit content size
                                "metadata": json.dumps(metadata) if isinstance(metadata, dict) else str(metadata),
                                "indexed_at": datetime.now().isoformat(),
                                "filename": metadata.get('filename', 'Unknown') if isinstance(metadata, dict) else 'Unknown',
                                "source": metadata.get('source', 'unknown') if isinstance(metadata, dict) else 'unknown'
                                # Note: Not including $vector for now - would need OpenAI embeddings
                            }
                            
                            batch_docs.append(doc_data)
                            
                        except Exception as doc_error:
                            st.warning(f"Failed to prepare document {i} in batch {batch_num}: {str(doc_error)}")
                            failed_inserts += 1
                    
                    # Insert batch
                    if batch_docs:
                        try:
                            result = self.collection.insert_many(batch_docs)
                            successful_inserts += len(batch_docs)
                            st.info(f"✅ Inserted batch {batch_num + 1}/{len(document_batches)} ({len(batch_docs)} documents)")
                        except Exception as batch_error:
                            st.error(f"Failed to insert batch {batch_num}: {str(batch_error)}")
                            failed_inserts += len(batch_docs)
                
                except Exception as batch_prep_error:
                    st.error(f"Error preparing batch {batch_num}: {str(batch_prep_error)}")
                    failed_inserts += len(batch)
            
            return {
                'successful': successful_inserts,
                'failed': failed_inserts,
                'total': len(documents) if documents else 0
            }
            
        except Exception as e:
            st.error(f"Error inserting documents: {str(e)}")
            return {
                'successful': 0, 
                'failed': len(documents) if documents else 0, 
                'total': len(documents) if documents else 0
            }
    
    def search_documents(self, query: str, top_k: int = 5, 
                        response_mode: str = "compact") -> Dict:
        """Search documents in Astra DB"""
        try:
            if not self.collection:
                return {
                    'response': "Document search not available. Collection not initialized.",
                    'sources': [],
                    'query': query,
                    'timestamp': datetime.now().isoformat()
                }
            
            # Basic text search (since we don't have vector embeddings yet)
            try:
                # Search by content or filename containing query terms
                search_filter = {
                    "$or": [
                        {"content": {"$regex": query}},
                        {"filename": {"$regex": query}}
                    ]
                }
                
                results = self.collection.find(filter=search_filter, limit=top_k)
                
                sources = []
                for result in results:
                    sources.append({
                        'text': result.get('content', '')[:500] + "...",
                        'metadata': json.loads(result.get('metadata', '{}')) if result.get('metadata') else {},
                        'filename': result.get('filename', 'Unknown'),
                        'score': 0.8  # Placeholder score
                    })
                
                response_text = f"Found {len(sources)} documents related to '{query}'. "
                if sources:
                    response_text += f"Top result is from {sources[0].get('filename', 'Unknown file')}."
                
                return {
                    'response': response_text,
                    'sources': sources,
                    'query': query,
                    'timestamp': datetime.now().isoformat(),
                    'search_type': 'text_search'
                }
                
            except Exception as search_error:
                st.warning(f"Search error: {str(search_error)}")
                return {
                    'response': f"Search functionality is available but encountered an error: {str(search_error)}",
                    'sources': [],
                    'query': query,
                    'timestamp': datetime.now().isoformat()
                }
            
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
                    # Get document count
                    count_result = self.collection.estimated_document_count()
                    stats['document_count'] = count_result
                    
                    # Get sample document to check structure
                    sample = self.collection.find_one({})
                    stats['has_documents'] = bool(sample)
                    
                except Exception as count_error:
                    st.warning(f"Could not get collection stats: {str(count_error)}")
                    stats['document_count'] = 'Unknown'
                    stats['has_documents'] = False
            else:
                stats['document_count'] = 0
                stats['has_documents'] = False
            
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
            'client_initialized': bool(self.client),
            'db_initialized': bool(self.db),
            'collection_initialized': bool(self.collection)
        }
    
    def cleanup(self):
        """Cleanup resources"""
        try:
            # Clean up any resources if needed
            self.collection = None
            self.db = None
            self.client = None
        except Exception as e:
            st.warning(f"Error during cleanup: {str(e)}")
