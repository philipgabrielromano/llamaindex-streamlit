# utils/astra_client.py
import streamlit as st
import os
from typing import List, Dict, Optional
from datetime import datetime

# Use astrapy directly instead of LlamaIndex vector store
try:
    import astrapy
    ASTRA_AVAILABLE = True
except ImportError:
    ASTRA_AVAILABLE = False
    astrapy = None

# Try LlamaIndex imports as fallback
try:
    from llama_index import Document
    DOCUMENT_AVAILABLE = True
except ImportError:
    try:
        from llama_index.core import Document
        DOCUMENT_AVAILABLE = True
    except ImportError:
        Document = None
        DOCUMENT_AVAILABLE = False

class AstraClient:
    """Handles Astra DB operations using astrapy directly"""
    
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
            self.client = astrapy.DataAPIClient(self.token)
            self.db = self.client.get_database_by_api_endpoint(self.endpoint)
            
            # Try to get or create collection
            try:
                self.collection = self.db.get_collection(self.collection_name)
            except:
                # Collection might not exist, create it
                self.collection = self.db.create_collection(
                    self.collection_name,
                    dimension=1536,  # OpenAI ada-002 dimension
                    metric="cosine"
                )
                
        except Exception as e:
            st.error(f"Failed to initialize Astra DB: {str(e)}")
    
    def test_connection(self) -> bool:
        """Test Astra DB connection"""
        try:
            if not self.db:
                return False
            
            # Try to list collections as a connection test
            collections = self.db.list_collection_names()
            return True
            
        except Exception as e:
            st.error(f"Astra DB connection test failed: {str(e)}")
            return False
    
    def insert_documents(self, documents) -> Dict[str, int]:
        """Insert documents into Astra DB"""
        try:
            if not self.collection:
                raise Exception("Astra DB collection not initialized")
            
            successful_inserts = 0
            failed_inserts = 0
            
            for doc in documents:
                try:
                    # Generate embedding (placeholder - you'd use OpenAI here)
                    doc_data = {
                        "_id": f"doc_{successful_inserts}_{int(datetime.now().timestamp())}",
                        "content": doc.get('content', '') if isinstance(doc, dict) else str(doc),
                        "metadata": doc.get('metadata', {}) if isinstance(doc, dict) else {},
                        "$vector": [0.0] * 1536  # Placeholder vector
                    }
                    
                    # Insert document
                    self.collection.insert_one(doc_data)
                    successful_inserts += 1
                    
                except Exception as e:
                    st.warning(f"Failed to insert document: {str(e)}")
                    failed_inserts += 1
            
            return {
                'successful': successful_inserts,
                'failed': failed_inserts,
                'total': len(documents)
            }
            
        except Exception as e:
            st.error(f"Error inserting documents: {str(e)}")
            return {'successful': 0, 'failed': len(documents), 'total': len(documents)}
    
    def search_documents(self, query: str, top_k: int = 5, 
                        response_mode: str = "compact") -> Dict:
        """Search documents in Astra DB"""
        try:
            if not self.collection:
                raise Exception("Astra DB collection not initialized")
            
            # For now, return a placeholder response
            # In a full implementation, you'd generate query embeddings and do vector search
            
            result = {
                'response': f"Search functionality is being implemented. Query: '{query}'",
                'sources': [],
                'query': query,
                'timestamp': datetime.now().isoformat()
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
