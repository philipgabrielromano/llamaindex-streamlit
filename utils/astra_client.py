# utils/astra_client.py
import streamlit as st
import os
from typing import List, Dict, Optional
from datetime import datetime

from llama_index.core import VectorStoreIndex, Document
from llama_index.vector_stores.astra_db import AstraDBVectorStore

class AstraClient:
    """Handles Astra DB vector store operations"""
    
    def __init__(self):
        self.token = os.getenv("ASTRA_DB_TOKEN")
        self.endpoint = os.getenv("ASTRA_DB_ENDPOINT")
        self.collection_name = os.getenv("ASTRA_COLLECTION_NAME", "documents")
        
        if not all([self.token, self.endpoint]):
            raise ValueError("Missing required Astra DB configuration")
        
        self.vector_store = None
        self.index = None
        self._initialize_store()
    
    def _initialize_store(self):
        """Initialize Astra DB vector store"""
        try:
            self.vector_store = AstraDBVectorStore(
                token=self.token,
                api_endpoint=self.endpoint,
                collection_name=self.collection_name,
                embedding_dimension=1536,  # OpenAI ada-002 dimension
            )
            
            self.index = VectorStoreIndex.from_vector_store(self.vector_store)
            
        except Exception as e:
            st.error(f"Failed to initialize Astra DB: {str(e)}")
    
    def test_connection(self) -> bool:
        """Test Astra DB connection"""
        try:
            if not self.vector_store:
                return False
            
            # Try a simple operation
            # This would depend on the specific Astra DB client methods available
            return True
            
        except Exception as e:
            st.error(f"Astra DB connection test failed: {str(e)}")
            return False
    
    def insert_documents(self, documents: List[Document]) -> Dict[str, int]:
        """Insert documents into Astra DB"""
        try:
            if not self.index:
                raise Exception("Astra DB index not initialized")
            
            successful_inserts = 0
            failed_inserts = 0
            
            for doc in documents:
                try:
                    # Add processing timestamp
                    doc.metadata['indexed_at'] = datetime.now().isoformat()
                    
                    # Insert document
                    self.index.insert(doc)
                    successful_inserts += 1
                    
                except Exception as e:
                    st.warning(f"Failed to insert document {doc.metadata.get('filename', 'Unknown')}: {str(e)}")
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
            if not self.index:
                raise Exception("Astra DB index not initialized")
            
            # Configure query engine
            query_engine = self.index.as_query_engine(
                similarity_top_k=top_k,
                response_mode=response_mode
            )
            
            # Execute query
            response = query_engine.query(query)
            
            # Format response
            result = {
                'response': str(response),
                'sources': [],
                'query': query,
                'timestamp': datetime.now().isoformat()
            }
            
            # Add source information if available
            if hasattr(response, 'source_nodes') and response.source_nodes:
                for node in response.source_nodes:
                    source_info = {
                        'text': node.text,
                        'score': getattr(node, 'score', 0),
                        'metadata': node.metadata
                    }
                    result['sources'].append(source_info)
            
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
            # This would require direct Astra DB API calls
            # For now, return placeholder stats
            return {
                'document_count': 'N/A',
                'collection_name': self.collection_name,
                'status': 'active' if self.vector_store else 'inactive',
                'last_updated': datetime.now().isoformat()
            }
            
        except Exception as e:
            st.error(f"Error getting collection stats: {str(e)}")
            return {
                'document_count': 'Error',
                'collection_name': self.collection_name,
                'status': 'error',
                'last_updated': datetime.now().isoformat()
            }
    
    def delete_documents(self, document_ids: List[str]) -> Dict[str, int]:
        """Delete documents by ID"""
        try:
            # This would require implementing document deletion
            # Placeholder implementation
            return {
                'deleted': 0,
                'failed': len(document_ids),
                'total': len(document_ids)
            }
            
        except Exception as e:
            st.error(f"Error deleting documents: {str(e)}")
            return {
                'deleted': 0,
                'failed': len(document_ids),
                'total': len(document_ids)
            }
    
    def validate_configuration(self) -> Dict[str, bool]:
        """Validate Astra DB configuration"""
        return {
            'token': bool(self.token),
            'endpoint': bool(self.endpoint),
            'collection_name': bool(self.collection_name),
            'vector_store_initialized': bool(self.vector_store),
            'index_initialized': bool(self.index)
        }
