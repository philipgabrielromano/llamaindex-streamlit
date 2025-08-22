# utils/document_processor.py
# utils/document_processor.py (Updated imports)
import streamlit as st
import hashlib
import mimetypes
from datetime import datetime
from typing import List, Dict, Optional, Tuple
import PyPDF2
import docx
import pandas as pd
import io
from pathlib import Path
import json

# Try to import additional libraries
try:
    import pptx
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False

try:
    from bs4 import BeautifulSoup
    BS4_AVAILABLE = True
except ImportError:
    BS4_AVAILABLE = False

from llama_index.core import Document
from llama_index.core.text_splitter import RecursiveCharacterTextSplitter

# Rest of your DocumentProcessor class remains the same...

class DocumentProcessor:
    """Handles document processing, chunking, and text extraction"""
    
    def __init__(self, chunk_size: int = 1000, chunk_overlap: int = 200):
        self.chunk_size = chunk_size
        self.chunk_overlap = chunk_overlap
        self.text_splitter = RecursiveCharacterTextSplitter(
            chunk_size=chunk_size,
            chunk_overlap=chunk_overlap,
            separators=["\n\n", "\n", ". ", " ", ""]
        )
    
    def process_uploaded_file(self, uploaded_file) -> Optional[Document]:
        """Process a single uploaded file and return a Document object"""
        try:
            content = uploaded_file.read()
            file_type = self._get_file_type(uploaded_file.name)
            
            # Extract text based on file type
            if file_type == 'pdf':
                text = self._extract_pdf_text(content)
            elif file_type == 'docx':
                text = self._extract_docx_text(content)
            elif file_type == 'txt':
                text = content.decode('utf-8')
            else:
                st.error(f"Unsupported file type: {file_type}")
                return None
            
            # Create document with metadata
            document = Document(
                text=text,
                metadata={
                    'filename': uploaded_file.name,
                    'file_type': file_type,
                    'file_size': len(content),
                    'processed_at': datetime.now().isoformat(),
                    'chunk_size': self.chunk_size,
                    'chunk_overlap': self.chunk_overlap,
                    'document_hash': self._generate_hash(content),
                    'source': 'manual_upload'
                }
            )
            
            return document
            
        except Exception as e:
            st.error(f"Error processing file {uploaded_file.name}: {str(e)}")
            return None
    
    def process_sharepoint_documents(self, documents: List[Dict]) -> List[Document]:
        """Process documents from SharePoint"""
        processed_docs = []
        
        for doc_info in documents:
            try:
                # Create document from SharePoint info
                document = Document(
                    text=doc_info.get('content', ''),
                    metadata={
                        'filename': doc_info.get('filename', 'Unknown'),
                        'sharepoint_id': doc_info.get('id'),
                        'modified_date': doc_info.get('modified'),
                        'file_path': doc_info.get('file_path'),
                        'processed_at': datetime.now().isoformat(),
                        'chunk_size': self.chunk_size,
                        'chunk_overlap': self.chunk_overlap,
                        'source': 'sharepoint'
                    }
                )
                processed_docs.append(document)
                
            except Exception as e:
                st.error(f"Error processing SharePoint document {doc_info.get('filename', 'Unknown')}: {str(e)}")
        
        return processed_docs
    
    def _extract_pdf_text(self, content: bytes) -> str:
        """Extract text from PDF content"""
        try:
            pdf_file = io.BytesIO(content)
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            
            text = ""
            for page_num, page in enumerate(pdf_reader.pages):
                try:
                    text += page.extract_text() + "\n"
                except Exception as e:
                    st.warning(f"Could not extract text from page {page_num + 1}")
            
            return text.strip()
            
        except Exception as e:
            raise Exception(f"PDF extraction failed: {str(e)}")
    
    def _extract_docx_text(self, content: bytes) -> str:
        """Extract text from DOCX content"""
        try:
            docx_file = io.BytesIO(content)
            doc = docx.Document(docx_file)
            
            text = ""
            for paragraph in doc.paragraphs:
                text += paragraph.text + "\n"
            
            # Extract text from tables
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        text += cell.text + " "
                    text += "\n"
            
            return text.strip()
            
        except Exception as e:
            raise Exception(f"DOCX extraction failed: {str(e)}")
    
    def _get_file_type(self, filename: str) -> str:
        """Get file type from filename"""
        return Path(filename).suffix.lower().lstrip('.')
    
    def _generate_hash(self, content: bytes) -> str:
        """Generate MD5 hash of file content"""
        return hashlib.md5(content).hexdigest()
    
    def chunk_text(self, text: str) -> List[str]:
        """Manually chunk text if needed"""
        return self.text_splitter.split_text(text)
    
    def get_document_stats(self, document: Document) -> Dict:
        """Get statistics about a document"""
        text_length = len(document.text)
        word_count = len(document.text.split())
        estimated_chunks = max(1, text_length // self.chunk_size)
        
        return {
            'text_length': text_length,
            'word_count': word_count,
            'estimated_chunks': estimated_chunks,
            'filename': document.metadata.get('filename', 'Unknown')
        }
