# config.py
import os
from dataclasses import dataclass
from typing import List, Optional

@dataclass
class SharePointConfig:
    client_id: str
    client_secret: str
    tenant_id: str
    site_name: str
    
    @classmethod
    def from_env(cls):
        return cls(
            client_id=os.getenv("SHAREPOINT_CLIENT_ID"),
            client_secret=os.getenv("SHAREPOINT_CLIENT_SECRET"),
            tenant_id=os.getenv("SHAREPOINT_TENANT_ID"),
            site_name=os.getenv("SHAREPOINT_SITE_NAME")
        )

@dataclass
class AstraConfig:
    token: str
    endpoint: str
    collection_name: str = "documents"
    
    @classmethod
    def from_env(cls):
        return cls(
            token=os.getenv("ASTRA_DB_TOKEN"),
            endpoint=os.getenv("ASTRA_DB_ENDPOINT"),
            collection_name=os.getenv("ASTRA_COLLECTION_NAME", "documents")
        )

@dataclass
class OpenAIConfig:
    api_key: str
    model: str = "gpt-3.5-turbo"
    embedding_model: str = "text-embedding-ada-002"
    
    @classmethod
    def from_env(cls):
        return cls(
            api_key=os.getenv("OPENAI_API_KEY"),
            model=os.getenv("OPENAI_MODEL", "gpt-3.5-turbo"),
            embedding_model=os.getenv("OPENAI_EMBEDDING_MODEL", "text-embedding-ada-002")
        )
