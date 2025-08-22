# utils/sharepoint_client.py (Updated imports)
import streamlit as st
import os
import requests
from datetime import datetime, timedelta
from typing import List, Dict, Optional
import json

# Updated import for SharePoint
try:
    from llama_index.readers.microsoft_sharepoint import SharePointReader
    SHAREPOINT_AVAILABLE = True
except ImportError:
    try:
        # Fallback import path
        from llama_index_readers_microsoft_sharepoint import SharePointReader
        SHAREPOINT_AVAILABLE = True
    except ImportError:
        SharePointReader = None
        SHAREPOINT_AVAILABLE = False
        st.error("SharePoint reader not available. Please check package installation.")

# Rest of your SharePointClient class remains the same...
