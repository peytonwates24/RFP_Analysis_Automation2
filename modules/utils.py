# modules/utils.py
import streamlit as st
import pandas as pd
from .config import logger
import uuid
import datetime


def normalize_columns(df):
    column_mapping = {
        'bid_id': 'Bid ID',
        'business_group': 'Business Group',
        'product_type': 'Product Type',
        'incumbent': 'Incumbent',
        'baseline_price': 'Baseline Price',
        'bid_supplier_name': 'Supplier Name',
        'bid_supplier_capacity': 'Supplier Capacity',
        'bid_price': 'Bid Price',
        'supplier_name': 'Supplier Name',
        'bid_volume': 'Bid Volume',
        'facility': 'Facility'
    }
    return df.rename(columns=column_mapping)

# Example implementation in modules/utils.py

# Function to validate uploaded file
def validate_uploaded_file(uploaded_file) -> bool:
    """
    Validates the uploaded Excel file.
    
    Criteria:
    - Must contain all required columns.
    """
    try:
        # Attempt to read the Excel file
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        
        # Define required columns based on the Supabase table
        required_columns = [
            'bid_id',
            'supplier_name',
            'facility',
            'baseline_price',
            'current_price',
            'bid_volume',
            'bid_price',
            'supplier_capacity'
            # Add other required columns
        ]
        
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            st.error(f"Missing columns: {', '.join(missing_columns)}")
            logger.error(f"Uploaded file '{uploaded_file.name}' is missing columns: {', '.join(missing_columns)}")
            return False
        
        return True
    except Exception as e:
        st.error(f"Error reading the uploaded file: {e}")
        logger.error(f"Error reading the uploaded file '{uploaded_file.name}': {e}")
        return False

# Function to generate a unique file name
def generate_unique_filename(original_filename):
    unique_id = uuid.uuid4().hex
    if '.' in original_filename:
        name, extension = original_filename.rsplit('.', 1)
        return f"{name}_{unique_id}.{extension}"
    else:
        return f"{original_filename}_{unique_id}"
