# modules/utils.py
import streamlit as st
import pandas as pd
from .config import logger
from decimal import Decimal

from decimal import Decimal

def normalize_columns(df):
    column_mapping = {
        'bid_id': 'Bid ID',
        'business_group': 'Business Group',
        'product_type': 'Product Type',
        'incumbent': 'Incumbent',
        'baseline_price': 'Baseline Price',
        'current_price': 'Current Price',
        'bid_supplier_name': 'Supplier Name',
        'bid_supplier_capacity': 'Supplier Capacity',
        'bid_price': 'Bid Price',
        'supplier_name': 'Supplier Name',
        'bid_volume': 'Bid Volume',
        'facility': 'Facility'
    }

    # Instead of float, specify Decimal for price columns:
    known_dtypes = {
        'Bid ID': str,
        'Business Group': str,
        'Product Type': str,
        'Incumbent': str,
        'Baseline Price': Decimal, 
        'Supplier Name': str,
        'Supplier Capacity': Decimal,
        'Bid Price': Decimal,      
        'Bid Volume': Decimal,
        'Facility': str,
        'Current Price': Decimal
    }

    df.columns = [col.strip().lower() for col in df.columns]
    df = df.rename(columns=column_mapping)
    df.columns = [col.strip().title() for col in df.columns]

    for col in df.columns:
        if col in known_dtypes:
            desired_dtype = known_dtypes[col]
        else:
            desired_dtype = str

        if desired_dtype == str:
            df[col] = df[col].astype(str).str.strip()
        elif desired_dtype == Decimal:
            # Convert to string first, strip whitespace, then Decimal
            df[col] = df[col].astype(str).str.strip().apply(lambda x: Decimal(x) if x not in ["", "nan"] else None)
        else:
            df[col] = pd.to_numeric(df[col], errors='coerce').astype(desired_dtype, errors='ignore')
    return df


def validate_uploaded_file(file):
    if not file:
        st.error("No file uploaded. Please upload an Excel file.")
        return False
    if not file.name.endswith('.xlsx'):
        st.error("Invalid file type. Please upload an Excel file (.xlsx).")
        return False

# Function to generate a unique file name
def generate_unique_filename(original_filename):
    unique_id = uuid.uuid4().hex
    if '.' in original_filename:
        name, extension = original_filename.rsplit('.', 1)
        return f"{name}_{unique_id}.{extension}"
    else:
        return f"{original_filename}_{unique_id}"
