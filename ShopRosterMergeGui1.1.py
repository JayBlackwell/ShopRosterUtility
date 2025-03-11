import streamlit as st
import pandas as pd
import numpy as np
import io
import base64

def get_download_link(df, filename, text):
    """Generate a download link for a DataFrame"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    b64 = base64.b64encode(output.getvalue()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">{text}</a>'
    return href

def process_member_data(df):
    """Process member data to merge IDs and remove duplicates based on names and emails"""
    # Track processing statistics
    stats = {
        "total_records": len(df),
        "unique_names": 0,
        "name_matches_found": 0,
        "name_ids_copied": 0,
        "name_records_removed": 0,
        "unique_emails": 0,
        "email_matches_found": 0,
        "email_ids_copied": 0,
        "email_records_removed": 0
    }
    
    # Create progress bars and status text for both processes
    name_progress_bar = st.progress(0)
    name_status_text = st.empty()
    
    # Create name keys for matching
    name_status_text.text("Creating name keys for matching...")
    df['FullName'] = df['First Name'].str.strip().str.lower() + ' ' + df['Last Name'].str.strip().str.lower()
    
    # Handle empty Member Card IDs
    # Convert empty strings AND whitespace-only strings to NaN
    df['Member Card ID'] = df['Member Card ID'].astype(str)
    df['Member Card ID'] = df['Member Card ID'].replace(r'^\s*$', np.nan, regex=True)
    df['Member Card ID'] = df['Member Card ID'].replace('nan', np.nan)
    df['Member Card ID'] = df['Member Card ID'].replace('None', np.nan)
    
    # Handle empty Email addresses
    # Convert empty strings AND whitespace-only strings to NaN
    df['Email'] = df['Email'].astype(str)
    df['Email'] = df['Email'].replace(r'^\s*$', np.nan, regex=True) 
    df['Email'] = df['Email'].replace('nan', np.nan)
    df['Email'] = df['Email'].replace('None', np.nan)
    
    # Normalize email addresses to lowercase
    df['Email'] = df['Email'].str.lower()
    
    # Get counts before processing
    empty_ids = df['Member Card ID'].isna() | (df['Member Card ID'].str.strip() == '')
    st.write(f"Initial records with Member Card ID: {len(df) - sum(empty_ids)}")
    st.write(f"Initial records without Member Card ID: {sum(empty_ids)}")
    
    # Create list to track which rows to keep
    rows_to_keep = list(range(len(df)))
    
    # Track changes for reporting - separate for name and email matches
    name_changes = []
    email_changes = []
    
    # STEP 1: Process based on names
    # Get unique names
    unique_names = df['FullName'].unique()
    stats["unique_names"] = len(unique_names)
    
    # Process each unique full name
    for i, name in enumerate(unique_names):
        # Update progress
        name_progress_bar.progress((i + 1) / len(unique_names))
        name_status_text.text(f"Processing {i+1} of {len(unique_names)} unique names: {name}")
        
        # Get indices of all rows with this name
        indices = df[df['FullName'] == name].index.tolist()
        
        # Skip if only one record with this name
        if len(indices) <= 1:
            continue
        
        # Check if we have both with and without IDs in this group
        has_id_indices = []
        no_id_indices = []
        
        for idx in indices:
            # Check if Member Card ID is empty or NaN
            id_value = df.loc[idx, 'Member Card ID']
            is_empty = pd.isna(id_value) or str(id_value).strip() == ''
            
            if not is_empty:
                has_id_indices.append(idx)
            else:
                no_id_indices.append(idx)
        
        # If we have both types, process them
        if has_id_indices and no_id_indices:
            stats["name_matches_found"] += 1
            
            # For each record without ID, copy from a record with ID
            for no_id_idx in no_id_indices:
                if has_id_indices:
                    # Get the first available ID record
                    has_id_idx = has_id_indices[0]  # Keep using the same ID record
                    member_id = df.loc[has_id_idx, 'Member Card ID']
                    
                    # Copy the ID to the record without one
                    df.loc[no_id_idx, 'Member Card ID'] = member_id
                    
                    # Track the change
                    stats["name_ids_copied"] += 1
                    name_changes.append({
                        'match_type': 'Name',
                        'identifier': name,
                        'no_id_row': no_id_idx + 2,  # +2 for Excel row number
                        'has_id_row': has_id_idx + 2,  # +2 for Excel row number
                        'id_copied': member_id
                    })
            
            # Mark all but one of the has_id records for removal
            # Keep the first one and remove the rest
            if len(has_id_indices) > 1:
                for has_id_idx in has_id_indices[1:]:
                    if has_id_idx in rows_to_keep:
                        rows_to_keep.remove(has_id_idx)
                        stats["name_records_removed"] += 1
    
    # Clear progress indicators for name processing
    name_progress_bar.empty()
    name_status_text.empty()
    
    # Get the intermediate dataframe after name-based deduplication
    intermediate_df = df.iloc[rows_to_keep].copy()
    
    # STEP 2: Process based on email addresses
    # Create new progress indicators for email processing
    email_progress_bar = st.progress(0)
    email_status_text = st.empty()
    email_status_text.text("Starting email-based deduplication...")
    
    # Get unique emails (excluding NaN/empty)
    valid_emails = intermediate_df['Email'].dropna()
    unique_emails = valid_emails[valid_emails != ''].unique()
    stats["unique_emails"] = len(unique_emails)
    
    # Track which rows to keep from the intermediate dataset
    intermediate_rows_to_keep = list(range(len(intermediate_df)))
    
    # Process each unique email
    for i, email in enumerate(unique_emails):
        # Update progress
        email_progress_bar.progress((i + 1) / len(unique_emails))
        email_status_text.text(f"Processing {i+1} of {len(unique_emails)} unique emails")
        
        # Get indices of all rows with this email
        indices = intermediate_df[intermediate_df['Email'] == email].index.tolist()
        
        # Skip if only one record with this email
        if len(indices) <= 1:
            continue
        
        # Check if we have both with and without IDs in this group
        has_id_indices = []
        no_id_indices = []
        
        for idx in indices:
            # Check if Member Card ID is empty or NaN
            id_value = intermediate_df.loc[idx, 'Member Card ID']
            is_empty = pd.isna(id_value) or str(id_value).strip() == ''
            
            if not is_empty:
                has_id_indices.append(idx)
            else:
                no_id_indices.append(idx)
        
        # If we have both types, process them
        if has_id_indices and no_id_indices:
            stats["email_matches_found"] += 1
            
            # For each record without ID, copy from a record with ID
            for no_id_idx in no_id_indices:
                if has_id_indices:
                    # Get the first available ID record
                    has_id_idx = has_id_indices[0]  # Keep using the same ID record
                    member_id = intermediate_df.loc[has_id_idx, 'Member Card ID']
                    
                    # Copy the ID to the record without one
                    intermediate_df.loc[no_id_idx, 'Member Card ID'] = member_id
                    
                    # Track the change
                    stats["email_ids_copied"] += 1
                    email_changes.append({
                        'match_type': 'Email',
                        'identifier': email,
                        'no_id_row': intermediate_df.index[no_id_idx] + 2,  # +2 for Excel row number
                        'has_id_row': intermediate_df.index[has_id_idx] + 2,  # +2 for Excel row number
                        'id_copied': member_id
                    })
            
            # Mark all but one record for removal (keep the first one with ID)
            keep_idx = has_id_indices[0] if has_id_indices else indices[0]
            for idx in indices:
                if idx != keep_idx and idx in intermediate_rows_to_keep:
                    intermediate_rows_to_keep.remove(idx)
                    stats["email_records_removed"] += 1
    
    # Clear progress indicators for email processing
    email_progress_bar.empty()
    email_status_text.empty()
    
    # Keep only the rows we want after email-based deduplication
    result_df = intermediate_df.iloc[intermediate_rows_to_keep].copy()
    
    # Remove the helper column
    result_df = result_df.drop(columns=['FullName'])
    
    # Combine the change records for reporting
    all_changes = name_changes + email_changes
    
    return result_df, all_changes, stats

# Set up the Streamlit app
st.set_page_config(page_title="Golf Shop Roster Utility", page_icon="solsticelogo.png", layout="wide")

# App title and description
st.title("Golf Shop Roster Utility")
st.markdown("© Solstice Solutions | all rights reserved")

# File uploader
st.write("Upload your Excel roster file")
uploaded_file = st.file_uploader("", type=['xlsx', 'xls'])

if uploaded_file is not None:
    # Load the data
    try:
        with st.spinner("Loading data..."):
            # First, try to identify the GGS_ID column by checking column names
            # Read the first few rows to get column names
            preview_df = pd.read_excel(uploaded_file, nrows=1)
            column_dtypes = {}
            
            # Look for columns that might contain IDs and ensure they're treated as strings
            for col in preview_df.columns:
                if any(id_term in col.lower() for id_term in ['id', 'ggs', 'member', 'card']):
                    column_dtypes[col] = str
            
            # Now read the full file with the specified dtypes
            df = pd.read_excel(uploaded_file, dtype=column_dtypes)
            
            # Additionally, convert any other columns that look like they contain large numeric IDs
            for col in df.columns:
                # Check a sample of values to see if they're large numbers
                sample = df[col].dropna().head(10)
                if sample.astype(str).str.len().mean() > 10 and pd.to_numeric(sample, errors='coerce').notna().all():
                    df[col] = df[col].astype(str)
        
        # Show a preview of the data
        st.subheader("Data Preview")
        st.dataframe(df.head())
        
        # Verify required columns exist
        required_columns = ['First Name', 'Last Name', 'Member Card ID', 'Email']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            st.error(f"Missing required columns: {', '.join(missing_columns)}")
            
            # Show available columns to help the user
            st.write("Available columns in your file:")
            st.write(", ".join(df.columns.tolist()))
            
            # Allow column mapping
            st.subheader("Column Mapping")
            st.write("Please map the required columns to your file's columns:")
            
            mapping = {}
            for req_col in required_columns:
                if req_col in missing_columns:
                    mapping[req_col] = st.selectbox(f"Select column for '{req_col}':", [""] + df.columns.tolist())
            
            if st.button("Apply Mapping"):
                # Rename columns according to mapping
                for req_col, file_col in mapping.items():
                    if file_col:
                        df = df.rename(columns={file_col: req_col})
                st.success("Column mapping applied!")
                st.experimental_rerun()
            
            st.stop()
            
        # Process button
        if st.button("Process Data"):
            with st.spinner("Processing data..."):
                # Process the data
                result_df, changes, stats = process_member_data(df)
                
                # Show statistics with two sections
                st.subheader("Processing Results")
                
                # Name-based statistics
                st.write("### Name-Based Deduplication")
                col1, col2, col3 = st.columns(3)
                col1.metric("Unique Names", stats["unique_names"])
                col2.metric("Matches Found", stats["name_matches_found"])
                col3.metric("IDs Copied", stats["name_ids_copied"])
                st.metric("Records Removed", stats["name_records_removed"])
                
                # Email-based statistics
                st.write("### Email-Based Deduplication")
                col1, col2, col3 = st.columns(3)
                col1.metric("Unique Emails", stats["unique_emails"])
                col2.metric("Matches Found", stats["email_matches_found"])
                col3.metric("IDs Copied", stats["email_ids_copied"])
                st.metric("Records Removed", stats["email_records_removed"])
                
                # Overall statistics
                st.write("### Overall Results")
                col1, col2 = st.columns(2)
                total_records_removed = stats["name_records_removed"] + stats["email_records_removed"]
                col1.metric("Initial Records", stats["total_records"])
                col2.metric("Final Records", stats["total_records"] - total_records_removed, f"-{total_records_removed}")
                
                # Show the changes made
                if changes:
                    st.subheader("Changes Made")
                    
                    # Convert to DataFrame for display
                    changes_df = pd.DataFrame(changes)
                    
                    # Separate name and email changes for display
                    if not changes_df.empty:
                        # Add a filter widget
                        match_type = st.selectbox(
                            "Filter by match type:", 
                            ["All", "Name", "Email"]
                        )
                        
                        if match_type == "All":
                            filtered_df = changes_df
                        else:
                            filtered_df = changes_df[changes_df['match_type'] == match_type]
                        
                        st.dataframe(filtered_df)
                        st.write(f"Showing {len(filtered_df)} changes out of {len(changes_df)} total changes.")
                else:
                    st.info("No matching profiles found to merge.")
                
                # Preview the result
                st.subheader("Result Preview")
                st.dataframe(result_df.head())
                
                # Download link
                st.subheader("Download Processed Data")
                filename = "processed_roster.xlsx"
                
                # Create an Excel writer using openpyxl
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False)
                    
                    # Access the worksheet
                    worksheet = writer.sheets['Sheet1']
                    
                    # Find ID columns and format them as text
                    id_columns = [i+1 for i, col in enumerate(result_df.columns) 
                                if any(id_term in col.lower() for id_term in ['id', 'ggs', 'member', 'card'])]
                    
                    for col_idx in id_columns:
                        for row in range(2, len(result_df) + 2):  # +2 for header and 1-based indexing
                            cell = worksheet.cell(row=row, column=col_idx)
                            cell.number_format = '@'  # Format as text
                
                # Get the download link
                b64 = base64.b64encode(output.getvalue()).decode()
                href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">Click here to download the processed Excel file</a>'
                st.markdown(href, unsafe_allow_html=True)
    
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
        st.exception(e)
