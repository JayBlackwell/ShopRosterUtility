import pandas as pd
import numpy as np

def main():
    try:
        # Get file paths
        input_file = input("Enter the path to your input Excel file: ")
        output_file = input("Enter the path for the output Excel file: ")
        
        # Load Excel file
        print(f"Loading {input_file}...")
        df = pd.read_excel(input_file)
        
        # Print information about the data
        print(f"Total records: {len(df)}")
        
        # Check for truly empty Member Card IDs (both NaN and empty strings)
        empty_ids = df['Member Card ID'].isna() | (df['Member Card ID'] == '')
        print(f"Records with Member Card ID: {len(df) - sum(empty_ids)}")
        print(f"Records with empty Member Card ID: {sum(empty_ids)}")
        
        # Create a modified dataframe where empty strings are treated as NaN
        df['Member Card ID'] = df['Member Card ID'].replace('', np.nan)
        
        # Create name keys for matching
        df['FullName'] = df['First Name'].str.strip().str.lower() + ' ' + df['Last Name'].str.strip().str.lower()
        
        # Create list to track which rows to keep
        rows_to_keep = list(range(len(df)))
        
        # Track changes
        changes = []
        
        # Process each unique full name
        for name in df['FullName'].unique():
            # Get indices of all rows with this name
            indices = df[df['FullName'] == name].index.tolist()
            
            # Skip if only one record with this name
            if len(indices) <= 1:
                continue
            
            # Check if we have both with and without IDs in this group
            has_id_indices = []
            no_id_indices = []
            
            for idx in indices:
                if pd.notna(df.loc[idx, 'Member Card ID']):
                    has_id_indices.append(idx)
                else:
                    no_id_indices.append(idx)
            
            # If we have both types, process them
            if has_id_indices and no_id_indices:
                # Print what we found for visibility
                print(f"\nFound match group for '{name}':")
                print(f"  With IDs: {len(has_id_indices)} records at rows {[i+2 for i in has_id_indices]}")
                print(f"  Without IDs: {len(no_id_indices)} records at rows {[i+2 for i in no_id_indices]}")
                
                # Look specifically at the records to verify
                for i, idx in enumerate(has_id_indices):
                    print(f"  ID record {i+1}: {df.loc[idx, 'First Name']} {df.loc[idx, 'Last Name']} - ID: {df.loc[idx, 'Member Card ID']}")
                for i, idx in enumerate(no_id_indices):
                    print(f"  No ID record {i+1}: {df.loc[idx, 'First Name']} {df.loc[idx, 'Last Name']}")
                
                # For each record without ID, copy from a record with ID
                for no_id_idx in no_id_indices:
                    if has_id_indices:
                        # Get the first available ID record
                        has_id_idx = has_id_indices.pop(0)
                        member_id = df.loc[has_id_idx, 'Member Card ID']
                        
                        # Copy the ID to the record without one
                        df.loc[no_id_idx, 'Member Card ID'] = member_id
                        
                        # Mark the source record for removal
                        if has_id_idx in rows_to_keep:
                            rows_to_keep.remove(has_id_idx)
                            
                        # Track the change
                        changes.append({
                            'name': name,
                            'no_id_row': no_id_idx + 2,  # +2 for Excel row
                            'has_id_row': has_id_idx + 2,  # +2 for Excel row
                            'id_copied': member_id
                        })
        
        # Keep only the rows we want
        result_df = df.iloc[rows_to_keep].copy()
        
        # Remove the helper column
        result_df = result_df.drop(columns=['FullName'])
        
        # Report on changes
        print(f"\nProcessed {len(changes)} matches:")
        for change in changes:
            print(f"  - For {change['name']}: Copied ID {change['id_copied']} from row {change['has_id_row']} to row {change['no_id_row']}")
        
        print(f"\nBefore: {len(df)} records")
        print(f"After: {len(result_df)} records")
        print(f"Removed: {len(df) - len(result_df)} records")
        
        # Save the result
        print(f"\nSaving to {output_file}...")
        result_df.to_excel(output_file, index=False)
        print("Done!")
        
    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
