# app.py - Streamlit Invoice Extraction App (Optimized Version)

import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import warnings
import base64
from datetime import datetime
import time
import gc
import sys

# Try to import openpyxl, show instructions if not installed
try:
    from openpyxl import load_workbook
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    st.error("⚠️ **openpyxl is not installed!** Please install it using: `pip install openpyxl`")

warnings.filterwarnings('ignore')

# Set page configuration
st.set_page_config(
    page_title="Invoice Data Extractor",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1E3A8A;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.5rem;
        color: #374151;
        margin-top: 1.5rem;
        margin-bottom: 1rem;
    }
    .success-box {
        background-color: #D1FAE5;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #10B981;
        margin: 1rem 0;
    }
    .info-box {
        background-color: #DBEAFE;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #3B82F6;
        margin: 1rem 0;
    }
    .warning-box {
        background-color: #FEF3C7;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #F59E0B;
        margin: 1rem 0;
    }
    .error-box {
        background-color: #FEE2E2;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #EF4444;
        margin: 1rem 0;
    }
    .metric-card {
        background-color: #F8FAFC;
        padding: 1rem;
        border-radius: 0.5rem;
        border: 1px solid #E2E8F0;
        text-align: center;
    }
    .stProgress > div > div > div > div {
        background-color: #3B82F6;
    }
    .install-instructions {
        background-color: #F3F4F6;
        padding: 1.5rem;
        border-radius: 0.5rem;
        border: 2px dashed #6B7280;
        margin: 1rem 0;
        font-family: monospace;
    }
    .sheet-info {
        background-color: #F8FAFC;
        padding: 0.5rem;
        border-radius: 0.25rem;
        margin: 0.25rem 0;
        font-size: 0.9rem;
    }
</style>
""", unsafe_allow_html=True)

# Check if openpyxl is available
if not OPENPYXL_AVAILABLE:
    st.markdown("""
    <div class="error-box">
    <h3>⚠️ Missing Dependencies</h3>
    <p>Please install the required packages:</p>
    <div class="install-instructions">
    pip install openpyxl pandas numpy streamlit
    </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# Initialize session state
if 'extracted_data' not in st.session_state:
    st.session_state.extracted_data = None
if 'processing_complete' not in st.session_state:
    st.session_state.processing_complete = False
if 'selected_sheets_count' not in st.session_state:
    st.session_state.selected_sheets_count = 0
if 'current_progress' not in st.session_state:
    st.session_state.current_progress = 0

# Optimized functions
def read_excel_sheet_fast(file_content, sheet_name):
    """Optimized function to read Excel sheet with memory management"""
    try:
        # Use pandas to read the sheet directly (faster for large files)
        df = pd.read_excel(
            io.BytesIO(file_content),
            sheet_name=sheet_name,
            header=None,
            engine='openpyxl'
        )
        return df
    except Exception as e:
        # Fallback to openpyxl if pandas fails
        try:
            wb = load_workbook(io.BytesIO(file_content), data_only=True, read_only=True)
            ws = wb[sheet_name]
            
            # Get data efficiently
            data = []
            for row in ws.iter_rows(values_only=True):
                data.append(list(row))
            
            df = pd.DataFrame(data)
            wb.close()
            return df
        except Exception as e2:
            st.warning(f"Could not read sheet '{sheet_name}': {str(e2)[:100]}")
            return pd.DataFrame()

def extract_invoice_data_fast(df, file_name, sheet_name):
    """Optimized extraction function with early returns"""
    result = {
        'file_name': file_name,
        'sheet_name': sheet_name,
        'customer_name': '',
        'document_date': '',
        'invoice_number': '',
        'order_number': '',
        'items_df': None,
        'subtotal': 0,
        'vat_amount': 0,
        'total_amount': 0
    }
    
    # Convert to string for searching (only first 30 rows for speed)
    df_sample = df.head(30).fillna('').astype(str)
    
    # 1. Extract Customer/Branch Name (optimized)
    for idx in range(min(10, len(df_sample))):
        for col in range(min(5, len(df_sample.columns))):
            cell_val = df_sample.iloc[idx, col]
            if 'Chandarana' in cell_val or 'Delis' in cell_val:
                # Try to extract full name
                name_match = re.search(r'(Chandarana[^\\d\\n]{0,50}(?:Branch|Mall|Naivasha)?)', cell_val, re.IGNORECASE)
                if name_match:
                    result['customer_name'] = name_match.group(1).strip()
                    break
        if result['customer_name']:
            break
    
    # 2. Extract Document Date (optimized for columns E,F,G)
    for idx in range(min(20, len(df))):
        for col in range(min(5, len(df.columns))):
            if pd.notna(df.iloc[idx, col]) and str(df.iloc[idx, col]).strip() == 'Document Date':
                # Check columns E,F,G (indices 4,5,6)
                for date_col in [4, 5, 6]:
                    if date_col < len(df.columns) and pd.notna(df.iloc[idx, date_col]):
                        date_str = str(df.iloc[idx, date_col])
                        date_match = re.search(r'(\d{1,2}[\.\/\-]\s*[A-Za-z]+\s*\d{4}|\d{1,2}[\.\/\-]\d{1,2}[\.\/\-]\d{2,4})', date_str)
                        if date_match:
                            result['document_date'] = date_match.group(1)
                            break
                break
    
    # 3. Extract Invoice Number
    for idx in range(min(20, len(df))):
        for col in range(min(10, len(df.columns))):
            cell_val = str(df.iloc[idx, col]) if pd.notna(df.iloc[idx, col]) else ''
            if 'Invoice No.' in cell_val:
                # Check adjacent cells
                for offset in [1, 2, -1, -2]:
                    check_col = col + offset
                    if 0 <= check_col < len(df.columns) and pd.notna(df.iloc[idx, check_col]):
                        inv_str = str(df.iloc[idx, check_col]).strip()
                        if inv_str and re.match(r'^\d+$', inv_str):
                            result['invoice_number'] = inv_str
                            break
                break
    
    # 4. Extract Order Number
    for idx in range(min(20, len(df))):
        for col in range(min(10, len(df.columns))):
            cell_val = str(df.iloc[idx, col]) if pd.notna(df.iloc[idx, col]) else ''
            if 'Order No.' in cell_val:
                for offset in [1, 2, -1, -2]:
                    check_col = col + offset
                    if 0 <= check_col < len(df.columns) and pd.notna(df.iloc[idx, check_col]):
                        order_str = str(df.iloc[idx, check_col]).strip()
                        if order_str and re.match(r'^\d+$', order_str):
                            result['order_number'] = order_str
                            break
                break
    
    # 5. Find items table (optimized)
    items_data = []
    header_found = False
    
    # Look for header row pattern
    for idx in range(len(df)):
        if header_found:
            break
            
        # Check if this row has item header indicators
        header_indicators = 0
        for col in range(min(15, len(df.columns))):
            if pd.notna(df.iloc[idx, col]):
                cell_str = str(df.iloc[idx, col]).lower()
                if any(keyword in cell_str for keyword in ['no.', 'description', 'qty', 'uom', 'unit price', 'vat', 'line amount']):
                    header_indicators += 1
        
        if header_indicators >= 3:  # Found header
            header_found = True
            
            # Extract items from following rows
            for item_idx in range(idx + 1, min(idx + 100, len(df))):  # Limit to 100 rows after header
                # Check if first column has item code
                if pd.notna(df.iloc[item_idx, 0]):
                    first_cell = str(df.iloc[item_idx, 0]).strip()
                    
                    # Check if it's an item code
                    if re.match(r'^[A-Z]{2,4}-\d{5}', first_cell):
                        item = {
                            'No.': first_cell,
                            'Description': str(df.iloc[item_idx, 1]) if 1 < len(df.columns) and pd.notna(df.iloc[item_idx, 1]) else '',
                            'Qty': df.iloc[item_idx, 2] if 2 < len(df.columns) and pd.notna(df.iloc[item_idx, 2]) else '',
                            'UoM': str(df.iloc[item_idx, 3]) if 3 < len(df.columns) and pd.notna(df.iloc[item_idx, 3]) else 'Kilograms'
                        }
                        
                        # Try to extract prices from remaining columns
                        for price_col in range(4, min(10, len(df.columns))):
                            if pd.notna(df.iloc[item_idx, price_col]):
                                cell_val = df.iloc[item_idx, price_col]
                                try:
                                    # Try to convert to number
                                    if isinstance(cell_val, (int, float)):
                                        if 'Unit Price' not in item:
                                            item['Unit Price Excl. VAT'] = float(cell_val)
                                        elif 'VAT %' not in item:
                                            item['VAT %'] = float(cell_val)
                                        elif 'Line Amount Excl. VAT' not in item:
                                            item['Line Amount Excl. VAT'] = float(cell_val)
                                except:
                                    pass
                        
                        items_data.append(item)
                    
                    # Stop if we hit totals
                    elif any(keyword in first_cell.lower() for keyword in ['subtotal', 'total', 'vat amount']):
                        break
    
    # Create items DataFrame if we found items
    if items_data:
        items_df = pd.DataFrame(items_data)
        
        # Ensure all required columns exist
        required_columns = ['No.', 'Description', 'Qty', 'UoM', 
                           'Unit Price Excl. VAT', 'VAT %', 'Line Amount Excl. VAT']
        
        for col in required_columns:
            if col not in items_df.columns:
                items_df[col] = ''
        
        # Clean numeric columns
        for col in ['Qty', 'Unit Price Excl. VAT', 'VAT %', 'Line Amount Excl. VAT']:
            if col in items_df.columns:
                items_df[col] = pd.to_numeric(items_df[col], errors='coerce')
        
        # Clean UoM
        if 'UoM' in items_df.columns:
            items_df['UoM'] = items_df['UoM'].apply(
                lambda x: str(x).strip().title() if pd.notna(x) and str(x).strip() else 'Kilograms'
            )
        
        result['items_df'] = items_df
        
        # Calculate totals
        if 'Line Amount Excl. VAT' in items_df.columns:
            result['subtotal'] = items_df['Line Amount Excl. VAT'].sum(skipna=True)
            result['vat_amount'] = result['subtotal'] * 0.16
            result['total_amount'] = result['subtotal'] + result['vat_amount']
    
    return result

def get_excel_download_link(df, filename):
    """Generate a download link for Excel file"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Extracted Data')
    b64 = base64.b64encode(output.getvalue()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">📥 Download Excel File</a>'
    return href

# Sidebar
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3135/3135715.png", width=100)
    st.title("Invoice Extractor")
    st.markdown("---")
    
    st.markdown("### 📋 How to Use:")
    st.markdown("""
    1. **Upload** your Excel invoice file
    2. **Configure** extraction settings
    3. **Process** the file
    4. **Download** the results
    """)
    
    st.markdown("---")
    st.markdown("### ⚙️ Settings")
    
    # Processing settings
    processing_mode = st.radio(
        "Processing Mode",
        ["All Sheets", "First N Sheets", "Last N Sheets", "Custom Range"],
        index=0
    )
    
    if processing_mode == "First N Sheets":
        n_sheets = st.slider("Number of sheets", min_value=1, max_value=100, value=10, step=1)
    elif processing_mode == "Last N Sheets":
        n_sheets = st.slider("Number of sheets", min_value=1, max_value=100, value=10, step=1)
    elif processing_mode == "Custom Range":
        col1, col2 = st.columns(2)
        with col1:
            start_sheet = st.number_input("Start sheet", min_value=1, value=1, step=1)
        with col2:
            end_sheet = st.number_input("End sheet", min_value=1, value=10, step=1)
    
    # Performance settings
    st.markdown("### ⚡ Performance")
    batch_size = st.slider("Sheets per batch", min_value=1, max_value=50, value=10, 
                          help="Process sheets in batches to avoid memory issues")
    show_progress = st.checkbox("Show progress", value=True)
    enable_optimization = st.checkbox("Enable fast mode", value=True, 
                                     help="Faster processing with reduced detail")
    
    st.markdown("---")
    st.markdown("### 📊 Features")
    st.markdown("""
    ✅ Extract from merged cells  
    ✅ Handle multiple date formats  
    ✅ Extract invoice & order numbers  
    ✅ Get all product details  
    ✅ Process multiple sheets  
    ✅ Clean & format data
    """)

# Main content
st.markdown('<h1 class="main-header">📄 Invoice Data Extraction Tool</h1>', unsafe_allow_html=True)
st.markdown("Extract invoice data from Excel files with ease")

# File upload section
st.markdown("### 📤 Upload Your Excel File")
uploaded_file = st.file_uploader(
    "Choose an Excel file",
    type=['xlsx', 'xls'],
    help="Upload Excel files containing invoice data"
)

if uploaded_file:
    st.markdown(f'<div class="success-box"><strong>File uploaded:</strong> {uploaded_file.name}</div>', unsafe_allow_html=True)
    
    try:
        # Get sheet names
        wb = load_workbook(io.BytesIO(uploaded_file.getvalue()), read_only=True, data_only=True)
        sheet_names = wb.sheetnames
        total_sheets = len(sheet_names)
        wb.close()
        
        st.markdown(f'<div class="info-box"><strong>Found:</strong> {total_sheets} sheets in the file</div>', unsafe_allow_html=True)
        
        # Determine which sheets to process
        if processing_mode == "All Sheets":
            selected_sheets = sheet_names
            st.session_state.selected_sheets_count = min(total_sheets, 100)  # Limit for safety
        elif processing_mode == "First N Sheets":
            selected_sheets = sheet_names[:min(n_sheets, total_sheets)]
            st.session_state.selected_sheets_count = len(selected_sheets)
        elif processing_mode == "Last N Sheets":
            selected_sheets = sheet_names[-min(n_sheets, total_sheets):]
            st.session_state.selected_sheets_count = len(selected_sheets)
        elif processing_mode == "Custom Range":
            start_idx = max(0, start_sheet - 1)
            end_idx = min(total_sheets, end_sheet)
            selected_sheets = sheet_names[start_idx:end_idx]
            st.session_state.selected_sheets_count = len(selected_sheets)
        
        # Show processing info
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Sheets", total_sheets)
        with col2:
            st.metric("Sheets to Process", st.session_state.selected_sheets_count)
        with col3:
            file_size_mb = len(uploaded_file.getvalue()) / (1024 * 1024)
            st.metric("File Size", f"{file_size_mb:.2f} MB")
        
        # Show first few sheet names
        if len(selected_sheets) <= 10:
            st.markdown("**Sheets to be processed:**")
            for sheet in selected_sheets:
                st.markdown(f'<div class="sheet-info">{sheet}</div>', unsafe_allow_html=True)
        else:
            st.markdown(f"**Sheets to be processed:** First 10 of {len(selected_sheets)}")
            for sheet in selected_sheets[:10]:
                st.markdown(f'<div class="sheet-info">{sheet}</div>', unsafe_allow_html=True)
            st.markdown(f"*... and {len(selected_sheets) - 10} more sheets*")
        
        # Process button
        if st.button("🚀 Start Extraction", type="primary", use_container_width=True):
            all_results = []
            progress_bar = None
            status_text = None
            results_container = st.container()
            
            if show_progress:
                progress_bar = st.progress(0)
                status_text = st.empty()
            
            with results_container:
                st.markdown("### 📝 Processing Started")
                
                start_time = time.time()
                processed_count = 0
                successful_sheets = 0
                
                # Process in batches to manage memory
                for batch_start in range(0, len(selected_sheets), batch_size):
                    batch_end = min(batch_start + batch_size, len(selected_sheets))
                    batch_sheets = selected_sheets[batch_start:batch_end]
                    
                    for i, sheet_name in enumerate(batch_sheets):
                        sheet_num = batch_start + i + 1
                        
                        if show_progress:
                            progress = sheet_num / len(selected_sheets)
                            progress_bar.progress(progress)
                            status_text.text(f"Processing sheet {sheet_num} of {len(selected_sheets)}: {sheet_name[:30]}...")
                        
                        try:
                            # Read sheet
                            df = read_excel_sheet_fast(uploaded_file.getvalue(), sheet_name)
                            
                            if df.empty:
                                continue
                            
                            # Extract data
                            if enable_optimization:
                                sheet_data = extract_invoice_data_fast(df, uploaded_file.name, sheet_name)
                            else:
                                # Use original extraction function if fast mode is disabled
                                sheet_data = extract_invoice_data_fast(df, uploaded_file.name, sheet_name)
                            
                            if sheet_data['items_df'] is not None and not sheet_data['items_df'].empty:
                                all_results.append(sheet_data)
                                successful_sheets += 1
                            
                            processed_count += 1
                            
                            # Clear memory
                            del df
                            gc.collect()
                            
                            # Small delay for UI update
                            time.sleep(0.01)
                            
                        except Exception as e:
                            st.warning(f"Sheet '{sheet_name}' skipped: {str(e)[:100]}")
                            continue
                
                end_time = time.time()
                processing_time = end_time - start_time
                
                # Clear progress indicators
                if show_progress:
                    progress_bar.empty()
                    status_text.empty()
                
                # Combine all results
                if all_results:
                    st.markdown(f'<div class="success-box">✅ Extraction Complete! Processed {successful_sheets} sheets in {processing_time:.1f} seconds</div>', unsafe_allow_html=True)
                    
                    # Show summary
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Sheets Processed", successful_sheets)
                    with col2:
                        total_items = sum(len(r['items_df']) for r in all_results)
                        st.metric("Total Items", total_items)
                    with col3:
                        customers = len(set(r['customer_name'] for r in all_results if r['customer_name']))
                        st.metric("Unique Customers", customers)
                    with col4:
                        st.metric("Processing Time", f"{processing_time:.1f}s")
                    
                    # Combine data
                    combined_list = []
                    for result in all_results:
                        items_df = result['items_df'].copy()
                        
                        # Add metadata columns
                        items_df['File Name'] = result['file_name']
                        items_df['Sheet Name'] = result['sheet_name']
                        items_df['Customer Name'] = result['customer_name']
                        items_df['Document Date'] = result['document_date']
                        items_df['Invoice Number'] = result['invoice_number']
                        items_df['Order Number'] = result['order_number']
                        items_df['Subtotal'] = result['subtotal']
                        items_df['VAT Amount'] = result['vat_amount']
                        items_df['Total Amount'] = result['total_amount']
                        
                        combined_list.append(items_df)
                    
                    # Combine all data
                    combined_df = pd.concat(combined_list, ignore_index=True)
                    
                    # Reorder columns
                    column_order = [
                        'File Name', 'Sheet Name', 'Customer Name', 'Document Date',
                        'Invoice Number', 'Order Number', 'Subtotal', 'VAT Amount', 'Total Amount',
                        'No.', 'Description', 'Qty', 'UoM',
                        'Unit Price Excl. VAT', 'VAT %', 'Line Amount Excl. VAT'
                    ]
                    
                    available_cols = [col for col in column_order if col in combined_df.columns]
                    combined_df = combined_df[available_cols]
                    
                    # Store in session state
                    st.session_state.extracted_data = combined_df
                    st.session_state.processing_complete = True
                    
                    # Display sample data
                    st.markdown("### 📋 Sample Data (First 10 rows)")
                    st.dataframe(combined_df.head(10), width='stretch')
                    
                    # Export options
                    st.markdown("### 📥 Export Options")
                    
                    # Generate filename with timestamp
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"extracted_invoices_{timestamp}.xlsx"
                    
                    # Download buttons
                    col1, col2 = st.columns(2)
                    with col1:
                        st.markdown("#### Excel Format")
                        st.markdown(get_excel_download_link(combined_df, filename), unsafe_allow_html=True)
                    
                    with col2:
                        st.markdown("#### CSV Format")
                        csv = combined_df.to_csv(index=False)
                        b64_csv = base64.b64encode(csv.encode()).decode()
                        csv_filename = f"extracted_invoices_{timestamp}.csv"
                        href_csv = f'<a href="data:file/csv;base64,{b64_csv}" download="{csv_filename}">📥 Download CSV File</a>'
                        st.markdown(href_csv, unsafe_allow_html=True)
                    
                    # Data preview
                    with st.expander("📊 Data Statistics"):
                        col1, col2 = st.columns(2)
                        with col1:
                            st.markdown("##### Column Information")
                            col_info = pd.DataFrame({
                                'Column': combined_df.columns,
                                'Non-Null Count': combined_df.notna().sum().values,
                                'Data Type': combined_df.dtypes.astype(str).values
                            })
                            st.dataframe(col_info, width='stretch')
                        
                        with col2:
                            st.markdown("##### File Summary")
                            summary_data = {
                                'Metric': ['Total Sheets', 'Total Items', 'Unique Customers', 'File Size'],
                                'Value': [
                                    successful_sheets,
                                    len(combined_df),
                                    combined_df['Customer Name'].nunique(),
                                    f"{sys.getsizeof(combined_df) / 1024 / 1024:.2f} MB"
                                ]
                            }
                            st.dataframe(pd.DataFrame(summary_data), width='stretch')
                
                else:
                    st.markdown('<div class="warning-box">⚠️ No data was extracted from any sheets. Please check your file format.</div>', unsafe_allow_html=True)
    
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        st.markdown('<div class="error-box">Please try again with a smaller file or fewer sheets.</div>', unsafe_allow_html=True)

else:
    # Show welcome message
    st.markdown("""
    <div class="info-box">
    <h3>Welcome to the Invoice Data Extractor!</h3>
    <p>This tool helps you extract invoice data from Excel files with the following features:</p>
    <ul>
        <li><strong>Smart Extraction:</strong> Handles merged cells and various formats</li>
        <li><strong>Multi-Sheet Processing:</strong> Process all sheets or selected ranges</li>
        <li><strong>Complete Data:</strong> Extracts customer info, dates, invoice numbers, and all product details</li>
        <li><strong>Easy Export:</strong> Download results as Excel or CSV files</li>
    </ul>
    <p>To get started, upload your Excel file using the uploader above.</p>
    </div>
    """, unsafe_allow_html=True)

# Footer
st.markdown("---")
st.markdown(
    """
    <div style="text-align: center; color: #666; font-size: 0.9rem;">
    <p>Invoice Data Extractor v1.0 | Built with Streamlit</p>
    <p>© 2024 | Extracts data from Excel invoices with precision</p>
    </div>
    """,
    unsafe_allow_html=True
)
