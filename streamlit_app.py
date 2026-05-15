import streamlit as st
import pdfplumber
import pandas as pd
import re
import os
from io import BytesIO

# Page configuration
st.set_page_config(
    page_title="Invoice Automation - PDF to Excel",
    page_icon="📄",
    layout="wide"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main {
        padding-top: 2rem;
    }
    .stButton>button {
        width: 100%;
        background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
        color: white;
        font-weight: 600;
        padding: 0.75rem;
        border-radius: 0.5rem;
        border: none;
    }
    .stButton>button:hover {
        box-shadow: 0 4px 12px rgba(59, 130, 246, 0.4);
    }
    .upload-box {
        border: 2px dashed #3b82f6;
        border-radius: 1rem;
        padding: 2rem;
        text-align: center;
        background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
    }
</style>
""", unsafe_allow_html=True)


def convert_to_number(text):
    """Convert string to number, removing commas."""
    if text is None or text == '':
        return None
    try:
        cleaned = str(text).replace(',', '').replace('$', '').strip()
        return float(cleaned)
    except (ValueError, AttributeError):
        return None


def sanitize_for_excel(text):
    """Clean text for Excel compatibility."""
    if text is None:
        return ''
    
    if not isinstance(text, str):
        text = str(text)
    
    cleaned = ''
    for char in text:
        code = ord(char)
        if 32 <= code <= 126 or code >= 160:
            if char not in ['\x00', '\x01', '\x02', '\x03', '\x04', '\x05', '\x06', '\x07', 
                           '\x08', '\x0b', '\x0c', '\x0e', '\x0f', '\x10', '\x11', '\x12',
                           '\x13', '\x14', '\x15', '\x16', '\x17', '\x18', '\x19', '\x1a',
                           '\x1b', '\x1c', '\x1d', '\x1e', '\x1f']:
                cleaned += char
        elif char in [' ', '\n', '\t']:
            cleaned += char
    
    if cleaned and cleaned[0] in ['=', '+', '-', '@']:
        cleaned = "'" + cleaned
    
    cleaned = cleaned.strip()
    
    if len(cleaned) > 32000:
        cleaned = cleaned[:32000] + '...'
    
    return cleaned


def extract_invoice_data(pdf_file):
    """Extract structured data from invoice PDF."""
    invoice_data = {
        'invoice_number': '',
        'date': '',
        'due_date': '',
        'vendor_name': '',
        'customer_name': '',
        'total_amount': '',
        'tax_amount': '',
        'subtotal': '',
        'items': []
    }
    
    try:
        with pdfplumber.open(pdf_file) as pdf:
            full_text = ''
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    full_text += text + '\n'
            
            # Invoice number patterns
            invoice_patterns = [
                r'Invoice\s*No\.\s*\n.*?(INV-[A-Z0-9\-]+)',  # Invoice No. \n ... INV-0020
                r'Invoice\s*No\.\s*([A-Z0-9\-]+)',  # Invoice No. INV-0020
                r'invoice\s+(TBI-\d{4}-\d+)',  # invoice TBI-2025-43
                r'Tax\s*Invoice\s*-\s*([\d]+)',
                r'Invoice\s*#\s*([\d]+)\b',
                r'Invoice\s*Number[:\s]*([A-Z0-9\-]+)',
                r'InvoiceNumber[:\s]*([A-Z]{3}-\d+)',
                r'\b(INV-\d+)\b',
                r'Invoice[:\s]*#?[:\s]*([A-Z0-9\-]+)',
            ]
            for pattern in invoice_patterns:
                match = re.search(pattern, full_text, re.IGNORECASE)
                if match:
                    invoice_data['invoice_number'] = match.group(1)
                    break
            
            # Date patterns
            date_patterns = [
                r'date\s+((?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4})',  # date November 13, 2025
                r'InvoiceDate[:\s]*(\d{1,2}[A-Za-z]{3}\d{4})',
                r'Invoice\s*Date[:\s]*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})',
                r'Date[:\s]*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})',
                r'\b(\d{1,2}[/-]\d{1,2}[/-]\d{4})\b'
            ]
            for pattern in date_patterns:
                match = re.search(pattern, full_text, re.IGNORECASE)
                if match:
                    invoice_data['date'] = match.group(1)
                    break
            
            # Due date patterns
            due_patterns = [
                r'Due\s*Date[:\s]*(\d{1,2}\s*[A-Za-z]+\s*\d{4})',
                r'Due[:\s]*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})',
                r'Payment\s*Due[:\s]*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})'
            ]
            for pattern in due_patterns:
                match = re.search(pattern, full_text, re.IGNORECASE)
                if match:
                    invoice_data['due_date'] = match.group(1)
                    break
            
            # Total amount patterns
            total_patterns = [                r'TOTAL\s+AUD\s+\$\s*([\d,]+\.\d{2})',  # TOTAL AUD $ 1,164.63
                r'Total\s+USD\s+\$\s*([\d,]+\.\d{2})',  # Total USD $ 20,179.45                r'Invoice\s*Totals\s+([\d,]+\.?\d*)\s+[\d,]+\.?\d*\s+[\d,]+\.?\d*\s+[\d,]+\.?\d*\s+([\d,]+\.?\d*)',
                r'Total\s*Amount\s*In\s*AUD[:\s]*([\d,]+\.?\d*)',
                r'NetAmountDue[:\s]*([\d,]+\.?\d*)',
                r'TOTAL\s*AUD[:\s]*([\d,]+\.?\d*)',
                r'Total\s*Amount[:\s]*\$?([\d,]+\.?\d*)',
                r'Amount\s*Due[:\s]*\$?([\d,]+\.?\d*)',
                r'TOTAL[:\s]*\$?([\d,]+\.?\d*)'
            ]
            for pattern in total_patterns:
                match = re.search(pattern, full_text, re.IGNORECASE)
                if match:
                    invoice_data['total_amount'] = match.group(1)
                    break
            
            # Tax/GST patterns
            tax_patterns = [
                r'TOTAL\s+GST\s+10%\s+\$\s*([\d,]+\.\d{2})',  # TOTAL GST 10% $ 105.88
                r'Invoice\s*Totals\s+[\d,]+\.?\d*\s+[\d,]+\.?\d*\s+[\d,]+\.?\d*\s+([\d,]+\.?\d*)',
                r'Tax\s*Amount\s*\(\d+%\)\s*In\s*AUD[:\s]*([\d,]+\.?\d*)',
                r'GST@\d+%[:\s]*([\d,]+\.?\d*)',
                r'TOTAL\s*GST\s*\d+%[:\s]*([\d,]+\.?\d*)',
                r'GST[:\s]*\$?([\d,]+\.?\d*)',
                r'Tax[:\s]*\$?([\d,]+\.?\d*)'
            ]
            for pattern in tax_patterns:
                match = re.search(pattern, full_text, re.IGNORECASE)
                if match:
                    invoice_data['tax_amount'] = match.group(1)
                    break
            
            # Subtotal patterns
            subtotal_patterns = [
                r'Subtotal\s+\$\s*([\d,]+\.\d{2})',  # Subtotal $ 20,179.45
                r'Invoice\s*Totals\s+([\d,]+\.?\d*)',
                r'Sub\s*Total\s*In\s*AUD[:\s]*([\d,]+\.?\d*)',
                r'GrossAmount[:\s]*([\d,]+\.?\d*)',
                r'Subtotal[:\s]*\$?([\d,]+\.?\d*)',
                r'Sub[\s-]*Total[:\s]*\$?([\d,]+\.?\d*)'
            ]
            for pattern in subtotal_patterns:
                match = re.search(pattern, full_text, re.IGNORECASE)
                if match:
                    invoice_data['subtotal'] = match.group(1)
                    break
            
            # Extract vendor/customer info
            lines = full_text.split('\n')
            for i, line in enumerate(lines[:15]):
                if any(suffix in line for suffix in ['PtyLtd', 'Pty Ltd', 'Ltd', 'LLC', 'Inc', 'Corporation']):
                    if not invoice_data['vendor_name']:
                        vendor = re.sub(r'([a-z])([A-Z])', r'\1 \2', line.strip())
                        invoice_data['vendor_name'] = vendor[:50]
            
            # Extract line items
            items_section_match = re.search(
                r'(?:Code\s+Description|Description.*?(?:Amount|Price|Net\s*Cost)).*?(?:\n|$)(.*?)(?:Payment\s*Terms|Sub\s*Total|Subtotal|TOTAL|Invoice\s*Totals|Tax\s*Amount)',
                full_text,
                re.IGNORECASE | re.DOTALL
            )
            
            if items_section_match:
                items_text = items_section_match.group(1)
            else:
                items_text = full_text
            
            item_pattern = r'([A-Za-z][A-Za-z0-9\s\-\(\),\.\']+?)[\s\-]+(\d+\.?\d*)\s+([\d,]+\.?\d+)\s+\d+%\s+([\d,]+\.?\d+)'
            items = re.findall(item_pattern, items_text)
            
            # Webjet format fallback
            if not items:
                webjet_pattern = r'(Hotel|Service Fee|Flight|Car|Transfer|Merchant Fee)\s+([\d,]+\.?\d+)\s+([\d,]+\.?\d+)\s+([\d,]+\.?\d+)\s+([\d,]+\.?\d+)\s+([\d,]+\.?\d+)'
                webjet_items = re.findall(webjet_pattern, full_text)
                for item in webjet_items:
                    invoice_data['items'].append({
                        'description': item[0],
                        'quantity': '1',
                        'price': item[5]
                    })
            
            # TBI format: PV-01 5.73 km $235.00 /km $1,346.55
            if not items and not invoice_data['items']:
                tbi_pattern = r'([A-Z]{2}-\d{2})\s+([\d\.]+\s+km)\s+\$[\d,]+\.?\d+\s+/km\s+\$([\d,]+\.?\d+)'
                tbi_items = re.findall(tbi_pattern, full_text)
                for item in tbi_items:
                    invoice_data['items'].append({
                        'description': item[0],
                        'quantity': item[1],
                        'price': item[2]
                    })
            
            # NPgeo single-line format: Palm Valley Repro Project: 6h 25m $165/hr 10% $1,058.75
            if not items and not invoice_data['items']:
                npgeo_pattern = r'(Palm\s+Valley[^:]+):\s+([\d]+h\s+[\d]+m)\s+\$[\d,]+/hr\s+\d+%\s+\$([\d,]+\.?\d+)'
                npgeo_items = re.findall(npgeo_pattern, full_text)
                for item in npgeo_items:
                    invoice_data['items'].append({
                        'description': item[0].strip(),
                        'quantity': item[1],
                        'price': item[2]
                    })
            
            for item in items[:20]:
                if len(item) >= 4:
                    desc = item[0].strip()
                    desc = re.sub(r'[\-\s\n]+$', '', desc)
                    desc = re.sub(r'\s+', ' ', desc)
                    
                    desc_lower = desc.lower()
                    skip_keywords = ['sub total', 'subtotal', 'total amount', 'invoice total', 'payment terms',
                                   'gst', 'net amount', 'header comments', 'tax amount', 'gross amount',
                                   'service fee totals', 'hotel totals', 'invoice totals']
                    
                    if (len(desc) > 3 and 
                        not any(keyword in desc_lower for keyword in skip_keywords) and
                        not desc.startswith(('ABN', 'Invoice', 'TAX', 'Header', 'Customer', 'Plant'))):
                        invoice_data['items'].append({
                            'description': desc,
                            'quantity': item[1],
                            'price': item[3]
                        })
            
            # Table extraction fallback
            if not invoice_data['items']:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if table and len(table) > 1:
                            for row in table[1:]:
                                if row and len(row) >= 2:
                                    desc = str(row[0]) if row[0] is not None and str(row[0]).strip() else ''
                                    if not desc and len(row) > 1:
                                        desc = str(row[1]) if row[1] is not None else ''
                                    
                                    qty = str(row[2]) if len(row) > 2 and row[2] is not None else ''
                                    if not qty:
                                        qty = str(row[1]) if len(row) > 1 and row[1] is not None else ''
                                    
                                    price = str(row[-1]) if row[-1] is not None else ''
                                    
                                    desc_lower = desc.lower()
                                    skip_keywords = ['sub total', 'subtotal', 'total amount', 'invoice total',
                                                   'tax amount', 'gst', 'payment terms', 'net amount',
                                                   'gross amount', 'total gst']
                                    
                                    if (desc and desc.strip() and 
                                        not any(keyword in desc_lower for keyword in skip_keywords)):
                                        invoice_data['items'].append({
                                            'description': desc,
                                            'quantity': qty,
                                            'price': price
                                        })
    
    except Exception as e:
        st.error(f"Error extracting data: {str(e)}")
    
    return invoice_data


def create_excel_from_invoices(invoice_data_list):
    """Create Excel file from multiple invoice data."""
    summary_data = []
    for data in invoice_data_list:
        summary_data.append({
            'Invoice Number': sanitize_for_excel(data.get('invoice_number', '')),
            'Date': sanitize_for_excel(data.get('date', '')),
            'Due Date': sanitize_for_excel(data.get('due_date', '')),
            'Vendor': sanitize_for_excel(data.get('vendor_name', '')),
            'Customer': sanitize_for_excel(data.get('customer_name', '')),
            'Subtotal': convert_to_number(data.get('subtotal', '')),
            'Tax': convert_to_number(data.get('tax_amount', '')),
            'Total Amount': convert_to_number(data.get('total_amount', ''))
        })
    
    items_data = []
    for data in invoice_data_list:
        invoice_num = sanitize_for_excel(data.get('invoice_number', ''))
        items = data.get('items', [])
        
        if items and isinstance(items, list):
            for item in items:
                if isinstance(item, dict):
                    desc = sanitize_for_excel(item.get('description', ''))
                    qty = convert_to_number(item.get('quantity', ''))
                    price = convert_to_number(item.get('price', ''))
                    
                    if desc or qty or price:
                        items_data.append({
                            'Invoice Number': invoice_num,
                            'Description': desc,
                            'Quantity': qty,
                            'Price': price
                        })
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if summary_data:
            df_summary = pd.DataFrame(summary_data)
            df_summary.to_excel(writer, sheet_name='Invoice Summary', index=False)
        
        if items_data:
            df_items = pd.DataFrame(items_data)
            df_items.to_excel(writer, sheet_name='Line Items', index=False)
        else:
            df_empty = pd.DataFrame(columns=['Invoice Number', 'Description', 'Quantity', 'Price'])
            df_empty.to_excel(writer, sheet_name='Line Items', index=False)
    
    output.seek(0)
    return output


# Main app
st.title("📄 Invoice Automation")
st.markdown("### Convert your PDF invoices to Excel spreadsheets instantly")

# File uploader
uploaded_files = st.file_uploader(
    "Drop PDF invoices here or click to browse",
    type=['pdf'],
    accept_multiple_files=True,
    help="Upload one or more PDF invoices to extract data"
)

if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} file(s) uploaded")
    
    # Show uploaded files
    with st.expander("📋 Uploaded Files", expanded=True):
        for file in uploaded_files:
            st.write(f"📄 {file.name} ({file.size / 1024:.1f} KB)")
    
    # Process button
    if st.button("🚀 Process Invoices & Download Excel", type="primary"):
        with st.spinner("Processing your invoices..."):
            invoice_data_list = []
            
            # Progress bar
            progress_bar = st.progress(0)
            for idx, file in enumerate(uploaded_files):
                invoice_data = extract_invoice_data(file)
                invoice_data_list.append(invoice_data)
                progress_bar.progress((idx + 1) / len(uploaded_files))
            
            # Create Excel
            excel_file = create_excel_from_invoices(invoice_data_list)
            
            st.success("✅ Processing complete!")
            
            # Download button
            st.download_button(
                label="📥 Download Excel File",
                data=excel_file,
                file_name="invoices.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # Show summary
            st.subheader("📊 Extraction Summary")
            cols = st.columns(4)
            
            total_items = sum(len(d['items']) for d in invoice_data_list)
            cols[0].metric("Invoices", len(invoice_data_list))
            cols[1].metric("Line Items", total_items)
            
            total_amount = sum(float(d['total_amount'].replace(',', '')) 
                             for d in invoice_data_list 
                             if d['total_amount'])
            cols[2].metric("Total Amount", f"${total_amount:,.2f}")
            
            extracted_count = sum(1 for d in invoice_data_list if d['invoice_number'])
            cols[3].metric("Extracted", f"{extracted_count}/{len(invoice_data_list)}")

# Features section
st.markdown("---")
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown("### ⚡ Fast Processing")
    st.write("Instantly extract data from multiple invoices")

with col2:
    st.markdown("### 📊 Excel Output")
    st.write("Organized sheets with summary and line items")

with col3:
    st.markdown("### 🔒 Secure")
    st.write("Your files are processed securely")

with col4:
    st.markdown("### 🤖 AI-Powered")
    st.write("Smart extraction of invoice data")

# Footer
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #64748b;'>Built with Streamlit & Python | "
    "<a href='https://github.com/yourusername/invoice-automation' style='color: #3b82f6;'>View on GitHub</a></div>",
    unsafe_allow_html=True
)
