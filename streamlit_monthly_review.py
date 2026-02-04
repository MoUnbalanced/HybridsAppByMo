import streamlit as st
import openpyxl
from pathlib import Path
import pandas as pd

# Page configuration
st.set_page_config(
    page_title="Monthly Review Portal",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for glassy design
st.markdown("""
<style>
    * {
        margin: 0;
        padding: 0;
    }
    
    html, body, [data-testid="stAppViewContainer"] {
        background: linear-gradient(135deg, #0f172a 0%, #1a3a3a 50%, #0f172a 100%);
        color: #e0f2fe;
    }
    
    [data-testid="stMainBlockContainer"] {
        padding-top: 2rem;
    }
    
    /* Header styling */
    .header-container {
        text-align: center;
        margin-bottom: 3rem;
        padding: 2rem;
        background: rgba(15, 23, 42, 0.5);
        border: 1px solid rgba(16, 185, 129, 0.2);
        border-radius: 16px;
        backdrop-filter: blur(10px);
    }
    
    .header-title {
        font-size: 2.5rem;
        font-weight: 900;
        background: linear-gradient(135deg, #6ee7b7 0%, #facc15 50%, #6ee7b7 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        margin-bottom: 0.5rem;
        letter-spacing: -1px;
    }
    
    .header-subtitle {
        color: #cbd5e1;
        font-size: 0.95rem;
        font-weight: 300;
        letter-spacing: 1px;
    }
    
    /* Glass morphism effect */
    .glass-card {
        background: rgba(15, 23, 42, 0.5);
        border: 1px solid rgba(16, 185, 129, 0.2);
        border-radius: 12px;
        backdrop-filter: blur(10px);
        padding: 1.5rem;
        margin-bottom: 1rem;
        transition: all 0.3s ease;
    }
    
    .glass-card:hover {
        background: rgba(15, 23, 42, 0.7);
        border-color: rgba(16, 185, 129, 0.4);
        box-shadow: 0 8px 32px rgba(16, 185, 129, 0.1);
    }
    
    /* Request cards */
    .request-card {
        background: rgba(15, 23, 42, 0.4);
        border: 1px solid rgba(16, 185, 129, 0.15);
        border-radius: 10px;
        padding: 1.25rem;
        margin-bottom: 0.75rem;
        transition: all 0.3s ease;
    }
    
    .request-card:hover {
        background: rgba(15, 23, 42, 0.6);
        border-color: rgba(16, 185, 129, 0.4);
        box-shadow: 0 4px 16px rgba(16, 185, 129, 0.1);
    }
    
    .card-title {
        color: #6ee7b7;
        font-size: 1.15rem;
        font-weight: 600;
        margin-bottom: 0.5rem;
    }
    
    .card-meta {
        color: #94a3b8;
        font-size: 0.9rem;
        margin-bottom: 0.5rem;
    }
    
    .card-time {
        color: #facc15;
        font-size: 0.85rem;
        margin: 0.5rem 0;
    }
    
    .card-notes {
        color: #cbd5e1;
        font-size: 0.8rem;
        font-style: italic;
        margin-top: 0.5rem;
    }
    
    /* Buttons */
    .stButton > button {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%);
        color: white;
        border: none;
        border-radius: 8px;
        font-weight: 600;
        font-size: 0.9rem;
        padding: 0.5rem 1.25rem;
        transition: all 0.3s ease;
        cursor: pointer;
    }
    
    .stButton > button:hover {
        background: linear-gradient(135deg, #6ee7b7 0%, #10b981 100%);
        box-shadow: 0 4px 12px rgba(16, 185, 129, 0.3);
        transform: translateY(-2px);
    }
    
    .gold-button > button {
        background: linear-gradient(135deg, #eab308 0%, #ca8a04 100%);
    }
    
    .gold-button > button:hover {
        background: linear-gradient(135deg, #facc15 0%, #eab308 100%);
        box-shadow: 0 4px 12px rgba(234, 179, 8, 0.3);
        transform: translateY(-2px);
    }
    
    /* Upload area */
    .uploadedFile {
        color: #6ee7b7 !important;
    }
    
    [data-testid="stFileUploadDropzone"] {
        border: 2px dashed rgba(16, 185, 129, 0.3) !important;
        background: rgba(15, 23, 42, 0.3) !important;
        border-radius: 12px !important;
    }
    
    [data-testid="stFileUploadDropzone"]:hover {
        border-color: rgba(16, 185, 129, 0.6) !important;
        background: rgba(15, 23, 42, 0.5) !important;
    }
    
    /* Stats boxes */
    .stat-box {
        background: rgba(15, 23, 42, 0.5);
        border: 1px solid rgba(16, 185, 129, 0.2);
        border-radius: 12px;
        padding: 1.5rem;
        text-align: center;
        backdrop-filter: blur(10px);
    }
    
    .stat-number {
        font-size: 2rem;
        font-weight: 900;
        background: linear-gradient(135deg, #6ee7b7 0%, #facc15 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }
    
    .stat-label {
        color: #94a3b8;
        font-size: 0.85rem;
        margin-top: 0.5rem;
        letter-spacing: 1px;
    }
    
    /* Expander */
    [data-testid="stExpander"] {
        background: rgba(15, 23, 42, 0.4) !important;
        border: 1px solid rgba(16, 185, 129, 0.2) !important;
        border-radius: 10px !important;
    }
    
    [data-testid="stExpander"] > div {
        background: rgba(15, 23, 42, 0.5) !important;
    }
    
    /* Success message */
    .stSuccess {
        background: rgba(16, 185, 129, 0.1) !important;
        border-left: 4px solid #10b981 !important;
    }
    
    .stSuccess > div {
        color: #6ee7b7 !important;
    }
    
    /* Info message */
    .stInfo {
        background: rgba(100, 150, 200, 0.1) !important;
        border-left: 4px solid #3b82f6 !important;
    }
    
    .stInfo > div {
        color: #93c5fd !important;
    }
    
    /* Columns */
    [data-testid="column"] {
        padding: 0 0.5rem;
    }
    
    /* Text styling */
    h1, h2, h3 {
        color: #e0f2fe !important;
    }
    
    /* Scrollbar */
    ::-webkit-scrollbar {
        width: 8px;
    }
    
    ::-webkit-scrollbar-track {
        background: rgba(15, 23, 42, 0.5);
    }
    
    ::-webkit-scrollbar-thumb {
        background: #10b981;
        border-radius: 4px;
    }
    
    ::-webkit-scrollbar-thumb:hover {
        background: #6ee7b7;
    }
    
    /* Footer */
    .footer {
        text-align: center;
        padding: 2rem;
        color: #94a3b8;
        border-top: 1px solid rgba(16, 185, 129, 0.1);
        margin-top: 3rem;
        font-size: 0.9rem;
    }
</style>
""", unsafe_allow_html=True)

# Helper function to generate message
def generate_message(request_data):
    return f"""We'd love to allocate a Monthly Review lesson to you! 
Here are the details:
👤 Student: {request_data['student_name']}
📚 Year and Subject: Year {request_data['year']} - {request_data['subject']}
🕒 Preferred Time Slot: {request_data['time_slot']}
📝 Topics or Notes: {request_data['notes']}
Please let us know if you're available to take this class.
✅ If you're happy with the time slot, we'll proceed with the setup.
⏳ If not, kindly suggest an alternative time and we'll confirm with the parent.
Looking forward to your response!"""

# Function to load requests from Excel
def load_requests(file_path):
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    
    # Column indices (0-based)
    name_idx = 0
    year_idx = 5
    subject_idx = 7
    requested_time_idx = 13
    available_times_idx = 14
    notes_idx = 16
    
    # Process each data row (starting from row 6)
    data_rows = list(ws.iter_rows(min_row=6, values_only=True))
    
    requests = []
    for row in data_rows:
        if row[0] is None:
            continue
        
        student_name = row[name_idx] if name_idx < len(row) else ""
        year = row[year_idx] if year_idx < len(row) else ""
        subject = row[subject_idx] if subject_idx < len(row) else ""
        requested_time = row[requested_time_idx] if requested_time_idx < len(row) else ""
        available_times = row[available_times_idx] if available_times_idx < len(row) else ""
        notes = row[notes_idx] if notes_idx < len(row) else ""
        
        time_slot = requested_time if requested_time else available_times
        
        requests.append({
            "student_name": student_name,
            "year": year,
            "subject": subject,
            "time_slot": time_slot,
            "notes": notes if notes else "No specific notes"
        })
    
    return requests

# Initialize session state
if 'requests' not in st.session_state:
    st.session_state.requests = []
if 'uploaded_file' not in st.session_state:
    st.session_state.uploaded_file = None

# Header
st.markdown("""
<div class="header-container">
    <div class="header-title">📚 Monthly Review Portal</div>
    <div class="header-subtitle">Created By Mohammed Abdelwahed</div>
</div>
""", unsafe_allow_html=True)

# Upload section
st.markdown('<div class="glass-card">', unsafe_allow_html=True)
uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx'], label_visibility="collapsed")
st.markdown('</div>', unsafe_allow_html=True)

if uploaded_file is not None:
    try:
        # Load requests
        st.session_state.requests = load_requests(uploaded_file)
        st.session_state.uploaded_file = uploaded_file.name
        
        # Stats
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown(f"""
            <div class="stat-box">
                <div class="stat-number">{len(st.session_state.requests)}</div>
                <div class="stat-label">TOTAL REQUESTS</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class="stat-box">
                <div class="stat-number">{st.session_state.uploaded_file[:20]}</div>
                <div class="stat-label">FILE LOADED</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown(f"""
            <div class="stat-box">
                <div class="stat-number">✓</div>
                <div class="stat-label">READY TO EXPORT</div>
            </div>
            """, unsafe_allow_html=True)
        
        # Export button
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        col1, col2 = st.columns([3, 1])
        with col2:
            st.markdown('<div class="gold-button">', unsafe_allow_html=True)
            if st.button("⬇️ Export All", use_container_width=True):
                messages = []
                for request in st.session_state.requests:
                    messages.append(generate_message(request))
                    messages.append("\n" + "="*80 + "\n")
                
                export_text = "\n".join(messages)
                st.download_button(
                    label="Download",
                    data=export_text,
                    file_name="monthly_review_messages.txt",
                    mime="text/plain",
                    use_container_width=True
                )
            st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Requests grid
        st.markdown("### Student Requests")
        
        for idx, request in enumerate(st.session_state.requests):
            st.markdown("""
            <div class="request-card">
            """, unsafe_allow_html=True)
            
            col1, col2 = st.columns([4, 1])
            
            with col1:
                st.markdown(f'<div class="card-title">{request["student_name"]}</div>', unsafe_allow_html=True)
                st.markdown(f'<div class="card-meta">Year {request["year"]} • {request["subject"]}</div>', unsafe_allow_html=True)
                st.markdown(f'<div class="card-time">🕒 {request["time_slot"]}</div>', unsafe_allow_html=True)
                st.markdown(f'<div class="card-notes">{request["notes"]}</div>', unsafe_allow_html=True)
            
            with col2:
                col2_1, col2_2 = st.columns(2)
                with col2_1:
                    if st.button("View", key=f"view_{idx}", use_container_width=True):
                        st.session_state[f"expand_{idx}"] = True
                
                with col2_2:
                    if st.button("Copy", key=f"copy_{idx}", use_container_width=True):
                        message = generate_message(request)
                        st.info(f"✅ Message for {request['student_name']} copied to clipboard!")
                        st.code(message, language="text")
            
            # Expandable message view
            if st.session_state.get(f"expand_{idx}", False):
                st.markdown('<div class="glass-card">', unsafe_allow_html=True)
                st.text_area(
                    label="Formatted Message",
                    value=generate_message(request),
                    height=250,
                    disabled=True,
                    key=f"message_{idx}"
                )
                
                col1, col2 = st.columns(2)
                with col1:
                    if st.button(f"📋 Copy Message", key=f"copy_full_{idx}", use_container_width=True):
                        st.success(f"Message copied to clipboard!")
                
                with col2:
                    if st.button(f"✕ Close", key=f"close_{idx}", use_container_width=True):
                        st.session_state[f"expand_{idx}"] = False
                        st.rerun()
                
                st.markdown('</div>', unsafe_allow_html=True)
            
            st.markdown('</div>', unsafe_allow_html=True)
    
    except Exception as e:
        st.error(f"Error loading file: {str(e)}")

else:
    # Empty state
    st.markdown("""
    <div class="glass-card" style="text-align: center; padding: 3rem;">
        <div style="color: #6ee7b7; font-size: 2rem; margin-bottom: 1rem;">📁</div>
        <div style="color: #e0f2fe; font-size: 1.2rem; margin-bottom: 0.5rem; font-weight: 600;">No requests yet</div>
        <div style="color: #94a3b8;">Upload an Excel file to get started</div>
    </div>
    """, unsafe_allow_html=True)

# Footer
st.markdown("""
<div class="footer">
    <p>Monthly Review Portal v1.0 • Built with Streamlit</p>
    <p style="font-size: 0.85rem;">© 2026 - All Rights Reserved</p>
</div>
""", unsafe_allow_html=True)
