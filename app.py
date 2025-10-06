"""
Manuscript Formatter Pro - Main Application
Professional DOCX Formatting for Evidance Health Sciences
"""

import streamlit as st
import io
from pathlib import Path
from document_processor import DocumentProcessor
from formatter import JournalFormatter
import traceback

# Page configuration
st.set_page_config(
    page_title="Manuscript Formatter Pro",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        text-align: center;
        padding: 2rem 0;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-radius: 10px;
        margin-bottom: 2rem;
    }
    .success-box {
        padding: 1.5rem;
        background-color: #d4edda;
        border-left: 5px solid #28a745;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .info-box {
        padding: 1.5rem;
        background-color: #d1ecf1;
        border-left: 5px solid #0c5460;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .warning-box {
        padding: 1.5rem;
        background-color: #fff3cd;
        border-left: 5px solid #856404;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .error-box {
        padding: 1.5rem;
        background-color: #f8d7da;
        border-left: 5px solid #721c24;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .stButton>button {
        width: 100%;
        background-color: #667eea;
        color: white;
        font-weight: bold;
        padding: 0.75rem;
        border-radius: 5px;
    }
    .stButton>button:hover {
        background-color: #764ba2;
    }
</style>
""", unsafe_allow_html=True)

def main():
    # Header
    st.markdown("""
    <div class="main-header">
        <h1>📄 Manuscript Formatter Pro</h1>
        <p style="font-size: 1.2rem; margin-top: 0.5rem;">
            Professional DOCX Formatting for Evidance Health Sciences
        </p>
    </div>
    """, unsafe_allow_html=True)

    # Sidebar
    with st.sidebar:
        st.image("https://via.placeholder.com/300x100/667eea/ffffff?text=Evidance+Health+Sciences", 
                 use_container_width=True)
        st.markdown("---")
        st.markdown("### 🎯 Features")
        st.markdown("""
        - ✅ Smart DOCX parsing
        - ✅ Citation detection
        - ✅ Auto figure/table placement
        - ✅ Professional formatting
        - ✅ Journal-ready output
        """)
        st.markdown("---")
        st.markdown("### ⚙️ Settings")
        
        # Formatting options
        font_size = st.selectbox(
            "Body Text Font Size",
            [10, 11, 12],
            index=1
        )
        
        line_spacing = st.selectbox(
            "Line Spacing",
            ["Single", "1.5 lines", "Double"],
            index=2
        )
        
        figure_width = st.slider(
            "Figure Width (inches)",
            min_value=3.0,
            max_value=7.0,
            value=6.0,
            step=0.5
        )
        
        st.markdown("---")
        st.markdown("### 📖 Instructions")
        st.info("""
        1. Upload your manuscript DOCX
        2. Review detected elements
        3. Adjust settings if needed
        4. Click 'Format Document'
        5. Download formatted DOCX
        """)

    # Main content area
    col1, col2 = st.columns([2, 1])

    with col1:
        st.markdown("### 📤 Upload Manuscript")
        uploaded_file = st.file_uploader(
            "Choose a DOCX file",
            type=['docx'],
            help="Upload your draft manuscript with figures and tables"
        )

        if uploaded_file is not None:
            try:
                # Save uploaded file temporarily
                file_bytes = uploaded_file.read()
                
                st.markdown('<div class="success-box">✅ File uploaded successfully!</div>', 
                           unsafe_allow_html=True)
                
                # Display file info
                st.markdown(f"**Filename:** {uploaded_file.name}")
                st.markdown(f"**Size:** {len(file_bytes) / 1024:.2f} KB")
                
                # Process button
                if st.button("🚀 Format Document", type="primary"):
                    with st.spinner("Processing document... This may take a minute."):
                        try:
                            # Create progress bar
                            progress_bar = st.progress(0)
                            status_text = st.empty()
                            
                            # Step 1: Initialize processor
                            status_text.text("Step 1/5: Initializing processor...")
                            progress_bar.progress(10)
                            processor = DocumentProcessor(io.BytesIO(file_bytes))
                            
                            # Step 2: Extract content
                            status_text.text("Step 2/5: Extracting document content...")
                            progress_bar.progress(30)
                            extracted_data = processor.extract_all_content()
                            
                            # Step 3: Detect citations
                            status_text.text("Step 3/5: Detecting citations...")
                            progress_bar.progress(50)
                            citations = processor.detect_citations()
                            
                            # Step 4: Format document
                            status_text.text("Step 4/5: Applying formatting...")
                            progress_bar.progress(70)
                            formatter = JournalFormatter(
                                font_size=font_size,
                                line_spacing=line_spacing,
                                figure_width=figure_width
                            )
                            formatted_doc = formatter.format_document(
                                extracted_data,
                                citations,
                                processor
                            )
                            
                            # Step 5: Save
                            status_text.text("Step 5/5: Finalizing document...")
                            progress_bar.progress(90)
                            
                            output_buffer = io.BytesIO()
                            formatted_doc.save(output_buffer)
                            output_buffer.seek(0)
                            
                            progress_bar.progress(100)
                            status_text.text("✅ Processing complete!")
                            
                            # Success message
                            st.markdown("""
                            <div class="success-box">
                                <h3>✨ Document Formatted Successfully!</h3>
                                <p>Your manuscript has been professionally formatted and is ready to download.</p>
                            </div>
                            """, unsafe_allow_html=True)
                            
                            # Download button
                            output_filename = uploaded_file.name.replace('.docx', '_FORMATTED.docx')
                            st.download_button(
                                label="📥 Download Formatted Document",
                                data=output_buffer.getvalue(),
                                file_name=output_filename,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                type="primary"
                            )
                            
                            # Show statistics
                            with st.expander("📊 Processing Statistics"):
                                col_a, col_b, col_c = st.columns(3)
                                with col_a:
                                    st.metric("Figures Found", len(extracted_data.get('figures', [])))
                                with col_b:
                                    st.metric("Tables Found", len(extracted_data.get('tables', [])))
                                with col_c:
                                    st.metric("Citations Detected", len(citations))
                            
                        except Exception as e:
                            st.markdown(f"""
                            <div class="error-box">
                                <h3>❌ Processing Error</h3>
                                <p><strong>Error:</strong> {str(e)}</p>
                                <p>Please check your document format and try again.</p>
                            </div>
                            """, unsafe_allow_html=True)
                            
                            with st.expander("🔍 Technical Details"):
                                st.code(traceback.format_exc())
                            
            except Exception as e:
                st.markdown(f"""
                <div class="error-box">
                    <h3>❌ File Upload Error</h3>
                    <p><strong>Error:</strong> {str(e)}</p>
                </div>
                """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("### ℹ️ About This Tool")
        st.markdown("""
        This application automatically formats your manuscript according to 
        professional medical journal standards used by leading publications.
        
        **What it does:**
        - Extracts figures and tables
        - Detects citation locations
        - Places visuals at citations
        - Applies consistent formatting
        - Generates publication-ready output
        
        **Formatting Standards:**
        - Times New Roman font
        - Structured headings
        - Proper margins (1 inch)
        - Professional spacing
        - Centered figures/tables
        - Caption formatting
        """)
        
        st.markdown("---")
        st.markdown("### 🛠️ Technical Info")
        st.markdown("""
        **Supported Citation Formats:**
        - Figure 1, Fig. 1, Figure 1A
        - Table 1, Tab. 1
        - (**Figure 1**), (Fig. 1)
        
        **Requirements:**
        - DOCX format only
        - Figures as images
        - Tables in document
        - Max file size: 50MB
        """)

    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 2rem 0;">
        <p><strong>Manuscript Formatter Pro</strong> v1.0.0</p>
        <p>Built for Evidance Health Sciences</p>
        <p style="font-size: 0.9rem;">© 2024 All rights reserved</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
