import streamlit as st
import pandas as pd
from utils import generate_and_send_payslip

# Custom CSS for professional styling
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');
        
        * {
            font-family: 'Inter', sans-serif;
        }
        
        .main {
            max-width: 1000px;
            padding: 2rem;
        }
        
        .header {
            color: #1a237e;
            border-bottom: 2px solid #e3f2fd;
            padding-bottom: 1.5rem;
            margin-bottom: 2.5rem;
            background: linear-gradient(to right, #ffffff, #e3f2fd);
            padding: 2rem;
            border-radius: 10px;
            box-shadow: 0 2px 12px rgba(0,0,0,0.05);
        }
        
        .stButton>button {
            background: linear-gradient(45deg, #1e88e5, #1976d2);
            color: white;
            border-radius: 8px;
            padding: 0.6rem 1.2rem;
            border: none;
            font-weight: 600;
            transition: all 0.3s ease;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        .stButton>button:hover {
            background: linear-gradient(45deg, #1976d2, #1565c0);
            transform: translateY(-2px);
            box-shadow: 0 4px 15px rgba(25, 118, 210, 0.2);
        }
        
        .stButton>button:active {
            transform: translateY(0);
        }
        
        .stAlert {
            border-radius: 8px;
            border: none;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        }
        
        .stDataFrame {
            border-radius: 10px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.08);
        }
        
        .file-uploader {
            border: 2px dashed #1e88e5;
            border-radius: 10px;
            padding: 2.5rem;
            text-align: center;
            margin-bottom: 2.5rem;
            background: #f8f9fa;
            transition: all 0.3s ease;
        }
        
        .file-uploader:hover {
            border-color: #1976d2;
            background: #ffffff;
            box-shadow: 0 4px 15px rgba(0,0,0,0.05);
        }
        
        .email-section {
            background: linear-gradient(145deg, #ffffff, #f8f9fa);
            padding: 2rem;
            border-radius: 10px;
            margin: 1.5rem 0;
            box-shadow: 0 4px 15px rgba(0,0,0,0.05);
            border: 1px solid #e3f2fd;
        }
        
        .metric-card {
            background: white;
            padding: 1.5rem;
            border-radius: 10px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.05);
            border: 1px solid #e3f2fd;
            transition: all 0.3s ease;
        }
        
        .metric-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(0,0,0,0.08);
        }
        
        /* Custom styling for select boxes */
        .stSelectbox [data-baseweb="select"] {
            border-radius: 8px;
        }
        
        /* Custom styling for text inputs */
        .stTextInput>div>div>input {
            border-radius: 8px;
        }
        
        .footer {
            margin-top: 3rem;
            padding-top: 2rem;
            border-top: 1px solid #e3f2fd;
            text-align: center;
            color: #546e7a;
        }
    </style>
""", unsafe_allow_html=True)

# App configuration
st.set_page_config(
    page_title="ENA Coach Payslip Generator",
    layout="centered",
    page_icon="💼",
    initial_sidebar_state="collapsed"
)

# Header with modern design
st.markdown("""
    <div class="header">
        <div style="display: flex; align-items: center; gap: 1rem;">
            <h1 style="margin: 0; font-size: 2.2rem; font-weight: 600;">
                <span style="color: #1e88e5;">ENA</span> Coach Payslip Generator
            </h1>
        </div>
        <p style="color: #546e7a; margin: 0.8rem 0 0; font-size: 1.1rem;">
            Generate and distribute professional payslips efficiently and securely
        </p>
    </div>
""", unsafe_allow_html=True)

# File upload section with enhanced UI
st.markdown("""
    <div class="file-uploader">
        <div style="margin-bottom: 1rem;">
            <span style="font-size: 2rem;">📄</span>
        </div>
    </div>
""", unsafe_allow_html=True)
uploaded_file = st.file_uploader(
    "Upload Payroll Excel File", 
    type=["xlsx"],
    help="Please upload an Excel file containing employee payroll data"
)

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        if 'Month' not in df.columns or 'Email' not in df.columns:
            st.error("❌ The uploaded file is missing required columns: 'Month' and 'Email' must be present.")
        else:
            # Data processing with enhanced visuals
            months = sorted(df['Month'].dropna().unique())
            st.markdown("### 📅 Select Processing Month")
            selected_month = st.selectbox(
                "Select the month for payslip generation", 
                months,
                help="Choose the month for which you want to generate and send payslips"
            )
            filtered_df = df[df['Month'] == selected_month]

            # Enhanced metrics display
            st.markdown('<div style="display: flex; gap: 1rem; margin: 2rem 0;">', unsafe_allow_html=True)
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("""
                    <div class="metric-card">
                        <h3 style="margin: 0; color: #1e88e5; font-size: 2rem;">{}</h3>
                        <p style="margin: 0.5rem 0 0; color: #546e7a;">Employees</p>
                    </div>
                """.format(len(filtered_df)), unsafe_allow_html=True)
            with col2:
                st.markdown("""
                    <div class="metric-card">
                        <h3 style="margin: 0; color: #1e88e5; font-size: 2rem;">{}</h3>
                        <p style="margin: 0.5rem 0 0; color: #546e7a;">Processing Month</p>
                    </div>
                """.format(selected_month), unsafe_allow_html=True)

            # Enhanced data preview
            with st.expander("📋 Preview Employee Data", expanded=False):
                st.dataframe(
                    filtered_df.style.format({
                        'Basic Salary': '{:,.2f}',
                        'Net Salary': '{:,.2f}',
                        'Overtime': '{:,.2f}'
                    }).set_properties(**{
                        'background-color': '#f8f9fa',
                        'color': '#2c3e50',
                        'border-color': '#e3f2fd'
                    })
                )

            # Enhanced email settings section
            st.markdown('<div class="email-section">', unsafe_allow_html=True)
            st.markdown("### 📧 Email Configuration")
            
            col1, col2 = st.columns(2)
            with col1:
                sender_email = st.text_input(
                    "Sender Email Address",
                    placeholder="your.email@enacoach.com",
                    help="Enter the email address that will be used to send payslips"
                )
            with col2:
                sender_password = st.text_input(
                    "Email App Password",
                    type="password",
                    placeholder="Enter your email app password",
                    help="For security, use an app-specific password"
                )
            
            # Enhanced test email option
            test_email = st.checkbox(
                "🔍 Send a test payslip to myself first",
                help="Verify the payslip format by sending a test to your email"
            )
            st.markdown('</div>', unsafe_allow_html=True)

            # Enhanced action button
            if st.button("🚀 Generate & Send Payslips", key="send_button"):
                if not sender_email or not sender_password:
                    st.error("⚠️ Please provide both email and password to proceed.")
                else:
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    if test_email:
                        test_row = filtered_df.iloc[0].copy()
                        test_row['Email'] = sender_email
                        try:
                            generate_and_send_payslip(test_row, sender_email, sender_password, selected_month)
                            st.success("✅ Test payslip sent successfully! Please check your email.")
                            proceed = st.checkbox("✓ I've verified the test email and want to proceed with sending to all employees")
                            if not proceed:
                                st.stop()
                        except Exception as e:
                            st.error(f"❌ Test email failed: {str(e)}")
                            st.stop()
                    
                    # Process all employees with enhanced progress tracking
                    success_count = 0
                    for i, (index, row) in enumerate(filtered_df.iterrows()):
                        try:
                            progress = int((i + 1) / len(filtered_df) * 100)
                            progress_bar.progress(progress)
                            status_text.markdown(f"⏳ Processing {i+1}/{len(filtered_df)}: **{row['Name']}**")
                            
                            generate_and_send_payslip(row, sender_email, sender_password, selected_month)
                            success_count += 1
                        except Exception as e:
                            st.error(f"❌ Failed to send to {row['Email']}: {str(e)}")
                    
                    progress_bar.empty()
                    status_text.empty()
                    
                    if success_count == len(filtered_df):
                        st.balloons()
                    st.success(f"✨ Successfully sent {success_count} out of {len(filtered_df)} payslips!")

    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")

# Enhanced footer
st.markdown("---")
st.markdown("""
    <div class="footer">
        <p style="font-size: 1rem;">
            <strong>ENA COACH LTD</strong><br>
            <span style="font-size: 0.9rem;">Payslip Generator v1.0</span>
        </p>
        <p style="font-size: 0.8rem; margin-top: 0.5rem;">
            Built with ❤️ for efficiency and security
        </p>
    </div>
""", unsafe_allow_html=True)
