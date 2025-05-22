import streamlit as st
import pandas as pd
from utils import generate_and_send_payslip

# --- 1. Set Light Black Background (Add at the very top) ---
st.markdown("""
    <style>
        /* Main background */
        .stApp {
            background-color: #f0f2f6;  /* Light black/gray */
        }
        
        /* Content containers */
        .main .block-container, 
        .st-emotion-cache-1y4p8pa {
            background-color: white;
            border-radius: 10px;
            padding: 2rem;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        
        /* Login container */
        .login-container {
            background-color: white !important;
        }
        
        /* File uploader */
        .file-uploader {
            background-color: white !important;
        }
        
        /* Email section */
        .email-section {
            background-color: #f8f9fa !important;
        }
    </style>
""", unsafe_allow_html=True)
# --- Login Check (Added at the top, preserves your original structure) ---
def check_login():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False
    
    if not st.session_state.logged_in:
        # Login Page Branding
        st.markdown("""
        <style>
            .login-container {
                max-width: 500px;
                margin: 0 auto;
                padding: 2rem;
                border-radius: 10px;
                box-shadow: 0 4px 12px rgba(0,0,0,0.1);
                background: white;
            }
            .login-title {
                color: #2c3e50;
                text-align: center;
                margin-bottom: 1.5rem;
            }
            .login-logo {
                text-align: center;
                font-size: 2.5rem;
                margin-bottom: 1rem;
            }
            .login-footer {
                text-align: center;
                margin-top: 2rem;
                color: #7f8c8d;
                font-size: 0.8rem;
            }
        </style>
        
        <div class="login-container">
            <div class="login-logo">💼</div>
            <h2 class="login-title">ENA COACH LTD</h2>
            <h3 style="text-align: center; color: #3498db; margin-bottom: 2rem;">Payslip Generator Portal</h3>
        """, unsafe_allow_html=True)

        col1, col2 = st.columns([1, 2])
        with col1:
            username = st.text_input("Username", key="username")
        with col2:
            password = st.text_input("Password", type="password", key="password")
        
        if st.button("Login", key="login-btn"):
            if username == "admin" and password == "1234":
                st.session_state.logged_in = True
                st.rerun()
            else:
                st.error("Invalid credentials. Please try again.")
        
        st.markdown("""
            <div class="login-footer">
                <p>Restricted access for authorized personnel only</p>
                <p>© 2023 ENA COACH LTD. All rights reserved</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
        st.stop()

check_login()  # Blocks app if not logged in  # Blocks app if not logged in

# --- Your Original App Code (Unchanged Below) ---
# App configuration
st.set_page_config(
    page_title="Payslip Generator",
    layout="centered",
    page_icon="💼"
)

# Custom CSS for professional styling
st.markdown("""
    <style>
        .main {
            max-width: 900px;
            padding: 2rem;
        }
        .header {
            color: #2c3e50;
            border-bottom: 1px solid #e0e0e0;
            padding-bottom: 1rem;
            margin-bottom: 2rem;
        }
        .stButton>button {
            background-color: #3498db;
            color: white;
            border-radius: 5px;
            padding: 0.5rem 1rem;
            border: none;
            font-weight: 500;
            transition: all 0.3s;
        }
        .stButton>button:hover {
            background-color: #2980b9;
            transform: translateY(-1px);
            box-shadow: 0 2px 5px rgba(0,0,0,0.2);
        }
        .stAlert {
            border-radius: 5px;
        }
        .stDataFrame {
            border-radius: 5px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        .file-uploader {
            border: 2px dashed #3498db;
            border-radius: 5px;
            padding: 2rem;
            text-align: center;
            margin-bottom: 2rem;
        }
        .email-section {
            background-color: #f8f9fa;
            padding: 1.5rem;
            border-radius: 5px;
            margin: 1rem 0;
        }
    </style>
""", unsafe_allow_html=True)

# Header with logo (you can replace with your actual logo)
st.markdown("""
    <div class="header">
        <div style="display: flex; align-items: center; gap: 1rem;">
            <h1 style="margin: 0;">ENA Coach Payslip Generator</h1>
        </div>
        <p style="color: #7f8c8d; margin: 0.5rem 0 0;">Generate and send professional payslips to your employees</p>
    </div>
""", unsafe_allow_html=True)

# Add logout button (New addition at the top of your app body)
if st.sidebar.button("Logout"):
    st.session_state.logged_in = False
    st.rerun()

# File upload section
st.markdown('<div class="file-uploader">', unsafe_allow_html=True)
uploaded_file = st.file_uploader(
    "Upload Payroll Excel File", 
    type=["xlsx"],
    help="Please upload an Excel file with columns: Month, Email, Name, Basic Salary, etc."
)
st.markdown('</div>', unsafe_allow_html=True)

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        if 'Month' not in df.columns or 'Email' not in df.columns:
            st.error("❌ Missing required columns: 'Month' or 'Email' must be present in the file.")
        else:
            # Data processing
            months = sorted(df['Month'].dropna().unique())
            selected_month = st.selectbox(
                "Select Month to Send Payslips", 
                months,
                help="Select the month for which you want to generate payslips"
            )
            filtered_df = df[df['Month'] == selected_month]

            # Display summary
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Employees Found", len(filtered_df))
            with col2:
                st.metric("Selected Month", selected_month)

            # Data preview with expander
            with st.expander("View Employee Data", expanded=False):
                st.dataframe(filtered_df.style.format({
                    'Basic Salary': '{:,.2f}',
                    'Net Salary': '{:,.2f}',
                    'Overtime': '{:,.2f}'
                }))

            # Email settings section
            st.markdown('<div class="email-section">', unsafe_allow_html=True)
            st.subheader("📧 Email Settings")
            
            col1, col2 = st.columns(2)
            with col1:
                sender_email = st.text_input(
                    "Sender Email",
                    placeholder="your.email@company.com",
                    help="The email address that will send the payslips"
                )
            with col2:
                sender_password = st.text_input(
                    "App Password",
                    type="password",
                    placeholder="Your email app password",
                    help="For Gmail, this is an app-specific password"
                )
            
            # Test email option
            test_email = st.checkbox(
                "Send test email to myself first",
                help="Send a sample payslip to your email before sending to all employees"
            )
            st.markdown('</div>', unsafe_allow_html=True)

            # Action button
            if st.button("📨 Send Payslips", key="send_button"):
                if not sender_email or not sender_password:
                    st.error("❌ Email and password are required to send payslips.")
                else:
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    if test_email:
                        # Send test email
                        test_row = filtered_df.iloc[0].copy()
                        test_row['Email'] = sender_email
                        try:
                            generate_and_send_payslip(test_row, sender_email, sender_password, selected_month)
                            st.success(f"✅ Test payslip sent to {sender_email}. Please verify before sending to all employees.")
                            if not st.checkbox("I've verified the test email and want to proceed"):
                                st.stop()
                        except Exception as e:
                            st.error(f"❌ Error sending test email: {e}")
                            st.stop()
                    
                    # Process all employees
                    success_count = 0
                    for i, (index, row) in enumerate(filtered_df.iterrows()):
                        try:
                            progress = int((i + 1) / len(filtered_df) * 100)
                            progress_bar.progress(progress)
                            status_text.text(f"Processing {i+1}/{len(filtered_df)}: {row['Name']}")
                            
                            generate_and_send_payslip(row, sender_email, sender_password, selected_month)
                            success_count += 1
                        except Exception as e:
                            st.error(f"❌ Error sending to {row['Email']}: {str(e)}")
                    
                    progress_bar.empty()
                    status_text.empty()
                    st.success(f"✅ Successfully sent {success_count} out of {len(filtered_df)} payslips!")

    except Exception as e:
        st.error(f"❌ Failed to process file: {str(e)}")

# Footer
st.markdown("---")
st.markdown("""
    <div style="text-align: center; color: #7f8c8d; font-size: 0.9rem;">
        <p>ENA COACH LTD • Payslip Generator v1.0</p>
    </div>
""", unsafe_allow_html=True)
