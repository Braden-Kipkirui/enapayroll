import streamlit as st
import pandas as pd
from utils import generate_and_send_payslip
import time
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="ENA Coach Payslip Generator",
    page_icon="📩",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
    <style>
    .main {
        padding: 2rem;
    }
    .stButton>button {
        width: 100%;
        margin-top: 1rem;
    }
    .upload-section {
        border: 2px dashed #cccccc;
        border-radius: 5px;
        padding: 2rem;
        text-align: center;
        margin: 2rem 0;
    }
    .success-message {
        padding: 1rem;
        background-color: #d4edda;
        border-radius: 5px;
        margin: 1rem 0;
    }
    </style>
""", unsafe_allow_html=True)

# Sidebar for company info and instructions
with st.sidebar:
    st.title("ℹ️ Information")
    st.image("https://placeholder.com/150x150", caption="Upload company logo")
    
    st.markdown("""
    ### Instructions
    1. Upload your Excel file containing employee data
    2. Select the month for payslip generation
    3. Enter your email credentials
    4. Review the data and send payslips
    
    ### Required Excel Columns
    - Name
    - Email
    - Month
    - Basic Salary
    - Overtime
    - Allowance
    - PAYE Tax
    - SHA
    - NSSF
    - Penalties
    - Deductions
    - Net Salary
    
    ### Need Help?
    Contact support at support@enacoach.co.ke
    """)

# Main content
st.title("📩 ENA Coach Payslip Generator & Email Sender")
st.markdown("Generate and send professional payslips to your employees securely.")

# File upload section with drag and drop
st.markdown('<div class="upload-section">', unsafe_allow_html=True)
uploaded_file = st.file_uploader(
    "Drag and drop your payroll Excel file here",
    type=["xlsx"],
    help="Upload an Excel file containing employee payroll data"
)
st.markdown('</div>', unsafe_allow_html=True)

if uploaded_file:
    with st.spinner('Processing file...'):
        try:
            df = pd.read_excel(uploaded_file)
            required_columns = ['Month', 'Email', 'Name', 'Basic Salary', 'Net Salary']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                st.error(f"Missing required columns: {', '.join(missing_columns)}")
            else:
                # Create two columns for the layout
                col1, col2 = st.columns([2, 1])
                
                with col1:
                    st.success("✅ File uploaded successfully!")
                    months = sorted(df['Month'].dropna().unique())
                    current_month = datetime.now().strftime("%B %Y")
                    default_month_index = months.index(current_month) if current_month in months else 0
                    selected_month = st.selectbox(
                        "Select Month to Send Payslips",
                        months,
                        index=default_month_index,
                        help="Choose the month for which you want to generate payslips"
                    )
                    
                    filtered_df = df[df['Month'] == selected_month]
                    st.metric("Number of Employees", len(filtered_df))
                    
                    # Display preview with pagination
                    st.subheader("📋 Data Preview")
                    page_size = 5
                    page_number = st.number_input("Page", min_value=1, max_value=(len(filtered_df) // page_size) + 1, value=1)
                    start_idx = (page_number - 1) * page_size
                    end_idx = start_idx + page_size
                    st.dataframe(filtered_df.iloc[start_idx:end_idx], use_container_width=True)
                
                with col2:
                    st.subheader("📧 Email Settings")
                    with st.form("email_settings"):
                        sender_email = st.text_input(
                            "Sender Email",
                            help="Enter your Gmail address"
                        )
                        sender_password = st.text_input(
                            "App Password",
                            type="password",
                            help="Enter your Gmail App Password (not your regular password)"
                        )
                        
                        submit_button = st.form_submit_button("📨 Send Payslips")
                        
                        if submit_button:
                            if not sender_email or not sender_password:
                                st.error("Email and password are required.")
                            else:
                                progress_bar = st.progress(0)
                                status_text = st.empty()
                                
                                for index, row in filtered_df.iterrows():
                                    try:
                                        progress = (index + 1) / len(filtered_df)
                                        status_text.text(f"Processing: {row['Name']}")
                                        
                                        generate_and_send_payslip(row, sender_email, sender_password, selected_month)
                                        progress_bar.progress(progress)
                                        
                                        st.success(f"✅ Sent to {row['Email']}")
                                        time.sleep(0.5)  # Small delay for better UX
                                        
                                    except Exception as e:
                                        st.error(f"❌ Error sending to {row['Email']}: {str(e)}")
                                
                                status_text.text("✨ All payslips processed!")

        except Exception as e:
            st.error(f"Failed to process file: {str(e)}")
            st.info("Please make sure your Excel file is properly formatted and contains all required columns.")
