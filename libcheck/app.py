import streamlit as st
import pandas as pd
import requests
import time
import io

# 1. Page Configuration
st.set_page_config(
    page_title="HKPL School Library Matcher",
    page_icon="📚",
    layout="centered"
)

# 2. App Styling & Header
st.title("📚 HKPL Catalog Bulk Matcher")
st.markdown("""
This application automates checking whether your school's books are available in the **Hong Kong Public Library (HKPL)** system using their ISBNs.
""")
st.markdown("---")

# 3. File Upload Component
uploaded_file = st.file_uploader("Step 1: Upload your School Library Excel file (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Read the Excel sheet
        df = pd.read_excel(uploaded_file)
        st.success("Excel file uploaded successfully!")
        
        # Step 2: Let the user select the exact column holding the ISBNs
        columns = df.columns.tolist()
        isbn_column = st.selectbox("Step 2: Select the column that contains your ISBNs:", columns)
        
        st.markdown("---")
        st.write("### Preview of Uploaded Data (First 5 Rows)")
        st.dataframe(df.head())
        
        # Step 3: Trigger the execution loop
        if st.button("🚀 Start Automatic HKPL Verification"):
            results = []
            
            # Interactive UI Progress Elements
            progress_bar = st.progress(0)
            status_text = st.empty()
            total_books = len(df)
            
            # Start timer
            start_time = time.time()
            
            for index, row in df.iterrows():
                # Clean up the ISBN string (removing trailing decimals or spaces)
                raw_isbn = str(row[isbn_column]).strip()
                isbn = raw_isbn.split('.')[0] if '.' in raw_isbn else raw_isbn
                
                status_text.text(f"Checking book {index + 1} of {total_books} (ISBN: {isbn})...")
                
                # Check for empty/invalid values
                if not isbn or isbn.lower() == 'nan' or len(isbn) < 9:
                    results.append("Invalid/Missing ISBN")
                    progress_bar.progress((index + 1) / total_books)
                    continue
                
                # HKPL Web Catalog Query URL
                url = f"https://webcat.hkpl.gov.hk/search/query?term_1={isbn}&field_1=isbn&theme=WEB"
                
                try:
                    # Spoof standard browser headers to avoid instant blocking
                    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
                    response = requests.get(url, headers=headers, timeout=10)
                    
                    # Analyze HKPL system response text
                    if "No record found" in response.text or "沒有找到符合的紀錄" in response.text:
                        results.append("No")
                    else:
                        results.append("Yes")
                        
                except Exception:
                    results.append("Error Connection Timeout")
                
                # Critical safety pause: respect HKPL servers so your app's IP doesn't get banned
                time.sleep(1.2)
                
                # Progress Update
                progress_bar.progress((index + 1) / total_books)
            
            # Calculate metrics
            elapsed_time = round(time.time() - start_time, 1)
            status_text.success(f"🎉 Verification Complete in {elapsed_time} seconds!")
            
            # Inject results into a copy of the dataframe
            output_df = df.copy()
            output_df['In HKPL Catalog'] = results
            
            # Show interactive final results preview
            st.markdown("---")
            st.write("### Verification Results Preview")
            st.dataframe(output_df.head(10))
            
            # Re-compile data stream back into an downloadable Excel spreadsheet
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                output_df.to_excel(writer, index=False)
            processed_data = output.getvalue()
            
            # Provide Browser Download Button
            st.download_button(
                label="📥 Download Updated Excel Sheet",
                data=processed_data,
                file_name="hkpl_checked_library_list.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"An error occurred reading the file: {e}")
