import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from docx import Document
import docx.shared
import datetime

# Function to generate and return a Word document with multiple graphs and logo
def generate_report_with_logo(data, client_name, logo_path):
    doc = Document()
    
    # Ensure date is in datetime format and extract month
    data['Date OF Purchase'] = pd.to_datetime(data['Date OF Purchase'], format="%d/%m/%Y")
    data['Month'] = data['Date OF Purchase'].dt.to_period('M')
    data['eBay Price'] = data['eBay Price'].astype(float)
    data['Ali Express Price'] = data['Ali Express Price'].astype(float)

    # Add logo to every page's header
    section = doc.sections[0]
    header = section.header
    header_paragraph = header.paragraphs[0]
    header_paragraph.alignment = 1  # Center alignment
    run = header_paragraph.add_run()
    run.add_picture(logo_path, width=docx.shared.Inches(2.0))  # Adjust logo size as needed

    # Calculating total metrics
    total_sales = data['eBay Price'].sum()
    total_profit_per_sale = data['Profit Per Sale'].str.extract(r'(\d+.\d+)').dropna()[0].astype(float).sum()

    # Create graphs
    fig_list = []

    # Plot 1: Number of Sales per Month
    monthly_sales_count = data.groupby('Month').size()
    fig1, ax1 = plt.subplots()
    ax1.bar(monthly_sales_count.index.astype(str), monthly_sales_count, color="purple")
    ax1.set_title("Number of Sales per Month")
    ax1.set_xlabel("Month")
    ax1.set_ylabel("Number of Sales")
    ax1.tick_params(axis='x', rotation=45)
    fig_list.append(fig1)

    # Plot 2: Sum of Sales per Month
    monthly_sales_sum = data.groupby('Month')['eBay Price'].sum()
    fig2, ax2 = plt.subplots()
    ax2.bar(monthly_sales_sum.index.astype(str), monthly_sales_sum, color="blue")
    ax2.set_title("Sum of Sales per Month")
    ax2.set_xlabel("Month")
    ax2.set_ylabel("Sum of Sales (eBay Price)")
    ax2.tick_params(axis='x', rotation=45)
    fig_list.append(fig2)

    # Plot 3: Sum of Profit per Month
    monthly_profit_sum = data['Profit Per Sale'].str.extract(r'(\d+.\d+)').astype(float)
    monthly_profit_sum['Month'] = data['Month']
    monthly_profit_sum = monthly_profit_sum.groupby('Month')[0].sum()
    fig3, ax3 = plt.subplots()
    ax3.bar(monthly_profit_sum.index.astype(str), monthly_profit_sum, color="green")
    ax3.set_title("Sum of Profit per Month")
    ax3.set_xlabel("Month")
    ax3.set_ylabel("Sum of Profit")
    ax3.tick_params(axis='x', rotation=45)
    fig_list.append(fig3)

    # Plot 4: Profit Per Sale Over Time
    fig4, ax4 = plt.subplots()
    ax4.plot(data['Date OF Purchase'], data['Profit Per Sale'].str.extract(r'(\d+.\d+)').astype(float), color="orange")
    ax4.set_title("Profit Per Sale Over Time")
    ax4.set_xlabel("Date of Purchase")
    ax4.set_ylabel("Profit Per Sale ($)")
    ax4.tick_params(axis='x', rotation=45)
    fig_list.append(fig4)

    # Plot 5: eBay Price Over Time
    fig5, ax5 = plt.subplots()
    ax5.plot(data['Date OF Purchase'], data['eBay Price'], color="blue")
    ax5.set_title("eBay Price Over Time")
    ax5.set_xlabel("Date of Purchase")
    ax5.set_ylabel("eBay Price ($)")
    ax5.tick_params(axis='x', rotation=45)
    fig_list.append(fig5)

    # Plot 6: Average Profit per Product per Month
    avg_profit_per_product = monthly_profit_sum / monthly_sales_count
    fig6, ax6 = plt.subplots()
    ax6.plot(avg_profit_per_product.index.astype(str), avg_profit_per_product, color="red")
    ax6.set_title("Average Profit per Product per Month")
    ax6.set_xlabel("Month")
    ax6.set_ylabel("Average Profit")
    ax6.tick_params(axis='x', rotation=45)
    fig_list.append(fig6)

    # Add summary to the Word document
    doc.add_heading(f'Investment Report for {client_name}', 0)
    doc.add_paragraph(f"Total Sales: ${total_sales:.2f}")
    doc.add_paragraph(f"Total Profit (Per Sale): ${total_profit_per_sale:.2f}")

    # Save each graph into the Word document
    for i, fig in enumerate(fig_list):
        img_stream = BytesIO()
        fig.savefig(img_stream, format='png')
        img_stream.seek(0)
        
        # Add to Word document with titles
        doc.add_paragraph(f'Graph {i+1}:')
        doc.add_picture(img_stream)

    # Save the document to a buffer and return it
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    return buffer

# Streamlit app
st.title('Investment Report - Benda Management LTD')

st.write("Upload a CSV file with your sales data, and get a Word document with analysis and graphs including your logo.")

uploaded_file = st.file_uploader("Choose a CSV file", type="csv")

if uploaded_file:
    # Read the CSV into a DataFrame
    data = pd.read_csv(uploaded_file)
    
    # Ask the user for their name
    client_name = st.text_input("Enter the client's name", value="Client")
    
    # Provide the path to the logo
    logo_path = "/Users/benda/Desktop/BendaLTD/Logos/BENDA - logo black.png"  # Make sure the logo file is in the same directory
    
    if st.button("Generate Report"):
        # Generate the report as a Word document with the logo
        report_buffer = generate_report_with_logo(data, client_name, logo_path)
        
        # Download link for the report
        st.download_button(
            label="Download Report",
            data=report_buffer,
            file_name=f"{client_name}_Investment_Report_{datetime.date.today()}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
