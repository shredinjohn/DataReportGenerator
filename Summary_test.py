import streamlit as st
import pandas as pd
import openai
from dotenv import load_dotenv
from openai import OpenAI
load_dotenv()
import os

st.title("Data Report Generator 📈")


#Variable Declaration

sales_data = ""
target_data = ""

# Importing the datasets
sales_Data_name = st.file_uploader("Upload your Data File in XLSX format", type=['xlsx'])
target_Data_name = st.file_uploader("Upload your Target file in XLSX format", type=['xlsx'])
  
if st.button("Generate Report"):
    with st.spinner("Loading .. 🔃"):
      if sales_Data_name is not None and target_Data_name is not None:
        try:
            sales_data = pd.read_excel(sales_Data_name)
            target_data = pd.read_excel(target_Data_name)
        except Exception as e:
            st.write("Error reading excel files - ", e)  
      st.title("Data Report ")
      # Convert the 'Date' columns to datetime
      sales_data['Date'] = pd.to_datetime(sales_data['Date'])
      target_data['Date'] = pd.to_datetime(target_data['Date'])

      # Display the first few rows of each dataset to understand their structure
      # sales_data.head(), target_data.head()

      # Filter sales data
      sales= sales_data[(sales_data['Date'].dt.year == 2023) & (sales_data['Date'].dt.month == 12)]

      # Calculate total sales
      total_sales = sales['Sales_FC'].sum()


      # Filter target data for November 2023 and get the Company Total
      target= target_data[(target_data['Date'].dt.year == 2023) & (target_data['Date'].dt.month == 11)]
      company_total_target = target[target['Attributes'] == 'Company Total']['Value'].sum()



      print(company_total_target) 
      difference_sales_target =total_sales-company_total_target
      print(difference_sales_target )
      sales_to_target_ratio = total_sales / company_total_target
      print(sales_to_target_ratio)


          # Store the variables
      data = {
        'Total Sales': total_sales,
        'Company Total Target': company_total_target,
        'Difference (Sales - Target)': difference_sales_target ,
        'Sales/Target Ratio': sales_to_target_ratio 
      }

      
      prompt = f"""
      Here is the sales data for November 2023:
      Total Sales: {data['Total Sales']}
      Company Total Target: {data['Company Total Target']}
      Difference (Sales - Target): {data['Difference (Sales - Target)']}


      Please provide a brief two line observation based on the summary of this data.
      """

      client= openai.OpenAI(api_key=os.getenv("OPEN_API_KEY"))

      completion = client.chat.completions.create(
        model="gpt-4o",
        messages=[
          {"role": "system", "content": "You are a data analyst with ability to generate quality observation from the given data"},
          {"role": "user", "content": prompt}
        ]
      )


      # Print the observation
      observation=completion.choices[0].message.content

      st.write(observation)

      df = pd.DataFrame(list(data.items()), columns=['Metric', 'Value'])

      st.write(df)

      output_path = r"C:\\Users\\Joseph\\Desktop\\TEST\\SUMMARY.xlsx"
      df.to_excel(output_path, index=False)
      st.download_button(label="Download Summary", data=output_path, file_name='SUMMARY.xlsx', mime='application/vnd.ms-excel')
