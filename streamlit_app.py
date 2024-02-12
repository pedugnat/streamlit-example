import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import xlwings as xw

st.set_option('deprecation.showPyplotGlobalUse', False)


# Set style
sns.set_style("whitegrid")

def generate_line_plot(output_values):
    x = np.linspace(0, 10, 10)
    y = output_values   
    plt.figure(figsize=(8, 6))
    plt.plot(x, y, color='blue', linewidth=2)
    plt.xlabel('X', fontsize=14)
    plt.ylabel('Y', fontsize=14)
    plt.title('Line Plot', fontsize=16)
    plt.xticks(fontsize=12)
    plt.yticks(fontsize=12)
    plt.grid(True)
    st.pyplot()

def main():
    st.title('Line Plot Generator')
    st.write("""
    Adjust the parameters to generate a line plot:
    """)
    
    with st.sidebar:
        st.subheader('Parameters')
        slope = st.slider('Slope', min_value=-10.0, max_value=10.0, value=1.0, step=0.1, format="%.1f")
        intercept = st.slider('Intercept', min_value=-10.0, max_value=10.0, value=0.0, step=0.1, format="%.1f")
    
    st.write(f'**Slope:** {slope}, **Intercept:** {intercept}')

    wb = xw.Book('dumb_model.xlsx', mode='i')

    # Access the inputs sheet
    inputs_sheet = wb.sheets['inputs']

    # Edit the values in the inputs sheet
    inputs_sheet.range('C2').value = slope
    inputs_sheet.range('C3').value = intercept

    # Calculate the workbook to update the outputs
    wb.app.calculate()

    # Access the outputs sheet
    outputs_sheet = wb.sheets['outputs']

    # Read the values from the outputs sheet
    output_values = outputs_sheet.range('C3:C12').value
    print(output_values)

    # Close the workbook
    wb.close()
    
    generate_line_plot(output_values)

if __name__ == "__main__":
    main()  


