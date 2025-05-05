import streamlit as st
import pandas as pd
from io import BytesIO
import graphviz

# Define default values for the airline industry case
default_problem_statement = "Inconsistent baggage handling time at the airport."
default_goal_statement = "Reduce baggage handling time to under 20 minutes for 95% of flights."
default_scope = "Baggage handling processes at all domestic terminals."
default_business_case = "Reducing baggage handling time will improve customer satisfaction and reduce operational costs."
default_suppliers = "Baggage handlers, Ground staff"
default_inputs = "Passenger baggage, Conveyor belts"
default_process_steps = "Check-in, Sorting, Loading, Unloading, Delivery to carousel"
default_outputs = "Baggage delivered to passengers"
default_customers = "Passengers, Airline staff"

# Streamlit app layout
st.title("Six Sigma Define Phase")

st.subheader("Project Charter", divider='orange')

# Project Charter Inputs
problem_statement = st.text_area(
    "Problem Statement", 
    value=default_problem_statement, 
    help="Describe the problem you are addressing."
)
goal_statement = st.text_area(
    "Goal Statement", 
    value=default_goal_statement, 
    help="Define the specific, measurable goal of the project."
)
scope = st.text_area(
    "Scope", 
    value=default_scope, 
    help="Describe the boundaries of the project."
)
business_case = st.text_area(
    "Business Case", 
    value=default_business_case, 
    help="Explain why this project is important for the business."
)

st.subheader("SIPOC", divider='orange')

# SIPOC Diagram Inputs
suppliers = st.text_area(
    "Suppliers", 
    value=default_suppliers, 
    help="List the suppliers involved in the process."
)
inputs = st.text_area(
    "Inputs", 
    value=default_inputs, 
    help="List the inputs needed for the process."
)
process_steps = st.text_area(
    "Process Steps", 
    value=default_process_steps, 
    help="Describe the key steps of the process."
)
outputs = st.text_area(
    "Outputs", 
    value=default_outputs, 
    help="List the outputs of the process."
)
customers = st.text_area(
    "Customers", 
    value=default_customers, 
    help="Identify the customers who receive the outputs."
)

# Review Button
if st.button("Review Project Charter and SIPOC Diagram"):
    st.subheader("Review")
    st.header("Project Charter", divider='orange')
    st.write(f"**Problem Statement:** {problem_statement}")
    st.write(f"**Goal Statement:** {goal_statement}")
    st.write(f"**Scope:** {scope}")
    st.write(f"**Business Case:** {business_case}")
    
    st.header("SIPOC", divider='orange')
    st.write(f"**Suppliers:** {suppliers}")
    st.write(f"**Inputs:** {inputs}")
    st.write(f"**Process Steps:** {process_steps}")
    st.write(f"**Outputs:** {outputs}")
    st.write(f"**Customers:** {customers}")

    # Generate a compact SIPOC Diagram
    dot = graphviz.Digraph(comment='SIPOC Diagram', format='png')
    
    # Add SIPOC categories as nodes and align them horizontally
    dot.node('Suppliers', 'Suppliers', shape='rect', style='filled', color='lightblue')
    dot.node('Inputs', 'Inputs', shape='rect', style='filled', color='lightblue')
    dot.node('Process Steps', 'Process Steps', shape='rect', style='filled', color='lightblue')
    dot.node('Outputs', 'Outputs', shape='rect', style='filled', color='lightblue')
    dot.node('Customers', 'Customers', shape='rect', style='filled', color='lightblue')

    # Add elements within each category and align them vertically under each category
    for i, supplier in enumerate(suppliers.split(',')):
        dot.node(f'Supplier_{i}', supplier.strip(), shape='ellipse')
        dot.edge(f'Supplier_{i}', 'Suppliers')

    for i, input_item in enumerate(inputs.split(',')):
        dot.node(f'Input_{i}', input_item.strip(), shape='ellipse')
        dot.edge(f'Input_{i}', 'Inputs')

    for i, step in enumerate(process_steps.split(',')):
        dot.node(f'Step_{i}', step.strip(), shape='ellipse')
        dot.edge(f'Step_{i}', 'Process Steps')

    for i, output in enumerate(outputs.split(',')):
        dot.node(f'Output_{i}', output.strip(), shape='ellipse')
        dot.edge(f'Output_{i}', 'Outputs')

    for i, customer in enumerate(customers.split(',')):
        dot.node(f'Customer_{i}', customer.strip(), shape='ellipse')
        dot.edge(f'Customer_{i}', 'Customers')
    
    # Render the SIPOC diagram
    st.graphviz_chart(dot.source)

# Function to export data to Excel
def export_to_excel():
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')

    # Project Charter Data
    project_charter_data = {
        "Section": ["Problem Statement", "Goal Statement", "Scope", "Business Case"],
        "Content": [problem_statement, goal_statement, scope, business_case]
    }
    project_charter_df = pd.DataFrame(project_charter_data)

    # SIPOC Data
    sipoc_data = {
        "SIPOC": ["Suppliers", "Inputs", "Process Steps", "Outputs", "Customers"],
        "Details": [suppliers, inputs, process_steps.replace("\n", ", "), outputs, customers]
    }
    sipoc_df = pd.DataFrame(sipoc_data)

    # Write data to Excel
    project_charter_df.to_excel(writer, sheet_name='Project Charter', index=False)
    sipoc_df.to_excel(writer, sheet_name='SIPOC', index=False)

    # Format Excel sheets
    workbook  = writer.book
    worksheet1 = writer.sheets['Project Charter']
    worksheet2 = writer.sheets['SIPOC']

    # Set column widths and add formatting
    for sheet in [worksheet1, worksheet2]:
        sheet.column_dimensions['A'].width = 20
        sheet.column_dimensions['B'].width = 50

    # Save Excel file
    writer.close()
    output.seek(0)

    return output

# Export Button
if st.button("Export Project Charter and SIPOC Excel"):
    excel_file = export_to_excel()
    st.download_button(
        label="Download Excel",
        data=excel_file,
        file_name='Project_Charter_SIPOC.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    st.success("Your project charter and SIPOC have been exported as 'Project_Charter_SIPOC.xlsx'!")
