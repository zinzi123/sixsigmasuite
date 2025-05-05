import streamlit as st
import pandas as pd
import graphviz
from fpdf import FPDF
from io import BytesIO
from PIL import Image
import tempfile
from datetime import datetime

# Function to generate SIPOC Diagram (vertically aligned)
def generate_sipoc_diagram(sipoc_inputs):
    dot = graphviz.Digraph()

    # Sequentially add categories and their respective items, aligning them vertically
    previous_category = None
    categories = ["Suppliers", "Inputs", "Process", "Outputs", "Customers"]

    for category in categories:
        with dot.subgraph() as s:
            s.attr(rank='same')
            s.node(category, category, shape='box', style='filled', color='lightgrey')

            for item in sipoc_inputs.get(category, []):
                s.node(item, item)
                s.edge(category, item)
                if previous_category:
                    dot.edge(previous_category, category)

        previous_category = category

    return dot

# Function to generate Project Charter Excel file
def generate_project_charter_excel(charter_data, sipoc_inputs):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')

    # Save Project Charter with borders and headings
    df_charter = pd.DataFrame(charter_data.items(), columns=["Category", "Details"])
    df_charter.to_excel(writer, index=False, sheet_name="Project Charter")

    workbook = writer.book
    worksheet = writer.sheets['Project Charter']
    cell_format = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'top', 'text_wrap': True})
    header_format = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'valign': 'top', 'bg_color': '#D9D9D9'})
    
    worksheet.set_column('A:B', 30, cell_format)
    worksheet.set_row(0, None, header_format)

    # Save SIPOC Diagram
    sipoc_df = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in sipoc_inputs.items()]))
    sipoc_df.to_excel(writer, index=False, sheet_name="SIPOC Diagram")

    writer.close()
    processed_data = output.getvalue()
    return processed_data

# Function to generate PDF file with borders and margins
def generate_pdf(charter_data, sipoc_inputs, sipoc_graph):
    pdf = FPDF()
    pdf.set_left_margin(10)
    pdf.set_right_margin(10)
    pdf.add_page()

    # Add Date and Title with Borders
    pdf.set_font("Arial", 'B', size=14)
    pdf.cell(0, 10, txt=f"Six Sigma Project Charter", border=1, ln=True, align='C')

    # Add Date
    today = datetime.today().strftime('%Y-%m-%d')
    pdf.set_font("Arial", size=12)
    pdf.cell(0, 10, txt=f"Report Date: {today}", ln=True, align='R')

    # Add Project Charter Details without special characters
    pdf.set_font("Arial", size=12)
    for key, value in charter_data.items():
        pdf.ln(5)
        pdf.multi_cell(0, 10, txt=f"- {key}: {value}", border=0)

    # Add SIPOC Diagram
    pdf.add_page()
    pdf.cell(0, 10, txt="SIPOC Diagram", border=1, ln=True, align='C')

    # Convert Graphviz to Image
    sipoc_image = sipoc_graph.pipe(format='png')
    image = Image.open(BytesIO(sipoc_image))

    # Use a temporary file instead of a fixed path
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile:
        image.save(tmpfile.name)
        pdf.image(tmpfile.name, x=10, y=30, w=180)

    return pdf.output(dest='S').encode('latin1', errors='replace')

# Streamlit App
st.title("Six Sigma Project Charter & SIPOC Diagram Creator")

# Initialize session state
if 'chat_history' not in st.session_state:
    st.session_state.chat_history = []
    st.session_state.charter_data = {}
    st.session_state.sipoc_inputs = {}
    st.session_state.step = 'introduction'

# Introduction
if st.session_state.step == 'introduction':
    st.session_state.step = 'charter'

# Chatbot Conversation Flow
if st.session_state.step == 'charter':
    charter_questions = {
        "Problem Statement": "What is the problem statement for your Six Sigma project?",
        "Scope": "What is the scope of this project?",
        "Goal Statement": "What is the goal of the project?",
        "Team Members": "Who are the team members involved in this project?",
        "Timeline": "What is the project timeline?",
        "Business Case": "What is the business case for this project?",
    }
    default_answers = {
        "Problem Statement": "Baggage handling times are inconsistent, leading to customer dissatisfaction.",
        "Scope": "The project will focus on improving the baggage handling process at the airport.",
        "Goal Statement": "Reduce baggage handling time variance by 30% within 6 months.",
        "Team Members": "John Doe, Jane Smith, Airport Operations Team.",
        "Timeline": "6 months from project initiation.",
        "Business Case": "Improving baggage handling times will increase customer satisfaction and reduce costs associated with delays."
    }

    current_question_idx = len(st.session_state.charter_data)
    if current_question_idx < len(charter_questions):
        current_question_key = list(charter_questions.keys())[current_question_idx]
        current_question_text = charter_questions[current_question_key]

        with st.form(key=current_question_key):
            st.write(current_question_text)
            user_input = st.text_input("Your Answer", value=default_answers.get(current_question_key, ""))
            submitted = st.form_submit_button("Submit")

            if submitted:
                st.session_state.charter_data[current_question_key] = user_input if user_input else default_answers[current_question_key]
                st.experimental_rerun()

    else:
        st.session_state.step = 'sipoc'
        st.experimental_rerun()

elif st.session_state.step == 'sipoc':
    sipoc_questions = {
        "Suppliers": "List the Suppliers (comma-separated)",
        "Inputs": "List the Inputs (comma-separated)",
        "Process": "Describe the Process",
        "Outputs": "List the Outputs (comma-separated)",
        "Customers": "List the Customers (comma-separated)"
    }
    default_sipoc_answers = {
        "Suppliers": "Baggage handlers, Conveyor belt manufacturers",
        "Inputs": "Baggage, Conveyor belts, Handling staff",
        "Process": "Baggage is collected, sorted, and delivered to the correct flight.",
        "Outputs": "Sorted and delivered baggage",
        "Customers": "Passengers, Airlines"
    }

    current_question_idx = len(st.session_state.sipoc_inputs)
    if current_question_idx < len(sipoc_questions):
        current_question_key = list(sipoc_questions.keys())[current_question_idx]
        current_question_text = sipoc_questions[current_question_key]

        with st.form(key=current_question_key):
            st.write(current_question_text)
            user_input = st.text_input("Your Answer", value=default_sipoc_answers.get(current_question_key, ""))
            submitted = st.form_submit_button("Submit")

            if submitted:
                st.session_state.sipoc_inputs[current_question_key] = user_input.split(",") if user_input else default_sipoc_answers[current_question_key].split(",")
                st.experimental_rerun()

    else:
        st.session_state.step = 'generate'
        st.experimental_rerun()

# Generate Outputs
if st.session_state.step == 'generate':
    sipoc_graph = generate_sipoc_diagram(st.session_state.sipoc_inputs)
    excel_data = generate_project_charter_excel(st.session_state.charter_data, st.session_state.sipoc_inputs)
    pdf_data = generate_pdf(st.session_state.charter_data, st.session_state.sipoc_inputs, sipoc_graph)

    st.success("Generation complete!")
    st.download_button(label="Download Project Charter & SIPOC (Excel)", data=excel_data, file_name="Project_Charter_SIPOC.xlsx")
    st.download_button(label="Download Project Charter & SIPOC (PDF)", data=pdf_data, file_name="Project_Charter_SIPOC.pdf")
    st.graphviz_chart(sipoc_graph)
