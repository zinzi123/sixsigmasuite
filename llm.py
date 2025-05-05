import streamlit as st
from openai import OpenAI

# Initialize the OpenAI client


# Function to interact with GPT-4
def generate_response(messages):
    response = client.chat.completions.create(
        model="gpt-4",
        messages=messages
    )
    return response.choices[0].message.content

# Streamlit App
st.set_page_config(page_title="Six Sigma Project Charter Bot", page_icon="🤖", layout="centered")

st.header("Six Sigma Project Charter Bot")
st.write("Your assistant for creating Six Sigma project charters. I’ll guide you through the process and summarize all details for you.")

# Initialize session state to store messages
if "messages" not in st.session_state:
    st.session_state.messages = [
        {"role": "system", "content": "You are a Six Sigma Project Chartering Assistant. Ask the user for the different details required to be input in a Six Sigma project charter, including Project Name, Project Manager, Project Sponsor, Project Team Members, Start Date, Expected Completion Date, Estimated Cost, Estimated Savings, Project Overview - Problem/Issue, Purpose of Project, Business Case, Goals/Metrics, Project Scope - Within Scope, Without Scope, Add today's date. Ask all these questions one by one from the user briefly. Clarify if the input from user is not clear. Once all the questions have been asked, summarize all the inputs given by the user with the respective titles and as bullet points. The user may type it in casual language; rephrase it to sound professional."},
        {"role": "assistant", "content": "Type '6' to start preparing the Six Sigma Charter"}
    ]

if "input_text" not in st.session_state:
    st.session_state.input_text = ""  # This will store the input text

def add_message(role, content):
    st.session_state.messages.append({"role": role, "content": content})

def display_chat():
    avatars = {"user": "user", "assistant": "assistant"}
    for message in st.session_state.messages:
        # Handle 'system' role separately, without an avatar
        if message["role"] == "system":
            st.write(message["content"])
        else:
            with st.chat_message(avatars[message["role"]]):
                st.write(message["content"])

# Display the chat history
display_chat()

# User input
prompt = st.text_input("You:", value=st.session_state.input_text, key="prompt", on_change=lambda: None)

if prompt:
    add_message("user", prompt)

    # Generate a response based on the entire conversation history
    response = generate_response(st.session_state.messages)

    add_message("assistant", response)

    # Clear the input field by resetting the session state variable
    st.session_state.input_text = ""  # Clear the stored text without modifying the widget key directly
    st.experimental_rerun()  # Rerun to display new messages
