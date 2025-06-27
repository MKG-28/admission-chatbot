import streamlit as st
from chatbot import AdmissionChatbot

# Initialize chatbot
bot = AdmissionChatbot("data/Chatbot Questions & Answers.xlsx")

# Streamlit UI
st.set_page_config(page_title="🎓 Kepler Admission Assistant", layout="centered")
st.title("🎓 Kepler College Admission Assistant")
st.markdown("Ask me anything about admissions, programs, orientation, and more!")

# Chat history
if "messages" not in st.session_state:
    st.session_state.messages = []

# Display chat messages
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

# User input
if prompt := st.chat_input("Type your question here..."):
    # Add user message
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    # Get bot response
    response = bot.get_response(prompt)
    st.session_state.messages.append({"role": "assistant", "content": response})
    with st.chat_message("assistant"):
        st.markdown(response)