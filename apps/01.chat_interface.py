import openai
from dotenv import load_dotenv
import os
import streamlit as st

load_dotenv()

openai.api_key = os.getenv("OPENAI_API_KEY")
openai.azure_endpoint = os.getenv("AZURE_ENDPOINT")
openai.api_type = os.getenv("OPENAI_API_TYPE")
openai.api_version = os.getenv("OPENAI_API_VERSION")

# TEST.....
# print("*" * 20)
# print("OpenAI API Key:", openai.api_key)

def get_openai_client(messages):

    try:
        response = openai.chat.completions.create(
            model="gpt-4.1-mini",
            messages=messages
        )
        return response.choices[0].message.content

    except Exception as e:
        st.error(f"OpenAI API 호출 중 오류가 발생했습니다.:  {e}")
        return f"오류: {e}"

# 채팅 기록의 초기화
st.title("Caht with OpenAI")
st.write("GPT와 대화해 보세요")

if "messages" not in st.session_state:
    st.session_state.messages = []

# 채팅 메시지 표시
for message in st.session_state.messages:
    st.chat_message(message["role"]).write(message["content"])

if user_input := st.chat_input("메시지를 입력하세요"):
    st.session_state.messages.append({"role": "user", "content": user_input})
    st.chat_message("user").write(user_input)

    with st.spinner("GPT가 응답하는 중..."):
        assistant_response = get_openai_client(st.session_state.messages)

    st.session_state.messages.append({"role": "assistant", "content": assistant_response})
    st.chat_message("assistant").write(assistant_response)

