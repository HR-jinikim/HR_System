import streamlit as st
import pandas as pd
from datetime import datetime
import io
import time
from docx import Document

# OpenAI 라이브러리 체크
try:
    from openai import OpenAI
    openai_installed = True
except ImportError:
    openai_installed = False

# 1. 페이지 설정 (화면을 꽉 차게 씀)
st.set_page_config(page_title="채용 서포트 시스템", layout="wide")

# --- 데이터 로드 ---
@st.cache_data
def load_data():
    try:
        df = pd.read_excel("data.xlsx")
        data = {}
        for index, row in df.iterrows():
            data[row['직무명']] = {
                "jd": row['JD'],
                "questions": {
                    "Level 1": row['질문_Lv1'],
                    "Level 2": row['질문_Lv2'],
                    "Level 3": row['질문_Lv3'],
                    "Level 4": row['질문_Lv4']
                }
            }
        return data
    except Exception as e:
        return None

jd_data = load_data()

# --- 워드 파일 생성 함수 ---
def create_word_file(position, level, comments, result, question):
    doc = Document()
    doc.add_heading('면접 결과 리포트', 0)
    doc.add_heading('1. 기본 정보', level=1)
    doc.add_paragraph(f'면접 일시: {datetime.now().strftime("%Y-%m-%d %H:%M")}')
    doc.add_paragraph(f'지원 포지션: {position}')
    doc.add_heading('2. 역량 평가', level=1)
    doc.add_paragraph(f'평가 레벨: {level}')
    doc.add_paragraph(f'질문 내용: {question}')
    doc.add_heading('3. 면접관 코멘트', level=1)
    doc.add_paragraph(comments)
    doc.add_heading('4. 종합 결과', level=1)
    doc.add_paragraph(f'채용 추천 여부: {result}')
    bio = io.BytesIO()
    doc.save(bio)
    return bio

# --- 메인 화면 시작 ---

st.title("🤝 면접 서포트 어시스턴트")
st.markdown("---")

# 엑셀 파일 체크 (파일 없으면 여기서 경고 띄움)
if jd_data is None:
    st.error("🚨 'data.xlsx' 파일을 찾을 수 없거나 형식이 잘못되었습니다! 폴더에 파일이 있는지 확인해주세요.")
    st.stop()

# 직무 선택
selected_position = st.selectbox("진행할 면접 포지션을 선택하세요:", list(jd_data.keys()))
st.markdown("---")

# ★★★ 여기가 핵심! 화면을 3개로 나눕니다 ★★★
# 비율조절: 왼쪽(1) : 가운데(1.2) : 오른쪽(0.8 - AI용)
col1, col2, col3 = st.columns([1, 1.2, 0.8])

# [1구역: 왼쪽] JD
with col1:
    st.info(f"📋 {selected_position} JD")
    # 내용이 길면 스크롤 생기도록 높이 고정 (height=600)
    with st.container(height=600):
        st.markdown(str(jd_data[selected_position]["jd"]).replace("\n", "  \n"))

# [2구역: 가운데] 평가표
with col2:
    st.success("📝 면접 평가")
    with st.container(height=600): # 높이를 맞춰서 깔끔하게
        st.write("#### 1. 역량 레벨 체크")
        level = st.radio("레벨 선택", ["Level 1", "Level 2", "Level 3", "Level 4"], horizontal=True)
        
        current_question = jd_data[selected_position]['questions'][level]
        st.warning(f"💡 **질문 가이드:**\n\n{current_question}")
        
        st.markdown("---")
        
        st.write("#### 2. 면접관 코멘트")
        comments = st.text_area("상세 의견 작성", height=100)
        
        result = st.radio("최종 결과", ["채용 추천 (Pass)", "보류/불합격 (Fail)"], horizontal=True)
        
        st.markdown("---")
        
        # 다운로드 버튼
        word_file = create_word_file(selected_position, level, comments, result, current_question)
        st.download_button(
            label="📥 결과 리포트 다운로드",
            data=word_file.getvalue(),
            file_name=f"면접결과_{selected_position}_{datetime.now().strftime('%Y%m%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )

# [3구역: 오른쪽] AI 도우미
with col3:
    st.markdown("### 🤖 AI 도우미")
    
    # 채팅창을 박스 안에 가둠 (깔끔하게)
    with st.container(height=600, border=True):
        # API 키 입력
        api_key = st.text_input("🔑 API Key 입력", type="password", placeholder="없으면 체험판 모드")
        
        if not api_key:
            st.caption("※ 키가 없으면 체험판 봇이 응답합니다.")

        # 채팅 기록 초기화
        if "messages" not in st.session_state:
            st.session_state["messages"] = [{"role": "assistant", "content": "안녕하세요! 무엇을 도와드릴까요?"}]

        # 대화 내용 표시
        for msg in st.session_state.messages:
            st.chat_message(msg["role"]).write(msg["content"])

        # 입력창 (채팅창 하단에 고정됨)
        if prompt := st.chat_input("질문을 입력하세요..."):
            # 사용자 메시지
            st.session_state.messages.append({"role": "user", "content": prompt})
            st.chat_message("user").write(prompt)

            # AI 응답
            msg = ""
            if not api_key:
                time.sleep(1)
                msg = "📢 [체험판] 키가 입력되지 않았습니다.\n\n(실제라면 여기서 똑똑한 답변을 해줍니다!)"
            else:
                if openai_installed:
                    try:
                        client = OpenAI(api_key=api_key)
                        response = client.chat.completions.create(
                            model="gpt-3.5-turbo",
                            messages=st.session_state.messages
                        )
                        msg = response.choices[0].message.content
                    except Exception as e:
                        msg = f"❌ 오류: {e}"
                else:
                    msg = "OpenAI 설치가 필요합니다."

            st.session_state.messages.append({"role": "assistant", "content": msg})
            st.chat_message("assistant").write(msg)