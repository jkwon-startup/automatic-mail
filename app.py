import streamlit as st
import smtplib
import requests
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.mime.image import MIMEImage
import mimetypes

# 구글 스프레드시트 API를 통해 시트 데이터 가져오기
def get_sheet_data(api_key, spreadsheet_id, sheet_range):
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}/values/{sheet_range}?key={api_key}"
    response = requests.get(url)
    if response.status_code == 200:
        return response.json().get("values", [])
    else:
        st.error(f"스프레드시트 데이터를 가져오는 데 실패했습니다. 상태 코드: {response.status_code}")
        return []

# 이메일 발송 함수 (이미지 본문 삽입 및 파일 첨부 포함)
def send_email(smtp_user, smtp_password, recipient, subject, body, image=None, attached_file=None):
    try:
        # 이메일 메시지 생성
        msg = MIMEMultipart()
        msg['From'] = smtp_user
        msg['To'] = recipient
        msg['Subject'] = subject

        # 이메일 본문에서 줄 바꿈 처리를 위해 \n을 <br>로 변환
        body_with_line_breaks = body.replace('\n', '<br>')
        msg.attach(MIMEText(body_with_line_breaks, 'html'))

        # 이미지 첨부 (본문 삽입)
        if image is not None:
            mime_type, _ = mimetypes.guess_type(image.name)
            if mime_type is None:
                mime_type = "image/png"  # 기본값으로 PNG 사용
            mime_main, mime_sub = mime_type.split('/')
            img_data = image.read()
            image_mime = MIMEImage(img_data, _subtype=mime_sub)
            image_mime.add_header('Content-ID', '<image1>')
            msg.attach(image_mime)

        # 파일 첨부
        if attached_file is not None:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attached_file.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename= {attached_file.name}')
            msg.attach(part)

        # Daum SMTP 서버 설정
        smtp_server = "smtp.daum.net"
        smtp_port = 465

        # 이메일 전송
        with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
            server.login(smtp_user, smtp_password)
            server.sendmail(smtp_user, recipient, msg.as_string())
    except Exception as e:
        st.error(f"{recipient}에게 이메일을 보내는 중 오류 발생: {e}")

# Streamlit UI 구성
st.title("대량 이메일 전송 시스템")

# Daum 이메일 설정 입력
st.sidebar.subheader("Daum 계정 설정")
smtp_user = st.sidebar.text_input("Daum 이메일 주소")
smtp_password = st.sidebar.text_input("Daum 비밀번호", type="password")
send_interval = st.sidebar.slider("전송 간격 (초)", 1, 10, 2)
daily_limit = st.sidebar.number_input("일일 전송 제한", min_value=1, max_value=1000, value=500)

# 구글 스프레드시트 API 설정
st.sidebar.subheader("구글 스프레드시트 설정")
api_key = st.sidebar.text_input("Google API 키")
spreadsheet_id = st.sidebar.text_input("스프레드시트 ID")
sheet_range = st.sidebar.text_input("시트 범위 (예: Sheet1!B2:B499)")

# 이메일 작성
st.subheader("이메일 작성")
subject = st.text_input("이메일 제목")
body = st.text_area("이메일 내용 (이미지를 삽입하려면 아래의 이미지를 업로드하세요)")

# 본문에 삽입할 이미지 업로드
uploaded_image = st.file_uploader("본문에 삽입할 이미지 선택 (옵션)", type=['png', 'jpg', 'jpeg'])

# 파일 첨부 (옵션)
attached_file = st.file_uploader("파일 첨부 (옵션)", type=['txt', 'pdf', 'png', 'jpg', 'jpeg', 'doc', 'docx'])

# 미리보기
if st.button("미리보기"):
    st.markdown(f"### 제목: {subject}")
    st.markdown(f"### 내용: \n {body.replace('\n', '<br>')}", unsafe_allow_html=True)
    if uploaded_image:
        st.image(uploaded_image)
    if attached_file:
        st.markdown(f"첨부 파일: {attached_file.name}")

# 대량 메일 전송
if st.button("대량 메일 전송"):
    if not subject or not body:
        st.error("이메일 제목과 내용을 입력하세요.")
    elif not smtp_user or not smtp_password or not api_key or not spreadsheet_id or not sheet_range:
        st.error("모든 설정을 입력하세요.")
    else:
        # 구글 스프레드시트 데이터 가져오기
        data = get_sheet_data(api_key, spreadsheet_id, sheet_range)
        if not data:
            st.error("스프레드시트에서 데이터를 가져오지 못했습니다.")
        else:
            sent_count = 0
            for row in data:
                if sent_count >= daily_limit:
                    st.warning("일일 전송 제한에 도달했습니다.")
                    break
                recipient = row[0]  # 첫 번째 열이 이메일 주소라고 가정

                # 이미지가 있을 경우 이메일 본문에 삽입
                email_body = body.replace('\n', '<br>')  # 줄 바꿈 변환
                if uploaded_image is not None:
                    email_body += '<br><img src="cid:image1">'

                send_email(smtp_user, smtp_password, recipient, subject, email_body, uploaded_image, attached_file)
                time.sleep(send_interval)  # 전송 간격 적용
                sent_count += 1
                st.success(f"{recipient}에게 이메일 전송 완료!")
            
            st.sidebar.subheader("전송 상태")
            st.sidebar.text(f"오늘 전송된 이메일 수: {sent_count}")
