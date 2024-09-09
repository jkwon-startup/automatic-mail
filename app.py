import streamlit as st
import smtplib
import requests
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.mime.image import MIMEImage

# 구글 스프레드시트 API를 통해 시트 데이터 가져오기
def get_sheet_data(api_key, spreadsheet_id, sheet_range):
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}/values/{sheet_range}?key={api_key}"
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json().get("values", [])
        if not data:
            st.error("스프레드시트에서 데이터를 찾을 수 없습니다. 시트 범위를 확인하세요.")
        return data
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

        # 이메일 본문
        msg.attach(MIMEText(body, 'html'))

        # 이미지 첨부 (본문 삽입)
        if image is not None:
            img_data = image.read()
            image_mime = MIMEImage(img_data)
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
sheet_range = st.sidebar.text_input("시트 범위 (예: Sheet1!A2:B499)")

# 이메일 작성
st.subheader("이메일 작성")
subject = st.text_input("이메일 제목")
body_template = st.text_area("이메일 내용 (본문 상단에 {이름}님!! 으로 이름이 추가됩니다)")

# 본문에 삽입할 이미지 업로드
uploaded_image = st.file_uploader("본문에 삽입할 이미지 선택 (옵션)", type=['png', 'jpg', 'jpeg'])

# 파일 첨부 (옵션)
attached_file = st.file_uploader("파일 첨부 (옵션)", type=['txt', 'pdf', 'png', 'jpg', 'jpeg', 'doc', 'docx'])

# 미리보기
if st.button("미리보기"):
    if "{이름}" not in body_template:
        st.warning("미리보기에서는 {이름} 변수를 사용해야만 이름을 확인할 수 있습니다.")
    st.markdown(f"### 제목: {subject}")
    # 줄바꿈을 반영한 미리보기
    st.markdown(f"### 내용 (이름 적용): \n {body_template.replace('{이름}', '홍길동').replace('\n', '<br>')}")
    if uploaded_image:
        st.image(uploaded_image)
    if attached_file:
        st.markdown(f"첨부 파일: {attached_file.name}")

# 대량 메일 전송
if st.button("대량 메일 전송"):
    if not subject or not body_template:
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
            for idx, row in enumerate(data, start=2):
                if len(row) < 2:
                    st.warning(f"{idx}행에 이메일 주소만 존재하고 이름이 없습니다: {row}")
                    recipient_name = "고객님"  # 이름이 없을 경우 기본값으로 '고객님' 사용
                    recipient_email = row[0].strip()  # 이메일 주소만 있는 경우
                else:
                    recipient_name = row[0].strip()  # A열의 이름 (여백 제거)
                    recipient_email = row[1].strip()  # B열의 이메일 (여백 제거)

                if not recipient_email:
                    st.warning(f"{idx}행에 올바르지 않은 이메일 주소가 있습니다: {row}")
                    continue

                if sent_count >= daily_limit:
                    st.warning("일일 전송 제한에 도달했습니다.")
                    break

                # {이름} 변수를 사용하여 동적으로 본문에 이름을 추가하고 줄바꿈을 처리
                email_body = body_template.replace("{이름}", recipient_name).replace("\n", "<br>")

                # 이미지가 있을 경우 이메일 본문에 삽입
                if uploaded_image is not None:
                    email_body += '<br><img src="cid:image1">'

                # 이메일 전송
                send_email(smtp_user, smtp_password, recipient_email, subject, email_body, uploaded_image, attached_file)
                
                # 전송 간격 적용
                time.sleep(send_interval)
                
                sent_count += 1
                st.success(f"{recipient_email} ({recipient_name})에게 이메일 전송 완료!")
            
            st.sidebar.subheader("전송 상태")
            st.sidebar.text(f"오늘 전송된 이메일 수: {sent_count}")

