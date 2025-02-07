import streamlit as st
import google.oauth2.credentials
import google_auth_oauthlib.flow
import googleapiclient.discovery
import base64
import json
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
import os
import time
import mimetypes

# Gmail APIのスコープ
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

# Streamlit Cloud の Secrets から OAuth 設定を取得
OAUTH_CONFIG = {
    "web": {
        "client_id": st.secrets["oauth"]["client_id"],
        "client_secret": st.secrets["oauth"]["client_secret"],
        "redirect_uris": [st.secrets["oauth"]["redirect_uri"]],
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token"
    }
}

# Google認証のリダイレクトURL (Streamlit CloudのURL)
REDIRECT_URI = st.secrets["oauth"]["redirect_uri"]

# セッション変数の初期化
if "credentials" not in st.session_state:
    st.session_state.credentials = None

st.title("📧 一斉送信 Gmail アプリ")

# 📂 Excelテンプレートのダウンロード機能
if st.button("📥 Excel テンプレートをダウンロード"):
    sample_df = pd.DataFrame({
        "to_email": ["user1@example.com", "user2@example.com"],
        "変数1": ["田中", "鈴木"],
        "変数2": ["商品A", "商品B"],
        "添付ファイル": ["file1.pdf, file2.png", "file3.docx"]  # カンマ区切りで複数ファイル対応
    })
    sample_df.to_csv("sample_template.xlsx", index=False)
    st.download_button(label="📂 サンプルExcelをダウンロード", data=sample_df.to_csv().encode(), file_name="sample_template.xlsx", mime="text/csv")

# URLのクエリパラメータを取得（認証コードが含まれているかチェック）
query_params = st.query_params
auth_code = query_params.get("code")

# 1️⃣ 認証コードがURLにある場合、自動で処理
if auth_code and st.session_state.credentials is None:
    try:
        flow = google_auth_oauthlib.flow.Flow.from_client_config(OAUTH_CONFIG, SCOPES)
        flow.redirect_uri = REDIRECT_URI
        flow.fetch_token(code=auth_code)

        # 認証情報をセッションに保存
        st.session_state.credentials = flow.credentials

        # クエリパラメータを削除してページをリロード
        st.query_params.clear()
        st.rerun()
    except Exception as e:
        st.error(f"認証に失敗しました: {e}")

# 2️⃣ 認証していない場合、Googleログインボタンを表示
if st.session_state.credentials is None:
    st.write("Googleログインが必要です")

    flow = google_auth_oauthlib.flow.Flow.from_client_config(OAUTH_CONFIG, SCOPES)
    flow.redirect_uri = REDIRECT_URI
    auth_url, _ = flow.authorization_url(prompt="consent")

    st.markdown(f'<a href="{auth_url}" target="_blank">🔗 Googleにログイン</a>', unsafe_allow_html=True)

# 3️⃣ 認証完了後、メール送信画面を表示
if st.session_state.credentials:
    service = googleapiclient.discovery.build("gmail", "v1", credentials=st.session_state.credentials)
    st.success("✅ Google 認証が完了しました！")

    uploaded_file = st.file_uploader("📂 送信リスト (Excel)", type=["xlsx"])
    attachment_files = st.file_uploader("📎 添付ファイル (複数選択可)", type=["pdf", "docx", "xlsx", "png", "jpg"], accept_multiple_files=True)

    attachment_dict = {file.name: file.getvalue() for file in attachment_files} if attachment_files else {}

    if uploaded_file:
        df = pd.read_excel(uploaded_file)

        if "to_email" not in df.columns:
            st.error("❌ 'to_email' カラムが必要です")
        else:
            st.write("📊 アップロードされたデータ:")
            st.dataframe(df.head())

            subject_template = st.text_input("✉ 件名テンプレート", "【お知らせ】{変数1}様へのご案内")
            body_template = st.text_area("📩 本文テンプレート", "こんにちは、{変数1}様。\n\nお世話になっております。\n\n詳細: {変数2}")

            if st.button("🚀 メール送信"):
                success_count = 0
                error_count = 0
                errors = []

                with st.status("📨 メール送信中...", expanded=True) as status:
                    for index, row in df.iterrows():
                        try:
                            to_email = row["to_email"]
                            subject = subject_template
                            body = body_template

                            for col in df.columns:
                                subject = subject.replace(f"{{{col}}}", str(row[col]))
                                body = body.replace(f"{{{col}}}", str(row[col]))

                            msg = MIMEMultipart()
                            msg["To"] = to_email
                            msg["Subject"] = subject
                            msg.attach(MIMEText(body, "plain"))

                            if "添付ファイル" in df.columns and pd.notna(row["添付ファイル"]):
                                attach_names = row["添付ファイル"].split(",")  # カンマ区切り対応
                                for attach_name in attach_names:
                                    attach_name = attach_name.strip()
                                    if attach_name in attachment_dict:
                                        file_data = attachment_dict[attach_name]
                                        mime_type, _ = mimetypes.guess_type(attach_name)
                                        if mime_type is None:
                                            mime_type = "application/octet-stream"
                                        main_type, sub_type = mime_type.split("/", 1)
                                        attachment_part = MIMEBase(main_type, sub_type)
                                        attachment_part.set_payload(file_data)
                                        encoders.encode_base64(attachment_part)
                                        attachment_part.add_header("Content-Disposition", f"attachment; filename={attach_name}")
                                        msg.attach(attachment_part)

                            raw_msg = base64.urlsafe_b64encode(msg.as_bytes()).decode()
                            service.users().messages().send(userId="me", body={"raw": raw_msg}).execute()

                            success_count += 1
                            st.write(f"✅ {index+1}. {to_email} へ送信成功")

                            time.sleep(1)

                        except Exception as e:
                            error_count += 1
                            errors.append(f"❌ {index+1}. {to_email} - エラー: {e}")
                            st.write(errors[-1])

                    status.update(label=f"📩 {success_count}件送信完了 / 失敗 {error_count}件", state="complete", expanded=False)

    if st.button("🔒 ログアウト"):
        st.session_state.credentials = None
        st.rerun()
