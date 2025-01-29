import streamlit as st
import pandas as pd
import google.oauth2.credentials
import google_auth_oauthlib.flow
import googleapiclient.discovery
import base64
from email.mime.text import MIMEText
import os
import time

# OAuth 2.0 クライアントIDのJSONファイル
CLIENT_SECRETS_FILE = "client_secret.json"

# Gmail APIのスコープ
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

# Google認証のリダイレクトURL (Streamlitのポートに合わせる)
REDIRECT_URI = "https://bulkmailer-mp.streamlit.app/"

# セッション変数の初期化
if "credentials" not in st.session_state:
    st.session_state.credentials = None

st.title("📧 一斉送信 Gmail アプリ")

# URLのクエリパラメータを取得（認証コードが含まれているかチェック）
query_params = st.query_params
auth_code = query_params.get("code")

# 1️⃣ 認証コードがURLにある場合、自動で処理
if auth_code and st.session_state.credentials is None:
    try:
        flow = google_auth_oauthlib.flow.Flow.from_client_secrets_file(
            CLIENT_SECRETS_FILE, SCOPES
        )
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
    if st.button("Googleにログイン"):
        flow = google_auth_oauthlib.flow.Flow.from_client_secrets_file(
            CLIENT_SECRETS_FILE, SCOPES
        )
        flow.redirect_uri = REDIRECT_URI
        auth_url, _ = flow.authorization_url(prompt="consent")

        # Google認証ページへリダイレクト
        st.markdown(f'<meta http-equiv="refresh" content="0;URL={auth_url}">', unsafe_allow_html=True)

# 3️⃣ 認証完了後、メール送信画面を表示
if st.session_state.credentials:
    service = googleapiclient.discovery.build("gmail", "v1", credentials=st.session_state.credentials)
    st.success("✅ Google 認証が完了しました！")

    # Excel ファイルのアップロード
    uploaded_file = st.file_uploader("📂 送信リスト (Excel)", type=["xlsx"])

    if uploaded_file:
        df = pd.read_excel(uploaded_file)

        # 必須カラムがあるか確認
        if "to_email" not in df.columns:
            st.error("❌ 'to_email' カラムが必要です")
        else:
            st.write("📊 アップロードされたデータ:")
            st.dataframe(df.head())

            # 件名と本文テンプレートを入力
            subject_template = st.text_input("✉ 件名テンプレート", "【お知らせ】{変数1}様へのご案内")
            body_template = st.text_area(
                "📩 本文テンプレート",
                "こんにちは、{変数1}様。\n\nお世話になっております。\n\n以下のご案内をお送りします。\n\n詳細: {変数2}\n\nよろしくお願いいたします。",
            )

            # メール一斉送信
            if st.button("🚀 メール送信"):
                success_count = 0
                error_count = 0
                errors = []

                with st.status("📨 メール送信中...", expanded=True) as status:
                    for index, row in df.iterrows():
                        try:
                            to_email = row["to_email"]

                            # 変数を置き換え
                            subject = subject_template
                            body = body_template
                            for col in df.columns:
                                subject = subject.replace(f"{{{col}}}", str(row[col]))
                                body = body.replace(f"{{{col}}}", str(row[col]))

                            # メール送信処理
                            msg = MIMEText(body)
                            msg["to"] = to_email
                            msg["subject"] = subject
                            raw_msg = base64.urlsafe_b64encode(msg.as_bytes()).decode()

                            service.users().messages().send(
                                userId="me", body={"raw": raw_msg}
                            ).execute()

                            success_count += 1
                            st.write(f"✅ {index+1}. {to_email} へ送信成功")

                            # API制限を避けるために少し待機
                            time.sleep(1)

                        except Exception as e:
                            error_count += 1
                            errors.append(f"❌ {index+1}. {to_email} - エラー: {e}")
                            st.write(errors[-1])

                    # 送信結果を表示
                    status.update(label="📩 送信完了", state="complete", expanded=False)
                    st.success(f"✅ 成功: {success_count}件")
                    st.error(f"❌ 失敗: {error_count}件") if error_count > 0 else None

                    if errors:
                        with st.expander("📋 エラー詳細"):
                            for error in errors:
                                st.write(error)

    # ログアウトボタン
    if st.button("🔒 ログアウト"):
        st.session_state.credentials = None
        st.rerun()
