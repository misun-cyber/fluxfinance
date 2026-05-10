import streamlit as st
import google.generativeai as genai
import win32com.client
import pandas as pd
import pythoncom
import os
from datetime import datetime

# 1. 고정 설정 (API 키 및 화면 레이아웃)
GOOGLE_API_KEY = "AIzaSyAU9a__r3TZbpixPcl-83FthgKastppvlg"
st.set_page_config(page_title="EW FINANCE Adviser Note Helper", layout="wide")

# 2. 섹션/키워드 규칙 기본값
if 'rules_df' not in st.session_state:
    st.session_state['rules_df'] = pd.DataFrame([
        {"Section": "1. Disclosure delivery", "Keywords": "Welcome email, Disclosure, Scope of Service"},
        {"Section": "2. Initial consultation & strategy", "Keywords": "Initial meeting, Strategy discussion, Consultation"},
        {"Section": "3. Fact find drafting (internal)", "Keywords": "Analyse documents, Fact find draft, Internal review"},
        {"Section": "4. Follow-up call (document chase)", "Keywords": "Updating customer, Document chase, Missing info"},
        {"Section": "5. Final document receipt & submission", "Keywords": "SOW confirm, Final documents, Submitted to lender"},
        {"Section": "6. Lender query (adviser declaration)", "Keywords": "Initial response, Adviser declaration, Lender query"},
        {"Section": "7. Lender RFI & strategy adjustment", "Keywords": "RFI, Negotiation note, Strategy adjustment, Further info"},
        {"Section": "8. Initial conditional approval", "Keywords": "Initial approval, Conditional approval, Approval letter"},
        {"Section": "9. Rate review & amount increase", "Keywords": "Rate confirm, Amount increase, Rate review"},
        {"Section": "10. Revised conditional approval", "Keywords": "Revised approval, Amended letter"},
        {"Section": "11. Final approval (post-valuation)", "Keywords": "Final approval, Valuation clear, Unconditional"},
        {"Section": "12. Compliance check", "Keywords": "Vendor finance inquiry, Final condition check, Compliance"}
    ])

tabs = st.tabs(["🔍 데이터 분석 및 수집", "⚙️ 섹션/키워드 규칙 설정"])

with tabs[1]:
    st.header("⚙️ 업무 표준 섹션 관리")
    edited_rules = st.data_editor(st.session_state['rules_df'], num_rows="dynamic", use_container_width=True)
    if st.button("✅ 규칙 저장"):
        st.session_state['rules_df'] = edited_rules
        st.success("규칙이 저장되었습니다.")

with tabs[0]:
    st.title("📧 Adviser Note 자동 완성 헬퍼")
    
    col_e, col_r = st.columns(2)
    with col_e:
        target_emails = st.text_input("고객 이메일", value="John.Bondoc@nz.harveynorman.com")
    with col_r:
        ref_numbers = st.text_input("Reference 번호", value="I-1362129")

    def ask_ai(date, subject, body, rules_context):
        try:
            genai.configure(api_key=GOOGLE_API_KEY)
            model = genai.GenerativeModel('gemini-pro')
            prompt = f"규칙:\n{rules_context}\n\n이메일(제목:{subject}, 본문:{body[:1500]})\n위 내용을 보고 'Section: [번호. 섹션명] | Summary: [요약]'으로 작성해줘."
            return model.generate_content(prompt).text
        except: return "분류 실패"

    if st.button("🚀 검색 및 AI 분석 시작", use_container_width=True):
        pythoncom.CoInitialize()
        rules_str = st.session_state['rules_df'].to_string(index=False)
        
        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            all_data = []
            e_list = [e.strip().lower() for e in target_emails.split(",") if e.strip()]
            r_list = [r.strip() for r in ref_numbers.split(",") if r.strip()]

            with st.spinner("아웃룩 검색 중..."):
                for store in outlook.Stores:
                    root = store.GetRootFolder()
                    for folder in root.Folders:
                        f_name = folder.Name.lower()
                        direction = "Received" if any(k in f_name for k in ["inbox", "받은"]) else "Sent" if any(k in f_name for k in ["sent", "보낸"]) else None
                        if not direction: continue

                        for msg in folder.Items:
                            try:
                                subj, bdy = msg.Subject, getattr(msg, 'Body', '')
                                snd = getattr(msg, 'SenderEmailAddress', '').lower()
                                to_address = getattr(msg, 'To', '').lower()

                                if any(t in snd or t in to_address for t in e_list) or any(r in subj or r in bdy for r in r_list):
                                    analysis = ask_ai(msg.SentOn, subj, bdy, rules_str)
                                    all_data.append({
                                        "Date": msg.SentOn.strftime("%Y-%m-%d %H:%M"),
                                        "Direction": direction,
                                        "AI Analysis": analysis,
                                        "From": getattr(msg, 'SenderName', snd),
                                        "To": to_address,
                                        "Subject": subj,
                                        "Details": bdy # 원문 보존
                                    })
                            except: continue
            
            if all_data:
                # 데이터를 세션에 저장하여 클릭 시 참조 가능하게 함
                st.session_state['search_results'] = pd.DataFrame(all_data).sort_values(by="Date", ascending=False)
                st.success(f"✅ {len(all_data)}건의 기록을 찾았습니다.")
            else:
                st.warning("검색 결과가 없습니다.")
        except Exception as e:
            st.error(f"오류: {e}")

    # --- 결과 출력 및 본문 확인 로직 ---
    if 'search_results' in st.session_state:
        df = st.session_state['search_results']
        
        st.write("💡 아래 행을 클릭하면 하단에 이메일 원문이 나타납니다.")
        # on_select="rerun"과 selection_mode="single-row"를 통해 클릭 감지
        selection = st.dataframe(
            df,
            column_order=("Date", "Direction", "AI Analysis", "From", "To", "Subject"),
            hide_index=True,
            use_container_width=True,
            on_select="rerun",
            selection_mode="single-row"
        )

        # 행이 선택되었을 때 본문(Details) 표시
        if selection.selection.rows:
            selected_idx = selection.selection.rows[0]
            selected_data = df.iloc[selected_idx]
            
            st.markdown("---")
            st.subheader(f"📄 이메일 상세 본문 ({selected_data['Date']})")
            # 텍스트 영역에 본문 출력
            st.text_area(
                label="내용 복사(Adviser Note용)", 
                value=selected_data['Details'], 
                height=450
            )

            # 엑셀 저장 안내
            path = os.path.join(os.path.expanduser("~"), "Desktop", f"EW_Note_{datetime.now().strftime('%m%d_%H%M')}.xlsx")
            df.to_excel(path, index=False)
            st.info(f"📍 전체 데이터가 바탕화면 엑셀 파일로 저장되었습니다.")