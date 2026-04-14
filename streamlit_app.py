import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(
    page_title="TVAL出稿量集計ツール",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

MASTER_PATH = "エリア集計マスタ.xlsx"
MASTER_SHEET = "master"


def inject_global_css():
    st.markdown("""
    <style>
        html, body, [class*="css"] {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
        }

        .stApp {
            background:
                radial-gradient(circle at top left, rgba(30, 58, 138, 0.18), transparent 26%),
                radial-gradient(circle at top right, rgba(15, 23, 42, 0.12), transparent 22%),
                linear-gradient(180deg, #eef2ff 0%, #f8fafc 18%, #f8fafc 100%);
        }

        .block-container {
            max-width: 1380px;
            padding-top: 1.5rem;
            padding-bottom: 2rem;
            padding-left: 2rem;
            padding-right: 2rem;
        }

        [data-testid="stHeader"] {
            background: rgba(0, 0, 0, 0);
        }

        section[data-testid="stSidebar"] {
            background: rgba(255, 255, 255, 0.78);
            backdrop-filter: blur(18px);
            border-right: 1px solid rgba(226, 232, 240, 0.95);
        }

        .hero-wrap {
            margin-bottom: 22px;
        }

        .hero-card {
            position: relative;
            overflow: hidden;
            background: linear-gradient(135deg, #0b1220 0%, #102a6b 55%, #1d4ed8 100%);
            border-radius: 26px;
            padding: 30px 34px;
            box-shadow: 0 22px 55px rgba(15, 23, 42, 0.22);
            color: #ffffff;
            border: 1px solid rgba(255,255,255,0.08);
        }

        .hero-card::before {
            content: "";
            position: absolute;
            top: -60px;
            right: -20px;
            width: 220px;
            height: 220px;
            background: radial-gradient(circle, rgba(255,255,255,0.16) 0%, rgba(255,255,255,0.02) 60%, transparent 75%);
            border-radius: 999px;
        }

        .hero-card::after {
            content: "";
            position: absolute;
            bottom: -100px;
            left: -40px;
            width: 260px;
            height: 260px;
            background: radial-gradient(circle, rgba(255,255,255,0.10) 0%, rgba(255,255,255,0.02) 60%, transparent 75%);
            border-radius: 999px;
        }

        .hero-eyebrow {
            font-size: 12px;
            letter-spacing: 0.16em;
            text-transform: uppercase;
            opacity: 0.8;
            margin-bottom: 10px;
            font-weight: 700;
        }

        .hero-title {
            font-size: 32px;
            font-weight: 800;
            line-height: 1.2;
            margin-bottom: 10px;
            letter-spacing: -0.02em;
        }

        .hero-subtitle {
            font-size: 14px;
            line-height: 1.8;
            opacity: 0.94;
            max-width: 780px;
        }

        .glass-card {
            background: rgba(255,255,255,0.82);
            backdrop-filter: blur(18px);
            border: 1px solid rgba(226, 232, 240, 0.9);
            border-radius: 22px;
            padding: 22px 22px 18px 22px;
            box-shadow: 0 16px 40px rgba(15, 23, 42, 0.08);
            margin-bottom: 18px;
        }

        .glass-card-tight {
            background: rgba(255,255,255,0.84);
            backdrop-filter: blur(18px);
            border: 1px solid rgba(226, 232, 240, 0.92);
            border-radius: 20px;
            padding: 18px 18px 16px 18px;
            box-shadow: 0 14px 34px rgba(15, 23, 42, 0.07);
            margin-bottom: 16px;
        }

        .section-title {
            font-size: 18px;
            font-weight: 800;
            color: #0f172a;
            margin-bottom: 6px;
            letter-spacing: -0.01em;
        }

        .section-caption {
            font-size: 13px;
            color: #64748b;
            margin-bottom: 12px;
            line-height: 1.7;
        }

        .sidebar-title {
            font-size: 18px;
            font-weight: 800;
            color: #0f172a;
            margin-bottom: 6px;
        }

        .sidebar-note {
            font-size: 13px;
            color: #475569;
            line-height: 1.8;
        }

        .login-shell {
            min-height: 76vh;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .login-card {
            width: 100%;
            max-width: 430px;
            background: rgba(255,255,255,0.82);
            backdrop-filter: blur(24px);
            border: 1px solid rgba(226,232,240,0.95);
            border-radius: 26px;
            padding: 32px 30px 26px 30px;
            box-shadow: 0 24px 60px rgba(15,23,42,0.14);
        }

        .login-brand {
            font-size: 12px;
            letter-spacing: 0.16em;
            text-transform: uppercase;
            color: #64748b;
            font-weight: 800;
            margin-bottom: 8px;
        }

        .login-title {
            font-size: 28px;
            line-height: 1.2;
            font-weight: 800;
            color: #0f172a;
            margin-bottom: 8px;
            letter-spacing: -0.02em;
        }

        .login-sub {
            font-size: 13px;
            color: #64748b;
            line-height: 1.8;
            margin-bottom: 18px;
        }

        div[data-testid="stMetric"] {
            background: rgba(255,255,255,0.84);
            backdrop-filter: blur(14px);
            border: 1px solid rgba(226,232,240,0.95);
            border-radius: 20px;
            padding: 14px 16px;
            box-shadow: 0 14px 30px rgba(15,23,42,0.07);
        }

        div[data-testid="stMetricLabel"] {
            color: #64748b;
            font-weight: 700;
        }

        div[data-testid="stMetricValue"] {
            color: #0f172a;
            font-weight: 800;
            letter-spacing: -0.02em;
        }

        div[data-testid="stFileUploader"] {
            background: rgba(255,255,255,0.55);
            border-radius: 16px;
            border: 1px dashed #cbd5e1;
            padding: 8px;
        }

        div[data-baseweb="input"] {
            background: rgba(248, 250, 252, 0.9);
            border-radius: 14px;
        }

        div[data-baseweb="input"] > div {
            border: 1px solid #dbe4f0 !important;
            border-radius: 14px !important;
            background: rgba(248,250,252,0.96) !important;
            min-height: 46px;
        }

        .stTextInput label {
            font-weight: 700 !important;
            color: #334155 !important;
        }

        .stButton > button,
        .stDownloadButton > button,
        div[data-testid="stFormSubmitButton"] > button {
            width: 100%;
            min-height: 46px;
            border-radius: 14px !important;
            border: 1px solid rgba(30, 64, 175, 0.18) !important;
            background: linear-gradient(135deg, #0f172a 0%, #1d4ed8 100%) !important;
            color: white !important;
            font-weight: 800 !important;
            box-shadow: 0 12px 26px rgba(29, 78, 216, 0.24);
        }

        .stButton > button:hover,
        .stDownloadButton > button:hover,
        div[data-testid="stFormSubmitButton"] > button:hover {
            filter: brightness(1.03);
            transform: translateY(-1px);
        }

        .plain-button .stButton > button {
            background: #ffffff !important;
            color: #0f172a !important;
            border: 1px solid #dbe4f0 !important;
            box-shadow: none !important;
        }

        div[data-testid="stDataFrame"] {
            border-radius: 18px;
            overflow: hidden;
            border: 1px solid rgba(226,232,240,0.95);
        }

        .tiny-note {
            font-size: 12px;
            color: #64748b;
            line-height: 1.7;
        }

        hr {
            border-color: #e2e8f0 !important;
        }
    </style>
    """, unsafe_allow_html=True)


def build_output_excel(summary_df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="集計", index=False)
    return output.getvalue()


def check_login():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False

    if st.session_state.logged_in:
        return True

    inject_global_css()

    left, center, right = st.columns([1.2, 1, 1.2])
    with center:
        st.markdown('<div class="login-shell">', unsafe_allow_html=True)
        st.markdown('<div class="login-card">', unsafe_allow_html=True)
        st.markdown('<div class="login-brand">YOMIKO INTERNAL TOOL</div>', unsafe_allow_html=True)
        st.markdown('<div class="login-title">TVAL Dashboard</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="login-sub">'
            '社内向けの出稿量集計ツールです。<br>'
            '配布されたID / パスワードを入力してログインしてください。'
            '</div>',
            unsafe_allow_html=True
        )

        with st.form("login_form", clear_on_submit=False):
            user_id = st.text_input("ID", placeholder="IDを入力")
            password = st.text_input("パスワード", type="password", placeholder="パスワードを入力")
            submitted = st.form_submit_button("ログイン")

        if submitted:
            if user_id == st.secrets["APP_ID"] and password == st.secrets["APP_PASSWORD"]:
                st.session_state.logged_in = True
                st.rerun()
            else:
                st.error("IDまたはパスワードが違います。")

        st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    return False


def main():
    inject_global_css()

    st.markdown('<div class="hero-wrap">', unsafe_allow_html=True)
    st.markdown("""
    <div class="hero-card">
        <div class="hero-eyebrow">YOMIKO INTERNAL TOOL</div>
        <div class="hero-title">📊 TVAL出稿量集計ダッシュボード</div>
        <div class="hero-subtitle">
            input Excel をアップロードすると、F列のエリアを対応エリアへ変換し、
            Y列の数値を対応エリアごとに自動集計します。
            毎月の定型業務を、より速く・迷いなく処理するための社内用ツールです。
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    with st.sidebar:
        st.markdown('<div class="sidebar-title">操作パネル</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="sidebar-note">'
            '1. input Excel をアップロード<br>'
            '2. 集計結果を確認<br>'
            '3. output.xlsx をダウンロード'
            '</div>',
            unsafe_allow_html=True
        )

        st.divider()

        uploaded_file = st.file_uploader(
            "input Excel",
            type=["xlsx"],
            help="集計対象のExcelファイルを選択してください。"
        )

        st.divider()

        st.markdown("### 条件")
        st.markdown(
            '<div class="sidebar-note">'
            '・F列：エリア<br>'
            '・Y列：集計対象の数値<br>'
            '・マスタ：エリア集計マスタ.xlsx / master シート'
            '</div>',
            unsafe_allow_html=True
        )

        st.divider()

        st.markdown('<div class="plain-button">', unsafe_allow_html=True)
        if st.button("ログアウト", use_container_width=True):
            st.session_state.logged_in = False
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_file is None:
        col1, col2 = st.columns([1.2, 1])

        with col1:
            st.markdown('<div class="glass-card">', unsafe_allow_html=True)
            st.markdown('<div class="section-title">アップロード待機中</div>', unsafe_allow_html=True)
            st.markdown('<div class="section-caption">左側の操作パネルから input Excel を選択してください。</div>', unsafe_allow_html=True)
            st.info("ファイルをアップロードすると、この画面に集計結果が表示されます。")
            st.markdown("""
            **このツールでできること**
            - エリア名を対応エリアへ自動変換
            - 未対応エリアを除外して集計
            - output.xlsx をその場でダウンロード
            """)
            st.markdown('</div>', unsafe_allow_html=True)

        with col2:
            st.markdown('<div class="glass-card">', unsafe_allow_html=True)
            st.markdown('<div class="section-title">処理フロー</div>', unsafe_allow_html=True)
            st.markdown('<div class="section-caption">毎月の定型業務を短時間で処理できます。</div>', unsafe_allow_html=True)
            st.markdown("""
            ① input Excel を選択  
            ② 対応エリアへ変換  
            ③ 対応エリアごとに集計  
            ④ output.xlsx をダウンロード
            """)
            st.markdown('</div>', unsafe_allow_html=True)
        return

    try:
        df = pd.read_excel(uploaded_file)
        master = pd.read_excel(MASTER_PATH, sheet_name=MASTER_SHEET)

        df.columns = df.columns.str.strip()
        master.columns = master.columns.str.strip()

        input_area_col = df.columns[5]   # F列
        value_col = df.columns[24]       # Y列
        master_key_col = "エリア"
        master_value_col = "対応エリア"

        df[input_area_col] = (
            df[input_area_col]
            .astype(str)
            .str.strip()
            .str.replace(",", "、", regex=False)
        )

        master[master_key_col] = (
            master[master_key_col]
            .astype(str)
            .str.strip()
            .str.replace(",", "、", regex=False)
        )

        area_dict = dict(zip(master[master_key_col], master[master_value_col]))
        df["対応エリア"] = df[input_area_col].map(area_dict)

        unmatched_count = int(df["対応エリア"].isna().sum())

        df = df[df["対応エリア"].notna()].copy()
        df[value_col] = pd.to_numeric(df[value_col], errors="coerce").fillna(0)

        summary = (
            df.groupby("対応エリア")[value_col]
            .sum()
            .reset_index()
            .rename(columns={value_col: "合計値"})
            .sort_values("合計値", ascending=False)
        )

        total_value = float(summary["合計値"].sum()) if not summary.empty else 0
        area_count = int(summary["対応エリア"].nunique()) if not summary.empty else 0
        top_area = summary.iloc[0]["対応エリア"] if not summary.empty else "-"
        top_value = float(summary.iloc[0]["合計値"]) if not summary.empty else 0

        m1, m2, m3, m4 = st.columns(4)
        with m1:
            st.metric("対応エリア数", f"{area_count:,}")
        with m2:
            st.metric("合計値", f"{total_value:,.0f}")
        with m3:
            st.metric("未対応件数", f"{unmatched_count:,}")
        with m4:
            st.metric("最大エリア", top_area, f"{top_value:,.0f}" if top_area != "-" else None)

        left, right = st.columns([1.45, 1])

        with left:
            st.markdown('<div class="glass-card">', unsafe_allow_html=True)
            st.markdown('<div class="section-title">集計結果</div>', unsafe_allow_html=True)
            st.markdown('<div class="section-caption">対応エリアごとの集計結果です。</div>', unsafe_allow_html=True)

            display_summary = summary.copy()
            display_summary["合計値"] = display_summary["合計値"].map(lambda x: f"{x:,.0f}")
            st.dataframe(display_summary, use_container_width=True, hide_index=True)

            st.markdown('</div>', unsafe_allow_html=True)

        with right:
            st.markdown('<div class="glass-card">', unsafe_allow_html=True)
            st.markdown('<div class="section-title">ダウンロード</div>', unsafe_allow_html=True)
            st.markdown('<div class="section-caption">集計結果のみを Excel 形式で出力します。</div>', unsafe_allow_html=True)

            excel_data = build_output_excel(summary)

            st.download_button(
                label="📥 output.xlsx をダウンロード",
                data=excel_data,
                file_name="output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)
            st.success("集計処理が完了しました。")

            with st.expander("補足情報"):
                st.write("未対応件数は、エリア集計マスタに存在しないエリア名の件数です。")
                st.write("未対応データは集計結果に含めていません。")
                st.write("列位置は F列＝エリア、Y列＝数値 を前提にしています。")

            st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">上位10エリア</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-caption">合計値が大きい順に表示しています。</div>', unsafe_allow_html=True)

        top10 = summary.head(10).set_index("対応エリア")
        if not top10.empty:
            st.bar_chart(top10["合計値"])
        else:
            st.info("表示できるデータがありません。")

        st.markdown('</div>', unsafe_allow_html=True)

    except Exception as e:
        st.error(f"エラーが発生しました: {e}")


if check_login():
    main()
