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


# -----------------------------
# 共通CSS
# -----------------------------
st.markdown("""
<style>
    .main {
        background-color: #f6f8fb;
    }

    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        padding-left: 2rem;
        padding-right: 2rem;
        max-width: 1400px;
    }

    .hero-card {
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%);
        padding: 28px 32px;
        border-radius: 20px;
        color: white;
        box-shadow: 0 10px 30px rgba(15, 23, 42, 0.18);
        margin-bottom: 24px;
    }

    .hero-title {
        font-size: 30px;
        font-weight: 700;
        margin-bottom: 8px;
        letter-spacing: 0.2px;
    }

    .hero-subtitle {
        font-size: 14px;
        opacity: 0.92;
        line-height: 1.6;
    }

    .section-card {
        background: white;
        padding: 22px 22px 18px 22px;
        border-radius: 18px;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.06);
        border: 1px solid #e5e7eb;
        margin-bottom: 18px;
    }

    .section-title {
        font-size: 18px;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 6px;
    }

    .section-caption {
        font-size: 13px;
        color: #64748b;
        margin-bottom: 10px;
    }

    .metric-card {
        background: white;
        border-radius: 18px;
        padding: 18px 18px;
        border: 1px solid #e5e7eb;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.05);
    }

    .sidebar-note {
        font-size: 13px;
        color: #475569;
        line-height: 1.7;
    }

    .small-note {
        font-size: 12px;
        color: #64748b;
    }

    div[data-testid="stMetric"] {
        background: white;
        border: 1px solid #e5e7eb;
        padding: 14px 16px;
        border-radius: 18px;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.05);
    }

    div[data-testid="stMetricLabel"] {
        color: #64748b;
        font-weight: 600;
    }

    div[data-testid="stMetricValue"] {
        color: #0f172a;
        font-weight: 800;
    }

    .stDownloadButton > button {
        width: 100%;
        border-radius: 12px;
        height: 46px;
        font-weight: 700;
    }

    .stButton > button {
        border-radius: 12px;
        font-weight: 700;
    }

    .stFileUploader {
        border-radius: 14px;
    }
</style>
""", unsafe_allow_html=True)


# -----------------------------
# ログイン
# -----------------------------
def check_login():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False

    if st.session_state.logged_in:
        return True

    left, center, right = st.columns([1, 1.2, 1])

    with center:
        st.markdown("""
        <div class="hero-card" style="margin-top: 40px;">
            <div class="hero-title">📊 TVAL出稿量集計ツール</div>
            <div class="hero-subtitle">
                社内向けの簡易集計アプリです。<br>
                IDとパスワードを入力してログインしてください。
            </div>
        </div>
        """, unsafe_allow_html=True)

        with st.container():
            st.markdown('<div class="section-card">', unsafe_allow_html=True)
            st.markdown('<div class="section-title">ログイン</div>', unsafe_allow_html=True)
            st.markdown('<div class="section-caption">配布されたID / パスワードを入力してください。</div>', unsafe_allow_html=True)

            with st.form("login_form"):
                user_id = st.text_input("ID", placeholder="例：yomiko")
                password = st.text_input("パスワード", type="password", placeholder="パスワードを入力")
                submitted = st.form_submit_button("ログイン")

            if submitted:
                correct_id = st.secrets["APP_ID"]
                correct_password = st.secrets["APP_PASSWORD"]

                if user_id == correct_id and password == correct_password:
                    st.session_state.logged_in = True
                    st.rerun()
                else:
                    st.error("IDまたはパスワードが違います。")

            st.markdown('</div>', unsafe_allow_html=True)

    return False


# -----------------------------
# Excel出力
# -----------------------------
def build_output_excel(summary_df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="集計", index=False)
    return output.getvalue()


# -----------------------------
# 本体
# -----------------------------
def main():
    # ヘッダー
    st.markdown("""
    <div class="hero-card">
        <div class="hero-title">📊 TVAL出稿量集計ツール</div>
        <div class="hero-subtitle">
            input Excel をアップロードすると、F列のエリアを対応エリアへ変換し、
            Y列の数値を対応エリアごとに自動集計します。
        </div>
    </div>
    """, unsafe_allow_html=True)

    # サイドバー
    with st.sidebar:
        st.markdown("## 操作パネル")
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

        st.markdown("### 利用条件")
        st.markdown(
            '<div class="sidebar-note">'
            '・F列：エリア<br>'
            '・Y列：集計対象の数値<br>'
            '・マスタ：エリア集計マスタ.xlsx の master シート'
            '</div>',
            unsafe_allow_html=True
        )

        st.divider()

        if st.button("ログアウト", use_container_width=True):
            st.session_state.logged_in = False
            st.rerun()

    # 初期画面
    if uploaded_file is None:
        left, right = st.columns([1.1, 1])

        with left:
            st.markdown('<div class="section-card">', unsafe_allow_html=True)
            st.markdown('<div class="section-title">はじめに</div>', unsafe_allow_html=True)
            st.markdown('<div class="section-caption">左の操作パネルからファイルをアップロードしてください。</div>', unsafe_allow_html=True)

            st.info("input Excel をアップロードすると、集計結果がこの画面に表示されます。")
            st.markdown("""
            **このツールでできること**
            - エリア名を対応エリアに自動変換
            - 未対応エリアを除外して集計
            - output.xlsx をそのままダウンロード
            """)
            st.markdown('</div>', unsafe_allow_html=True)

        with right:
            st.markdown('<div class="section-card">', unsafe_allow_html=True)
            st.markdown('<div class="section-title">想定フロー</div>', unsafe_allow_html=True)
            st.markdown('<div class="section-caption">毎月の定型業務を短時間で処理できます。</div>', unsafe_allow_html=True)

            st.markdown("""
            ① input Excel を選択  
            ② 自動で対応エリアへ変換  
            ③ 対応エリアごとに集計  
            ④ output.xlsx をダウンロード
            """)
            st.markdown('</div>', unsafe_allow_html=True)
        return

    try:
        # 読み込み
        df = pd.read_excel(uploaded_file)
        master = pd.read_excel(MASTER_PATH, sheet_name=MASTER_SHEET)

        df.columns = df.columns.str.strip()
        master.columns = master.columns.str.strip()

        input_area_col = df.columns[5]   # F列
        value_col = df.columns[24]       # Y列
        master_key_col = "エリア"
        master_value_col = "対応エリア"

        # 表記ゆれ補正
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

        # 対応エリア付与
        area_dict = dict(zip(master[master_key_col], master[master_value_col]))
        df["対応エリア"] = df[input_area_col].map(area_dict)

        unmatched_count = int(df["対応エリア"].isna().sum())

        # 未対応除外
        df = df[df["対応エリア"].notna()].copy()

        # 数値整形
        df[value_col] = pd.to_numeric(df[value_col], errors="coerce").fillna(0)

        # 集計
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

        # KPI
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("対応エリア数", f"{area_count:,}")
        with c2:
            st.metric("合計値", f"{total_value:,.0f}")
        with c3:
            st.metric("未対応件数", f"{unmatched_count:,}")
        with c4:
            st.metric("最大エリア", f"{top_area}", f"{top_value:,.0f}" if top_area != "-" else None)

        # レイアウト
        left, right = st.columns([1.5, 1])

        with left:
            st.markdown('<div class="section-card">', unsafe_allow_html=True)
            st.markdown('<div class="section-title">集計結果</div>', unsafe_allow_html=True)
            st.markdown('<div class="section-caption">対応エリアごとの集計結果です。</div>', unsafe_allow_html=True)

            display_summary = summary.copy()
            display_summary["合計値"] = display_summary["合計値"].map(lambda x: f"{x:,.0f}")
            st.dataframe(display_summary, use_container_width=True, hide_index=True)

            st.markdown('</div>', unsafe_allow_html=True)

        with right:
            st.markdown('<div class="section-card">', unsafe_allow_html=True)
            st.markdown('<div class="section-title">ダウンロード</div>', unsafe_allow_html=True)
            st.markdown('<div class="section-caption">集計結果のみを Excel で出力します。</div>', unsafe_allow_html=True)

            excel_data = build_output_excel(summary)

            st.download_button(
                label="📥 output.xlsx をダウンロード",
                data=excel_data,
                file_name="output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            st.markdown("<br>", unsafe_allow_html=True)

            st.success("集計処理が完了しました。")

            with st.expander("補足情報"):
                st.write("未対応件数は、エリア集計マスタに存在しないエリア名の件数です。")
                st.write("未対応データは集計結果に含めていません。")
                st.write("列位置は F列＝エリア、Y列＝数値 を前提にしています。")

            st.markdown('</div>', unsafe_allow_html=True)

        # グラフ
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
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