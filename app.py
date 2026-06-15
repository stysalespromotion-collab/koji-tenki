import streamlit as st
import io
from transfer import transfer_dekidaka, transfer_kanryo

st.set_page_config(
    page_title="工事書式転記システム",
    page_icon="📋",
    layout="centered"
)

st.title("📋 工事書式転記システム")
st.caption("見積決定申請書の内容を各書式へ自動転記します")

st.divider()

# ── ① 見積決定ファイルのアップロード
st.subheader("① 見積決定申請書をアップロード")
mitsumori_file = st.file_uploader(
    "見積決定.xlsx を選択してください",
    type=["xlsx"],
    key="mitsumori"
)

st.divider()

# ── ② 書式選択
st.subheader("② 書式を選択")
shoshiki = st.radio(
    "転記先の書式",
    ["出来高（途中回）", "工事完了（最終明細）"],
    horizontal=True
)

st.divider()

# ── ③ 回数入力 & 追加ファイル
if shoshiki == "出来高（途中回）":
    st.subheader("③ 回数を入力")
    kaime = st.number_input("今回は何回目の出来高ですか？", min_value=1, step=1, value=1)

    st.divider()
    st.subheader("④ 転記実行")

    if mitsumori_file is None:
        st.warning("見積決定申請書をアップロードしてください")
    else:
        if st.button("転記してExcelをダウンロード", type="primary", use_container_width=True):
            with st.spinner("転記中..."):
                try:
                    result = transfer_dekidaka(
                        mitsumori_file.read(),
                        kaime=int(kaime)
                    )
                    st.success(f"✅ 転記完了！（第{kaime}回 出来高）")
                    st.download_button(
                        label="📥 出来高.xlsx をダウンロード",
                        data=result,
                        file_name=f"出来高_{kaime}回目.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"エラーが発生しました: {e}")

else:  # 工事完了（最終明細）
    st.subheader("③ 回数と過去の出来高ファイルをアップロード")
    kaime = st.number_input("今回は何回目の完了ですか？", min_value=1, step=1, value=1)

    st.markdown("**過去の出来高ファイルをすべてアップロード**（下請支払の累計計算に使用）")
    dekidaka_files = st.file_uploader(
        "出来高_1回目.xlsx、出来高_2回目.xlsx… を選択（複数可）",
        type=["xlsx"],
        accept_multiple_files=True,
        key="dekidaka"
    )

    if dekidaka_files:
        st.info(f"📁 {len(dekidaka_files)}件の出来高ファイルを読み込みました")
        for f in dekidaka_files:
            st.caption(f"　・{f.name}")

    st.divider()
    st.subheader("④ 転記実行")

    ready = mitsumori_file is not None
    if not ready:
        st.warning("見積決定申請書をアップロードしてください")
    elif len(dekidaka_files) == 0:
        st.warning("過去の出来高ファイルをアップロードしてください（下請支払の累計計算に必要です）")
        if st.button("出来高ファイルなしで転記する", use_container_width=True):
            with st.spinner("転記中..."):
                try:
                    result = transfer_kanryo(
                        mitsumori_file.read(),
                        dekidaka_bytes_list=[],
                        kaime=int(kaime)
                    )
                    st.success(f"✅ 転記完了！（第{kaime}回 工事完了）")
                    st.download_button(
                        label="📥 工事完了.xlsx をダウンロード",
                        data=result,
                        file_name=f"工事完了_{kaime}回目.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"エラーが発生しました: {e}")
    else:
        if st.button("転記してExcelをダウンロード", type="primary", use_container_width=True):
            with st.spinner("転記中..."):
                try:
                    dekidaka_bytes_list = [f.read() for f in dekidaka_files]
                    result = transfer_kanryo(
                        mitsumori_file.read(),
                        dekidaka_bytes_list=dekidaka_bytes_list,
                        kaime=int(kaime)
                    )
                    st.success(f"✅ 転記完了！（第{kaime}回 工事完了・下請支払累計{len(dekidaka_files)}回分）")
                    st.download_button(
                        label="📥 工事完了.xlsx をダウンロード",
                        data=result,
                        file_name=f"工事完了_{kaime}回目.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"エラーが発生しました: {e}")

st.divider()
st.caption("ダウンロードしたExcelファイルはすべてのセルを手修正できます")
