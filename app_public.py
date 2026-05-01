from pathlib import Path

import streamlit as st
from openpyxl import load_workbook


BASE_DIR = Path(__file__).resolve().parent

EXCEL_FILE = BASE_DIR / "pitching_form_training.xlsx"

ASSETS_DIR = BASE_DIR / "assets"
GIF_DIR = ASSETS_DIR / "eval_gif"
OK_NG_DIR = ASSETS_DIR / "ok_ng"


CHECKS = [
    {
        "id": "C01",
        "phase": "ワインドアップ",
        "title": "軸足安定",
        "criteria": {
            0: "・骨盤水平・体幹ブレなし\n・支持脚で安定して立位保持できる",
            1: "・骨盤がわずかに傾く\n・または体幹の軽度左右ブレあり",
            2: "・骨盤の横シフトが大きい\n・または明らかな体幹動揺がある",
        },
    },
    {
        "id": "C02",
        "phase": "アーリーコッキング",
        "title": "軸足膝前方移動",
        "criteria": {
            0: "・軸足膝がつま先と同程度\n・またはやや後方に位置している",
            1: "・軸足膝がつま先よりわずかに前方へ出る",
            2: "・軸足膝がつま先より明らかに前方へ突出する\n・前方移動優位になっている",
        },
    },
    {
        "id": "C03",
        "phase": "アーリーコッキング",
        "title": "軸足アライメント",
        "criteria": {
            0: "・股関節・膝・足部のラインが一直線\n・Knee inがみられない",
            1: "・膝がわずかに内側へ入る\n・ただしラインの崩れは軽度",
            2: "・膝が明らかに内側へ崩れる\n・股関節・膝・足部の一直線が保てない",
        },
    },
    {
        "id": "C04",
        "phase": "レイトコッキング",
        "title": "ステップ足膝屈曲",
        "criteria": {
            0: "・接地時の膝屈曲が適度\n・約40〜60°程度",
            1: "・接地時の膝屈曲がやや大きい\n・沈み込みがやや目立つ",
            2: "・接地時に膝が深く曲がりすぎる\n・約100°以上の過度屈曲",
        },
    },
    {
        "id": "C05",
        "phase": "レイトコッキング",
        "title": "ステップ足膝外側動揺",
        "criteria": {
            0: "・ステップ足が進行方向に対して直線的に接地している",
            1: "・ステップ足がわずかに外側へ開く\n・または回転時に軽度の外逃げがある",
            2: "・ステップ足が明らかに外側へ逃げる\n・回転運動中に支持が不安定になる",
        },
    },
    {
        "id": "C06",
        "phase": "レイトコッキング",
        "title": "テイクバック位置",
        "criteria": {
            0: "・テイクバックでボールが頭部後方に位置する\n・正面から見て頭の後ろに隠れている",
            1: "・ボールがやや横に見える\n・頭部後方への引き込みがやや不足する",
            2: "・ボールが横または前方に明らかに見える",
        },
    },
    {
        "id": "C07",
        "phase": "アクセラレーション",
        "title": "Cカーブ形成",
        "criteria": {
            0: "・肩最大外旋位で体全体がCカーブを描く",
            1: "・Cカーブがやや不十分",
            2: "・Cカーブが明らかに形成できていない\n・体のラインが一直線または前屈傾向",
        },
    },
    {
        "id": "C08",
        "phase": "アクセラレーション",
        "title": "ステップ足膝伸展",
        "criteria": {
            0: "・ステップ足接地後からリリースまで膝が伸展する",
            1: "・膝伸展がやや不十分\n・接地後の膝伸び上がりが弱い",
            2: "・膝屈曲位を維持したままリリースする",
        },
    },
    {
        "id": "C09",
        "phase": "アクセラレーション",
        "title": "肘抜け",
        "criteria": {
            0: "・頭部の横でリリースできる\n・肩・肘・手のラインが一直線に近い",
            1: "・肘がやや先行する\n・リリース位置がわずかに前方へずれる",
            2: "・肘が明らかに先行する\n・頭部より前方でリリースし肘抜けがみられる",
        },
    },
    {
        "id": "C10",
        "phase": "フォロースルー",
        "title": "上体回旋角度",
        "criteria": {
            0: "・両肩を結んだ線がセンターとキャッチャー方向を向く",
            1: "・回旋はあるがやや不十分\n・肩のラインが目標方向まで回りきらない",
            2: "・回旋が途中で止まる\n・肩のラインが明らかに目標方向を向かない",
        },
    },
    {
        "id": "C11",
        "phase": "フォロースルー",
        "title": "軸足けり上げ",
        "criteria": {
            0: "・軸足が後方へ振り上がる\n・足部が腰のラインよりも高い",
            1: "・軸足が振り上がるが腰のラインよりも低い",
            2: "・軸足が残る、もしくは引きずる\n・蹴り上げがほとんどみられない",
        },
    },
]


def show_gif(check_id: str) -> None:
    gif_path = GIF_DIR / f"{check_id}.gif"

    if gif_path.exists():
        st.image(str(gif_path), use_container_width=True)
    else:
        st.info(f"{check_id} GIFなし：{gif_path.name}")


def show_ok_ng(check_id: str) -> None:
    ok_path = OK_NG_DIR / f"{check_id}ok.png"
    ng_path = OK_NG_DIR / f"{check_id}ng.png"

    col_ok, col_ng = st.columns(2)

    with col_ok:
        if ok_path.exists():
            st.image(str(ok_path), use_container_width=True)
        else:
            st.warning(f"{check_id} OK画像なし")

    with col_ng:
        if ng_path.exists():
            st.image(str(ng_path), use_container_width=True)
        else:
            st.warning(f"{check_id} NG画像なし")


def update_excel_scores(scores: dict) -> None:
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(f"Excelファイルが見つかりません: {EXCEL_FILE}")

    wb = load_workbook(EXCEL_FILE)
    ws = wb["評価入力"]

    headers = [cell.value for cell in ws[1]]

    check_id_col = headers.index("チェックID") + 1
    score_col = headers.index("スコア") + 1

    for row in range(2, ws.max_row + 1):
        check_id = ws.cell(row=row, column=check_id_col).value

        if check_id in scores:
            ws.cell(row=row, column=score_col).value = scores[check_id]

    wb.save(EXCEL_FILE)


st.set_page_config(
    page_title="投球フォーム評価",
    page_icon="⚾",
    layout="wide",
)

st.title("投球フォーム評価シート")
st.caption("GIF・OK/NG画像・評価基準を確認しながら、C01〜C11を0・1・2で評価します。")

st.markdown("---")

scores = {}

for check in CHECKS:
    cid = check["id"]

    with st.container(border=True):
        st.markdown(f"## {cid}　{check['title']}")
        st.caption(check["phase"])

        col_gif, col_img, col_score = st.columns([1.15, 1.55, 1.25])

        with col_gif:
            show_gif(cid)

        with col_img:
            show_ok_ng(cid)

        with col_score:
            st.markdown("### 評価基準")

            for score_value, comment in check["criteria"].items():
                st.markdown(f"**{score_value}点**")
                st.markdown(comment.replace("\n", "  \n"))

            selected_score = st.radio(
                "スコアを選択",
                options=[0, 1, 2],
                horizontal=True,
                key=f"score_{cid}",
            )

            scores[cid] = selected_score


st.markdown("---")
st.subheader("評価サマリー")

total_score = sum(scores.values())
problem_ids = [cid for cid, score in scores.items() if score >= 1]

col1, col2, col3 = st.columns(3)

with col1:
    st.metric("総合スコア", total_score)

with col2:
    st.metric("チェック項目数", len(problem_ids))

with col3:
    st.metric("最大スコア", 22)

if problem_ids:
    st.warning("チェックあり：" + " / ".join(problem_ids))
else:
    st.success("全項目0点です。大きな問題はありません。")


st.markdown("---")

if st.button("レポート作成", type="primary"):
    try:
        st.write("Excelにスコアを書き込み中...")
        update_excel_scores(scores)
        st.success("Excel更新完了")

        st.write("レポート生成中...")

        from report_generator_public import generate_report

        result = generate_report()

        st.success("レポート完成！")

        pptx_path = result.get("pptx")
        pdf_path = result.get("pdf")

        if pptx_path and Path(pptx_path).exists():
            with open(pptx_path, "rb") as f:
                st.download_button(
                    label="PPTXをダウンロード",
                    data=f,
                    file_name="pitching_report_auto.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
        else:
            st.warning("PPTXが見つかりません。")

        if pdf_path and Path(pdf_path).exists():
            with open(pdf_path, "rb") as f:
                st.download_button(
                    label="PDFをダウンロード",
                    data=f,
                    file_name="pitching_report_auto.pdf",
                    mime="application/pdf",
                )
        else:
            st.info("PDFはクラウドでは作成されないため、PPTXをダウンロードしてください。")

        if result.get("url"):
            st.caption(f"QRリンク: {result['url']}")

    except Exception as e:
        st.error("レポート生成でエラーが発生しました")
        st.exception(e)