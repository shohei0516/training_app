from pathlib import Path
from html import escape

import streamlit as st
from openpyxl import load_workbook


BASE_DIR = Path(__file__).resolve().parent

EXCEL_FILE = BASE_DIR / "pitching_form_training.xlsx"

ASSETS_DIR = BASE_DIR / "assets"
GIF_DIR = ASSETS_DIR / "eval_gif"
OK_NG_DIR = ASSETS_DIR / "ok_ng"

OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)


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


def generate_ppt_style_html_report(scores: dict) -> Path:
    total_score = sum(scores.values())
    problem_checks = [c for c in CHECKS if scores.get(c["id"], 0) >= 1]

    phase_names = [
        "ワインドアップ",
        "アーリーコッキング",
        "レイトコッキング",
        "アクセラレーション",
        "フォロースルー",
    ]

    phase_problem_map = {phase: False for phase in phase_names}

    for check in problem_checks:
        phase_problem_map[check["phase"]] = True

    if total_score == 0:
        main_comment = "全体として大きなフォーム上の問題は確認されませんでした。現在のフォームを維持しながら、継続的なコンディショニングを行いましょう。"
    elif total_score <= 6:
        main_comment = "一部のフェーズで軽度のフォーム課題が確認されました。該当フェーズを中心に、動作の安定性と再現性を高めるトレーニングが推奨されます。"
    else:
        main_comment = "複数のフェーズでフォーム課題が確認されました。局所的な修正だけでなく、下肢・体幹・上肢の連動性を高める段階的なトレーニングが推奨されます。"

    phase_html = ""

    for phase in phase_names:
        star = "★" if phase_problem_map[phase] else ""
        phase_html += f"""
        <div class="phase-box">
            <div class="phase-image">
                <div class="skeleton">投球画像</div>
                <div class="star">{star}</div>
            </div>
            <div class="phase-name">{escape(phase)}</div>
        </div>
        """

    score_rows_html = ""

    for check in CHECKS:
        cid = check["id"]
        score = scores.get(cid, 0)
        row_class = "problem-row" if score >= 1 else ""
        score_rows_html += f"""
        <tr class="{row_class}">
            <td>{escape(cid)}</td>
            <td>{escape(check["title"])}</td>
            <td>{score}</td>
        </tr>
        """

    if problem_checks:
        problem_text = ""
        for check in problem_checks:
            cid = check["id"]
            score = scores.get(cid, 0)
            problem_text += f"""
            <div class="problem-item">
                <b>{escape(cid)} {escape(check["title"])}</b>
                <span>スコア {score}</span>
            </div>
            """
    else:
        problem_text = """
        <div class="problem-item">
            <b>チェック項目なし</b>
            <span>全項目0点</span>
        </div>
        """

    html = f"""
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>投球フォーム評価レポート</title>

<style>
    body {{
        margin: 0;
        padding: 0;
        background: #e5e7eb;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        color: #111827;
    }}

    .page-wrap {{
        width: 100%;
        overflow-x: auto;
        padding: 16px;
        box-sizing: border-box;
    }}

    .report {{
        width: 1180px;
        height: 680px;
        background: #ffffff;
        margin: 0 auto;
        border-radius: 18px;
        box-shadow: 0 12px 30px rgba(0,0,0,0.18);
        position: relative;
        overflow: hidden;
    }}

    .top-bar {{
        height: 82px;
        background: #111827;
        color: white;
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 0 34px;
        box-sizing: border-box;
    }}

    .logo {{
        width: 92px;
        height: 48px;
        border: 2px solid white;
        border-radius: 10px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 800;
        font-size: 22px;
        letter-spacing: 1px;
    }}

    .title-block h1 {{
        margin: 0;
        font-size: 28px;
        letter-spacing: 1px;
    }}

    .title-block p {{
        margin: 5px 0 0;
        font-size: 13px;
        opacity: 0.85;
    }}

    .score-badge {{
        text-align: right;
    }}

    .score-badge .label {{
        font-size: 13px;
        opacity: 0.85;
    }}

    .score-badge .value {{
        font-size: 32px;
        font-weight: 800;
        line-height: 1.1;
    }}

    .phase-area {{
        position: absolute;
        left: 32px;
        top: 104px;
        width: 1116px;
        height: 190px;
        display: grid;
        grid-template-columns: repeat(5, 1fr);
        gap: 14px;
    }}

    .phase-box {{
        border: 2px solid #d1d5db;
        border-radius: 16px;
        overflow: hidden;
        background: #f9fafb;
        position: relative;
    }}

    .phase-image {{
        height: 138px;
        background: linear-gradient(135deg, #f3f4f6, #ffffff);
        display: flex;
        align-items: center;
        justify-content: center;
        position: relative;
    }}

    .skeleton {{
        width: 120px;
        height: 72px;
        border: 2px dashed #9ca3af;
        border-radius: 12px;
        color: #6b7280;
        font-size: 14px;
        display: flex;
        align-items: center;
        justify-content: center;
    }}

    .star {{
        position: absolute;
        top: 8px;
        right: 10px;
        color: #dc2626;
        font-size: 40px;
        font-weight: 900;
        text-shadow: 0 2px 4px rgba(0,0,0,0.25);
    }}

    .phase-name {{
        height: 52px;
        background: #111827;
        color: white;
        font-size: 16px;
        font-weight: 700;
        display: flex;
        align-items: center;
        justify-content: center;
        text-align: center;
    }}

    .bottom-area {{
        position: absolute;
        left: 32px;
        top: 322px;
        width: 1116px;
        height: 326px;
        display: grid;
        grid-template-columns: 340px 1fr 250px;
        gap: 18px;
    }}

    .panel {{
        border: 2px solid #d1d5db;
        border-radius: 16px;
        background: #ffffff;
        overflow: hidden;
    }}

    .panel-title {{
        background: #f3f4f6;
        padding: 10px 14px;
        font-size: 17px;
        font-weight: 800;
        border-bottom: 2px solid #d1d5db;
    }}

    .score-table {{
        width: 100%;
        border-collapse: collapse;
        font-size: 13px;
    }}

    .score-table th {{
        background: #f9fafb;
        padding: 7px;
        border-bottom: 1px solid #d1d5db;
        text-align: left;
    }}

    .score-table td {{
        padding: 6px 7px;
        border-bottom: 1px solid #e5e7eb;
    }}

    .score-table td:nth-child(3) {{
        text-align: center;
        font-weight: 800;
        font-size: 16px;
    }}

    .problem-row {{
        background: #fef2f2;
        color: #991b1b;
        font-weight: 700;
    }}

    .comment-box {{
        padding: 16px;
        font-size: 17px;
        line-height: 1.75;
    }}

    .problem-list {{
        margin-top: 14px;
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 8px;
    }}

    .problem-item {{
        border: 1px solid #fecaca;
        background: #fff1f2;
        border-radius: 10px;
        padding: 9px 10px;
        font-size: 13px;
    }}

    .problem-item b {{
        display: block;
        color: #991b1b;
    }}

    .problem-item span {{
        color: #7f1d1d;
        font-size: 12px;
    }}

    .qr-zone {{
        padding: 16px;
        text-align: center;
    }}

    .qr-dummy {{
        width: 150px;
        height: 150px;
        border: 3px solid #111827;
        margin: 8px auto 12px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 800;
        font-size: 20px;
        background:
            linear-gradient(90deg, #111827 10px, transparent 10px) 0 0 / 30px 30px,
            linear-gradient(#111827 10px, transparent 10px) 0 0 / 30px 30px,
            white;
        color: #111827;
    }}

    .qr-zone p {{
        font-size: 14px;
        line-height: 1.55;
        margin: 0;
    }}

    .footer {{
        position: absolute;
        bottom: 10px;
        left: 34px;
        right: 34px;
        display: flex;
        justify-content: space-between;
        color: #6b7280;
        font-size: 12px;
    }}

    .mobile-note {{
        max-width: 1180px;
        margin: 10px auto 0;
        color: #4b5563;
        font-size: 13px;
    }}

    @media print {{
        body {{
            background: white;
        }}

        .page-wrap {{
            padding: 0;
            overflow: visible;
        }}

        .report {{
            box-shadow: none;
            border-radius: 0;
        }}

        .mobile-note {{
            display: none;
        }}
    }}
</style>
</head>

<body>
<div class="page-wrap">
    <div class="report">

        <div class="top-bar">
            <div class="logo">FGC</div>

            <div class="title-block">
                <h1>投球フォーム評価レポート</h1>
                <p>Pitching Form Report & Training Program</p>
            </div>

            <div class="score-badge">
                <div class="label">TOTAL SCORE</div>
                <div class="value">{total_score} / 22</div>
            </div>
        </div>

        <div class="phase-area">
            {phase_html}
        </div>

        <div class="bottom-area">

            <div class="panel">
                <div class="panel-title">フォーム評価スコア</div>
                <table class="score-table">
                    <thead>
                        <tr>
                            <th>ID</th>
                            <th>評価項目</th>
                            <th>点</th>
                        </tr>
                    </thead>
                    <tbody>
                        {score_rows_html}
                    </tbody>
                </table>
            </div>

            <div class="panel">
                <div class="panel-title">総合コメント</div>
                <div class="comment-box">
                    {escape(main_comment)}
                    <div class="problem-list">
                        {problem_text}
                    </div>
                </div>
            </div>

            <div class="panel">
                <div class="panel-title">自主トレーニング</div>
                <div class="qr-zone">
                    <div class="qr-dummy">QR</div>
                    <p>
                        チェックされた項目に対応する<br>
                        自主トレーニングシートを<br>
                        QRコードから確認します。
                    </p>
                    <p style="margin-top:12px; font-weight:700;">
                        1点・2点ともに<br>
                        該当IDのトレーニングを<br>
                        すべて出力
                    </p>
                </div>
            </div>

        </div>

        <div class="footer">
            <div>Fukui General Clinic / FGC</div>
            <div>Generated by Pitching Training System</div>
        </div>

    </div>

    <div class="mobile-note">
        ※スマホでは横長レポートとして表示されます。必要に応じて横スクロールしてください。
    </div>
</div>
</body>
</html>
"""

    html_path = OUTPUT_DIR / "pitching_report_ppt_style.html"
    html_path.write_text(html, encoding="utf-8")
    return html_path


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
st.subheader("出力設定")

output_type = st.radio(
    "出力形式を選択してください",
    ["PPT風HTML", "PPTレポート", "両方"],
    horizontal=True,
)

if st.button("レポート作成", type="primary"):
    try:
        st.write("Excelにスコアを書き込み中...")
        update_excel_scores(scores)
        st.success("Excel更新完了")

        if output_type in ["PPT風HTML", "両方"]:
            st.write("PPT風HTMLレポートを作成中...")

            html_path = generate_ppt_style_html_report(scores)

            st.success("PPT風HTMLレポート完成！")

            with open(html_path, "rb") as f:
                st.download_button(
                    label="PPT風HTMLをダウンロード",
                    data=f,
                    file_name="pitching_report_ppt_style.html",
                    mime="text/html",
                )

            with st.expander("PPT風HTMLプレビュー"):
                st.components.v1.html(
                    html_path.read_text(encoding="utf-8"),
                    height=760,
                    scrolling=True,
                )

        if output_type in ["PPTレポート", "両方"]:
            st.write("PPTレポート生成中...")

            from report_generator_public import generate_report

            result = generate_report()

            st.success("PPTレポート完成！")

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