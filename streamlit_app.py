"""Google スプレッドシート AI操作アプリ(Streamlit版)。

自然言語の指示から Claude が操作内容(現在のシートを更新/新しいシートに書き込む/
シートを削除する/シートを作成する)を判断し、確認の上で反映する。

一度スプレッドシートを読み込んだ後は、指示を変えるたびに
URLを入力し直す必要はない(セッション内でシートを保持する)。
同じURLでの操作はセッション内で履歴として保持され、
「元に戻して」等の指示やボタン操作で、指定した時点まで取り消せる。
取り消した操作は「やり直す」ボタンで再度適用できる。

既存データの扱い:
- 指示に書き込み位置(セル)の指定がない場合は、シート全体を再構成する指示として
  扱い、シート全体をクリアしてから書き込む(デフォルト、従来通り)。
- 「H1セルに書いて」のように書き込み位置が明示された指示の場合は、
  その位置にのみ書き込み、それ以外の既存データはクリアしない。
"""

import os
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components
from dotenv import load_dotenv

import sheets_core

# 起動時のカレントディレクトリに関係なく、このファイルと同じ場所の .env を読み込む
# (別ディレクトリから `streamlit run` した場合に ANTHROPIC_API_KEY が読めず
#  401 invalid x-api-key になるのを防ぐ)。
load_dotenv(Path(__file__).with_name(".env"))
load_dotenv()  # カレントディレクトリ側の .env も一応探す(既存の値は上書きしない)

st.set_page_config(page_title="Sheets AI Assistant", page_icon="\U0001f4ca", layout="centered")

# Streamlit標準の「Press Enter to submit form」等のフォーム注記を非表示にする
# (enter_to_submit=False を指定していても環境によって表示が残ることがあるための保険)。
st.markdown(
    '<style>[data-testid="InputInstructions"] { display: none !important; }</style>',
    unsafe_allow_html=True,
)


def _bind_ctrl_enter(button_text: str):
    """指定したテキストのボタンを Ctrl(またはCmd)+Enter でクリックできるようにする最小限のJS。"""
    components.html(
        f"""
        <script>
        const doc = window.parent.document;
        doc.addEventListener('keydown', function(e) {{
            if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {{
                for (const btn of doc.querySelectorAll('button')) {{
                    if (btn.innerText.trim() === {button_text!r}) {{
                        btn.click();
                        break;
                    }}
                }}
            }}
        }});
        </script>
        """,
        height=0,
    )


@st.cache_resource
def get_client():
    return sheets_core.get_client()


def _show_format_failures():
    """複数書式の一括適用で一部だけ失敗した場合に、どの範囲が失敗したかを表示する。"""
    failures = st.session_state.pop("format_failures", None)
    if not failures:
        return
    with st.expander(f"⚠️ {len(failures)}件の書式適用に失敗しました(内訳を表示)", expanded=True):
        for item in failures:
            st.write(f"- `{item['range']}`: {item['error']}")


def _show_api_key_warning():
    """APIキーが未設定・不正な形式の場合に、画面上部に警告を出す。

    ここで st.stop() はしない(シートの読み込みや履歴確認など、
    Claudeを使わない操作は引き続き行えるようにするため)。
    """
    message = sheets_core.api_key_status_message()
    if message:
        st.warning(message)


def _flash(kind: str, message: str):
    st.session_state["flash"] = (kind, message)


def _show_flash():
    flash = st.session_state.pop("flash", None)
    if not flash:
        return
    kind, message = flash
    getattr(st, kind)(message)


def require_login() -> bool:
    if st.session_state.get("authenticated"):
        return True

    st.title("Sheets AI Assistant")
    st.caption("自然言語でGoogleスプレッドシートを操作するデモアプリ")

    with st.form("login_form"):
        password = st.text_input("パスワード", type="password")
        submitted = st.form_submit_button("ログイン", type="primary")

    if submitted:
        expected = os.environ.get("WEB_APP_PASSWORD")
        if not expected:
            st.error("環境変数 WEB_APP_PASSWORD が設定されていません。")
        elif password == expected:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("パスワードが違います。")
    return False


def _reset_sheet_state():
    for key in (
        "sheet_loaded", "spreadsheet_url", "worksheet_input", "df",
        "spreadsheet_title", "worksheet_title", "pending_action", "history", "redo_stack",
        "header_offset",
    ):
        st.session_state.pop(key, None)


def load_sheet_section():
    """スプレッドシートの読み込みUI。読み込み済みならスキップして再利用する。"""
    loaded = st.session_state.get("sheet_loaded")

    if loaded:
        st.success(
            f"読み込み済み: {st.session_state['spreadsheet_title']} / "
            f"{st.session_state['worksheet_title']} "
            f"({st.session_state['df'].shape[0]}行 x {st.session_state['df'].shape[1]}列)"
        )
        col1, col2 = st.columns(2)
        with col1:
            if st.button("データを再読み込み"):
                _load_sheet(st.session_state["spreadsheet_url"], st.session_state["worksheet_input"], keep_history=True)
        with col2:
            if st.button("別のシートを読み込む"):
                _reset_sheet_state()
                st.rerun()
        return

    st.subheader("1. スプレッドシートを読み込む")
    with st.form("load_sheet_form"):
        spreadsheet_url = st.text_input(
            "スプレッドシートのURL", placeholder="https://docs.google.com/spreadsheets/d/..."
        )
        worksheet_input = st.text_input("シート名(省略可、未入力なら1枚目)")
        service_email = sheets_core.get_service_account_email()
        if service_email:
            st.caption(f"事前にサービスアカウントのメールアドレス「{service_email}」へ「編集者」共有をしてください。")
        else:
            st.caption("事前にサービスアカウントのメールアドレスへ「編集者」共有をしてください。")
        submitted = st.form_submit_button("読み込む", type="primary")

    if submitted:
        if not spreadsheet_url.strip():
            st.error("スプレッドシートのURLを入力してください。")
        else:
            _load_sheet(spreadsheet_url.strip(), worksheet_input.strip() or None, keep_history=False)


def _load_sheet(spreadsheet_url: str, worksheet_name: str | None, keep_history: bool):
    try:
        with st.spinner("読み込み中..."):
            client = get_client()
            sh = sheets_core.open_spreadsheet(client, spreadsheet_url)
            df, ws = sheets_core.load_dataframe(sh, worksheet_name)
    except Exception as exc:  # noqa: BLE001
        st.error(f"読み込みに失敗しました: {exc}")
        return

    st.session_state["sheet_loaded"] = True
    st.session_state["spreadsheet_url"] = spreadsheet_url
    st.session_state["worksheet_input"] = worksheet_name
    st.session_state["df"] = df
    st.session_state["spreadsheet_title"] = sh.title
    st.session_state["worksheet_title"] = ws.title
    st.session_state.pop("pending_action", None)
    if not keep_history:
        st.session_state["history"] = []
        st.session_state["redo_stack"] = []
        st.session_state["header_offset"] = 0
    st.rerun()


def _get_sh():
    client = get_client()
    return sheets_core.open_spreadsheet(client, st.session_state["spreadsheet_url"])


def _refresh_current_df(sh):
    ws = sh.worksheet(st.session_state["worksheet_title"])
    header_row = 1 + st.session_state.get("header_offset", 0)
    df, ws = sheets_core.load_dataframe(sh, ws.title, header_row=header_row)
    st.session_state["df"] = df


def _push_history(entry: dict):
    st.session_state.setdefault("history", []).append(entry)
    st.session_state["redo_stack"] = []  # 新しい操作をしたらやり直し履歴は破棄する


def _format_ops_of(entry: dict) -> list[dict]:
    """apply_format の履歴から書式操作のリストを取り出す。

    複数操作に対応する前に積まれた履歴(単一range形式)も同じ形に揃えて返す。
    """
    ops = entry.get("ops")
    if ops:
        return ops
    return [{
        "range": entry["range"],
        "format": entry["format_spec"],
        "prev_format": entry.get("prev_format"),
    }]


def _revert_entry(sh, entry: dict):
    """undo: entry の操作を取り消して直前の状態に戻す。"""
    action = entry["action"]
    if action in ("write_new", "create_sheet"):
        sheets_core.delete_worksheet_if_exists(sh, entry["worksheet_title"])
    elif action == "write_current":
        ws = sh.worksheet(entry["worksheet_title"])
        old_values = entry["undo_snapshot"]["values"]
        ws.clear()
        if old_values:
            ws.update(old_values)
        if "prev_header_offset" in entry:
            st.session_state["header_offset"] = entry["prev_header_offset"]
    elif action == "delete_sheet":
        for target in entry["targets"]:
            sheets_core.recreate_worksheet(sh, target["title"], target["undo_snapshot"])
    elif action == "format_cells":
        sheets_core.delete_conditional_format(sh, entry["sheet_id"], index=0)
    elif action == "insert_summary":
        ws = sh.worksheet(entry["worksheet_title"])
        sheets_core.remove_summary_block(ws, entry["row_count"])
        st.session_state["header_offset"] = max(st.session_state.get("header_offset", 0) - entry["row_count"], 0)
    elif action == "apply_format":
        ws = sh.worksheet(entry["worksheet_title"])
        # 複数操作をまとめて適用した履歴は、適用時と逆順に、範囲ごとの適用前書式へ戻す。
        for op in reversed(_format_ops_of(entry)):
            restore_spec = op.get("prev_format") or sheets_core.DEFAULT_RESET_FORMAT
            sheets_core.apply_cell_format(ws, op["range"], restore_spec)


def _reapply_entry(sh, entry: dict):
    """redo: 一度取り消した entry の操作をもう一度適用する。

    動的リスト(数式)の場合は、静止した値ではなく数式を再度書き込むことで
    ライブな(自動更新される)状態を復元する。
    """
    action = entry["action"]
    if action == "write_current":
        ws = sh.worksheet(entry["worksheet_title"])
        if entry.get("dynamic"):
            sheets_core.write_formula(ws, entry["formula"], mode=entry.get("mode", "full"), range_start=entry.get("range", "A1"))
        else:
            ws.clear()
            redo_values = entry["redo_snapshot"]["values"]
            if redo_values:
                ws.update(redo_values)
        if "new_header_offset" in entry:
            st.session_state["header_offset"] = entry["new_header_offset"]
    elif action in ("write_new", "create_sheet"):
        if entry.get("dynamic"):
            sheets_core.recreate_worksheet(sh, entry["worksheet_title"], sheets_core.blank_snapshot())
            new_ws = sh.worksheet(entry["worksheet_title"])
            sheets_core.write_formula(new_ws, entry["formula"], mode="partial", range_start=entry.get("range", "A1"))
        else:
            sheets_core.recreate_worksheet(sh, entry["worksheet_title"], entry["redo_snapshot"])
    elif action == "delete_sheet":
        for target in entry["targets"]:
            sheets_core.delete_worksheet_if_exists(sh, target["title"])
    elif action == "format_cells":
        ws = sh.worksheet(entry["worksheet_title"])
        result = sheets_core.add_conditional_format(sh, ws, entry["range"], entry["formula"], entry["color"])
        entry["sheet_id"] = result["sheet_id"]
    elif action == "insert_summary":
        ws = sh.worksheet(entry["worksheet_title"])
        sheets_core.insert_summary_block(ws, entry["rows"])
        st.session_state["header_offset"] = st.session_state.get("header_offset", 0) + entry["row_count"]
    elif action == "apply_format":
        ws = sh.worksheet(entry["worksheet_title"])
        sheets_core.apply_format_ops(ws, _format_ops_of(entry))


def _undo_one() -> dict | None:
    history = st.session_state.get("history", [])
    if not history:
        return None
    entry = history.pop()
    sh = _get_sh()
    _revert_entry(sh, entry)
    if entry["action"] in ("write_current", "insert_summary") and entry.get("worksheet_title") == st.session_state["worksheet_title"]:
        _refresh_current_df(sh)
    st.session_state.setdefault("redo_stack", []).append(entry)
    return entry


def do_undo_last():
    entry = _undo_one()
    if entry is None:
        _flash("warning", "取り消せる操作がありません。")
    else:
        _flash("success", f"操作を取り消しました: {entry['description']}")
    st.session_state.pop("pending_action", None)
    st.rerun()


def do_undo_to(keep_n: int):
    reverted = 0
    while len(st.session_state.get("history", [])) > keep_n:
        _undo_one()
        reverted += 1
    if reverted == 0:
        _flash("warning", "取り消せる操作がありません。")
    else:
        _flash("success", f"{reverted}件の操作を取り消しました。")
    st.session_state.pop("pending_action", None)
    st.rerun()


def do_redo_last():
    redo_stack = st.session_state.get("redo_stack", [])
    if not redo_stack:
        _flash("warning", "やり直せる操作がありません。")
        st.rerun()
        return
    entry = redo_stack.pop()
    sh = _get_sh()
    try:
        _reapply_entry(sh, entry)
    except Exception as exc:  # noqa: BLE001
        redo_stack.append(entry)
        _flash("error", f"やり直しに失敗しました: {exc}")
        st.rerun()
        return
    st.session_state.setdefault("history", []).append(entry)
    if entry["action"] in ("write_current", "insert_summary") and entry.get("worksheet_title") == st.session_state["worksheet_title"]:
        _refresh_current_df(sh)
    _flash("success", f"操作をやり直しました: {entry['description']}")
    st.session_state.pop("pending_action", None)
    st.rerun()


def history_section():
    history = st.session_state.get("history", [])
    redo_stack = st.session_state.get("redo_stack", [])

    st.subheader(f"操作履歴 ({len(history)}件)")
    if not history:
        st.caption("まだ操作はありません。")

    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("直前の操作を取り消す", disabled=not history, key="btn_undo_last"):
            do_undo_last()
    with col2:
        if st.button("取り消しをやり直す", disabled=not redo_stack, key="btn_redo_last"):
            do_redo_last()
    with col3:
        if st.button("すべての操作を取り消す", disabled=not history, key="btn_undo_all"):
            do_undo_to(0)

    if history:
        with st.expander("指定した時点まで戻す"):
            checkpoint_options = ["(初期状態まで全て取り消す)"] + [
                f"{i}. {h['description']}" for i, h in enumerate(history, start=1)
            ]
            selected = st.selectbox("この時点まで戻す", checkpoint_options, index=len(checkpoint_options) - 1)
            if st.button("選択した時点まで戻す", key="btn_undo_to_checkpoint"):
                if selected == checkpoint_options[0]:
                    do_undo_to(0)
                else:
                    keep_n = checkpoint_options.index(selected)
                    do_undo_to(keep_n)


def _render_undo_pending(pending: dict):
    """確認UIを描画し、ボタンが押された場合は実行する関数(未押下ならNone)を返す。

    ここでは実行せずに返すだけにすることで、呼び出し側が実行直前に
    確認UIをクリアしてから実行でき、処理中の二重描画を防げる。
    """
    history = st.session_state.get("history", [])
    if pending["mode"] == "undo_all":
        if not history:
            st.info("取り消せる操作がありません。")
            st.session_state.pop("pending_action", None)
            return None
        st.warning(f"これまでの操作(全{len(history)}件)をすべて取り消します。新しく作成したシートは削除され、上書きしたシートは元の内容に戻ります。")
        for i, h in enumerate(history, start=1):
            st.write(f"{i}. {h['description']}")
        if st.button("すべて取り消す", type="primary", key="btn_confirm_undo_all"):
            return lambda: do_undo_to(0)
        return None
    else:  # undo_last
        if not history:
            st.info("取り消せる操作がありません。")
            st.session_state.pop("pending_action", None)
            return None
        st.warning(f"直前の操作を取り消します: 「{history[-1]['description']}」")
        if st.button("取り消す", type="primary", key="btn_confirm_undo_last"):
            return lambda: do_undo_last()
        return None


def _resolve_delete_targets(plan: dict, sh) -> tuple[list[str], list[str], list[str]]:
    """delete_sheet の対象を解決する。戻り値: (削除対象, 見つからなかった名前, 保護されて除外された名前)"""
    worksheet_titles = [w.title for w in sh.worksheets()]
    targets = plan.get("target_sheets_resolved")
    if targets is None:
        targets = sheets_core.parse_target_sheets(plan.get("target_sheet", ""))

    invalid = [t for t in targets if t not in worksheet_titles]
    current_title = st.session_state["worksheet_title"]
    protected = [t for t in targets if t == current_title]
    valid_targets = [t for t in targets if t in worksheet_titles and t != current_title]
    return valid_targets, invalid, protected


def _render_plan_pending(plan: dict, sh):
    """確認UIを描画し、実行ボタンが押された場合は実行する関数(未押下ならNone)を返す。"""
    action = plan["action"]
    target = plan.get("target_sheet") or ""

    summary = plan.get("summary")

    if action == "write_current":
        existing_titles = {w.title for w in sh.worksheets()}
        target_title = target if target and target in existing_titles else st.session_state["worksheet_title"]
        range_label = "空いているセル" if str(plan.get("range", "")).strip().upper() == "AUTO" else plan.get("range", "A1")
        if summary:
            st.info(summary)
        elif plan.get("mode") == "partial":
            st.info(f"「{target_title}」の「{range_label}」を起点に書き込みます(それ以外の既存データは変更しません)。")
        else:
            st.info(f"「{target_title}」のデータを更新します。")
        if plan.get("dynamic"):
            st.caption("元データの変化に自動追従する数式として書き込みます(動的リスト)。")
            st.code(plan["code"], language=None)
        else:
            st.code(plan["code"], language="python")
    elif action == "write_new":
        label = target or "(自動的に名前が付けられます)"
        if summary:
            st.info(summary)
        else:
            st.info(f"新しいシート「{label}」に結果を書き込みます(元のシートは変更されません)。")
        if plan.get("dynamic"):
            st.caption("元データの変化に自動追従する数式として書き込みます(動的リスト)。")
            st.code(plan["code"], language=None)
        else:
            st.code(plan["code"], language="python")
    elif action == "delete_sheet":
        valid_targets, invalid, protected = _resolve_delete_targets(plan, sh)

        if invalid:
            st.error(f"次のシートが見つかりませんでした: {', '.join(invalid)}")
            worksheet_titles = [w.title for w in sh.worksheets()]
            selectable = [t for t in worksheet_titles if t != st.session_state["worksheet_title"]]
            if not selectable:
                st.info("現在開いているシート以外に削除可能なシートがありません。")
                return None
            chosen = st.multiselect("削除するシートを選び直してください", selectable, key="delete_target_override")
            if chosen and st.button("選択したシートを削除する", type="primary", key="btn_execute_delete_override"):
                override_plan = dict(plan)
                override_plan["target_sheets_resolved"] = chosen
                return lambda: _execute_plan(override_plan)
            return None

        if protected:
            st.warning(f"現在開いているシート「{', '.join(protected)}」は削除できないため対象から除外しました。")

        if not valid_targets:
            st.info("削除できるシートがありません。")
            return None

        st.warning(f"以下のシート({len(valid_targets)}件)を削除します。この操作は後から取り消せます。")
        for t in valid_targets:
            st.write(f"- {t}")
        plan = dict(plan)
        plan["target_sheets_resolved"] = valid_targets
    elif action == "create_sheet":
        label = target or "(自動的に名前が付けられます)"
        st.info(f"新しい空のシート「{label}」を作成します。")
    elif action == "format_cells":
        st.info(f"現在開いているシート「{st.session_state['worksheet_title']}」に条件付き書式(背景色ハイライト)を追加します。")
        st.caption(f"条件式: {plan['code']}")
        st.markdown(
            f'<div style="display:inline-block;width:1em;height:1em;background:{plan["color"]};'
            f'border:1px solid #999;vertical-align:middle;margin-right:0.4em;"></div>'
            f'背景色: {plan["color"]}',
            unsafe_allow_html=True,
        )
    elif action == "insert_summary":
        rows = sheets_core.parse_summary_rows(plan["code"])
        st.info(f"現在開いているシート「{st.session_state['worksheet_title']}」の最上部に、集計欄({len(rows)}項目)を挿入します。既存データは下に押し下げられます。")
        for label, formula in rows:
            st.write(f"- {label}: `{formula}`")
    elif action == "apply_format":
        ops = sheets_core.parse_format_ops(plan["code"], plan.get("range") or "")
        if ops:
            st.info(
                f"現在開いているシート「{st.session_state['worksheet_title']}」に、"
                f"次の{len(ops)}件の書式をまとめて適用します(値は変更しません)。"
            )
            for i, op in enumerate(ops, 1):
                st.write(f"{i}. `{op['range']}` → {sheets_core.describe_format_spec(op['format'])}")
            with st.expander("生成された内容(JSON)を表示"):
                st.code(plan["code"], language="json")
        else:
            range_label = plan.get("range") or "(現在のデータ範囲全体)"
            st.info(f"現在開いているシート「{st.session_state['worksheet_title']}」の「{range_label}」に書式を適用します(値は変更しません)。")
            st.code(plan["code"], language=None)

    if action in ("write_current", "write_new") and not plan.get("dynamic") and not sheets_core.looks_safe(plan["code"]):
        st.error("安全でない可能性のあるコードが検出されたため実行できません。")
        return None

    with st.form("execute_form"):
        submitted = st.form_submit_button("実行する", type="primary")
    _bind_ctrl_enter("実行する")
    if submitted:
        return lambda: _execute_plan(plan)
    return None


def _render_pending_action():
    pending = st.session_state.get("pending_action")
    if not pending:
        return

    st.subheader("3. 内容を確認して実行")
    placeholder = st.empty()

    # 確認UIをまず placeholder の中に描画する。ボタンが押されて実行が
    # 必要になった場合は action に実行関数を入れて返す(この時点ではまだ実行しない)。
    with placeholder.container():
        if pending["type"] == "undo":
            action = _render_undo_pending(pending)
        else:
            try:
                sh = _get_sh()
            except Exception as exc:  # noqa: BLE001
                st.error(f"スプレッドシートへの接続に失敗しました: {exc}")
                action = None
            else:
                action = _render_plan_pending(pending, sh)

    if action is not None:
        # 実行直前に確認UI(情報・コード・フォーム)を消してから実行することで、
        # 処理中に確認UIと進捗表示が二重に表示されるのを防ぐ。
        placeholder.empty()
        action()


def _apply_default_formatting_safe(ws, result, start_row: int, start_col: int) -> None:
    """書き込んだ結果に既定の書式(通貨/数値は右揃え、文字列は左揃え、見出しは太字)を適用する。

    書式付けはあくまで見た目の補助であり、失敗してもデータの書き込み自体は
    成功しているため、ここで例外を投げてAI自動修正リトライを誘発しないよう握りつぶす。
    """
    try:
        sheets_core.apply_default_formatting(ws, result, start_row=start_row, start_col=start_col)
    except Exception:  # noqa: BLE001
        pass


def _execute_plan_once(plan: dict, flash_success):
    """1回分の実行本体。失敗時は例外を投げる(呼び出し元の _execute_plan がリトライを制御する)。"""
    action = plan["action"]
    target = plan.get("target_sheet") or ""
    instruction = plan.get("instruction", "")

    sh = _get_sh()
    with st.spinner("実行中..."):
            # 実行の直前に必ず最新のシート構造・セルデータを取得し直す
            # (生成時から時間が経っていたり、他で手作業編集されていてもキャッシュが残らないように)。
            if action in ("write_current", "write_new") and not plan.get("dynamic"):
                _refresh_current_df(sh)
                other_sheets = sheets_core.load_all_dataframes(sh, exclude=st.session_state["worksheet_title"])

            if action == "write_current":
                # TARGET_SHEETが実在する別のシートを指していればそちらを対象にし、
                # 指示にシート名の指定がなければ最初に読み込んだシートを対象にする。
                existing_titles = {w.title for w in sh.worksheets()}
                target_title = target if target and target in existing_titles else st.session_state["worksheet_title"]
                is_loaded_sheet = target_title == st.session_state["worksheet_title"]

                ws = sh.worksheet(target_title)
                before_values = ws.get_all_values()
                mode = plan.get("mode", "full")
                range_start = plan.get("range", "A1")
                prev_offset = st.session_state.get("header_offset", 0) if is_loaded_sheet else 0
                if str(range_start).strip().upper() == "AUTO":
                    range_start = sheets_core.find_empty_range_start(ws, prev_offset)
                if plan.get("dynamic"):
                    sheets_core.write_formula(ws, plan["code"], mode=mode, range_start=range_start)
                else:
                    result = sheets_core.run_generated_code(plan["code"], st.session_state["df"], other_sheets)
                    sheets_core.apply_write(ws, result, mode=mode, range_start=range_start)
                    fmt_row, fmt_col = (0, 0) if mode == "full" else sheets_core.parse_a1_start(range_start)
                    _apply_default_formatting_safe(ws, result, fmt_row, fmt_col)
                after_values = ws.get_all_values()
                # full書き込みはシート全体を再構成するため、以前挿入した集計行のオフセットは失われる。
                new_offset = 0 if mode == "full" else prev_offset
                history_entry = {
                    "action": "write_current",
                    "worksheet_title": ws.title,
                    "description": f"「{ws.title}」を更新: {instruction}",
                    "undo_snapshot": {"values": before_values},
                    "redo_snapshot": {"values": after_values},
                    "dynamic": bool(plan.get("dynamic")),
                    "formula": plan["code"] if plan.get("dynamic") else None,
                    "mode": mode,
                    "range": range_start,
                }
                if is_loaded_sheet:
                    # 対象が「最初に読み込んだシート」の場合のみ、集計行の挿入状況を追跡する。
                    st.session_state["header_offset"] = new_offset
                    history_entry["prev_header_offset"] = prev_offset
                    history_entry["new_header_offset"] = new_offset
                _push_history(history_entry)
                if is_loaded_sheet:
                    _refresh_current_df(sh)
                flash_success(f"「{ws.title}」のデータを更新しました。")

            elif action == "write_new":
                if plan.get("dynamic"):
                    new_title = sheets_core.create_blank_worksheet(sh, title=target or None)
                    new_ws = sh.worksheet(new_title)
                    range_start = plan.get("range", "A1")
                    if str(range_start).strip().upper() == "AUTO":
                        range_start = sheets_core.find_empty_range_start(new_ws)
                    sheets_core.write_formula(new_ws, plan["code"], mode="partial", range_start=range_start)
                    _push_history({
                        "action": "write_new",
                        "worksheet_title": new_title,
                        "description": f"新しいシート「{new_title}」に動的リストを作成: {instruction}",
                        "dynamic": True,
                        "formula": plan["code"],
                        "range": range_start,
                    })
                    flash_success(f"新しいシート「{new_title}」に動的リストを作成しました。")
                else:
                    result = sheets_core.run_generated_code(plan["code"], st.session_state["df"], other_sheets)
                    new_title = sheets_core.write_result(
                        sh, st.session_state["worksheet_title"], result, title=target or None
                    )
                    _apply_default_formatting_safe(sh.worksheet(new_title), result, 0, 0)
                    _push_history({
                        "action": "write_new",
                        "worksheet_title": new_title,
                        "description": f"新しいシート「{new_title}」に書き込み: {instruction}",
                        "redo_snapshot": sheets_core.snapshot_from_result(result),
                    })
                    flash_success(f"新しいシート「{new_title}」に結果を書き込みました。")

            elif action == "delete_sheet":
                targets = plan.get("target_sheets_resolved") or sheets_core.parse_target_sheets(target)
                targets_snapshot = []
                for title in targets:
                    snapshot = sheets_core.delete_worksheet_by_title(sh, title)
                    targets_snapshot.append({"title": title, "undo_snapshot": snapshot})
                _push_history({
                    "action": "delete_sheet",
                    "targets": targets_snapshot,
                    "description": f"「{', '.join(targets)}」を削除: {instruction}",
                })
                flash_success(f"{len(targets)}件のシートを削除しました: {', '.join(targets)}")

            elif action == "create_sheet":
                new_title = sheets_core.create_blank_worksheet(sh, title=target or None)
                _push_history({
                    "action": "create_sheet",
                    "worksheet_title": new_title,
                    "description": f"新しいシート「{new_title}」を作成: {instruction}",
                    "redo_snapshot": sheets_core.blank_snapshot(),
                })
                flash_success(f"新しいシート「{new_title}」を作成しました。")

            elif action == "format_cells":
                ws = sh.worksheet(st.session_state["worksheet_title"])
                offset = st.session_state.get("header_offset", 0)
                anchor_row = 2 + offset
                # 直前にシートの最新状態を取得し直してから範囲を計算する(他の操作でずれないように)。
                range_a1 = sheets_core.data_range_a1(ws, offset)
                formula = sheets_core.shift_formula_row_anchor(plan["code"], anchor_row)
                color = plan["color"]
                result = sheets_core.add_conditional_format(sh, ws, range_a1, formula, color)
                _push_history({
                    "action": "format_cells",
                    "worksheet_title": ws.title,
                    "sheet_id": result["sheet_id"],
                    "range": range_a1,
                    "formula": formula,
                    "color": color,
                    "description": f"「{ws.title}」に条件付き書式を追加: {instruction}",
                })
                flash_success(f"「{ws.title}」に条件付き書式を追加しました。")

            elif action == "insert_summary":
                ws = sh.worksheet(st.session_state["worksheet_title"])
                rows = sheets_core.parse_summary_rows(plan["code"])
                if not rows:
                    raise RuntimeError("集計項目を解析できませんでした。")
                sheets_core.insert_summary_block(ws, rows)
                st.session_state["header_offset"] = st.session_state.get("header_offset", 0) + len(rows)
                _push_history({
                    "action": "insert_summary",
                    "worksheet_title": ws.title,
                    "rows": rows,
                    "row_count": len(rows),
                    "description": f"「{ws.title}」の最上部に集計欄を挿入: {instruction}",
                })
                _refresh_current_df(sh)
                flash_success(f"「{ws.title}」の最上部に集計欄を挿入しました。")

            elif action == "apply_format":
                ws = sh.worksheet(st.session_state["worksheet_title"])
                offset = st.session_state.get("header_offset", 0)
                # 直前にシートの最新状態を取得し直してから範囲を決める(未指定時のデフォルト範囲用)。
                default_range = plan.get("range") or sheets_core.data_range_a1(ws, offset)
                ops = sheets_core.parse_format_ops(plan["code"], default_range)
                if not ops:
                    raise RuntimeError("適用する書式を解析できませんでした。")

                # undo用に、各範囲の適用前の書式を先に控えておく。
                for op in ops:
                    op["prev_format"] = sheets_core.get_cell_format(sh, ws, op["range"].split(":")[0])

                results = sheets_core.apply_format_ops(ws, ops)
                succeeded = [op for op, r in zip(ops, results) if r["ok"]]
                failed = [(op, r) for op, r in zip(ops, results) if not r["ok"]]

                if succeeded:
                    # 成功した操作だけを履歴に積む(失敗分をundoしようとしないため)。
                    ranges_label = "、".join(op["range"] for op in succeeded)
                    _push_history({
                        "action": "apply_format",
                        "worksheet_title": ws.title,
                        "ops": succeeded,
                        # 旧形式(単一range)の履歴との互換のため、代表値も残しておく。
                        "range": succeeded[0]["range"],
                        "format_spec": succeeded[0]["format"],
                        "prev_format": succeeded[0]["prev_format"],
                        "description": f"「{ws.title}」の{len(succeeded)}箇所に書式を適用: {instruction}",
                    })

                if failed and not succeeded:
                    detail = "、".join(f"{op['range']}({r['error']})" for op, r in failed)
                    raise RuntimeError(f"書式の適用に失敗しました: {detail}")

                if failed:
                    # 一部だけ失敗した場合はクラッシュさせず、成功分を残したまま内訳を表示する。
                    st.session_state["format_failures"] = [
                        {"range": op["range"], "error": r["error"]} for op, r in failed
                    ]
                    flash_success(
                        f"「{ws.title}」の{len(succeeded)}件の書式を適用しました"
                        f"({len(failed)}件は失敗しました)。"
                    )
                else:
                    flash_success(f"「{ws.title}」の{len(succeeded)}件の書式を適用しました({ranges_label})。")


def _execute_plan(plan: dict, attempt: int = 1):
    """_execute_plan_once を実行し、失敗した場合はエラー内容をAIにフィードバックして
    1回だけ自動修正・再実行を試みる(過度なリトライでAPIコストが膨らまないよう1回のみ)。
    """

    def flash_success(msg: str):
        if attempt > 1:
            _flash("success", f"エラーを自動修正して再実行しました。{msg}")
        else:
            _flash("success", msg)

    try:
        _execute_plan_once(plan, flash_success)
    except Exception as exc:  # noqa: BLE001
        if attempt > 1:
            st.error(f"自動修正を試みましたが、実行に失敗しました: {exc}")
            return
        try:
            sh = _get_sh()
            worksheet_titles = [w.title for w in sh.worksheets()]
            _refresh_current_df(sh)
            other_sheets = sheets_core.load_all_dataframes(sh, exclude=st.session_state["worksheet_title"])
            with st.spinner("エラーが発生しました。AIが内容を自動修正しています..."):
                fixed_plan = sheets_core.generate_fix_plan(
                    st.session_state["df"], plan.get("instruction", ""), plan.get("code", ""), str(exc),
                    worksheet_titles, st.session_state["worksheet_title"],
                    st.session_state.get("header_offset", 0), other_sheets,
                )
        except sheets_core.ApiKeyError as key_exc:
            st.error(f"実行に失敗しました: {exc}\n\n{key_exc}")
            return
        except Exception as fix_exc:  # noqa: BLE001
            st.error(f"実行に失敗しました: {exc}\n\n自動修正の生成にも失敗しました: {fix_exc}")
            return
        if fixed_plan.get("code_error"):
            st.error(fixed_plan["code_error"])
            return
        fixed_plan["type"] = "plan"
        fixed_plan["instruction"] = plan.get("instruction", "")
        _execute_plan(fixed_plan, attempt=2)
        return

    st.session_state.pop("pending_action", None)
    st.rerun()


def instruction_section():
    st.subheader("やってほしいことを入力")
    with st.form("instruction_form", enter_to_submit=False):
        instruction = st.text_input(
            "指示", placeholder="例: 商品名ごとの売上合計を出して", key="instruction_text"
        )
        st.caption("💡 指示通りにいかない場合は、「〇〇シートのA列」のように具体的に指定してください。")
        submitted = st.form_submit_button("内容を生成", type="primary")

    generate_placeholder = st.empty()

    if submitted:
        # まず直前の描画内容(古い確認UIの残り等)をこの位置から消してから、
        # 生成中の状態(エラー/スピナー)をここに描画する。処理中の二重表示を防ぐ。
        generate_placeholder.empty()
        if not instruction.strip():
            with generate_placeholder.container():
                st.error("指示を入力してください。")
        else:
            undo_intent = sheets_core.detect_undo_intent(instruction)
            if undo_intent:
                st.session_state["pending_action"] = {"type": "undo", "mode": undo_intent}
                st.rerun()
            else:
                plan = None
                try:
                    sh = _get_sh()
                    worksheet_titles = [w.title for w in sh.worksheets()]
                    # 生成の直前に必ず最新のシート構造・セルデータを取得し直す
                    # (ユーザーがスプレッドシート側で手作業編集していてもキャッシュが残らないように)。
                    _refresh_current_df(sh)
                    other_sheets = sheets_core.load_all_dataframes(sh, exclude=st.session_state["worksheet_title"])
                    with generate_placeholder.container():
                        with st.spinner("Claude に依頼中..."):
                            plan = sheets_core.generate_plan(
                                st.session_state["df"], instruction.strip(), worksheet_titles,
                                st.session_state["worksheet_title"],
                                st.session_state.get("header_offset", 0),
                                other_sheets,
                            )
                except sheets_core.ApiKeyError as exc:
                    with generate_placeholder.container():
                        st.error(str(exc))
                except Exception as exc:  # noqa: BLE001
                    with generate_placeholder.container():
                        st.error(f"生成に失敗しました: {exc}")

                if plan is not None:
                    if plan.get("code_error"):
                        with generate_placeholder.container():
                            st.error(plan["code_error"])
                    else:
                        plan["type"] = "plan"
                        plan["instruction"] = instruction.strip()
                        st.session_state["pending_action"] = plan
                        st.rerun()

    _render_pending_action()


def main():
    if not require_login():
        return

    st.title("Sheets AI Assistant")
    if st.button("ログアウト"):
        st.session_state.clear()
        st.rerun()

    _show_flash()
    _show_format_failures()
    _show_api_key_warning()
    load_sheet_section()

    if st.session_state.get("sheet_loaded"):
        history_section()
        instruction_section()


if __name__ == "__main__":
    main()
