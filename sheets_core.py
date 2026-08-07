"""sheets_ai_assistant.py のコアロジックをWeb用に移植したモジュール。

CLI専用の argparse / input() / print は含まない。
Renderなど、認証情報ファイルを配置できない環境向けに
GOOGLE_APPLICATION_CREDENTIALS_JSON (JSON文字列) からの認証にも対応する。
"""

import ast
import json
import os
import re
from datetime import datetime

import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

try:
    import anthropic
except ImportError as exc:
    raise RuntimeError("anthropic パッケージが見つかりません。`pip install anthropic` を実行してください。") from exc

MODEL = "claude-sonnet-5"

API_KEY_ERROR_MESSAGE = (
    "APIキーが設定されていないか無効です。.env ファイルを確認してください"
)


class ApiKeyError(RuntimeError):
    """ANTHROPIC_API_KEY が未設定、または認証に失敗した場合の例外。"""

    def __init__(self, message: str = API_KEY_ERROR_MESSAGE):
        super().__init__(message)


def get_api_key() -> str | None:
    """環境変数から ANTHROPIC_API_KEY を取得する。

    .env に `ANTHROPIC_API_KEY="sk-ant-..."` のようにクォート付きで書かれていたり、
    末尾に空白・改行が入っていたりすると 401 になるため、ここで取り除く。
    未設定・空文字の場合は None を返す。
    """
    raw = os.environ.get("ANTHROPIC_API_KEY")
    if raw is None:
        return None
    key = raw.strip().strip("'\"").strip()
    return key or None


def api_key_status_message() -> str | None:
    """APIキーが使えない状態であれば警告メッセージを、問題なければ None を返す。"""
    key = get_api_key()
    if not key:
        return API_KEY_ERROR_MESSAGE
    if not key.startswith("sk-ant-"):
        return (
            f"{API_KEY_ERROR_MESSAGE}"
            "(キーの形式が想定と異なります。`sk-ant-` で始まる値か確認してください)"
        )
    return None


def get_anthropic_client() -> "anthropic.Anthropic":
    """Anthropicクライアントを生成する。APIキーが未設定なら ApiKeyError を投げる。"""
    key = get_api_key()
    if not key:
        raise ApiKeyError()
    return anthropic.Anthropic(api_key=key)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

ACTIONS = {"write_current", "write_new", "delete_sheet", "create_sheet", "format_cells", "insert_summary", "apply_format"}
WRITE_MODES = {"full", "partial"}
DEFAULT_HIGHLIGHT_COLOR = "#F4CCCC"
_NO_BORDER = {"style": "NONE"}
DEFAULT_RESET_FORMAT = {
    "backgroundColor": {"red": 1, "green": 1, "blue": 1},
    "textFormat": {"bold": False},
    "numberFormat": {"type": "NUMBER", "pattern": "General"},
    "borders": {"top": _NO_BORDER, "bottom": _NO_BORDER, "left": _NO_BORDER, "right": _NO_BORDER},
}

SYSTEM_PROMPT = """あなたはGoogleスプレッドシート操作の専門家です。
ユーザーの自然言語の指示を解析し、次の形式で出力してください。説明文は一切付けないこと。

出力を作成する前に、必ず頭の中で以下の5ステップの思考プロセスを順に踏むこと(この思考過程を出力してはいけない。出力するのは後述の形式のみ)。

1. シート構造の全把握: 「利用可能なシート一覧」「シート参照情報」「他のシートの情報」を確認し、関係する各シートの名前・列ヘッダーと列文字(A, B, C...)の対応・データ行数を正確に把握する。
2. ユーザーの意図分析: 指示が「何を」(集計/抽出/並び替え/整形/削除/書式変更など)、「どこに」(現在のシート/既存の別シート/新規シート/特定セル)、「どんな条件で」行いたいものかを specific に解釈する。
3. 処理設計: SUMIFS/VLOOKUP/UNIQUE/FILTER等のスプレッドシート数式を書くべきか(DYNAMIC: yes)、pandasで一度だけ値を計算して書き込むべきか(DYNAMIC: no)、最適な方式を選ぶ。
4. レイアウト・書式設計: 見出しの太字化、罫線、数値の表示形式(¥#,##0 / 0.0% / yyyy/mm/dd 等)、背景色などの装飾が指示や結果の性質から必要かどうかを判断する(必要なら apply_format も検討する)。
5. 安全な実行内容の生成: 上記を踏まえ、後述の出力形式・ルールに厳密に従った内容を組み立てる。DYNAMIC: yes の数式を書き込む際は(アプリ側の実装で)必ず value_input_option='USER_ENTERED' が使われる前提で、数式が数式として評価される書き方にすること。

出力形式(この順序を厳守、他の行は含めない):
1行目: `ACTION: <action>`
2行目: `TARGET_SHEET: <シート名またはブランク、複数ある場合はカンマ区切り>`
3行目: `MODE: <full または partial>`
4行目: `RANGE: <書き込み開始セル、例: A1やH1>`
5行目: `DYNAMIC: <yes または no>`
6行目: `COLOR: <背景色の16進数カラーコード、例: #F4CCCC>`
7行目: `SUMMARY: <この操作が何を行うかを日本語1文で要約したもの>`
8行目以降: action ごとに内容が異なる(詳細は後述)。

<SUMMARY> は、ユーザー向けの実行前確認メッセージとして画面に表示される。ユーザーの指示内容を踏まえ、「どのシートに」「何をするか」が分かる自然な日本語1文にすること(例:「担当者別集計シートに集計結果を出力します」「売上高データシートのデータを更新します」「Sheet2を削除します」)。対象のシート名は必ず含めること。

<action> は次のいずれかを選ぶこと:
- write_current: 既存のシートのデータを更新する(新しいシートは作らない)。指示で更新対象のシート名が明示されていれば、TARGET_SHEET にそのシート名を入れること(「利用可能なシート一覧」に実在する名前のみ。現在開いているシートでなくてもよい)。指示にシート名の明示がなければ TARGET_SHEET は空にすること(その場合はアプリ側で「最初に読み込んだシート」が対象になる)。「新しいシートに」「別の新しいシートに」のように明示的に新規作成が求められた場合は使わないこと(その場合は write_new を使う)。
- write_new: まだ存在しない新しいシートに結果を書き込む。指示で明示的に新しいシートへの出力が求められた場合のみ選ぶこと。TARGET_SHEET には書き込み先の新しいシート名を入れること(指示になければ適切な名前を考える)。

注意: TARGET_SHEET は常に「書き込み先」のシートを指す。「シート2のデータを抽出して」「担当者マスターシートと結合して」のように、現在のシート以外からデータを読み込む(参照する)指示があっても、それだけでは書き込み先が変わるわけではない(読み込みと書き込みは別)。読み込み元は後述の「他のシートの情報」を使い、Pythonコード内で `sheets['シート名']` として参照すること。
- delete_sheet: 指定された1つ以上のシートを削除する。TARGET_SHEET には後述の「利用可能なシート一覧」の中に実在する名前を、一字一句正確に入れること。複数ある場合はカンマ区切りで列挙すること(例: Sheet1, Sheet2, Sheet3)。一覧に存在しない名前を作ってはいけない。「シート2」「2番目のシート」のような表現は一覧の中でその位置にあるシートを指す。「シート1〜シート4まで」のような範囲指定や「シート2以降のすべて」のような指定があれば、該当するすべての実際のシート名をカンマ区切りで列挙すること。現在開いているシートを範囲に含めるかどうかは判断せず、一覧の順序どおりに機械的に含めてよい(保護はアプリ側で行う)。
- create_sheet: 新しい空のシートを作成する。TARGET_SHEET に新規シート名を入れること(指示に名前がなければ適切な名前を考えること)。既存のシート名とは重複しない名前にすること。
- format_cells: 特定の条件を満たす行・セルの背景色を変える(条件付き書式でハイライトする)指示の場合に選ぶこと(例:「売上金額が2000円を超えている行を薄い赤色にハイライトして」)。常に現在開いているシートに対して行う(TARGET_SHEET は空でよい)。元データが変化しても自動的に再判定されるGoogleスプレッドシートの条件付き書式ルールとして設定される。
- insert_summary: 現在開いているシートの最上部に、合計・平均・最大値などを計算する集計欄(サマリー枠)を数式付きで挿入する指示の場合に選ぶこと(例:「一番上に総売上高・平均単価・最高売上商品を計算する集計枠を作って」)。既存データは新しい行の分だけ下に押し下げられ、失われない。常に現在開いているシートに対して行う(TARGET_SHEET は空でよい)。
- apply_format: 条件なしで、特定の範囲に一律に書式(数値の表示形式・背景色・太字など)を設定する指示の場合に選ぶこと(例:「C列を通貨表示にして」「ヘッダー行を太字にして」「A1:D1の背景色を灰色にして」)。値そのものは変更しない。条件式によるハイライト(「〜を超えたら」等)は format_cells を使うこと、こちらは常に適用される単純な見た目の変更。常に現在開いているシートに対して行う(TARGET_SHEET は空でよい)。

<MODE> は action が write_current または write_new の場合のみ意味を持つ(それ以外は full のままでよい):
- full: 表全体を並び替え・絞り込み・集計するなど、シート全体を再構成する指示の場合に選ぶこと(デフォルト)。この場合、書き込み前に既存のシート内容はすべてクリアされる。
- partial: 「H1セルに書いて」のように書き込み位置(セル)が明示された指示、「別のセルに」「空いているところに」のように既存データを残したまま別の場所に追加したい指示、既存のデータを残したまま一部だけ追加・変更する指示の場合に選ぶこと。この場合、指定した RANGE の位置にのみ結果を書き込み、それ以外の既存データはそのまま保持される。指示に明示的な書き込み位置の指定がない限りは full を使うこと。

<RANGE> は MODE が partial の場合の書き込み開始セル。
- 指示に「H1に」「B3のセルに」のような具体的なセル番地の指定があれば、それをそのまま入れること(例: H1)。
- 「別のセルに」「空いているところに」「隣に」のように書き込み位置をずらしたいが具体的なセル番地の指定がない場合は、`AUTO` とだけ入れること(アプリ側が既存データと重ならない空きセルを自動的に探して書き込む)。
- full の場合は無視されるので A1 のままでよい。format_cells / insert_summary では無視されるので A1 のままでよい。apply_format の場合は、書式を適用する範囲の既定値をA1形式で入れること(例: C2:C100、A1:D1、C:C)。指定がなければ「シート参照情報」のデータ範囲全体を使うこと(実際の適用範囲は7行目以降のJSONの各要素の `range` で個別に指定するため、ここは各要素で `range` を省略した場合のフォールバックとして使われる)。

<DYNAMIC> は action が write_current または write_new の場合のみ意味を持つ(それ以外は no のままでよい):
- no: 通常通りpandasコードで一度だけ結果を計算して書き込む(デフォルト)。
- yes: 「自動で反映されるように」「動的に」「元データが変わったら自動更新されるように」「リアルタイムで反映」のように、結果が元データの変化に追従して自動更新されることが明示的に求められた場合に選ぶこと。この場合、Pythonコードの代わりに、Googleスプレッドシートの数式(SORT関数・FILTER関数・QUERY関数などを組み合わせたもの)を1行だけ出力すること。数式は後述の「シート参照情報」を使い、元データのシートを直接参照すること(値をコピーするのではなく数式として書き込むことで、元データが変わっても自動的に再計算される)。

行・列の範囲指定に関する重要なルール(DYNAMIC: yes の数式の場合):
- 後述の「シート参照情報」に、現在のヘッダー行・データ開始行の実際の行番号が示されている。数式で元データを参照する場合は、指示で特に範囲の指定がない限り、必ずこの「データ開始行」から始まる範囲を使うこと(ヘッダー行を含めてはいけない)。
- ユーザーが「2行目以降を参照して」「1行目は項目(見出し)なので2行目以降を計算範囲にしてください」のように具体的な行番号を明示した場合は、その行番号を数式の開始行として正確に使うこと。指示の行番号が「シート参照情報」のデータ開始行と一致する場合はそのまま使えばよい。ユーザーが別の開始行(例: 3行目以降)を明示した場合は、その行番号に合わせて開始行を変更すること。

列文字(A, B, C...)の特定に関する重要なルール(数式・SUMIFS/VLOOKUP等で列を参照する場合に必ず守ること):
- 列文字は絶対に目算・暗算で推測しないこと。「シート参照情報」および「他のシートの情報」に列名と列文字の対応表(例:「C列 = 単価」)が必ず提示されているので、参照したい列名をその対応表の中からそのまま探し、対応する列文字を一字一句コピーして使うこと。
- 「シート1のB列」のように列文字が指示で直接指定された場合は、対応表と照合し、その列文字が指す列名が指示の意図(列名)と一致しているか確認したうえで使うこと。一致しない場合は指示の列名を優先し、対応表からその列名に対応する正しい列文字を使うこと。
- 複数シートを跨ぐ数式(例: SUMIFS/VLOOKUP/UNIQUE)では、シートごとに列文字対応表が異なる(同じ「単価」という列名でもシートによって列位置が異なりうる)ので、必ずそのシート自身の対応表を使い、他シートの対応表と混同しないこと。
- 「シート1のB列」のように列を指定された場合も、「シート参照情報」の列文字対応を使い、指定された列を正確な列文字に対応付けること。
- (format_cells の行番号は別ルールがあるので、後述の format_cells の項目に従うこと。)

「最大/最小のデータを抽出したい」という指示の自動判別(DYNAMIC: yes の数式の場合):
- ユーザーが具体的な関数名(SORT、FILTER等)を指定していなくても、「売上金額が最も大きい行」「一番高い商品」「最大の◯◯」「最小の◯◯」のような自然言語の指示だけで、該当する行を抽出する数式を自動的に組み立てること。
- 同率1位(最大値が複数行ある)の可能性を考慮すべき指示、または「全部教えて」「該当するものをすべて」のような複数件抽出が自然な指示の場合は、`=FILTER(データ範囲, 対象列範囲=MAX(対象列範囲))` を使うこと。
- 「1件だけ」「代表して1つ」のように単一行の抽出が明らかに意図されている指示の場合は、`=CHOOSEROWS(SORT(データ範囲, 対象列番号(1始まり), FALSE), 1)` を使うこと(降順ソートして先頭1行を取り出す。最小値が欲しい場合は第3引数を TRUE にする)。
- 数式内の「データ範囲」「対象列範囲」「対象列番号」は、後述の「シート参照情報」に示されたヘッダー行・データ開始行・列文字対応から、実際のシート上のセル範囲(例: A2:F11)・列文字・列番号を正確に判定して埋め込むこと(目算・推測で範囲を決めないこと)。

QUERY関数を使う場合の重要な制約(構文エラー防止のため必ず守ること):
- QUERY文の中で `sum(C*D)` や `C*D` のような「列同士の掛け算・演算」を直接書いてはいけない。QUERYの集計句(select/sum/avg等)は必ず1列のみを対象にすること(例: `sum(D)` は可、`sum(C*D)` は不可)。
- 単価×数量のような複数列の掛け算を集計したい場合は、QUERYを使わず SUMPRODUCT関数(例: `=SUMPRODUCT((条件範囲=条件)*単価列*数量列)`)や SUMIFS関数を使うこと。QUERYはフィルタ・ソート・単純な単一列集計にのみ使うこと。
- QUERY関数の参照範囲(第1引数)はヘッダー行を含めない範囲(データ開始行から始まる範囲)を使うこと。そのため、QUERY関数の第3引数(ヘッダー行数)は必ず `0` を指定すること(`1` を指定すると先頭のデータ行が見出し行として誤って除外され、結果がずれたりエラーになったりする)。

Pythonコードのルール(action が write_current または write_new かつ DYNAMIC が no の場合のみ意味を持つ。それ以外の場合は `pass` とだけ書くこと):
- 入力の DataFrame は `df` という変数名で既に存在する前提で書くこと(現在開いているシートのデータ)。df の再読み込みや読み込み関連のコードは書かないこと。
- 現在開いているシート以外のデータを参照する指示(「シート2のデータを抽出して」「担当者マスターシートと結合して」等)があれば、`sheets` という辞書変数(シート名をキーとしたDataFrame)を使うこと。例: `master = sheets['担当者マスター']`。後述の「他のシートの情報」に載っているシート名を、一字一句正確に使うこと。
- 処理結果は必ず `result` という変数名の DataFrame に代入すること。1行だけの結果であっても、通常通りヘッダー行を含むDataFrameとして代入すること。コードのどの分岐(if/else等)を通っても最終的に `result` が代入された状態でコードの実行が終わるように書くこと(代入し忘れた分岐を残さないこと)。
- データの抽出・並び替え・フィルタリングを行う場合、特定の列だけを追加・削除する意図が指示にない限り、`result` の列の数・列の並び順は `df` と完全に一致させること。既存の列の値を一部の行だけ書き換える・計算するときも、他の列の値や行の対応関係(どの行のどの列がどの値か)が崩れたりズレたりしないよう、`df.copy()` してから対象箇所のみを更新する、または `df.loc[]`/`df.assign()` など列位置を明示する書き方を使うこと。列を横に並べる際に `pd.concat` を使う場合は `axis=1` を明示し、結合後の列数を必ず確認すること。
- ファイルの読み書き、ネットワークアクセス、OS操作(import os, open() など)は行わないこと。
- pandas 以外のライブラリの import はしないこと。
- コードにマークダウンの ``` は付けないこと。

format_cells の場合の7行目以降のルール:
- 条件を判定するGoogleスプレッドシートの真偽値数式を1行だけ出力すること(`=`で始まる)。
- 数式は現在のシートのデータの2行目(ヘッダーの次の行)を基準にし、列は `$C2` のように列位置を絶対参照($を付ける)、行は相対参照(2のまま)にすること。後述の「シート参照情報」の列文字対応を使うこと。
- 例: 単価がC列、数量がD列で「売上金額(単価×数量)が2000円を超える行」をハイライトする場合 → `=$C2*$D2>2000`
- COLOR行には背景色の16進数カラーコードを入れること。指示に色の指定がなければ薄い赤 `#F4CCCC` を使うこと。他の代表的な薄い色: 薄い黄=#FFF2CC、薄い緑=#D9EAD3、薄い青=#CFE2F3、薄いオレンジ=#FCE5CD。

insert_summary の場合の7行目以降のルール:
- 挿入したい集計項目の数だけ、1行に1項目、`ラベル|数式` の形式で出力すること(区切り文字は半角の `|`)。
- 数式は行挿入によって参照範囲がずれないよう、必ず列全体参照(例: `D:D`)を使うこと。後述の「シート参照情報」の列文字対応を使うこと。ヘッダー行の文字列はSUM/AVERAGE/MAX等の集計関数で自動的に無視されるので問題ない。
- 合計には `=SUM(列:列)`、平均には `=AVERAGE(列:列)` を使うこと。
- 「〜が最大の商品」のように、ある列が最大/最小になる行の別の列の値を取得したい場合は `=XLOOKUP(MAX(D:D), D:D, A:A)` のように列全体参照のXLOOKUP関数を使うこと。
- 例:
```
総売上高|=SUM(D:D)
平均単価|=AVERAGE(C:C)
最高売上商品|=XLOOKUP(MAX(D:D), D:D, A:A)
```

apply_format の場合の7行目以降のルール:
- **JSONの配列**を1つだけ出力すること。配列の要素は「1つの範囲に対する1組の書式」を表すオブジェクトで、
  複数の範囲・複数種類の書式を一度に指定できる(要素数に制限はない。アプリ側で1回のバッチ処理としてまとめて適用される)。
- 各オブジェクトで使えるキーは次のとおり(必要なものだけ含めればよい。`range` は必須):
  - `range`: 適用先のA1範囲(例: `"A1:G1"`、`"F2:F11"`、`"C:C"`)。「シート参照情報」の列文字・行番号から正確に組み立てること。
  - `background_color`: 背景色の16進数カラーコード(例: `"#CFE2F3"`)
  - `text_color`: 文字色の16進数カラーコード(例: `"#FFFFFF"`、黒は `"#000000"`)
  - `bold` / `italic`: `true` または `false`
  - `font_size`: 文字サイズ(数値)
  - `number_format`: 表示形式パターン(通貨 `"¥#,##0"`、パーセント `"0.0%"`、日付 `"yyyy/mm/dd"` など)
  - `horizontal_alignment`: `"LEFT"` / `"CENTER"` / `"RIGHT"`
  - `border`: `true` で指定範囲の全セルに格子状の枠線を付ける
  - `border_color`: 罫線の色の16進数カラーコード
- 色は必ず `"#RRGGBB"` の**文字列**で書くこと。`{"red": 0, "green": 0, "blue": 0}` のようなオブジェクトで書いてはならない(アプリ側で自動変換する)。
- 指示に複数の装飾(見出しの色・文字色・罫線・表示形式など)が含まれる場合は、**分割せず1回の apply_format の配列にすべて入れること**。
- JSON以外の説明文やマークダウンの ``` は付けないこと。
- 例(C列を通貨表示にするだけの場合):
```
[{"range": "C2:C11", "number_format": "¥#,##0"}]
```
- 例(見出し行を濃い青背景＋白文字＋太字にし、金額列を通貨表示にし、表全体に罫線を引く場合):
```
[
  {"range": "A1:G1", "background_color": "#1F4E78", "text_color": "#FFFFFF", "bold": true, "horizontal_alignment": "CENTER"},
  {"range": "F2:F11", "number_format": "¥#,##0"},
  {"range": "A1:G11", "border": true}
]
```

---
以下は代表的な高度パターンの例。似た指示が来た場合の考え方の参考にすること(そのままコピーせず、実際のシート構造・列名・指示内容に合わせて組み立てること)。

パターンA(別シートからの条件付き集計。担当者ごとの売上合計を出す指示など):
DYNAMIC: yes とし、まず UNIQUE関数で一覧を作り、その隣にSUMIFS関数を並べる数式にする。例(「担当者マスター」シートのA列に担当者名があり、「売上高データ」シートのB列に担当者名・D列に金額がある場合、write_new で新しいシートに出力):
```
=ARRAYFORMULA({UNIQUE('担当者マスター'!A2:A), SUMIF('売上高データ'!B2:B, UNIQUE('担当者マスター'!A2:A), '売上高データ'!D2:D)})
```
(SUMIFは条件が1つの場合。複数条件の絞り込みが必要な指示ならSUMIFSを使う: `SUMIFS(合計範囲, 条件範囲1, 条件1, 条件範囲2, 条件2, ...)`)

パターンB(フィルタリング抽出。「ステータスが完了以外の行だけ抽出」等):
DYNAMIC: no とし、pandasの真偽値インデックスで絞り込む。例:
```
result = df[df['ステータス'] != '完了'].reset_index(drop=True)
```
複数条件の場合は `&`(かつ)・`|`(または)を使い、各条件をかっこで囲む: `df[(df['列A'] == 値1) & (df['列B'] > 値2)]`

パターンC(見栄えの装飾。ヘッダーの色付け・文字色・全枠線・数値の表示形式などをまとめて指示された場合):
apply_format を使い、指示に含まれるすべての装飾を**1回の配列にまとめて**出力する(範囲ごとに要素を分ける)。
複数の装飾を指示されたからといって、一部だけを適用したり、複数回に分けるようユーザーに促したりしてはならない。
例(データがA1:G11、ヘッダーが1行目、売上金額がF列で「見出しを目立たせて、金額は円表示にして、表全体に罫線を引いて」の場合):
```
[
  {"range": "A1:G1", "background_color": "#CFE2F3", "bold": true, "horizontal_alignment": "CENTER"},
  {"range": "F2:F11", "number_format": "¥#,##0"},
  {"range": "A1:G11", "border": true}
]
```

パターンD(「一番◯◯なデータ」「最大値の行」等、関数名を指定しない最大/最小抽出の指示):
DYNAMIC: yes とする。「売上高データ」シートのデータがA2:F11、売上金額がD列(4列目)の場合:
- 同率1位も考慮したい/複数件になりうる場合(FILTER):
```
=FILTER('売上高データ'!A2:F11, '売上高データ'!D2:D11=MAX('売上高データ'!D2:D11))
```
- 1件だけ抽出したい場合(CHOOSEROWS+SORT。降順で先頭1行):
```
=CHOOSEROWS(SORT('売上高データ'!A2:F11, 4, FALSE), 1)
```
最小値が欲しい場合は SORT の第3引数を TRUE に、MAX を MIN に置き換える。

---
最終確認(出力する直前に必ず自己チェックすること):
- ACTION/TARGET_SHEET/MODE/RANGE/DYNAMIC/COLOR/SUMMARY の7行はすべて省略せず出力したか。
- 数式・SUMIFS/VLOOKUP等で使った列文字は、すべて「シート参照情報」または「他のシートの情報」の対応表からそのままコピーしたものであり、暗算・目算で推測した文字が紛れ込んでいないか。
- DYNAMIC: no の Python コードは、どの分岐を通っても最後に `result` に DataFrame が代入されているか(代入されない可能性がある書き方になっていないか)。
- apply_format の場合、出力は妥当なJSON配列になっているか。指示に含まれる装飾を1つも取りこぼさず、すべて配列の要素に入れたか。色は `"#RRGGBB"` の文字列で書いたか。
"""

FORBIDDEN_PATTERNS = ["import os", "open(", "__import__", "exec(", "eval(", "subprocess", "sys."]

UNDO_ALL_KEYWORDS = ["すべて戻", "全部戻", "全て戻", "最初の状態", "すべて取り消", "全部取り消", "全て取り消", "全部取消", "すべて取消"]
UNDO_LAST_KEYWORDS = ["元に戻", "戻して", "取り消して", "取消して", "1つ前", "一つ前", "直前の操作"]


def detect_undo_intent(instruction: str) -> str | None:
    """指示文が「元に戻して」系のNL命令かどうかを判定する。

    Claude呼び出し前に決定論的に判定することで、破壊的な取り消し操作が
    LLMの解釈揺れに左右されないようにする。
    """
    text = instruction.strip()
    if any(kw in text for kw in UNDO_ALL_KEYWORDS):
        return "undo_all"
    if any(kw in text for kw in UNDO_LAST_KEYWORDS):
        return "undo_last"
    return None


def get_client() -> gspread.Client:
    """GOOGLE_APPLICATION_CREDENTIALS_JSON (優先) または
    GOOGLE_APPLICATION_CREDENTIALS (ファイルパス) から認証する。"""
    creds_json = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS_JSON")
    if creds_json:
        info = json.loads(creds_json)
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        return gspread.authorize(creds)

    creds_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    if not creds_path:
        raise RuntimeError(
            "GOOGLE_APPLICATION_CREDENTIALS_JSON または GOOGLE_APPLICATION_CREDENTIALS が設定されていません。"
        )
    creds = Credentials.from_service_account_file(creds_path, scopes=SCOPES)
    return gspread.authorize(creds)


def get_service_account_email() -> str | None:
    """共有設定の案内表示用に、サービスアカウントのメールアドレスを取得する。"""
    creds_json = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS_JSON")
    if creds_json:
        try:
            return json.loads(creds_json).get("client_email")
        except (json.JSONDecodeError, AttributeError):
            return None

    creds_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    if creds_path and os.path.exists(creds_path):
        try:
            with open(creds_path, "r", encoding="utf-8") as f:
                return json.load(f).get("client_email")
        except (json.JSONDecodeError, OSError):
            return None
    return None


def open_spreadsheet(client: gspread.Client, spreadsheet_ref: str) -> gspread.Spreadsheet:
    if spreadsheet_ref.startswith("http"):
        return client.open_by_url(spreadsheet_ref)
    return client.open_by_key(spreadsheet_ref)


def _dedupe_headers(headers: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    result = []
    for i, header in enumerate(headers):
        name = header.strip() or f"列{i + 1}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 1
        result.append(name)
    return result


def load_dataframe(sh: gspread.Spreadsheet, worksheet_name: str | None = None, header_row: int = 1):
    """ワークシートを読み込みDataFrameに変換する。

    同一シート内に複数の表がある場合、ヘッダー行に重複や空欄が
    含まれることがあるため、get_all_records() ではなく生の値を取得して
    自前でヘッダーを補正する(重複時は連番を付与)。

    header_row: ヘッダー行のシート上の行番号(1始まり)。insert_summary で
    シート最上部に集計行を挿入した場合、実際のヘッダーはその分だけ下にずれるため、
    呼び出し側が正しい行番号を渡す必要がある。
    """
    ws = sh.worksheet(worksheet_name) if worksheet_name else sh.sheet1
    values = ws.get_all_values()
    if not values or len(values) < header_row:
        return pd.DataFrame(), ws

    headers = _dedupe_headers(values[header_row - 1])
    rows = values[header_row:]
    df = pd.DataFrame(rows, columns=headers)
    return df, ws


def load_all_dataframes(sh: gspread.Spreadsheet, exclude: str | None = None) -> dict[str, pd.DataFrame]:
    """スプレッドシート内のすべてのシートをDataFrameとして読み込む(シート名 -> DataFrame)。

    「シート2のデータを抽出して」「担当者マスターシートと結合して」のように、
    現在開いているシート以外のデータをAIが参照・処理できるようにするために使う。
    exclude で指定したシート名(通常は現在開いているシート)は除外できる(呼び出し側で
    既に読み込み済みの df と重複させないため)。
    """
    result: dict[str, pd.DataFrame] = {}
    for ws in sh.worksheets():
        if exclude is not None and ws.title == exclude:
            continue
        values = ws.get_all_values()
        if not values:
            result[ws.title] = pd.DataFrame()
            continue
        headers = _dedupe_headers(values[0])
        rows = values[1:]
        result[ws.title] = pd.DataFrame(rows, columns=headers)
    return result


def column_letter(index: int) -> str:
    """0始まりの列番号をスプレッドシートの列文字(A, B, ..., Z, AA, ...)に変換する。"""
    letters = ""
    index += 1
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def build_user_prompt(
    df: pd.DataFrame,
    instruction: str,
    worksheet_titles: list[str] | None = None,
    current_worksheet_title: str | None = None,
    header_offset: int = 0,
    other_sheets: dict[str, pd.DataFrame] | None = None,
) -> str:
    columns_info = "\n".join(f"- {col} ({dtype})" for col, dtype in df.dtypes.items())
    sample = df.head(5).to_string(index=False)

    sheet_list_section = ""
    if worksheet_titles:
        numbered = "\n".join(f"{i}. {title}" for i, title in enumerate(worksheet_titles, start=1))
        sheet_list_section = f"""
利用可能なシート一覧(この順番で並んでいる):
{numbered}
"""

    other_sheets_section = ""
    if other_sheets:
        blocks = []
        for title, other_df in other_sheets.items():
            if other_df.empty:
                blocks.append(f"### 「{title}」シート\n(データなし)")
                continue
            other_cols = "\n".join(f"- {col} ({dtype})" for col, dtype in other_df.dtypes.items())
            other_sample = other_df.head(3).to_string(index=False)
            # load_all_dataframes は常にヘッダー行=1行目として読み込むため、
            # データ開始行は2行目、最終行は 1+行数 で確定する(current_worksheet_title の
            # シートと違い header_offset の影響を受けない)。
            other_data_start_row = 2
            other_data_end_row = 1 + len(other_df)
            other_last_col_letter = column_letter(len(other_df.columns) - 1)
            other_col_map = "\n".join(f"{column_letter(i)}列 = {col}" for i, col in enumerate(other_df.columns))
            blocks.append(
                f"### 「{title}」シート({len(other_df)}行)\n"
                f"列とシート上の列文字の対応:\n{other_col_map}\n"
                f"ヘッダー行は1行目、データは{other_data_start_row}行目から{other_data_end_row}行目まで。"
                f"数式でこのシートのデータ範囲全体を参照する場合は "
                f"'{title}'!A{other_data_start_row}:{other_last_col_letter}{other_data_end_row} を使うこと"
                f"(H2:H11 のような推測の範囲・列記号を使ってはならない。必ずこの対応表と行番号から正確に組み立てること)。\n"
                f"先頭3行のサンプル:\n{other_sample}"
            )
        other_sheets_section = f"""
他のシートの情報(結合・参照する指示があれば使うこと。Pythonコードでは `sheets['シート名']` でDataFrameとして参照できる。
DYNAMIC: yes の数式でこれらのシートを参照する場合も、列文字・行番号は必ず下記の対応表と実際の行数から算出し、推測しないこと):
{chr(10).join(blocks)}
"""

    ref_section = ""
    if current_worksheet_title:
        header_row = 1 + header_offset
        data_start_row = 2 + header_offset
        data_end_row = data_start_row + len(df) - 1 if len(df) > 0 else data_start_row
        last_col_letter = column_letter(len(df.columns) - 1)
        col_map = "\n".join(f"{column_letter(i)}列 = {col}" for i, col in enumerate(df.columns))
        ref_section = f"""
シート参照情報(DYNAMIC: yes / format_cells / insert_summary / apply_format で数式・範囲を書く場合に使うこと):
- 現在開いているシート名: {current_worksheet_title}
- ヘッダー行(見出し)はシート上の{header_row}行目、データは{data_start_row}行目から{data_end_row}行目まで(合計{len(df)}行、集計欄の挿入等により1行目からとは限らない)。列とシート上の列文字の対応:
{col_map}
- DYNAMIC: yes の数式でこのシートのデータ範囲を参照する場合は '{current_worksheet_title}'!A{data_start_row}:{last_col_letter} のように、シート名を引用符で囲み、ヘッダー行を除いた{data_start_row}行目からの範囲を使うこと。
- format_cells の数式は、後述のルールどおり常に「2行目」を基準にして書くこと(アプリ側で実際の行番号に自動的に補正するため、ここでは{data_start_row}行目ではなく2行目のままでよい)。
- insert_summary の数式は列全体参照(例: D:D)を使うので、この行番号は関係ない。
- apply_format のRANGEに範囲の指定がない場合は、A{data_start_row}:{last_col_letter}{data_end_row} のようにデータ範囲全体を使うこと(見出し行を含めない)。列単位の指示(例:「C列を」)であれば C{data_start_row}:C{data_end_row} のように該当列だけに絞ること。
"""

    return f"""列情報:
{columns_info}

先頭5行のサンプル:
{sample}
{sheet_list_section}{ref_section}{other_sheets_section}
ユーザーの指示:
{instruction}
"""


def strip_code_fence(text: str) -> str:
    """コードブロック(```python ... ```)があれば、その中身だけを抽出する。

    AIの応答に解説文が混じっていても、フェンスで囲まれた部分だけを取り出すことで
    余計なプロンプト文がコードとして実行されるのを防ぐ。フェンスが無い場合は
    そのまま返す(呼び出し側の構文検証で異常な出力を弾く)。
    """
    # 言語タグ(python / py / json など)は何が付いていても取り除く。
    match = re.search(r"```[a-zA-Z0-9_+-]*[ \t]*\n?(.*?)\n?```", text, re.DOTALL)
    return match.group(1) if match else text


def validate_python_code(code: str) -> str | None:
    """生成されたPythonコードが構文的に妥当か検証する。

    問題がなければ None、問題があればユーザー向けのエラーメッセージ文字列を返す。
    exec() する前に必ずこれを通し、不完全なコードや解説文混じりの出力が
    実行されてシートを壊すのを防ぐ。
    """
    if not code.strip():
        return "コードが空です。"
    try:
        ast.parse(code)
    except SyntaxError:
        return "適切なコードが生成できませんでした。指示文をもう少し具体的に入力して再試行してください。"
    return None


def parse_target_sheets(target_sheet: str) -> list[str]:
    return [t.strip() for t in target_sheet.split(",") if t.strip()]


def parse_plan(text: str) -> dict:
    text = text.strip()
    # 指示に反して応答全体がコードフェンスで囲まれることがあるため、先に除去する。
    if text.startswith("```"):
        text = re.sub(r"^```[a-zA-Z]*\n?", "", text)
        text = re.sub(r"\n?```\s*$", "", text)
        text = text.strip()

    lines = text.splitlines()
    idx = 0
    fields: dict[str, str] = {}
    prefixes = ("ACTION:", "TARGET_SHEET:", "MODE:", "RANGE:", "DYNAMIC:", "COLOR:", "SUMMARY:")

    while idx < len(lines) and lines[idx].strip().startswith(prefixes):
        key, _, value = lines[idx].strip().partition(":")
        fields[key.strip()] = value.strip()
        idx += 1

    code = "\n".join(lines[idx:]).strip()
    code = strip_code_fence(code).strip()

    action = fields.get("ACTION", "write_current")
    if action not in ACTIONS:
        action = "write_current"

    mode = fields.get("MODE", "full")
    if mode not in WRITE_MODES:
        mode = "full"

    color = fields.get("COLOR", "").strip() or DEFAULT_HIGHLIGHT_COLOR
    if not re.fullmatch(r"#[0-9a-fA-F]{6}", color):
        color = DEFAULT_HIGHLIGHT_COLOR

    dynamic = fields.get("DYNAMIC", "no").strip().lower() == "yes"

    code_error = None
    if action in ("write_current", "write_new") and not dynamic:
        code_error = validate_python_code(code)

    return {
        "action": action,
        "target_sheet": fields.get("TARGET_SHEET", ""),
        "mode": mode,
        "range": fields.get("RANGE", "").strip() or "A1",
        "dynamic": dynamic,
        "color": color,
        "summary": fields.get("SUMMARY", "").strip(),
        "code": code,
        "code_error": code_error,
    }


def _call_claude_for_plan(user_prompt: str, worksheet_titles: list[str] | None) -> dict:
    client = get_anthropic_client()
    try:
        message = client.messages.create(
            model=MODEL,
            max_tokens=1024,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_prompt}],
        )
    except anthropic.AuthenticationError as exc:  # 401 invalid x-api-key
        raise ApiKeyError() from exc
    except anthropic.PermissionDeniedError as exc:  # 403 (無効化されたキー等)
        raise ApiKeyError() from exc
    text = "".join(block.text for block in message.content if block.type == "text")
    plan = parse_plan(text)

    if worksheet_titles and plan["action"] == "delete_sheet":
        targets = parse_target_sheets(plan["target_sheet"])
        plan["target_sheets_resolved"] = targets
        plan["target_sheets_invalid"] = [t for t in targets if t not in worksheet_titles]

    return plan


def generate_plan(
    df: pd.DataFrame,
    instruction: str,
    worksheet_titles: list[str] | None = None,
    current_worksheet_title: str | None = None,
    header_offset: int = 0,
    other_sheets: dict[str, pd.DataFrame] | None = None,
) -> dict:
    user_prompt = build_user_prompt(
        df, instruction, worksheet_titles, current_worksheet_title, header_offset, other_sheets
    )
    return _call_claude_for_plan(user_prompt, worksheet_titles)


def generate_fix_plan(
    df: pd.DataFrame,
    instruction: str,
    failed_code: str,
    error_message: str,
    worksheet_titles: list[str] | None = None,
    current_worksheet_title: str | None = None,
    header_offset: int = 0,
    other_sheets: dict[str, pd.DataFrame] | None = None,
) -> dict:
    """直前の実行でエラーが発生した場合に、エラー内容と最新のシート状態を踏まえて
    修正した内容を1回だけ生成し直す(自己修復・リトライ機構)。
    """
    base_prompt = build_user_prompt(
        df, instruction, worksheet_titles, current_worksheet_title, header_offset, other_sheets
    )
    fix_prompt = f"""{base_prompt}

直前に生成した以下の内容を実行したところ、エラーが発生しました。エラーの原因を分析し、
修正した内容を同じ出力形式で出力してください(説明文は付けないこと)。

--- 直前に生成した内容 ---
{failed_code}

--- 発生したエラー ---
{error_message}
"""
    return _call_claude_for_plan(fix_prompt, worksheet_titles)


def looks_safe(code: str) -> bool:
    return not any(pattern in code for pattern in FORBIDDEN_PATTERNS)


def run_generated_code(code: str, df: pd.DataFrame, other_sheets: dict[str, pd.DataFrame] | None = None) -> pd.DataFrame:
    """生成されたpandasコードを実行する。

    other_sheets を渡すと `sheets` という辞書変数(シート名 -> DataFrame)として
    生成コードから参照できる(他のシートとの結合・参照を行う指示に対応するため)。
    """
    code_error = validate_python_code(code)
    if code_error:
        raise RuntimeError(code_error)

    local_vars = {"df": df, "pd": pd, "sheets": dict(other_sheets or {})}
    exec(code, {"pd": pd}, local_vars)
    result = local_vars.get("result")
    if result is None:
        # result が代入されなかった場合でも実行エラーで停止させず、
        # シート側には完了メッセージのみを書き込んでフォールバックする。
        result = pd.DataFrame({"結果": ["処理が完了しました(シートをご確認ください)。"]})
    if not isinstance(result, pd.DataFrame):
        result = pd.DataFrame({"結果": [result]})
    return result


def to_grid(result: pd.DataFrame) -> list[list]:
    return [result.columns.values.tolist()] + result.astype(object).where(result.notna(), "").values.tolist()


def snapshot_from_result(result: pd.DataFrame) -> dict:
    """write_new / redo 用: 書き込んだ内容から復元可能なスナップショットを作る。"""
    grid = to_grid(result)
    return {"values": grid, "rows": len(grid) + 10, "cols": len(result.columns) + 5}


def blank_snapshot() -> dict:
    """create_sheet の redo 用: 空シートのスナップショット。"""
    return {"values": [], "rows": 100, "cols": 20}


def apply_write(ws: gspread.Worksheet, result: pd.DataFrame, mode: str = "full", range_start: str = "A1") -> None:
    """既存シートに結果を書き込む。

    mode="full": シート全体をクリアしてから書き込む(表の並び替え・絞り込みなど、
    表全体を再構成する指示のデフォルト)。
    mode="partial": range_start のセルを起点に結果を書き込むだけで、
    それ以外の既存データはクリアしない(特定セル位置を指定した指示向け)。
    """
    if mode == "full":
        ws.clear()
        ws.update(to_grid(result))
    else:
        ws.update(to_grid(result), range_start or "A1")


def write_formula(ws: gspread.Worksheet, formula: str, mode: str = "full", range_start: str = "A1") -> None:
    """数式を書き込む(USER_ENTERED)。SORT/FILTER関数などが元データの変化に追従する
    「動的リスト」を作るために使う。値ではなく数式として解釈させる必要がある。
    """
    if mode == "full":
        ws.clear()
    ws.update([[formula]], range_start or "A1", value_input_option="USER_ENTERED")


def write_result(sh: gspread.Spreadsheet, source_ws_title: str, result: pd.DataFrame, title: str | None = None) -> str:
    """新しいシートに結果を書き込む。titleが指定されていればその名前を使う(重複時は連番を付与)。"""
    if not title:
        title = f"{source_ws_title}_result_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

    existing_titles = {w.title for w in sh.worksheets()}
    final_title = title
    suffix = 2
    while final_title in existing_titles:
        final_title = f"{title}_{suffix}"
        suffix += 1

    new_ws = sh.add_worksheet(title=final_title, rows=str(len(result) + 10), cols=str(len(result.columns) + 5))
    new_ws.update(to_grid(result))
    return final_title


def create_blank_worksheet(sh: gspread.Spreadsheet, title: str | None = None) -> str:
    """空の新規シートを作成する。"""
    if not title:
        title = f"新しいシート_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

    existing_titles = {w.title for w in sh.worksheets()}
    final_title = title
    suffix = 2
    while final_title in existing_titles:
        final_title = f"{title}_{suffix}"
        suffix += 1

    sh.add_worksheet(title=final_title, rows="100", cols="20")
    return final_title


def snapshot_worksheet(ws: gspread.Worksheet) -> dict:
    """undo用にシートの現在値・サイズを保存する。"""
    return {
        "values": ws.get_all_values(),
        "rows": ws.row_count,
        "cols": ws.col_count,
    }


def delete_worksheet_by_title(sh: gspread.Spreadsheet, title: str) -> dict:
    """シートを削除する前にスナップショットを取得し、削除後にそれを返す。"""
    ws = sh.worksheet(title)
    snapshot = snapshot_worksheet(ws)
    sh.del_worksheet(ws)
    return snapshot


def recreate_worksheet(sh: gspread.Spreadsheet, title: str, snapshot: dict) -> None:
    """delete_sheet の undo: 保存しておいたスナップショットからシートを復元する。"""
    existing_titles = {w.title for w in sh.worksheets()}
    if title in existing_titles:
        return
    rows = max(snapshot.get("rows", 1), 1)
    cols = max(snapshot.get("cols", 1), 1)
    new_ws = sh.add_worksheet(title=title, rows=str(rows), cols=str(cols))
    values = snapshot.get("values") or []
    if values:
        new_ws.update(values)


def delete_worksheet_if_exists(sh: gspread.Spreadsheet, title: str) -> None:
    """write_new / create_sheet の undo: 作成したシートを削除する。"""
    try:
        ws = sh.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        return
    sh.del_worksheet(ws)


# Google Sheets API の CellFormat 内で「色オブジェクト」が入る位置のキー名。
# ここに16進数カラーコードの文字列がそのまま入っていると API が 400 を返すため、
# normalize_format_colors() で RGB オブジェクトへ変換する。
COLOR_FIELD_KEYS = ("backgroundColor", "foregroundColor", "color")

_HEX_COLOR_RE = re.compile(r"^#?(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6})$")


def is_hex_color(value) -> bool:
    """`#000000` `#fff` `FFFFFF` のような16進数カラーコード文字列かどうか。"""
    return isinstance(value, str) and bool(_HEX_COLOR_RE.match(value.strip()))


def hex_to_rgb(hex_color: str) -> dict:
    """`#000000` 形式のカラーコードを Sheets API の RGB オブジェクトへ変換する。

    >>> hex_to_rgb("#000000")
    {'red': 0.0, 'green': 0.0, 'blue': 0.0}

    `#fff` のような3桁表記、`#` なし、大文字小文字の混在にも対応する。
    変換できない値を渡した場合は ValueError を送出する。
    """
    if not isinstance(hex_color, str):
        raise ValueError(f"カラーコードは文字列で指定してください: {hex_color!r}")
    value = hex_color.strip().lstrip("#")
    if len(value) == 3:  # #fff -> #ffffff
        value = "".join(ch * 2 for ch in value)
    if len(value) != 6:
        raise ValueError(f"カラーコードの形式が不正です: {hex_color!r}")
    try:
        channels = [int(value[i:i + 2], 16) / 255 for i in (0, 2, 4)]
    except ValueError as exc:
        raise ValueError(f"カラーコードの形式が不正です: {hex_color!r}") from exc
    return {"red": channels[0], "green": channels[1], "blue": channels[2]}


def rgb_to_hex(rgb: dict) -> str:
    """Sheets API の RGB オブジェクトを `#RRGGBB` 形式のカラーコードへ逆変換する。

    >>> rgb_to_hex({"red": 0.0, "green": 0.0, "blue": 0.0})
    '#000000'

    Sheets API は値が 0 のチャンネルをレスポンスから省略することがあるため、
    キーが欠けている場合は 0 とみなす。
    """
    if not isinstance(rgb, dict):
        raise ValueError(f"RGBオブジェクトは辞書で指定してください: {rgb!r}")
    parts = []
    for key in ("red", "green", "blue"):
        channel = rgb.get(key, 0) or 0
        parts.append(max(0, min(255, round(float(channel) * 255))))
    return "#{:02X}{:02X}{:02X}".format(*parts)


def normalize_color(value) -> dict:
    """カラー指定(16進数文字列 / RGBオブジェクト)を RGB オブジェクトに揃える。"""
    if isinstance(value, dict):
        return value
    return hex_to_rgb(value)


def normalize_format_colors(spec):
    """書式スペックを再帰的に走査し、色の位置にある16進数文字列を RGB オブジェクトへ変換する。

    Sheets API へ渡す直前にこれを通すことで、`{"backgroundColor": "#000000"}` の
    ような文字列指定が API に到達して 400 エラーになるのを防ぐ。
    """
    if isinstance(spec, list):
        return [normalize_format_colors(item) for item in spec]
    if not isinstance(spec, dict):
        return spec

    normalized = {}
    for key, value in spec.items():
        if key in COLOR_FIELD_KEYS and is_hex_color(value):
            normalized[key] = hex_to_rgb(value)
        else:
            normalized[key] = normalize_format_colors(value)
    return normalized


# 旧名(内部利用箇所との互換のため残す)
_hex_to_rgb = hex_to_rgb


def data_range_a1(ws: gspread.Worksheet, header_offset: int = 0) -> str:
    """ヘッダー行を除いたシート上のデータ範囲を返す。

    キャッシュ済みのDataFrameではなく、呼び出し時点で ws.get_all_values() を
    実行して最新の値を取得してから範囲を計算する(他の操作でシートが変化していても
    ずれた位置に書き込んでしまわないようにするため)。

    header_offset: シート最上部に挿入済みの集計行数(insert_summaryで増える)。
    ヘッダー行は 1+header_offset 行目、データは 2+header_offset 行目から始まる。
    """
    values = ws.get_all_values()
    first_row = 2 + header_offset
    if not values or len(values) < first_row:
        return f"A{first_row}:A{first_row}"
    ncols = max((len(row) for row in values), default=1)
    last_col = column_letter(max(ncols - 1, 0))
    last_row = len(values)
    return f"A{first_row}:{last_col}{last_row}"


def find_empty_range_start(ws: gspread.Worksheet, header_offset: int = 0) -> str:
    """具体的なセル位置の指定がない「別のセルに」「空いているところに」等の指示向けに、
    既存データと重ならない空きセルの開始位置(既存データの右隣の列)を自動的に探す。

    呼び出し時点で ws.get_all_values() を実行し、最新の状態を基準に計算する。
    """
    values = ws.get_all_values()
    start_row = 1 + header_offset
    if not values:
        return f"A{start_row}"
    max_col = max((len(row) for row in values), default=0)
    return f"{column_letter(max_col)}{start_row}"


CURRENCY_KEYWORDS = ("単価", "金額", "売上", "合計", "価格", "料金", "コスト", "予算", "円")


def _is_numeric_series(series: pd.Series) -> bool:
    if series.empty:
        return False
    cleaned = series.astype(str).str.strip().str.replace(",", "", regex=False).str.replace("¥", "", regex=False)
    cleaned = cleaned[cleaned != ""]
    if cleaned.empty:
        return False
    try:
        pd.to_numeric(cleaned)
        return True
    except (ValueError, TypeError):
        return False


def infer_column_formats(result: pd.DataFrame) -> dict[int, dict]:
    """列名・列の値から、列ごとに適用すべき既定の書式(通貨/数値は右揃え、文字列は左揃え)を推定する。

    戻り値は 0始まりの列インデックス -> gspread互換の書式辞書。
    列ズレを防ぐため、必ず result の列順・列数と1対1で対応させて使うこと。
    """
    formats: dict[int, dict] = {}
    for i, col in enumerate(result.columns):
        series = result.iloc[:, i]
        is_numeric = _is_numeric_series(series)
        if is_numeric and any(kw in str(col) for kw in CURRENCY_KEYWORDS):
            formats[i] = {"numberFormat": {"type": "CURRENCY", "pattern": "¥#,##0"}, "horizontalAlignment": "RIGHT"}
        elif is_numeric:
            formats[i] = {"horizontalAlignment": "RIGHT"}
        else:
            formats[i] = {"horizontalAlignment": "LEFT"}
    return formats


def apply_default_formatting(ws: gspread.Worksheet, result: pd.DataFrame, start_row: int = 1, start_col: int = 0) -> None:
    """データ書き込み後に、列の内容(数値/金額らしき列名かどうか)に応じた既定の書式
    (通貨列は右揃え+¥#,##0、その他数値列は右揃え、文字列列は左揃え)と、
    見出し行の太字を自動的に適用する。

    書式は result の列インデックスと1対1で対応させて範囲を計算するため、
    書き込んだデータと見た目上の列がズレることはない。
    """
    if result.empty or len(result.columns) == 0:
        return
    ncols = len(result.columns)
    header_row = start_row
    data_start = start_row + 1
    data_end = start_row + len(result)
    col_formats = infer_column_formats(result)

    for i in range(ncols):
        fmt = col_formats.get(i)
        if not fmt:
            continue
        col_letter_ = column_letter(start_col + i)
        ws.format(f"{col_letter_}{data_start}:{col_letter_}{data_end}", fmt)

    header_start = column_letter(start_col)
    header_end = column_letter(start_col + ncols - 1)
    ws.format(f"{header_start}{header_row}:{header_end}{header_row}", {"textFormat": {"bold": True}})


def parse_a1_start(range_start: str) -> tuple[int, int]:
    """A1形式の開始セル(例: "C5")を (0始まり行, 0始まり列) に変換する。"""
    row, col = gspread.utils.a1_to_rowcol(range_start or "A1")
    return row - 1, col - 1


def shift_formula_row_anchor(formula: str, anchor_row: int, base_row: int = 2) -> str:
    """AIが生成した「2行目基準」の数式を、実際のヘッダー位置に応じた行番号にずらす。

    insert_summary でシート最上部に集計行が挿入されている場合、実データは
    2行目より下から始まるため、`$C2` のような列参照の行番号を補正する必要がある。
    """
    if anchor_row == base_row:
        return formula
    return re.sub(rf"(\$[A-Za-z]+){base_row}(?!\d)", rf"\g<1>{anchor_row}", formula)


def add_conditional_format(
    sh: gspread.Spreadsheet, ws: gspread.Worksheet, range_a1: str, formula: str, color_hex: str = DEFAULT_HIGHLIGHT_COLOR
) -> dict:
    """条件付き書式ルールを追加する(元データが変化しても自動的に再判定される)。

    常にルールをインデックス0に追加する。undo/redoはLIFOのため、
    履歴の一番上に積まれた操作(=直前に追加したルール)は常にインデックス0にあることが保証される。
    """
    grid_range = gspread.utils.a1_range_to_grid_range(range_a1, ws.id)
    body = {
        "requests": [{
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [grid_range],
                    "booleanRule": {
                        "condition": {
                            "type": "CUSTOM_FORMULA",
                            "values": [{"userEnteredValue": formula}],
                        },
                        "format": {"backgroundColor": normalize_color(color_hex)},
                    },
                },
                "index": 0,
            }
        }]
    }
    sh.batch_update(body)
    return {"sheet_id": ws.id, "range": range_a1, "formula": formula, "color": color_hex}


def delete_conditional_format(sh: gspread.Spreadsheet, sheet_id: int, index: int = 0) -> None:
    """add_conditional_format の undo: 指定インデックスのルールを削除する。"""
    body = {"requests": [{"deleteConditionalFormatRule": {"sheetId": sheet_id, "index": index}}]}
    sh.batch_update(body)


def parse_summary_rows(code: str) -> list[tuple[str, str]]:
    """`ラベル|数式` 形式の行を (ラベル, 数式) のリストに変換する。"""
    rows = []
    for line in code.splitlines():
        line = line.strip()
        if not line or "|" not in line:
            continue
        label, _, formula = line.partition("|")
        label, formula = label.strip(), formula.strip()
        if label and formula:
            rows.append((label, formula))
    return rows


def insert_summary_block(ws: gspread.Worksheet, rows: list[tuple[str, str]]) -> None:
    """シート最上部に集計行(ラベル+数式)を挿入し、既存データを下に押し下げる。"""
    if not rows:
        return
    grid = [[label, formula] for label, formula in rows]
    ws.insert_rows(grid, row=1, value_input_option="USER_ENTERED")


def remove_summary_block(ws: gspread.Worksheet, row_count: int) -> None:
    """insert_summary_block の undo: 挿入した先頭 row_count 行を削除する。"""
    if row_count <= 0:
        return
    ws.delete_rows(1, row_count)


def _border_spec(color_hex: str | None = None) -> dict:
    border_style = {
        "style": "SOLID",
        "width": 1,
        "color": hex_to_rgb(color_hex) if color_hex else {"red": 0, "green": 0, "blue": 0},
    }
    return {side: dict(border_style) for side in ("top", "bottom", "left", "right")}


def _apply_format_key(spec: dict, key: str, value) -> None:
    """1つの書式キーを CellFormat 形式の辞書 `spec` に反映する(未知のキーは無視)。"""
    key = key.strip().upper()
    if isinstance(value, str):
        value = value.strip()
    if value is None or value == "":
        return

    truthy = value is True or (isinstance(value, str) and value.lower() == "true")
    falsy = value is False or (isinstance(value, str) and value.lower() == "false")

    if key == "NUMBER_FORMAT":
        spec["numberFormat"] = {"type": "NUMBER", "pattern": value}
    elif key == "BACKGROUND_COLOR":
        spec["backgroundColor"] = normalize_color(value)
    elif key == "TEXT_COLOR":
        spec.setdefault("textFormat", {})["foregroundColor"] = normalize_color(value)
    elif key == "BOLD" and (truthy or falsy):
        spec.setdefault("textFormat", {})["bold"] = truthy
    elif key == "ITALIC" and (truthy or falsy):
        spec.setdefault("textFormat", {})["italic"] = truthy
    elif key == "FONT_SIZE":
        spec.setdefault("textFormat", {})["fontSize"] = int(value)
    elif key == "HORIZONTAL_ALIGNMENT" and str(value).upper() in ("LEFT", "CENTER", "RIGHT"):
        spec["horizontalAlignment"] = str(value).upper()
    elif key == "BORDER" and truthy:
        spec["borders"] = _border_spec()
    elif key == "BORDER_COLOR":
        spec["borders"] = _border_spec(value if is_hex_color(value) else None)


def parse_format_spec(code: str) -> dict:
    """apply_format の `キー: 値` 形式の行を、gspread の format() に渡せる辞書に変換する。

    複数操作(JSON配列)には parse_format_ops() を使うこと。この関数は
    単一の書式ブロックを解釈する下位関数として残している。
    """
    spec: dict = {}
    for line in code.splitlines():
        line = line.strip()
        if not line or ":" not in line:
            continue
        key, _, value = line.partition(":")
        try:
            _apply_format_key(spec, key, value)
        except (ValueError, TypeError):
            # 1つのキーの値が不正でも、他のキーの適用は続ける。
            continue
    return spec


def _op_from_mapping(item: dict) -> dict:
    """JSONの1操作(辞書)を {"range": ..., "format": CellFormat} に変換する。"""
    spec: dict = {}
    range_a1 = ""
    for key, value in item.items():
        if key.strip().upper() == "RANGE":
            range_a1 = str(value).strip()
            continue
        try:
            _apply_format_key(spec, key, value)
        except (ValueError, TypeError):
            continue
    return {"range": range_a1, "format": spec}


def parse_format_ops(code: str, default_range: str = "") -> list[dict]:
    """apply_format の内容を「複数の書式操作のリスト」として解釈する。

    受け付ける形式:
    1. JSON配列(推奨。複数操作をまとめて指示する場合):
       `[{"range": "A1:G1", "background_color": "#CFE2F3", "bold": true}, ...]`
    2. JSONオブジェクト単体: `{"range": "A1:G1", "bold": true}`
    3. 従来の `キー: 値` 行形式(1操作のみ)。`RANGE:` 行が現れるたびに次の操作として扱う。

    range が空の操作には default_range を補う。書式が1つも解釈できなかった操作は捨てる。
    """
    text = strip_code_fence(code).strip()
    ops: list[dict] = []

    if text.startswith("[") or text.startswith("{"):
        try:
            parsed = json.loads(text)
        except json.JSONDecodeError:
            parsed = None
        if isinstance(parsed, dict):
            parsed = [parsed]
        if isinstance(parsed, list):
            ops = [_op_from_mapping(item) for item in parsed if isinstance(item, dict)]

    if not ops:
        # 従来の行形式。RANGE行を区切りとして複数ブロックに対応する。
        current: dict = {"range": "", "format": {}}
        for line in text.splitlines():
            line = line.strip()
            if not line or ":" not in line:
                continue
            key, _, value = line.partition(":")
            if key.strip().upper() == "RANGE":
                if current["format"]:
                    ops.append(current)
                current = {"range": value.strip(), "format": {}}
                continue
            try:
                _apply_format_key(current["format"], key, value)
            except (ValueError, TypeError):
                continue
        if current["format"]:
            ops.append(current)

    result = []
    for op in ops:
        if not op["format"]:
            continue
        op["range"] = op["range"] or default_range
        if not op["range"]:
            continue
        op["format"] = normalize_format_colors(op["format"])
        result.append(op)
    return result


def apply_format_ops(ws: gspread.Worksheet, ops: list[dict]) -> list[dict]:
    """複数の書式操作を1回の batchUpdate でまとめて適用する。

    まず全操作を requests リストにまとめて一括送信し(gspread の batch_format が
    repeatCell リクエストの配列を1回の batchUpdate で送る)、それが失敗した場合は
    1操作ずつ個別に適用し直して、どれが成功・失敗したかを切り分ける。

    戻り値は各操作の結果リスト:
    `[{"range": "A1:G1", "ok": True, "error": None}, ...]`
    1つ失敗しても例外は送出しない(呼び出し側が結果を見て表示する)。
    """
    if not ops:
        return []

    payload = [{"range": op["range"], "format": op["format"]} for op in ops]
    try:
        ws.batch_format(payload)
    except Exception as batch_exc:  # noqa: BLE001
        # 一括送信は1件でも不正があると全体が失敗するため、個別適用に切り替える。
        results = []
        for op in ops:
            try:
                ws.format(op["range"], op["format"])
            except Exception as exc:  # noqa: BLE001
                results.append({
                    "range": op["range"],
                    "ok": False,
                    "error": f"{type(exc).__name__}: {exc}",
                })
            else:
                results.append({"range": op["range"], "ok": True, "error": None})
        if all(not r["ok"] for r in results):
            # 全滅した場合は一括送信時のエラーの方が原因を表していることが多い。
            for r in results:
                r["error"] = r["error"] or str(batch_exc)
        return results

    return [{"range": op["range"], "ok": True, "error": None} for op in ops]


def describe_format_spec(spec: dict) -> str:
    """書式スペックを確認画面表示用の日本語1行に要約する。"""
    parts = []
    if "backgroundColor" in spec:
        parts.append(f"背景色 {rgb_to_hex(spec['backgroundColor'])}")
    text_format = spec.get("textFormat", {})
    if "foregroundColor" in text_format:
        parts.append(f"文字色 {rgb_to_hex(text_format['foregroundColor'])}")
    if text_format.get("bold"):
        parts.append("太字")
    if text_format.get("italic"):
        parts.append("斜体")
    if "fontSize" in text_format:
        parts.append(f"文字サイズ {text_format['fontSize']}")
    if "numberFormat" in spec:
        parts.append(f"表示形式 {spec['numberFormat'].get('pattern', '')}")
    if "horizontalAlignment" in spec:
        parts.append(f"横位置 {spec['horizontalAlignment']}")
    if "borders" in spec:
        parts.append("罫線")
    return " / ".join(parts) if parts else "(書式なし)"


def get_cell_format(sh: gspread.Spreadsheet, ws: gspread.Worksheet, cell_a1: str) -> dict:
    """指定セル1つの現在の書式を取得する(apply_formatのundo用に、範囲内の代表的な既存書式として使う)。"""
    try:
        meta = sh.fetch_sheet_metadata(params={
            "ranges": [f"'{ws.title}'!{cell_a1}"],
            "fields": "sheets.data.rowData.values.userEnteredFormat",
        })
        row_data = meta["sheets"][0]["data"][0].get("rowData", [])
        if row_data and row_data[0].get("values"):
            return row_data[0]["values"][0].get("userEnteredFormat", {})
    except (KeyError, IndexError):
        pass
    return {}


def apply_cell_format(ws: gspread.Worksheet, range_a1: str, format_spec: dict) -> None:
    """指定範囲に書式(数値の表示形式・背景色・太字など)を直接適用する。

    update() によるセルの値の書き込みとは別に、gspreadのformat()メソッドを
    使って見た目だけを変更する(条件付きではなく常に適用される一律の書式)。
    """
    if not format_spec:
        return
    # 呼び出し側やAIの出力に `"backgroundColor": "#000000"` のような
    # 文字列指定が混ざっていても、ここで RGB オブジェクトへ変換してから送る。
    ws.format(range_a1, normalize_format_colors(format_spec))
