# app_task_memo.py
# -*- coding: utf-8 -*-

"""
作業メモ用デスクトップアプリ（Windows 11 / Python 3.12）
- tkinterベース
- ログ日付の区切りは 04:00 AM
- UIサイズは通常/4倍のトグル
- 10レーンの (TAG / SUB_TAG / テキスト / 音声入力)
- LLM質問エリア（表示切替、OpenAI RESTを標準ライブラリで直接叩く）
- 設定/ログパスの明示、プレビュー付き
"""

# 1. import文（必要最小限 + 標準ライブラリ中心）
import sys
import os
import io
import json
import time
import threading
import traceback
import urllib.request
import urllib.error
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from typing import Any, Callable, Dict, List, Optional, Tuple

# Windows/COM（音声入力用；利用可能な場合のみ）
try:
    import win32com.client  # noqa: WPS433 - 許可モジュール
    import pythoncom  # noqa: WPS433 - pywin32の一部
    HAS_WIN32 = True
except Exception:
    HAS_WIN32 = False

# GUI
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import tkinter.font as tkfont


# 2. 定数定義（繰り返し登場する固定文字列や数値はここで定義）
APP_BASENAME: str = "app_task_memo"
APP_FILENAME: str = f"{APP_BASENAME}.py"
CFG_FILENAME: str = f"{APP_BASENAME}.yaml"

# ログ日付の区切り時刻（04:00）
CUTOFF_HOUR: int = 4

# フォント基準（固定サイズ運用）
BASE_FONT_FAMILY: str = "Yu Gothic UI"
BASE_FONT_SIZE: int = 10  # 通常時
BASE_FONT_SIZE_LARGE: int = 12  # サイズ変更時
MONO_FONT_FAMILY: str = "Consolas"
MONO_FONT_SIZE: int = 10
MONO_FONT_SIZE_LARGE: int = 12

# UIスケール（未使用：固定フォントで対応）

# レーン数
LANE_COUNT: int = 10

# LLM既定
DEFAULT_LLM_PROVIDER: str = "openai"
DEFAULT_LLM_MODEL: str = "gpt-4o-mini"
DEFAULT_LLM_MAX_TOKENS: int = 2000
DEFAULT_LLM_TIMEOUT: int = 60
DEFAULT_OPENAI_KEY_ENV: str = "OPENAI_API_KEY"

# テキストの角括弧付き最小幅（例: "[TAG____]"）
BRACKET_MIN_WIDTH: int = 8

# JST
TZ_JST = ZoneInfo("Asia/Tokyo")

# 曜日3文字
DOW3 = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]

# メニュー等の文言
TITLE_TEXT = "Task Memo（作業メモ）"
BTN_SCALE_TEXT = "サイズ変更"
BTN_REGISTER_TEXT = "登録"
LLM_TOGGLE_TEXT = "LLMエリアを表示"
LLM_SEND_TEXT = "送信"
VOICE_BTN_TEXT = "🎤"

# 例外処理用：出力先ラベル
ERROR_DIALOG_TITLE = "エラー"
INFO_DIALOG_TITLE = "情報"


# 3. 関数／クラス定義
def get_exception_trace() -> str:
    """例外のトレースバックを取得"""
    return traceback.format_exc()


def get_base_dir() -> str:
    """
    実行ファイル（py/exe）と同じディレクトリを返す。

    Examples:
        >>> base = get_base_dir()
        >>> os.path.exists(base)
        True
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def effective_date(now: Optional[datetime] = None) -> datetime:
    """
    04:00 を境に「その日付」を返す（時刻は00:00固定ではなく、返り値は同日のnow基準）。
    ログファイル名に使うのは日付部分のみ。

    Examples:
        >>> dt = datetime(2025, 8, 30, 3, 59, tzinfo=TZ_JST)
        >>> effective_date(dt).date().isoformat()
        '2025-08-29'
        >>> dt2 = datetime(2025, 8, 30, 4, 0, tzinfo=TZ_JST)
        >>> effective_date(dt2).date().isoformat()
        '2025-08-30'
    """
    if now is None:
        now = datetime.now(TZ_JST)
    if now.hour < CUTOFF_HOUR:
        return now - timedelta(days=1)
    return now


def bracket_pad(text: str, min_total: int = BRACKET_MIN_WIDTH) -> str:
    """
    角括弧込みで指定幅を確保する。

    Examples:
        >>> bracket_pad("A", 8)
        '[A      ]'
        >>> bracket_pad("ABCDEFGH", 8)
        '[ABCDEFGH]'
    """
    inside_width = max(0, min_total - 2)
    inner = f"{text:<{inside_width}}"
    if len(text) > inside_width:
        # はみ出す場合はそのまま全体長が広がる
        inner = text
    return f"[{inner}]"


def now_jst() -> datetime:
    """JSTの現在日時を返す。"""
    return datetime.now(TZ_JST)


def format_timestamp(dt: Optional[datetime] = None) -> str:
    """
    ログ用のタイムスタンプ "YYYY-MM-DD HH:MM" を返す。

    Examples:
        >>> s = format_timestamp(datetime(2025,8,30,16,11,tzinfo=TZ_JST))
        >>> s
        '2025-08-30 16:11'
    """
    if dt is None:
        dt = now_jst()
    return dt.strftime("%Y-%m-%d %H:%M")


def yyyymmdd_dow(dt: Optional[datetime] = None) -> str:
    """
    "YYYYMMDD_DOW" 形式を返す（曜日は英3文字固定）。
    ロケールに依存しないため、独自テーブルを利用。

    Examples:
        >>> yyyymmdd_dow(datetime(2025,8,30,tzinfo=TZ_JST))
        '20250830_SAT'
    """
    if dt is None:
        dt = now_jst()
    ymd = dt.strftime("%Y%m%d")
    dow = DOW3[dt.weekday()]
    return f"{ymd}_{dow}"


class SimpleYAML:
    """
    依存ライブラリ無しで扱える非常に単純化したYAMLパーサ/ダンパ。
    - 本アプリの想定設定（KEY:SCALAR / KEY:LIST / KEY:DICT with LIST/SCALAR）に限定。
    - 完全なYAML互換ではない点に留意。

    Examples:
        >>> text = "A: 1\\nB:\\n  - x\\n  - y\\nC:\\n  K: V\\n"
        >>> data = SimpleYAML.loads(text)
        >>> data["A"], data["B"], data["C"]["K"]
        ('1', ['x', 'y'], 'V')
    """

    @staticmethod
    def strip_quotes(s: str) -> str:
        s = s.strip()
        if (s.startswith('"') and s.endswith('"')) or (s.startswith("'") and s.endswith("'")):
            return s[1:-1]
        return s

    @staticmethod
    def loads(text: str) -> Dict[str, Any]:
        lines = text.replace("\t", "  ").splitlines()
        root: Dict[str, Any] = {}
        # stack要素: (indent, container, parent_container, key_in_parent)
        stack: List[Tuple[int, Any, Optional[Dict[str, Any]], Optional[str]]] = [(-1, root, None, None)]

        for raw in lines:
            if not raw.strip() or raw.strip().startswith("#"):
                continue
            indent = len(raw) - len(raw.lstrip(" "))
            line = raw.strip()

            # リスト項目
            if line.startswith("- "):
                value = SimpleYAML.strip_quotes(line[2:])
                # 現在インデント以上のものはスコープ外としてpop
                while stack and indent <= stack[-1][0]:
                    stack.pop()
                if not stack:
                    raise ValueError("Invalid YAML structure near list item.")

                cur_indent, cur_container, cur_parent, cur_key = stack[-1]
                # 直前が dict コンテナ（空の入れ物）なら、ここで list に置換
                if isinstance(cur_container, dict):
                    if cur_parent is not None and isinstance(cur_parent, dict) and cur_key is not None:
                        new_list: List[Any] = []
                        cur_parent[cur_key] = new_list
                        stack[-1] = (cur_indent, new_list, cur_parent, cur_key)
                    else:
                        raise ValueError("Invalid YAML structure near list item.")

                # 以降は必ずリストであること
                if not isinstance(stack[-1][1], list):
                    raise ValueError("Invalid YAML structure near list item.")
                stack[-1][1].append(value)
                continue

            # key: value or key:
            if ":" in line:
                key, sep, rest = line.partition(":")
                key = SimpleYAML.strip_quotes(key)
                value = rest.strip()

                # 適切な親を探す
                while stack and indent <= stack[-1][0]:
                    stack.pop()
                parent = stack[-1][1]

                if value == "":
                    # ネスト開始（dict or list）
                    # 次の行の先頭が "- " ならリスト、それ以外はdictと仮定
                    # ここでは一旦dictを作り、実際に次行で "- " が来たらリストに入替
                    container: Any = {}
                    parent[key] = container
                    stack.append((indent, container, parent, key))
                else:
                    parent[key] = SimpleYAML.strip_quotes(value)
            else:
                # 形式外は無視（本アプリの想定では到達しない）
                continue

            # 直後がリストであるケースのため、空dictのまま "- " を受けたらリストに置換
            # これは行単位では判定不可のため、実装簡略化のため次行処理時に判定。
            # 実運用では "KEY:\n  - item" で正しく扱える。

            # 後続で "- " を受け取ったときに上書きできるよう、容器がdictか確認。
            if isinstance(stack[-1][1], dict):
                # もし次行がリストなら、その時点で置換される（上のリスト処理でpopされる）
                pass

        # 辞書の値のうち、空dictで終わった箇所を空辞書のまま採用
        # KEY: の後に何も続かない場合は空dict扱い
        # （本アプリ設定では該当しない前提）
        return root

    @staticmethod
    def dumps(data: Dict[str, Any]) -> str:
        def dump_obj(obj: Any, indent: int = 0) -> str:
            sp = " " * indent
            if isinstance(obj, dict):
                out = []
                for k, v in obj.items():
                    if isinstance(v, (dict, list)):
                        out.append(f"{sp}{k}:")
                        out.append(dump_obj(v, indent + 2))
                    else:
                        out.append(f"{sp}{k}: {SimpleYAML._scalar(v)}")
                return "\n".join(out)
            if isinstance(obj, list):
                out = []
                for item in obj:
                    if isinstance(item, (dict, list)):
                        out.append(f"{sp}-")
                        out.append(dump_obj(item, indent + 2))
                    else:
                        out.append(f"{sp}- {SimpleYAML._scalar(item)}")
                return "\n".join(out)
            return f"{sp}{SimpleYAML._scalar(obj)}"

        return dump_obj(data)

    @staticmethod
    def _scalar(v: Any) -> str:
        s = str(v)
        if any(ch in s for ch in [":", "#", "-", '"', "'"]) or s.strip() != s:
            return json.dumps(s, ensure_ascii=False)
        return s


class ConfigManager:
    """
    設定ファイル(.yaml)のロード/初期化を司る。

    - 既定の雛形を自動生成（初回起動時）
    - 必要キーが無ければ既定値で補完
    - 参照用のプロパティを提供

    Examples:
        >>> cm = ConfigManager()
        >>> isinstance(cm.config, dict)
        True
    """

    def __init__(self) -> None:
        self.base_dir: str = get_base_dir()
        self.cfg_path: str = os.path.join(self.base_dir, CFG_FILENAME)
        self.config: Dict[str, Any] = {}
        self.ensure_config()

    def ensure_config(self) -> None:
        if not os.path.exists(self.cfg_path):
            default_cfg = self._default_config()
            with io.open(self.cfg_path, "w", encoding="utf-8") as f:
                f.write(SimpleYAML.dumps(default_cfg))
        self.load()

    def load(self) -> None:
        with io.open(self.cfg_path, "r", encoding="utf-8") as f:
            text = f.read()
        data = SimpleYAML.loads(text)

        # 文字列/数値の既定補完
        data.setdefault("LOG_DIR", os.path.join(self.base_dir, "logs"))
        data.setdefault("USER_NAME", "User")
        data.setdefault("MAIN_TAG", [])  # 任意割当
        data.setdefault("TAGS", {"Work": ["Coding", "Meeting"], "Life": ["Family", "Health"]})
        data.setdefault("BASE_PROMPTS", {"要約": "次の文章を要約してください。", "アイデア": "テーマからアイデアを列挙してください。"})
        data.setdefault("LLM_PROVIDER", DEFAULT_LLM_PROVIDER)
        data.setdefault("LLM_MODEL", DEFAULT_LLM_MODEL)
        data.setdefault("LLM_MAX_COMPLETION_TOKENS", str(DEFAULT_LLM_MAX_TOKENS))
        data.setdefault("LLM_TIMEOUT", str(DEFAULT_LLM_TIMEOUT))
        data.setdefault("OPENAI_API_KEY_ENV", DEFAULT_OPENAI_KEY_ENV)
        data.setdefault("OPENAI_API_KEY", "")

        # 数値系はint化（SimpleYAMLは文字列で返すため）
        try:
            data["LLM_MAX_COMPLETION_TOKENS"] = int(data["LLM_MAX_COMPLETION_TOKENS"])
        except Exception:
            data["LLM_MAX_COMPLETION_TOKENS"] = DEFAULT_LLM_MAX_TOKENS
        try:
            data["LLM_TIMEOUT"] = int(data["LLM_TIMEOUT"])
        except Exception:
            data["LLM_TIMEOUT"] = DEFAULT_LLM_TIMEOUT

        # ログディレクトリの解決
        # - 空（または不正型）の場合は実行ファイルと同じ階層にする
        # - 相対パスなら実行ファイルディレクトリ基準で解決
        raw_log_dir = data.get("LOG_DIR", "")
        if isinstance(raw_log_dir, str):
            raw_log_dir = raw_log_dir.strip()
            if raw_log_dir == "":
                resolved_log_dir = self.base_dir
            else:
                expanded = os.path.expandvars(os.path.expanduser(raw_log_dir))
                resolved_log_dir = expanded if os.path.isabs(expanded) else os.path.join(self.base_dir, expanded)
        else:
            resolved_log_dir = self.base_dir

        data["LOG_DIR"] = resolved_log_dir

        # ログディレクトリの作成
        os.makedirs(resolved_log_dir, exist_ok=True)

        self.config = data

    def _default_config(self) -> Dict[str, Any]:
        return {
            "LOG_DIR": os.path.join(get_base_dir(), "logs"),
            "USER_NAME": "User",
            "MAIN_TAG": [],
            "TAGS": {
                "Plan": ["Today", "ThisWeek", "Later"],
                "Work": ["Coding", "Meeting", "Review"],
                "Learn": ["Read", "Watch", "Practice"],
                "Life": ["Family", "Health", "Money"],
                "Idea": ["Product", "Note"],
            },
            "BASE_PROMPTS": {
                "要約": "次の文章を要約してください。",
                "アイデア": "次のテーマからアイデアを10個列挙してください。",
                "振り返り": "本日の作業ログから良かった点・改善点を抽出してください。",
            },
            "LLM_PROVIDER": "openai",
            "LLM_MODEL": "gpt-4o-mini",
            "LLM_MAX_COMPLETION_TOKENS": DEFAULT_LLM_MAX_TOKENS,
            "LLM_TIMEOUT": DEFAULT_LLM_TIMEOUT,
            "OPENAI_API_KEY_ENV": "OPENAI_API_KEY",
            "OPENAI_API_KEY": "",
        }

    # プロパティ系のヘルパ
    @property
    def tags_map(self) -> Dict[str, List[str]]:
        tags = self.config.get("TAGS", {})
        # dict順で使う（3.7+は保持される）
        return tags

    @property
    def main_tags(self) -> List[str]:
        return list(self.config.get("MAIN_TAG", []))[:LANE_COUNT]

    @property
    def base_prompts(self) -> Dict[str, str]:
        return self.config.get("BASE_PROMPTS", {})

    @property
    def user_name(self) -> str:
        return self.config.get("USER_NAME", "User")

    @property
    def log_dir(self) -> str:
        return self.config.get("LOG_DIR", os.path.join(self.base_dir, "logs"))

    # LLM設定
    @property
    def llm_provider(self) -> str:
        return self.config.get("LLM_PROVIDER", DEFAULT_LLM_PROVIDER)

    @property
    def llm_model(self) -> str:
        return self.config.get("LLM_MODEL", DEFAULT_LLM_MODEL)

    @property
    def llm_max_tokens(self) -> int:
        return int(self.config.get("LLM_MAX_COMPLETION_TOKENS", DEFAULT_LLM_MAX_TOKENS))

    @property
    def llm_timeout(self) -> int:
        return int(self.config.get("LLM_TIMEOUT", DEFAULT_LLM_TIMEOUT))

    @property
    def openai_api_key(self) -> str:
        key = self.config.get("OPENAI_API_KEY", "").strip()
        if key:
            return key
        env_name = self.config.get("OPENAI_API_KEY_ENV", DEFAULT_OPENAI_KEY_ENV)
        return os.environ.get(env_name, "").strip()


class LogManager:
    """
    ログファイルの管理（パス決定、作成、読み書き）。

    Examples:
        >>> lm = LogManager(ConfigManager())
        >>> path = lm.current_log_path
        >>> path.endswith(".log")
        True
    """

    def __init__(self, cfg: ConfigManager) -> None:
        self.cfg = cfg
        self._current_effective_date: datetime = effective_date()
        self._current_log_path: str = self._build_log_path(self._current_effective_date)
        self.ensure_logfile()

    def _build_log_path(self, dt: datetime) -> str:
        return os.path.join(
            self.cfg.log_dir,
            f"{APP_BASENAME}_{yyyymmdd_dow(dt)}.log",
        )

    @property
    def current_log_path(self) -> str:
        return self._current_log_path

    def rollover_if_needed(self) -> bool:
        """
        04:00境界越え等で日付が変わったらパスを更新し、Trueを返す。
        """
        eff = effective_date()
        if eff.date() != self._current_effective_date.date():
            self._current_effective_date = eff
            self._current_log_path = self._build_log_path(eff)
            self.ensure_logfile()
            return True
        return False

    def ensure_logfile(self) -> None:
        if not os.path.exists(self._current_log_path):
            with io.open(self._current_log_path, "w", encoding="utf-8_sig") as f:
                f.write(f"# {TITLE_TEXT} / {yyyymmdd_dow(self._current_effective_date)}\n")

    def append_lines(self, lines: List[str]) -> None:
        if not lines:
            return
        with io.open(self._current_log_path, "a", encoding="utf-8_sig", newline="\n") as f:
            for line in lines:
                f.write(line.rstrip("\n") + "\n")

    def read_text(self, tail_kb: int = 512) -> str:
        """
        ログの末尾を中心に軽量読み込み。
        """
        path = self._current_log_path
        try:
            size = os.path.getsize(path)
            start = max(0, size - tail_kb * 1024)
            with io.open(path, "r", encoding="utf-8_sig", errors="ignore") as f:
                if start:
                    f.seek(start)
                    # 途中行の途中から始まる場合を吸収
                    f.readline()
                return f.read()
        except Exception:
            return ""


class VoiceRecognizer:
    """
    Windows SAPI を用いた簡易音声入力（環境依存）。
    - 利用可能な場合のみ有効化。
    - Tkのイベントループと干渉しないよう、定期ポンピングでイベント処理。

    Usage:
        vr = VoiceRecognizer(on_text=callback)
        vr.start() / vr.stop()
    """

    def __init__(self, on_text: Callable[[str], None]) -> None:
        self.on_text = on_text
        self.enabled = HAS_WIN32
        self._context = None
        self._grammar = None
        self._running = False

        if not HAS_WIN32:
            return

        try:
            # In-Proc Recognizer（共有でも可）
            self._recognizer = win32com.client.Dispatch("SAPI.SpInprocRecognizer")
            self._context = self._recognizer.CreateRecoContext()
            self._context.EventInterests = 1 | 2  # EndStream | SoundStart/End等（簡略）
            self._events = win32com.client.WithEvents(self._context, self._ContextEvents)
            self._events.set_parent(self)
            self._grammar = self._context.CreateGrammar()
            self._grammar.DictationLoad()
        except Exception:
            self.enabled = False

    class _ContextEvents:  # noqa: D401
        """ISpeechRecoContextEvents の簡易ハンドラ"""

        def set_parent(self, parent: "VoiceRecognizer") -> None:
            self._parent = parent

        def OnRecognition(self, StreamNumber, StreamPosition, RecognitionType, Result) -> None:  # noqa: N802
            try:
                text = Result.PhraseInfo.GetText()
                if self._parent and self._parent.on_text:
                    self._parent.on_text(text)
            except Exception:
                pass

    def start(self) -> None:
        if not self.enabled or self._running:
            return
        try:
            self._grammar.DictationSetState(1)  # SGDSActive
            self._running = True
        except Exception:
            self.enabled = False
            self._running = False

    def stop(self) -> None:
        if not self.enabled or not self._running:
            return
        try:
            self._grammar.DictationSetState(0)  # SGDSInactive
            self._running = False
        except Exception:
            self._running = False

    @staticmethod
    def pump_once() -> None:
        """COMメッセージを1回だけ処理（Tkのafterで定期呼び出し）"""
        if not HAS_WIN32:
            return
        try:
            pythoncom.PumpWaitingMessages()
        except Exception:
            pass


class LLMClient:
    """
    OpenAI REST（chat.completions）を標準ライブラリで呼び出す簡易クライアント。

    Examples:
        >>> cli = LLMClient(api_key="sk-XXXX")
        >>> # 実呼び出しはネット環境が必要
    """

    def __init__(self, api_key: str, model: str, timeout: int = DEFAULT_LLM_TIMEOUT) -> None:
        self.api_key = api_key
        self.model = model
        self.timeout = timeout

    def ask(self, prompt: str, max_tokens: int = DEFAULT_LLM_MAX_TOKENS) -> str:
        """
        単一プロンプトを system/user 最小構成で問い合わせる。
        """
        if not self.api_key:
            raise RuntimeError("APIキーが未設定です。")

        url = "https://api.openai.com/v1/chat/completions"
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}",
        }
        payload = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": prompt},
            ],
            "max_tokens": max_tokens,
        }

        req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"), headers=headers, method="POST")
        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                data = json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            try:
                err = e.read().decode("utf-8")
            except Exception:
                err = str(e)
            raise RuntimeError(f"HTTPエラー: {e.code} {err}") from e
        except urllib.error.URLError as e:
            raise RuntimeError(f"接続エラー: {e.reason}") from e
        except Exception as e:
            raise RuntimeError(str(e)) from e

        try:
            return data["choices"][0]["message"]["content"]
        except Exception:
            raise RuntimeError("API応答の解析に失敗しました。")


class Lane:
    """
    1レーン分のUI（TAG/SUB_TAG/テキスト/🎤）と値取得ロジック。

    Examples:
        >>> # GUIコンテキスト内で利用
    """

    def __init__(
        self,
        master: tk.Widget,
        tags_map: Dict[str, List[str]],
        font: tkfont.Font,
        on_voice: Callable[[int], None],
        on_register: Callable[[int], None],
        lane_index: int,
    ) -> None:
        self.master = master
        self.tags_map = tags_map
        self.font = font
        self.on_voice = on_voice
        self.on_register = on_register
        self.index = lane_index

        self.var_tag = tk.StringVar(value="")
        self.var_sub = tk.StringVar(value="")
        self.var_text = tk.StringVar(value="")

        self._build()

    def _build(self) -> None:
        self.frm = ttk.Frame(self.master)
        self.frm.grid_columnconfigure(2, weight=1)

        self.cb_tag = ttk.Combobox(self.frm, textvariable=self.var_tag, state="readonly", width=16, style="App.TCombobox")
        self.cb_tag.grid(row=0, column=0, padx=4, pady=2, sticky="ew")

        self.cb_sub = ttk.Combobox(self.frm, textvariable=self.var_sub, state="readonly", width=16, style="App.TCombobox")
        self.cb_sub.grid(row=0, column=1, padx=4, pady=2, sticky="ew")

        # ログ入力フォームの横幅を拡張（2.5倍相当: 既定を概ね20とみなし50へ）
        # Windows IMEのプレエディット表示差異を避けるためにフォントを明示指定
        self.ent = ttk.Entry(self.frm, textvariable=self.var_text, width=50, style="App.TEntry", font=self.font)
        self.ent.grid(row=0, column=2, padx=4, pady=2, sticky="ew")
        # Enter/Numpad Enterで該当レーンを登録
        self.ent.bind("<Return>", self._on_entry_return)
        self.ent.bind("<KP_Enter>", self._on_entry_return)

        self.btn_voice = ttk.Button(self.frm, text=VOICE_BTN_TEXT, width=3, command=self._voice_click, style="App.TButton")
        self.btn_voice.grid(row=0, column=3, padx=2, pady=2)

        # 各行の登録ボタン
        self.btn_reg = ttk.Button(self.frm, text=BTN_REGISTER_TEXT, command=lambda: self.on_register(self.index), style="App.TButton")
        self.btn_reg.grid(row=0, column=4, padx=2, pady=2)

        # フォントはスタイル経由で適用済み

        # 値セット
        self.cb_tag["values"] = list(self.tags_map.keys())
        self.cb_tag.bind("<<ComboboxSelected>>", self._on_tag_change)

    def grid(self, **kwargs: Any) -> None:
        self.frm.grid(**kwargs)

    def set_default_tag(self, tag: str) -> None:
        tags = list(self.tags_map.keys())
        if tag in self.tags_map:
            self.var_tag.set(tag)
        elif tags:
            self.var_tag.set(tags[0])
        else:
            self.var_tag.set("-")
        self._refresh_subtags()

    def _on_tag_change(self, *_: Any) -> None:
        self._refresh_subtags()

    def _refresh_subtags(self) -> None:
        tag = self.var_tag.get()
        subs = self.tags_map.get(tag, [])
        if not subs:
            subs = ["-"]
        self.cb_sub["values"] = subs
        self.var_sub.set(subs[0])

    def _voice_click(self) -> None:
        # 親（MainApp）側で該当レーンのエントリに音声テキストを注入
        if self.on_voice:
            self.on_voice(self.index)

    def _on_entry_return(self, event: tk.Event) -> str:  # type: ignore[name-defined]
        """エントリ内でEnter押下時にこのレーンを登録する。"""
        try:
            if self.on_register:
                self.on_register(self.index)
        except Exception:
            pass
        # Entryでのデフォルト動作（ビープ等）を抑止
        return "break"

    def get_text_for_log(self) -> Optional[str]:
        """
        入力テキストが空なら None。
        あればログ書式の一部（[TAG][SUB]と本文）を返す。
        """
        text = self.var_text.get().strip()
        if not text:
            return None
        tag = self.var_tag.get().strip() or "-"
        sub = self.var_sub.get().strip() or "-"
        return f"{bracket_pad(tag)} {bracket_pad(sub)} {text}"

    def get_text_only(self) -> str:
        return self.var_text.get().strip()

    def append_text(self, extra: str) -> None:
        """音声入力などで末尾に追記"""
        cur = self.var_text.get()
        if cur and not cur.endswith(" "):
            cur += " "
        self.var_text.set(cur + extra)


class MainApp:
    """
    メインウィンドウとアプリ全体の制御。

    Examples:
        >>> # if __name__ == "__main__": MainApp().run()
    """

    def __init__(self) -> None:
        # 変数初期化（メイン処理冒頭で初期化ルール遵守）
        self.cfg = ConfigManager()
        self.logman = LogManager(self.cfg)
        self.is_large = False
        self.running = True

        self.root = tk.Tk()
        self.root.title(TITLE_TEXT)

        # フォント（名前付きフォントで管理）
        self.font_base = tkfont.Font(family=BASE_FONT_FAMILY, size=BASE_FONT_SIZE)
        self.font_mono = tkfont.Font(family=MONO_FONT_FAMILY, size=MONO_FONT_SIZE)

        # ttkスタイル（フォントはスタイルで適用）
        self.style = ttk.Style(self.root)
        self._configure_widget_styles()

        # 音声認識
        self.active_voice_lane: Optional[int] = None
        self.voice_recognizer = VoiceRecognizer(on_text=self._on_voice_text)

        # LLM
        self.llm_cli: Optional[LLMClient] = None
        self._prepare_llm_client()

        # UI構築
        self._build_menu()
        self._build_top_info()
        self._build_input_area()
        self._build_preview_area()
        self._build_bottom_paths()
        self._build_llm_area()

        self._refresh_log_preview()
        self._update_datetime_loop()
        self._pump_voice_loop()

    # ====== UI構築 ======
    def _reconfigure_fonts(self) -> None:
        base_size = BASE_FONT_SIZE_LARGE if self.is_large else BASE_FONT_SIZE
        mono_size = MONO_FONT_SIZE_LARGE if self.is_large else MONO_FONT_SIZE
        self.font_base.configure(size=base_size)
        self.font_mono.configure(size=mono_size)
        # スタイルに紐づくフォントも更新（named font を参照しているため再設定で確実化）
        self._configure_widget_styles()

    def _configure_widget_styles(self) -> None:
        try:
            pad_y = 6 if self.is_large else 2
            pad_x = 6 if self.is_large else 4
            # 既定スタイルにも反映（各所のボタン/ラベルが追従するように）
            self.style.configure("TButton", font=self.font_base, padding=(pad_x, pad_y))
            self.style.configure("TLabel", font=self.font_base)
            self.style.configure("TEntry", font=self.font_base, padding=(pad_x, pad_y))
            self.style.configure("TCombobox", font=self.font_base, padding=(pad_x, pad_y))
            self.style.configure("App.TButton", font=self.font_base, padding=(pad_x, pad_y))
            self.style.configure("App.TEntry", font=self.font_base, padding=(pad_x, pad_y))
            self.style.configure("App.TCombobox", font=self.font_base, padding=(pad_x, pad_y))
            # Tk標準フォントも同期（IMEプレエディットの表示を含め統一）
            try:
                tkfont.nametofont("TkDefaultFont").configure(
                    family=self.font_base.cget("family"),
                    size=self.font_base.cget("size"),
                )
                tkfont.nametofont("TkTextFont").configure(
                    family=self.font_base.cget("family"),
                    size=self.font_base.cget("size"),
                )
                tkfont.nametofont("TkFixedFont").configure(
                    family=self.font_mono.cget("family"),
                    size=self.font_mono.cget("size"),
                )
            except Exception:
                pass
        except Exception:
            pass

    def _build_menu(self) -> None:
        menubar = tk.Menu(self.root)

        m_file = tk.Menu(menubar, tearoff=0)
        m_file.add_command(label="設定ファイルを開く", command=self._open_config_folder)
        m_file.add_command(label="ログフォルダを開く", command=self._open_log_folder)
        m_file.add_separator()
        m_file.add_command(label="終了", command=self._on_exit)
        menubar.add_cascade(label="ファイル", menu=m_file)

        m_view = tk.Menu(menubar, tearoff=0)
        m_view.add_command(label=BTN_SCALE_TEXT, command=self._toggle_scale)
        menubar.add_cascade(label="表示", menu=m_view)

        self.root.config(menu=menubar)

    def _build_top_info(self) -> None:
        frm = ttk.Frame(self.root)
        frm.pack(fill="x", padx=8, pady=6)

        # 現在日時
        self.var_now = tk.StringVar(value="")
        lbl_now = ttk.Label(frm, textvariable=self.var_now, font=self.font_base)
        lbl_now.grid(row=0, column=0, sticky="w", padx=4)

        # サイズ変更ボタン（時計の右隣）
        btn_scale = ttk.Button(frm, text=BTN_SCALE_TEXT, command=self._toggle_scale)
        btn_scale.grid(row=0, column=1, sticky="w", padx=8)

    def _build_input_area(self) -> None:
        outer = ttk.LabelFrame(self.root, text="入力エリア（TAG / SUB_TAG / テキスト）")
        outer.pack(fill="x", padx=8, pady=6)

        self.lanes: List[Lane] = []
        tags_map = self.cfg.tags_map

        for i in range(LANE_COUNT):
            lane = Lane(
                outer,
                tags_map=tags_map,
                font=self.font_base,
                on_voice=self._voice_click_from_lane,
                on_register=self._on_register_lane,
                lane_index=i,
            )
            lane.grid(row=i, column=0, sticky="ew", padx=4, pady=1)
            self.lanes.append(lane)

        # 既定のTAG割当（MAIN_TAGがあれば優先、無ければ上から）
        tags_list = list(tags_map.keys())
        main = self.cfg.main_tags
        for i, lane in enumerate(self.lanes):
            if i < len(main) and main[i] in tags_map:
                lane.set_default_tag(main[i])
            else:
                lane.set_default_tag(tags_list[i] if i < len(tags_list) else "-")

    def _build_preview_area(self) -> None:
        outer = ttk.LabelFrame(self.root, text="ログ簡易表示（選択してコピー可能）")
        outer.pack(fill="both", expand=False, padx=8, pady=6)

        self.txt_preview = tk.Text(
            outer,
            height=12,
            wrap="none",
            font=self.font_mono,
        )
        yscroll = ttk.Scrollbar(outer, orient="vertical", command=self.txt_preview.yview)
        xscroll = ttk.Scrollbar(outer, orient="horizontal", command=self.txt_preview.xview)
        self.txt_preview.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        self.txt_preview.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")

        outer.grid_columnconfigure(0, weight=1)
        outer.grid_rowconfigure(0, weight=1)

    def _build_bottom_paths(self) -> None:
        frm = ttk.Frame(self.root)
        frm.pack(fill="x", padx=8, pady=(0, 6))

        # 設定ファイルパス
        self.var_cfg = tk.StringVar(value=f"設定: {self.cfg.cfg_path}")
        lbl_cfg = ttk.Label(frm, textvariable=self.var_cfg, font=self.font_base)
        lbl_cfg.grid(row=0, column=0, sticky="w", padx=4)

        # ログパス
        self.var_log = tk.StringVar(value=f"ログ: {self.logman.current_log_path}")
        lbl_log = ttk.Label(frm, textvariable=self.var_log, font=self.font_base)
        lbl_log.grid(row=1, column=0, sticky="w", padx=4, pady=(2, 0))

    def _build_llm_area(self) -> None:
        outer = ttk.LabelFrame(self.root, text="LLM質問エリア（表示切替可能）")
        outer.pack(fill="x", padx=8, pady=6)
        self.llm_frame_outer = outer

        # 表示切替（常に表示）
        self.var_llm_visible = tk.BooleanVar(value=True)
        chk = ttk.Checkbutton(outer, text=LLM_TOGGLE_TEXT, variable=self.var_llm_visible, command=self._update_llm_visible)
        chk.grid(row=0, column=0, sticky="w", padx=4, pady=2)

        # コンテンツコンテナ（チェックボックスは残す）
        content = ttk.Frame(outer)
        content.grid(row=1, column=0, columnspan=2, sticky="ew")
        self.llm_frame_content = content

        # BASE_PROMPTS 選択
        self.var_base_select = tk.StringVar(value="-")
        base_keys = ["-"] + list(self.cfg.base_prompts.keys())
        ttk.Label(content, text="プリセット:", font=self.font_base).grid(row=0, column=0, sticky="e", padx=4)
        # プリセットの横幅を2.5倍（既定24→60程度）
        self.cb_base = ttk.Combobox(content, textvariable=self.var_base_select, values=base_keys, state="readonly", width=60)
        self.cb_base.grid(row=0, column=1, sticky="w", padx=4, pady=2)
        self.cb_base.bind("<<ComboboxSelected>>", self._on_base_selected)

        # 自由入力
        ttk.Label(content, text="自由入力:", font=self.font_base).grid(row=1, column=0, sticky="ne", padx=4)
        self.txt_prompt = tk.Text(content, height=4, font=self.font_base, wrap="word")
        self.txt_prompt.grid(row=1, column=1, sticky="ew", padx=4, pady=2)

        # 送信ボタン・状態
        self.btn_send = ttk.Button(content, text=LLM_SEND_TEXT, command=self._on_llm_send)
        self.btn_send.grid(row=2, column=1, sticky="e", padx=4, pady=4)

        self.var_llm_status = tk.StringVar(value="")
        self.lbl_llm_status = ttk.Label(content, textvariable=self.var_llm_status, font=self.font_base)
        self.lbl_llm_status.grid(row=2, column=0, sticky="w", padx=4)

        # 回答表示（ログ同様のモノスペース）
        ttk.Label(content, text="回答:", font=self.font_base).grid(row=3, column=0, sticky="ne", padx=4, pady=(2, 6))
        self.txt_answer = tk.Text(content, height=6, font=self.font_mono, wrap="word")
        self.txt_answer.grid(row=3, column=1, sticky="ew", padx=4, pady=(2, 6))

        # レイアウト拡張
        content.grid_columnconfigure(1, weight=1)

        # 初期の活性/非活性
        self._update_llm_visible()
        self._update_send_button_state()

        # 入力変化監視
        self.txt_prompt.bind("<<Modified>>", self._on_prompt_modified)

    # ====== イベント/動作 ======
    def _on_prompt_modified(self, *_: Any) -> None:
        try:
            self.txt_prompt.edit_modified(False)
        except Exception:
            pass
        self._update_send_button_state()

    def _update_send_button_state(self) -> None:
        sel = self.var_base_select.get()
        free = self.txt_prompt.get("1.0", "end").strip()
        api_ok = bool(self.cfg.openai_api_key)
        if sel == "-" and not free:
            ok = False
        else:
            ok = api_ok
        self.btn_send.configure(state=("normal" if ok else "disabled"))
        if not api_ok:
            self.var_llm_status.set("OPENAI_API_KEYが未設定です。環境変数またはyamlで設定してください。")
        else:
            self.var_llm_status.set("")

    def _on_base_selected(self, *_: Any) -> None:
        key = self.var_base_select.get()
        if key != "-":
            base = self.cfg.base_prompts.get(key, "")
            self.txt_prompt.delete("1.0", "end")
            self.txt_prompt.insert("1.0", base)
            try:
                self.txt_prompt.grid_remove()
            except Exception:
                pass
        else:
            try:
                self.txt_prompt.grid()
            except Exception:
                pass
        self._update_send_button_state()

    def _update_llm_visible(self) -> None:
        vis = self.var_llm_visible.get()
        try:
            if vis:
                self.llm_frame_content.grid()
            else:
                self.llm_frame_content.grid_remove()
        except Exception:
            pass

    def _open_config_folder(self) -> None:
        path = self.cfg.cfg_path
        try:
            os.startfile(os.path.dirname(path))
        except Exception:
            messagebox.showinfo(INFO_DIALOG_TITLE, f"エクスプローラーで開けませんでした。\n{path}")

    def _open_log_folder(self) -> None:
        path = self.cfg.log_dir
        try:
            os.startfile(path)
        except Exception:
            messagebox.showinfo(INFO_DIALOG_TITLE, f"エクスプローラーで開けませんでした。\n{path}")

    def _toggle_scale(self) -> None:
        self.is_large = not self.is_large
        self._reconfigure_fonts()
        # 再描画トリガ（フォントは参照で反映されるが、余白等の再計算のため更新）
        self.root.update_idletasks()
        # コンテンツの自然サイズへ合わせてウィンドウサイズを再計算
        try:
            self.root.geometry("")
        except Exception:
            pass

    def _on_register(self) -> None:
        ts = format_timestamp(effective_date(now_jst()))
        out: List[str] = []
        for lane in self.lanes:
            part = lane.get_text_for_log()
            if part:
                out.append(f"{ts} {part}")
        if not out:
            messagebox.showinfo(INFO_DIALOG_TITLE, "入力テキストがありません。")
            return
        try:
            self.logman.append_lines(out)
            self._refresh_log_preview()
            # 成功時は各行のテキストだけクリア（タグは維持）
            for lane in self.lanes:
                lane.var_text.set("")
        except Exception:
            messagebox.showerror(ERROR_DIALOG_TITLE, get_exception_trace())

    def _on_register_lane(self, idx: int) -> None:
        if idx < 0 or idx >= len(self.lanes):
            return
        ts = format_timestamp(effective_date(now_jst()))
        lane = self.lanes[idx]
        part = lane.get_text_for_log()
        if not part:
            messagebox.showinfo(INFO_DIALOG_TITLE, "入力テキストがありません。")
            return
        try:
            self.logman.append_lines([f"{ts} {part}"])
            self._refresh_log_preview()
            lane.var_text.set("")
        except Exception:
            messagebox.showerror(ERROR_DIALOG_TITLE, get_exception_trace())

    def _refresh_log_preview(self) -> None:
        try:
            self.txt_preview.configure(state="normal")
            self.txt_preview.delete("1.0", "end")
            self.txt_preview.insert("1.0", self.logman.read_text())
            self.txt_preview.configure(state="normal")
        except Exception:
            # 読込失敗しても致命にはしない
            pass

    def _update_datetime_loop(self) -> None:
        try:
            now = now_jst()
            self.var_now.set(now.strftime("%Y-%m-%d %H:%M:%S"))
            # ロールオーバ判定
            if self.logman.rollover_if_needed():
                self.var_log.set(f"ログ: {self.logman.current_log_path}")
                self._refresh_log_preview()
        except Exception:
            # 表示の更新失敗は無視
            pass
        finally:
            self.root.after(500, self._update_datetime_loop)

    # ====== 音声入力 ======
    def _voice_click_from_lane(self, idx: int) -> None:
        if not self.voice_recognizer.enabled:
            messagebox.showinfo(INFO_DIALOG_TITLE, "音声入力はこの環境では利用できません。")
            return
        # トグル：同じレーンで押されたら停止、別レーンなら切替
        if self.active_voice_lane == idx:
            self.voice_recognizer.stop()
            self.active_voice_lane = None
            messagebox.showinfo(INFO_DIALOG_TITLE, f"音声入力を停止しました（レーン{idx + 1}）。")
        else:
            self.voice_recognizer.start()
            self.active_voice_lane = idx
            messagebox.showinfo(INFO_DIALOG_TITLE, f"音声入力を開始しました（レーン{idx + 1}）。話しかけてください。")

    def _on_voice_text(self, text: str) -> None:
        if self.active_voice_lane is None:
            return
        if 0 <= self.active_voice_lane < len(self.lanes):
            self.lanes[self.active_voice_lane].append_text(text)

    def _pump_voice_loop(self) -> None:
        try:
            # 音声入力が稼働中のときのみCOMメッセージを処理
            if self.voice_recognizer.enabled and self.active_voice_lane is not None:
                VoiceRecognizer.pump_once()
        except Exception:
            pass
        finally:
            self.root.after(50, self._pump_voice_loop)

    # ====== LLM ======
    def _prepare_llm_client(self) -> None:
        api_key = self.cfg.openai_api_key
        if api_key:
            self.llm_cli = LLMClient(api_key=api_key, model=self.cfg.llm_model, timeout=self.cfg.llm_timeout)
        else:
            self.llm_cli = None

    def _on_llm_send(self) -> None:
        if not self.llm_cli:
            self._update_send_button_state()
            return

        prompt = self.txt_prompt.get("1.0", "end").strip()
        if not prompt:
            # BASE_PROMPTSが選ばれていればそこから読めるが、ボタン活性条件でここには来ない想定
            key = self.var_base_select.get()
            prompt = self.cfg.base_prompts.get(key, "").strip()

        self.btn_send.configure(state="disabled", text="送信中...")
        self.var_llm_status.set("回答待ち...")
        self.txt_answer.delete("1.0", "end")

        def worker() -> None:
            try:
                ans = self.llm_cli.ask(prompt, max_tokens=self.cfg.llm_max_tokens)
            except Exception as e:
                ans = f"[エラー] {e}"

            def done() -> None:
                self.var_llm_status.set("")
                self.btn_send.configure(state="normal", text=LLM_SEND_TEXT)
                ts = format_timestamp(effective_date(now_jst()))
                user = self.cfg.user_name
                # 表示
                self.txt_answer.insert("1.0", ans)
                # ログ追記（LLMタグで記録者名入り）
                log_line = f"{ts} {bracket_pad('LLM')} {bracket_pad(user)} Q:{prompt} / A:{ans}"
                try:
                    self.logman.append_lines([log_line])
                    self._refresh_log_preview()
                except Exception:
                    pass

            self.root.after(0, done)

        threading.Thread(target=worker, daemon=True).start()

    # ====== 終了 ======
    def _on_exit(self) -> None:
        try:
            self.running = False
            if self.voice_recognizer:
                self.voice_recognizer.stop()
            self.root.destroy()
        except Exception:
            os._exit(0)  # noqa: WPS437

    # ====== 実行 ======
    def run(self) -> None:
        self.root.mainloop()


# 4. メイン処理
if __name__ == "__main__":
    try:
        app = MainApp()
        app.run()
    except Exception:
        messagebox.showerror(ERROR_DIALOG_TITLE, get_exception_trace())
