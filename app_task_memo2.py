# -*- coding: utf-8 -*-
# アプリ名: 0. 作業メモ用デスクトップアプリ２

"""
作業メモ用デスクトップアプリ（Windows 11 / Python 3.12）
- PySide6ベース（tkinterベースからの移行）
- ログ日付の区切りは 04:00 AM
- UIサイズは通常/4倍のトグル
- レーン数は設定可能（既定5）の (TAG / SUB_TAG / テキスト)
- LLM質問エリア（表示切替、OpenAI RESTを標準ライブラリで直接叩く）
- 設定/ログパスの明示、プレビュー付き
"""

# 1. import文（必要最小限 + 標準ライブラリ中心）
import sys
import os
import io
import json
import threading
import urllib.request
import urllib.error
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from typing import Any, Callable, Dict, List, Optional

# GUI (PySide6)
from PySide6 import QtCore, QtGui, QtWidgets


# 2. 定数定義（繰り返し登場する固定文字列や数値はここで定義）
APP_BASENAME: str = "app_task_memo"
APP_FILENAME: str = f"{APP_BASENAME}.py"
CFG_FILENAME: str = f"{APP_BASENAME}.yaml"

# ログ日付の区切り時刻（04:00）
CUTOFF_HOUR: int = 4

# フォント基準（固定サイズ運用）
BASE_FONT_FAMILY: str = "Yu Gothic UI"
BASE_FONT_SIZE: int = 10  # 通常時
BASE_FONT_SIZE_LARGE: int = 12  # 大
MONO_FONT_FAMILY: str = "Consolas"
MONO_FONT_SIZE: int = 10
MONO_FONT_SIZE_LARGE: int = 12
BASE_FONT_SIZE_LARGE_MORE: int = 18  # 特大
MONO_FONT_SIZE_LARGE_MORE: int = 16


# レーン数（既定値。実際の使用数は設定 YAML の LANE_COUNT で上書き可能）
DEFAULT_LANE_COUNT: int = 5

# LLM既定
DEFAULT_LLM_PROVIDER: str = "openai"
DEFAULT_LLM_MODEL: str = "gpt-4o-mini"
DEFAULT_LLM_MAX_TOKENS: int = 2000
DEFAULT_LLM_TIMEOUT: int = 60
DEFAULT_OPENAI_KEY_ENV: str = "OPENAI_API_KEY"

# テキストの角括弧付き最小幅（例: "[TAG____]")
BRACKET_MIN_WIDTH: int = 5

# JST
TZ_JST = ZoneInfo("Asia/Tokyo")

# 曜日3文字
DOW3 = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]
# 追加: 日本語曜日（表示用）
DOW_JA = ["月", "火", "水", "木", "金", "土", "日"]

# メニュー等の文言
TITLE_TEXT = "Task Memo（作業メモ）"
BTN_SCALE_TEXT = "サイズ変更"
BTN_REGISTER_TEXT = "登録"
LLM_TOGGLE_TEXT = "LLMエリアを表示"
LLM_SEND_TEXT = "送信"
# 追加: ログエクスプローラー
BTN_LOG_EXPLORER_TEXT = "ログエクスプローラー"
#

# 例外処理用：出力先ダイアログ
ERROR_DIALOG_TITLE = "エラー"
INFO_DIALOG_TITLE = "情報"


# 3. 関数／クラス定義
def get_exception_trace() -> str:
	"""例外のトレースバックを取得"""
	import traceback

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


def format_date_jp(dt: Optional[datetime] = None) -> str:
	"""
	日本語の年月日 + 曜日を返す。

	Examples:
		>>> s = format_date_jp(datetime(2025, 9, 7, tzinfo=TZ_JST))
		>>> s
		'《2025年9月7日(日)》'
	"""
	if dt is None:
		dt = now_jst()
	return f"《{dt.year}年{dt.month}月{dt.day}日({DOW_JA[dt.weekday()]})》"


class SimpleYAML:
	"""
	依存ライブラリ無しで扱える非常に単純化したYAMLパーサ/ダンパ。
	- 本アプリの想定設定（KEY:SCALAR / KEY:LIST / KEY:DICT with LIST/SCALAR）に限定。
	- 完全なYAML互換ではない点に留意。

	Examples:
		>>> text = "A: 1\nB:\n  - x\n  - y\nC:\n  K: V\n"
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
		stack: List[tuple[int, Any, Optional[Dict[str, Any]], Optional[str]]] = [(-1, root, None, None)]

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
				key, _sep, rest = line.partition(":")
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

		return root

	@staticmethod
	def dumps(data: Dict[str, Any]) -> str:
		def dump_obj(obj: Any, indent: int = 0) -> str:
			sp = " " * indent
			if isinstance(obj, dict):
				out: List[str] = []
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
		# レーン数（YAMLのスカラーは文字列になり得るため、後段で int 化）
		data.setdefault("LANE_COUNT", str(DEFAULT_LANE_COUNT))

		# 数値系はint化
		try:
			data["LLM_MAX_COMPLETION_TOKENS"] = int(data["LLM_MAX_COMPLETION_TOKENS"])
		except Exception:
			data["LLM_MAX_COMPLETION_TOKENS"] = DEFAULT_LLM_MAX_TOKENS
		try:
			data["LLM_TIMEOUT"] = int(data["LLM_TIMEOUT"])
		except Exception:
			data["LLM_TIMEOUT"] = DEFAULT_LLM_TIMEOUT
		# レーン数の int 化と簡易バリデーション（1以上）
		try:
			lane_cnt = int(data.get("LANE_COUNT", DEFAULT_LANE_COUNT))
			if lane_cnt < 1:
				lane_cnt = DEFAULT_LANE_COUNT
			data["LANE_COUNT"] = lane_cnt
		except Exception:
			data["LANE_COUNT"] = DEFAULT_LANE_COUNT

		# ログディレクトリの解決
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
		return list(self.config.get("MAIN_TAG", []))[: self.lane_count]

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

	@property
	def lane_count(self) -> int:
		"""入力レーン数（設定値。未設定/不正時は既定値）。"""
		try:
			return int(self.config.get("LANE_COUNT", DEFAULT_LANE_COUNT))
		except Exception:
			return DEFAULT_LANE_COUNT


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

	# 追加: 日付指定でパスを取得（副作用なし）
	def get_log_path_for_date(self, dt: datetime) -> str:
		return os.path.join(
			self.cfg.log_dir,
			f"{APP_BASENAME}_{yyyymmdd_dow(dt)}.log",
		)

	# 追加: 日付指定でテキストを読む（存在しなければ空文字）
	def read_text_for_date(self, dt: datetime, tail_kb: int = 512) -> str:
		path = self.get_log_path_for_date(dt)
		try:
			size = os.path.getsize(path)
			start = max(0, size - tail_kb * 1024)
			with io.open(path, "r", encoding="utf-8_sig", errors="ignore") as f:
				if start:
					f.seek(start)
					f.readline()
				return f.read()
		except Exception:
			return ""


class LogExplorerDialog(QtWidgets.QDialog):
	"""
	ログエクスプローラーウインドウ。

	- 「⇦」「⇨」で日付移動
 	- 中央に対象ログ日付を日本語表示（例: 《2025年9月7日(日)》）
 	- 下部にログ簡易表示
	"""

	def __init__(self, parent: QtWidgets.QWidget, logman: LogManager) -> None:
		super().__init__(parent)
		self.setWindowTitle("ログエクスプローラー")
		self.logman = logman
		self.cur_dt: datetime = effective_date(now_jst())

		root = QtWidgets.QWidget(self)
		lay = QtWidgets.QVBoxLayout(root)
		lay.setContentsMargins(8, 8, 8, 8)
		lay.setSpacing(8)
		self.setLayout(lay)

		# 上部: ナビゲーション
		top = QtWidgets.QWidget(root)
		top_l = QtWidgets.QGridLayout(top)
		top_l.setContentsMargins(0, 0, 0, 0)
		self.btn_prev = QtWidgets.QPushButton("⇦", top)
		self.btn_next = QtWidgets.QPushButton("⇨", top)
		self.lbl_date = QtWidgets.QLabel("", top)
		self.lbl_date.setAlignment(QtCore.Qt.AlignCenter)
		self.lbl_date.setFont(QtGui.QFont(BASE_FONT_FAMILY, BASE_FONT_SIZE))
		top_l.addWidget(self.btn_prev, 0, 0)
		top_l.addWidget(self.lbl_date, 0, 1)
		top_l.addWidget(self.btn_next, 0, 2)
		top_l.setColumnStretch(1, 1)
		lay.addWidget(top)

		# 本文: ログプレビュー
		self.txt = QtWidgets.QPlainTextEdit(root)
		self.txt.setReadOnly(False)
		self.txt.setLineWrapMode(QtWidgets.QPlainTextEdit.NoWrap)
		self.txt.setFont(QtGui.QFont(MONO_FONT_FAMILY, MONO_FONT_SIZE))
		lay.addWidget(self.txt)

		# シグナル
		self.btn_prev.clicked.connect(lambda: self._shift_date(-1))
		self.btn_next.clicked.connect(lambda: self._shift_date(1))

		# 初期表示
		self._refresh()
		self.resize(800, 600)

	def _shift_date(self, delta_days: int) -> None:
		self.cur_dt = self.cur_dt + timedelta(days=delta_days)
		self._refresh()

	def _refresh(self) -> None:
		self.lbl_date.setText(format_date_jp(self.cur_dt))
		text = self.logman.read_text_for_date(self.cur_dt)
		self.txt.setPlainText(text)


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


class LaneWidget(QtWidgets.QWidget):
	"""
	1レーン分のUI（TAG/SUB_TAG/テキスト）と値取得ロジック。

	Examples:
		>>> # GUIコンテキスト内で利用
	"""

	returnPressed = QtCore.Signal(int)

	def __init__(
		self,
		parent: QtWidgets.QWidget,
		tags_map: Dict[str, List[str]],
		on_register: Callable[[int], None],
		lane_index: int,
		base_font: QtGui.QFont,
	) -> None:
		super().__init__(parent)
		self.tags_map = tags_map
		self.on_register = on_register
		self.index = lane_index

		# フォント
		self.base_font = QtGui.QFont(base_font)

		# ウィジェット
		self.cb_tag = QtWidgets.QComboBox(self)
		self.cb_tag.setEditable(False)
		self.cb_tag.setMinimumContentsLength(10)
		self.cb_tag.setFont(self.base_font)

		self.cb_sub = QtWidgets.QComboBox(self)
		self.cb_sub.setEditable(False)
		self.cb_sub.setMinimumContentsLength(10)
		self.cb_sub.setFont(self.base_font)

		self.le_text = QtWidgets.QLineEdit(self)
		self.le_text.setFont(self.base_font)
		self.le_text.setMaxLength(10000)

		self.btn_reg = QtWidgets.QPushButton(BTN_REGISTER_TEXT, self)
		self.btn_reg.setFont(self.base_font)

		# レイアウト
		lay = QtWidgets.QGridLayout(self)
		lay.setContentsMargins(0, 0, 0, 0)
		pad_y = 6
		pad_x = 6
		lay.setHorizontalSpacing(pad_x)
		lay.setVerticalSpacing(pad_y)
		lay.addWidget(self.cb_tag, 0, 0)
		lay.addWidget(self.cb_sub, 0, 1)
		lay.addWidget(self.le_text, 0, 2)
		lay.addWidget(self.btn_reg, 0, 3)
		lay.setColumnStretch(2, 1)

		# 値セット
		self.cb_tag.addItems(list(self.tags_map.keys()))
		self.cb_tag.currentIndexChanged.connect(self._refresh_subtags)
		self._refresh_subtags()

		# シグナル
		self.le_text.returnPressed.connect(self._on_entry_return)
		self.btn_reg.clicked.connect(lambda: self.on_register(self.index))

	def set_default_tag(self, tag: str) -> None:
		tags = list(self.tags_map.keys())
		if tag in self.tags_map:
			self.cb_tag.setCurrentText(tag)
		elif tags:
			self.cb_tag.setCurrentText(tags[0])
		else:
			self.cb_tag.setCurrentText("-")
		self._refresh_subtags()

	def _refresh_subtags(self) -> None:
		tag = self.cb_tag.currentText()
		subs = self.tags_map.get(tag, [])
		if not subs:
			subs = ["-"]
		self.cb_sub.blockSignals(True)
		self.cb_sub.clear()
		self.cb_sub.addItems(subs)
		self.cb_sub.blockSignals(False)

	def _on_entry_return(self) -> None:
		try:
			if self.on_register:
				self.on_register(self.index)
		except Exception:
			pass

	def get_text_for_log(self) -> Optional[str]:
		"""
		入力テキストが空なら None。
		あればログ書式の一部（[TAG][SUB]と本文）を返す。
		"""
		text = self.le_text.text().strip()
		if not text:
			return None
		tag = self.cb_tag.currentText().strip() or "-"
		sub = self.cb_sub.currentText().strip() or "-"
		return f"{bracket_pad(tag)} {bracket_pad(sub)} {text}"

	def get_text_only(self) -> str:
		return self.le_text.text().strip()

	def append_text(self, extra: str) -> None:
		"""音声入力などで末尾に追記"""
		cur = self.le_text.text()
		if cur and not cur.endswith(" "):
			cur += " "
		self.le_text.setText(cur + extra)


class MainWindow(QtWidgets.QMainWindow):
	"""
	メインウィンドウとアプリ全体の制御。

	Examples:
		>>> # if __name__ == "__main__": run_app()
	"""

	def __init__(self) -> None:
		super().__init__()
		# 変数初期化（メイン処理冒頭で初期化ルール遵守）
		self.cfg = ConfigManager()
		self.logman = LogManager(self.cfg)
		self.scale_level = 0  # 0: 通常, 1: 大, 2: 特大

		self.setWindowTitle(TITLE_TEXT)

		# フォント
		self.font_base = QtGui.QFont(BASE_FONT_FAMILY, BASE_FONT_SIZE)
		self.font_mono = QtGui.QFont(MONO_FONT_FAMILY, MONO_FONT_SIZE)

		# ログエクスプローラー
		self._log_explorer: Optional[LogExplorerDialog] = None

		# UI構築
		self._build_menu()
		self._build_central()

		# LLM
		self.llm_cli: Optional[LLMClient] = None
		self._prepare_llm_client()

		# 初期表示
		self._refresh_log_preview()
		self._update_datetime()

		# タイマー
		self.timer = QtCore.QTimer(self)
		self.timer.setInterval(500)
		self.timer.timeout.connect(self._update_datetime)
		self.timer.start()

		# サイズ調整
		self.adjustSize()

	# ====== UI構築 ======
	def _reconfigure_fonts(self) -> None:
		if self.scale_level == 0:
			base_size = BASE_FONT_SIZE
			mono_size = MONO_FONT_SIZE
		elif self.scale_level == 1:
			base_size = BASE_FONT_SIZE_LARGE
			mono_size = MONO_FONT_SIZE_LARGE
		else:
			base_size = BASE_FONT_SIZE_LARGE_MORE
			mono_size = MONO_FONT_SIZE_LARGE_MORE
		self.font_base.setPointSize(base_size)
		self.font_mono.setPointSize(mono_size)

		# 再適用
		self._apply_fonts_recursive(self.centralWidget())

	def _apply_fonts_recursive(self, widget: Optional[QtWidgets.QWidget]) -> None:
		if widget is None:
			return
		if isinstance(widget, (QtWidgets.QLineEdit, QtWidgets.QComboBox, QtWidgets.QPushButton, QtWidgets.QLabel, QtWidgets.QTextEdit, QtWidgets.QPlainTextEdit)):
			if isinstance(widget, (QtWidgets.QTextEdit, QtWidgets.QPlainTextEdit)):
				widget.setFont(self.font_mono)
			else:
				widget.setFont(self.font_base)
		for child in widget.findChildren(QtWidgets.QWidget):
			self._apply_fonts_recursive(child)

	def _build_menu(self) -> None:
		menubar = self.menuBar()

		m_file = menubar.addMenu("ファイル")
		act_open_cfg = m_file.addAction("設定ファイルを開く")
		act_open_cfg.triggered.connect(self._open_config_folder)
		act_open_log = m_file.addAction("ログフォルダを開く")
		act_open_log.triggered.connect(self._open_log_folder)
		m_file.addSeparator()
		act_exit = m_file.addAction("終了")
		act_exit.triggered.connect(self._on_exit)

		m_view = menubar.addMenu("表示")
		act_scale = m_view.addAction(BTN_SCALE_TEXT)
		act_scale.triggered.connect(self._toggle_scale)

	def _build_central(self) -> None:
		root = QtWidgets.QWidget(self)
		self.setCentralWidget(root)

		vbox = QtWidgets.QVBoxLayout(root)
		vbox.setContentsMargins(8, 8, 8, 8)
		vbox.setSpacing(8)

		# 上段: 現在日時 / ログエクスプローラー / サイズ変更ボタン
		top = QtWidgets.QWidget(root)
		top_l = QtWidgets.QGridLayout(top)
		top_l.setContentsMargins(0, 0, 0, 0)
		self.lbl_now = QtWidgets.QLabel("", top)
		self.lbl_now.setFont(self.font_base)
		btn_explorer = QtWidgets.QPushButton(BTN_LOG_EXPLORER_TEXT, top)
		btn_explorer.setFont(self.font_base)
		btn_explorer.clicked.connect(self._open_log_explorer)
		btn_scale = QtWidgets.QPushButton(BTN_SCALE_TEXT, top)
		btn_scale.setFont(self.font_base)
		btn_scale.clicked.connect(self._toggle_scale)
		# 追加: 常に手前チェックボックス（サイズ変更ボタンの右）
		self.chk_topmost = QtWidgets.QCheckBox("常に手前", top)
		self.chk_topmost.setChecked(False)
		self.chk_topmost.toggled.connect(self._on_topmost_toggled)
		top_l.addWidget(self.lbl_now, 0, 0, 1, 1)
		top_l.addWidget(btn_explorer, 0, 1, 1, 1)
		top_l.addWidget(btn_scale, 0, 2, 1, 1)
		top_l.addWidget(self.chk_topmost, 0, 3, 1, 1)
		top_l.setColumnStretch(0, 1)
		vbox.addWidget(top)

		# 入力エリア
		input_group = QtWidgets.QGroupBox("入力エリア（TAG / SUB_TAG / テキスト）", root)
		input_l = QtWidgets.QVBoxLayout(input_group)
		self.lanes: List[LaneWidget] = []
		tags_map = self.cfg.tags_map
		lane_total = self.cfg.lane_count
		for i in range(lane_total):
			lane = LaneWidget(
				input_group,
				tags_map=tags_map,
				on_register=self._on_register_lane,
				lane_index=i,
				base_font=self.font_base,
			)
			input_l.addWidget(lane)
			self.lanes.append(lane)

		# 既定のTAG割当
		tags_list = list(tags_map.keys())
		main = self.cfg.main_tags
		for i, lane in enumerate(self.lanes):
			if i < len(main) and main[i] in tags_map:
				lane.set_default_tag(main[i])
			else:
				lane.set_default_tag(tags_list[i] if i < len(tags_list) else "-")

		vbox.addWidget(input_group)

		# ログ簡易表示
		prev_group = QtWidgets.QGroupBox("ログ簡易表示（選択してコピー可能）", root)
		prev_l = QtWidgets.QGridLayout(prev_group)
		self.txt_preview = QtWidgets.QPlainTextEdit(prev_group)
		self.txt_preview.setReadOnly(False)
		self.txt_preview.setLineWrapMode(QtWidgets.QPlainTextEdit.NoWrap)
		self.txt_preview.setFont(self.font_mono)
		prev_l.addWidget(self.txt_preview, 0, 0)
		vbox.addWidget(prev_group)

		# 下段: パス表示
		# 折りたたみトグル
		self.btn_paths = QtWidgets.QToolButton(root)
		self.btn_paths.setText("設定/ログのパスを表示")
		self.btn_paths.setCheckable(True)
		self.btn_paths.setChecked(False)
		self.btn_paths.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
		self.btn_paths.setArrowType(QtCore.Qt.RightArrow)
		vbox.addWidget(self.btn_paths)

		# 折りたたみコンテンツ
		self.paths_container = QtWidgets.QWidget(root)
		paths_l = QtWidgets.QGridLayout(self.paths_container)
		paths_l.setContentsMargins(0, 0, 0, 0)
		self.lbl_cfg = QtWidgets.QLabel(f"設定: {self.cfg.cfg_path}", self.paths_container)
		self.lbl_log = QtWidgets.QLabel(f"ログ: {self.logman.current_log_path}", self.paths_container)
		paths_l.addWidget(self.lbl_cfg, 0, 0, 1, 1)
		paths_l.addWidget(self.lbl_log, 1, 0, 1, 1)
		self.paths_container.setVisible(False)
		vbox.addWidget(self.paths_container)

		self.btn_paths.toggled.connect(self._on_toggle_paths)

		# LLM 質問エリア
		llm_group = QtWidgets.QGroupBox("LLM質問エリア（表示切替可能）", root)
		llm_l = QtWidgets.QGridLayout(llm_group)

		self.chk_llm_visible = QtWidgets.QCheckBox(LLM_TOGGLE_TEXT, llm_group)
		self.chk_llm_visible.setChecked(False)
		self.chk_llm_visible.toggled.connect(self._update_llm_visible)
		llm_l.addWidget(self.chk_llm_visible, 0, 0, 1, 2)

		self.llm_content = QtWidgets.QWidget(llm_group)
		llm_c_l = QtWidgets.QGridLayout(self.llm_content)

		llm_c_l.addWidget(QtWidgets.QLabel("プリセット:", self.llm_content), 0, 0)
		self.cb_base = QtWidgets.QComboBox(self.llm_content)
		base_keys = ["-"] + list(self.cfg.base_prompts.keys())
		self.cb_base.addItems(base_keys)
		self.cb_base.currentIndexChanged.connect(self._on_base_selected)
		self.cb_base.setMinimumContentsLength(40)
		llm_c_l.addWidget(self.cb_base, 0, 1)

		llm_c_l.addWidget(QtWidgets.QLabel("自由入力:", self.llm_content), 1, 0, QtCore.Qt.AlignTop)
		self.txt_prompt = QtWidgets.QTextEdit(self.llm_content)
		self.txt_prompt.textChanged.connect(self._update_send_button_state)
		llm_c_l.addWidget(self.txt_prompt, 1, 1)

		self.lbl_llm_status = QtWidgets.QLabel("", self.llm_content)
		self.btn_send = QtWidgets.QPushButton(LLM_SEND_TEXT, self.llm_content)
		self.btn_send.clicked.connect(self._on_llm_send)
		llm_c_l.addWidget(self.lbl_llm_status, 2, 0)
		llm_c_l.addWidget(self.btn_send, 2, 1, 1, 1, QtCore.Qt.AlignRight)

		llm_c_l.addWidget(QtWidgets.QLabel("回答:", self.llm_content), 3, 0, QtCore.Qt.AlignTop)
		self.txt_answer = QtWidgets.QPlainTextEdit(self.llm_content)
		self.txt_answer.setReadOnly(False)
		self.txt_answer.setFont(self.font_mono)
		llm_c_l.addWidget(self.txt_answer, 3, 1)

		llm_l.addWidget(self.llm_content, 1, 0, 1, 2)

		vbox.addWidget(llm_group)

		# 初期の活性/非活性
		self._update_llm_visible()
		self._update_send_button_state()

	# ====== イベント/動作 ======
	def _update_send_button_state(self) -> None:
		sel = self.cb_base.currentText()
		free = self.txt_prompt.toPlainText().strip()
		api_ok = bool(self.cfg.openai_api_key)
		if sel == "-" and not free:
			ok = False
		else:
			ok = api_ok
		self.btn_send.setEnabled(ok)
		if not api_ok:
			self.lbl_llm_status.setText("OPENAI_API_KEYが未設定です。環境変数またはyamlで設定してください。")
		else:
			self.lbl_llm_status.setText("")

	def _on_base_selected(self) -> None:
		key = self.cb_base.currentText()
		if key != "-":
			base = self.cfg.base_prompts.get(key, "")
			self.txt_prompt.blockSignals(True)
			self.txt_prompt.setPlainText(base)
			self.txt_prompt.blockSignals(False)
		self._update_send_button_state()

	def _update_llm_visible(self) -> None:
		vis = self.chk_llm_visible.isChecked()
		self.llm_content.setVisible(vis)

	def _on_toggle_paths(self, checked: bool) -> None:
		try:
			self.paths_container.setVisible(checked)
			self.btn_paths.setArrowType(QtCore.Qt.DownArrow if checked else QtCore.Qt.RightArrow)
		except Exception:
			pass

	def _open_log_explorer(self) -> None:
		try:
			if self._log_explorer is None or not self._log_explorer.isVisible():
				self._log_explorer = LogExplorerDialog(self, self.logman)
				self._apply_fonts_recursive(self._log_explorer)
				self._log_explorer.show()
			else:
				self._log_explorer.activateWindow()
				self._log_explorer.raise_()
		except Exception:
			QtWidgets.QMessageBox.critical(self, ERROR_DIALOG_TITLE, get_exception_trace())

	def _open_config_folder(self) -> None:
		path = self.cfg.cfg_path
		try:
			os.startfile(os.path.dirname(path))
		except Exception:
			QtWidgets.QMessageBox.information(self, INFO_DIALOG_TITLE, f"エクスプローラーで開けませんでした。\n{path}")

	def _open_log_folder(self) -> None:
		path = self.cfg.log_dir
		try:
			os.startfile(path)
		except Exception:
			QtWidgets.QMessageBox.information(self, INFO_DIALOG_TITLE, f"エクスプローラーで開けませんでした。\n{path}")

	def _toggle_scale(self) -> None:
		self.scale_level = (self.scale_level + 1) % 3
		self._reconfigure_fonts()
		self.updateGeometry()
		self.adjustSize()

	def _on_register(self) -> None:
		ts = format_timestamp(effective_date(now_jst()))
		out: List[str] = []
		for lane in self.lanes:
			part = lane.get_text_for_log()
			if part:
				out.append(f"{ts} {part}")
		if not out:
			QtWidgets.QMessageBox.information(self, INFO_DIALOG_TITLE, "入力テキストがありません。")
			return
		try:
			self.logman.append_lines(out)
			self._refresh_log_preview()
			self._scroll_preview_to_bottom()
			# 成功時は各行のテキストだけクリア（タグは維持）
			for lane in self.lanes:
				lane.le_text.setText("")
		except Exception:
			QtWidgets.QMessageBox.critical(self, ERROR_DIALOG_TITLE, get_exception_trace())

	def _on_register_lane(self, idx: int) -> None:
		if idx < 0 or idx >= len(self.lanes):
			return
		ts = format_timestamp(effective_date(now_jst()))
		lane = self.lanes[idx]
		part = lane.get_text_for_log()
		if not part:
			QtWidgets.QMessageBox.information(self, INFO_DIALOG_TITLE, "入力テキストがありません。")
			return
		try:
			self.logman.append_lines([f"{ts} {part}"])
			self._refresh_log_preview()
			self._scroll_preview_to_bottom()
			lane.le_text.setText("")
		except Exception:
			QtWidgets.QMessageBox.critical(self, ERROR_DIALOG_TITLE, get_exception_trace())

	def _refresh_log_preview(self) -> None:
		try:
			self.txt_preview.blockSignals(True)
			self.txt_preview.setPlainText(self.logman.read_text())
		finally:
			self.txt_preview.blockSignals(False)

	def _scroll_preview_to_bottom(self) -> None:
		try:
			self.txt_preview.moveCursor(QtGui.QTextCursor.End)
			self.txt_preview.ensureCursorVisible()
		except Exception:
			pass

	def _update_datetime(self) -> None:
		try:
			now = now_jst()
			self.lbl_now.setText(now.strftime("%Y-%m-%d %H:%M:%S"))
			# ロールオーバ判定
			if self.logman.rollover_if_needed():
				self.lbl_log.setText(f"ログ: {self.logman.current_log_path}")
				self._refresh_log_preview()
		except Exception:
			pass

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

		prompt = self.txt_prompt.toPlainText().strip()
		if not prompt:
			key = self.cb_base.currentText()
			prompt = self.cfg.base_prompts.get(key, "").strip()

		self.btn_send.setEnabled(False)
		self.btn_send.setText("送信中...")
		self.lbl_llm_status.setText("回答待ち...")
		self.txt_answer.setPlainText("")

		def worker() -> None:
			try:
				ans = self.llm_cli.ask(prompt, max_tokens=self.cfg.llm_max_tokens)  # type: ignore[union-attr]
			except Exception as e:  # noqa: BLE001
				ans = f"[エラー] {e}"

			def done() -> None:
				self.lbl_llm_status.setText("")
				self.btn_send.setEnabled(True)
				self.btn_send.setText(LLM_SEND_TEXT)
				ts = format_timestamp(effective_date(now_jst()))
				user = self.cfg.user_name
				# 表示
				self.txt_answer.setPlainText(ans)
				# ログ追記（LLMタグで記録者名入り）
				log_line = f"{ts} {bracket_pad('LLM')} {bracket_pad(user)} Q:{prompt} / A:{ans}"
				try:
					self.logman.append_lines([log_line])
					self._refresh_log_preview()
				except Exception:
					pass

			QtCore.QTimer.singleShot(0, done)

		threading.Thread(target=worker, daemon=True).start()

	# ====== 終了 ======
	def _on_exit(self) -> None:
		try:
			self.close()
		except Exception:
			os._exit(0)  # noqa: WPS437

	# ====== 常に手前 ======
	def _on_topmost_toggled(self, checked: bool) -> None:
		self.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint, bool(checked))
		# 反映のため再表示
		self.show()


def run_app() -> None:
	app = QtWidgets.QApplication(sys.argv)
	win = MainWindow()
	win.show()
	try:
		sys.exit(app.exec())
	except SystemExit:
		pass


# 4. メイン処理
if __name__ == "__main__":
	try:
		run_app()
	except Exception:
		# 例外時は標準エラーへ出す（GUIダイアログが使えない可能性もあるため）
		sys.stderr.write(get_exception_trace())
