"""
LiveChatter - YouTube Live Chat Reader & Summarizer

LiveChatter is a PyQt6 application that:
- Reads live YouTube chat (pytchat; optional chat-downloader fallback)
- Two chat modes: "Standard (read messages)" and "Summaries (periodic AI summaries)"
- Summary provider: OpenAI
- TTS options: System (accessible_output2: NVDA/JAWS/SAPI5), OpenAI TTS, Google Cloud TTS, Amazon Polly, ElevenLabs
- Sound packs via a 'sounds' directory
- Options dialog for API keys and sound settings
- Config persisted as JSON in a .livechatter folder under the current working directory

NOTE: External services require API keys and libraries installed.
"""
from __future__ import annotations

import sys
import os
import json
import re
import threading
import time
from typing import Optional, List, Dict, Any
import urllib.request, urllib.error, urllib.parse
try:
    import win32com.client as win32com_client
    import pythoncom
except Exception:
    win32com_client = None
    pythoncom = None

# -------- Optional imports with graceful fallbacks --------
try:
    import pytchat
except Exception:
    pytchat = None

try:
    from chat_downloader import ChatDownloader
except Exception:
    ChatDownloader = None

try:
    from accessible_output2.outputs.auto import Auto as AO2Auto
    from accessible_output2.outputs import sapi5
    # --- FIX FOR PYINSTALLER: Add dummy imports to find hidden modules ---
    try:
        from accessible_output2.outputs import nvda, jaws
    except ImportError:
        pass # Fine if not on Windows
    # -------------------------------------------------------------------
except Exception:
    AO2Auto = None
    sapi5 = None

# Optional: Windows SAPI5 voice control (pywin32)
try:
    import win32com.client as win32  # type: ignore
    import pythoncom  # type: ignore
except Exception:
    win32 = None
    pythoncom = None

try:
    import openai
except Exception:
    openai = None

try:
    import requests
except Exception:
    requests = None

try:
    from google.cloud import texttospeech as gcloud_tts
except Exception:
    gcloud_tts = None

try:
    import boto3
except Exception:
    boto3 = None

try:
    import elevenlabs
except Exception:
    elevenlabs = None

from PyQt6 import QtWidgets, QtCore, QtMultimedia

# ----------------- Config & Constants -----------------
APP_NAME = "LiveChatter"
CONFIG_DIR = os.path.join(os.getcwd(), ".livechatter")
CONFIG_PATH = os.path.join(CONFIG_DIR, "config.json")
SOUND_DIR = os.path.join(os.getcwd(), "sounds")


CHAT_MODES = [
    "Standard (read messages)",
    "Summaries (periodic)"
]

SUMMARY_PROVIDERS = [
    "OpenAI",
]

TTS_OPTIONS = [
    "System (screen reader/SAPI5)",
    "OpenAI TTS",
    "Google Cloud TTS",
    "Amazon Polly",
    "ElevenLabs",
]

SYSTEM_TTS_ENGINES = [
    "SAPI5 (Windows)",
    "Screen reader (Auto)",
]

DEFAULT_CONFIG = {
    "openai_api_key": "",
    "google_cloud_tts_api_key": "",
    "aws_access_key_id": "",
    "aws_secret_access_key": "",
    "aws_region": "us-east-1",
    "elevenlabs_api_key": "",
    "tts_voice": "",  # For SAPI5: voice description; for cloud: provider-specific id
    "sound_pack": "Default",
    "chat_mode": CHAT_MODES[0],
    "summary_provider": SUMMARY_PROVIDERS[0],
    "tts_option": TTS_OPTIONS[0],
    "system_tts_backend": SYSTEM_TTS_ENGINES[0],
    "summary_interval_minutes": 5,
    "quick_summary_count": 50,
    # Window state
    "window_x": None,
    "window_y": None,
    "window_w": 900,
    "window_h": 600,
    "window_max": False,
}

# ----------------- Config Helpers -----------------
def ensure_config_dir():
    os.makedirs(CONFIG_DIR, exist_ok=True)
    os.makedirs(SOUND_DIR, exist_ok=True)

def load_config() -> Dict[str, Any]:
    ensure_config_dir()
    cfg = DEFAULT_CONFIG.copy()
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            cfg.update(loaded)
        except Exception:
            pass
    return cfg

def save_config(cfg: Dict[str, Any]):
    ensure_config_dir()
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)

# ----------------- Sound Management -----------------
def list_sound_packs() -> List[str]:
    """Scans the sound directory for available packs."""
    packs = ["None"]
    if not os.path.isdir(SOUND_DIR):
        return packs
    for item in os.listdir(SOUND_DIR):
        if os.path.isdir(os.path.join(SOUND_DIR, item)):
            packs.append(item)
    return packs

class SoundManager:
    SOUND_FILES = ["chat", "donation", "error", "moderator", "start", "stop", "summary", "verified"]

    def __init__(self, cfg: Dict[str, Any]):
        self.cfg = cfg
        self._effects: Dict[str, QtMultimedia.QSoundEffect] = {}
        self.load_pack()

    def load_pack(self):
        self._effects = {}
        pack_name = self.cfg.get("sound_pack", "None")
        if pack_name == "None":
            return
        
        pack_path = os.path.join(SOUND_DIR, pack_name)
        if not os.path.isdir(pack_path):
            return
        
        for name in self.SOUND_FILES:
            sound_path = os.path.join(pack_path, f"{name}.wav")
            if os.path.exists(sound_path):
                effect = QtMultimedia.QSoundEffect()
                effect.setSource(QtCore.QUrl.fromLocalFile(sound_path))
                self._effects[name] = effect

    def play(self, sound_name: str):
        if sound_name in self._effects:
            self._effects[sound_name].play()


# ----------------- Utilities -----------------
def extract_youtube_video_id(url_or_id: str) -> Optional[str]:
    s = url_or_id.strip()
    if re.fullmatch(r"[A-Za-z0-9_-]{11}", s):
        return s
    for p in [r"youtu\.be/([A-Za-z0-9_-]{11})", r"v=([A-Za-z0-9_-]{11})", r"live/([A-Za-z0-9_-]{11})"]:
        m = re.search(p, s)
        if m:
            return m.group(1)
    return None

# ----------------- Native stderr suppression (for noisy SAPI/eSpeak outputs) -----------------
class _SuppressStderrFD:
    def __enter__(self):
        import os
        self._null = open(os.devnull, 'w')
        self._stderr_fd = os.dup(2)
        os.dup2(self._null.fileno(), 2)
        return self
    def __exit__(self, exc_type, exc, tb):
        import os
        try:
            os.dup2(self._stderr_fd, 2)
        finally:
            try:
                os.close(self._stderr_fd)
            except Exception:
                pass
            try:
                self._null.close()
            except Exception:
                pass

# ----------------- TTS (System primary; cloud stubs) -----------------
class TTSManager:
    def __init__(self, cfg: Dict[str, Any]):
        self.cfg = cfg
        self.system_out = None  # for Auto or accessible_output2.SAPI5
        self._init_system_tts()

    def _init_system_tts(self):
        choice = self.cfg.get("system_tts_backend", SYSTEM_TTS_ENGINES[0])
        # If user explicitly picked SAPI5, do NOT fall back to Auto (avoids COM probing spam)
        if choice.startswith("SAPI5"):
            # Prefer native pywin32 SAPI control (lets us pick voice)
            if win32 is not None and pythoncom is not None and sys.platform.startswith("win"):
                # We'll use thread-per-speak with COM init, so no object here
                self.system_out = "SAPI5_DIRECT"
                return
            # Fallback to accessible_output2's SAPI5 class if available
            if sapi5 is not None:
                try:
                    self.system_out = sapi5.SAPI5()
                    return
                except Exception:
                    self.system_out = None
            # Else none; we'll print fallback in speak()
            return
        # Screen reader Auto
        if AO2Auto is not None:
            try:
                self.system_out = AO2Auto()
            except Exception:
                self.system_out = None

    # ---- SAPI5 helpers ----
    @staticmethod
    def list_sapi5_voices() -> List[str]:
        names: List[str] = []
        if win32 is None or pythoncom is None or not sys.platform.startswith("win"):
            return names
        try:
            pythoncom.CoInitialize()
            sp = win32.Dispatch("SAPI.SpVoice")
            tokens = sp.GetVoices()
            for i in range(tokens.Count):
                names.append(tokens.Item(i).GetDescription())
        except Exception:
            pass
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        return names

    @staticmethod
    def _sapi5_say(text: str, voice_desc: str | None):
        # Speak in a background thread with its own COM apartment
        def _run():
            try:
                pythoncom.CoInitialize()
                sp = win32.Dispatch("SAPI.SpVoice")
                # Select voice by description if requested
                if voice_desc:
                    try:
                        toks = sp.GetVoices()
                        for i in range(toks.Count):
                            tok = toks.Item(i)
                            if tok.GetDescription() == voice_desc:
                                sp.Voice = tok
                                break
                    except Exception:
                        pass
                # 1 = SVSFlagsAsync would be async; use 0 for blocking but in this worker thread
                with _SuppressStderrFD():
                    sp.Speak(text, 0)
            except Exception:
                pass
            finally:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
        th = threading.Thread(target=_run, daemon=True)
        th.start()

    # ---- Public speak ----
    def speak(self, text: str):
        tts_opt = self.cfg.get("tts_option", TTS_OPTIONS[0])
        if tts_opt == "System (screen reader/SAPI5)":
            backend = self.cfg.get("system_tts_backend", SYSTEM_TTS_ENGINES[0])
            if backend.startswith("SAPI5"):
                if win32 is not None and pythoncom is not None and sys.platform.startswith("win"):
                    self._sapi5_say(text, self.cfg.get("tts_voice") or None)
                    return
                if self.system_out and self.system_out != "SAPI5_DIRECT":
                    try:
                        self.system_out.speak(text)
                        return
                    except Exception:
                        pass
                print("[System SAPI5 unavailable]", text)
                return
            # Screen reader (Auto)
            if self.system_out and self.system_out != "SAPI5_DIRECT":
                try:
                    self.system_out.speak(text)
                    return
                except Exception:
                    pass
            print("[System Auto TTS unavailable]", text)
            return
        # Cloud stubs (implement if you wire providers)
        print(f"[{tts_opt} stub]", text)

# ---- Voice listing helpers for cloud TTS providers ----

def list_sapi5_voices() -> List[str]:
    if not (sys.platform.startswith("win") and win32com_client and pythoncom):
        return []
    try:
        pythoncom.CoInitialize()
        try:
            spk = win32com_client.Dispatch("SAPI.SpVoice")
            voices = spk.GetVoices()
            names: List[str] = []
            for i in range(voices.Count):
                v = voices.Item(i)
                try:
                    names.append(str(v.GetDescription()))
                except Exception:
                    pass
            return names
        finally:
            pythoncom.CoUninitialize()
    except Exception:
        return []


def list_openai_tts_voices() -> List[str]:
    # OpenAI TTS voices are not enumerated by API; provide a common set.
    return [
        "alloy", "verse", "aria", "coral", "sage", "nova"
    ]

def list_gcloud_tts_voices(api_key: str) -> List[tuple[str, str]]:
    """Return list of (name, langs) from Google Cloud TTS via REST using an API key.
    If the key is missing or an error occurs, return []."""
    if not api_key:
        return []
    try:
        url = "https://texttospeech.googleapis.com/v1/voices?key=" + urllib.parse.quote(api_key)
        with urllib.request.urlopen(url, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        out: List[tuple[str, str]] = []
        for v in data.get("voices", []):
            name = v.get("name", "")
            langs = ",".join(v.get("languageCodes", []) or [])
            if name:
                out.append((name, langs))
        return out
    except Exception:
        return []

def list_polly_voices(aws_key: str, aws_secret: str, region: str) -> List[str]:
    if not (boto3 and aws_key and aws_secret and region):
        return []
    try:
        client = boto3.client(
            "polly",
            region_name=region,
            aws_access_key_id=aws_key,
            aws_secret_access_key=aws_secret,
        )
        resp = client.describe_voices()
        voices = [v.get("Id") for v in resp.get("Voices", []) if v.get("Id")]
        return sorted(set(voices))
    except Exception:
        return []

def list_elevenlabs_voices(api_key: str) -> List[tuple[str, str]]:
    if not (elevenlabs and api_key):
        return []
    try:
        # Try modern and legacy client shapes
        try:
            if hasattr(elevenlabs, "set_api_key"):
                elevenlabs.set_api_key(api_key)
            voices = getattr(elevenlabs, "voices", None)
            if callable(voices):
                vs = voices()
            else:
                vs = getattr(elevenlabs, "Voices").list()
        except Exception:
            return []
        out: List[tuple[str, str]] = []
        for v in vs:
            name = getattr(v, "name", None) or (v.get("name") if isinstance(v, dict) else None)
            vid = getattr(v, "voice_id", None) or (v.get("voice_id") if isinstance(v, dict) else None) or (v.get("id") if isinstance(v, dict) else None)
            if name and vid:
                out.append((str(name), str(vid)))
        return out
    except Exception:
        return []

# ----------------- Summarizer -----------------
class Summarizer:
    def __init__(self, cfg: Dict[str, Any]):
        self.cfg = cfg

    def summarize(self, messages: List[Dict[str, str]]) -> str:
        if not messages:
            return "It's been quiet. No new messages to summarize."

        # Construct the prompt based on the number of messages
        formatted_chat = "\n".join([f"{m.get('author', 'User')}: {m.get('text', '')}" for m in messages])
        
        # Re-introducing the humor to the prompt
        base_prompt = (
            "You are a witty and humorous assistant summarizing a YouTube live stream chat. "
            "Your summary will be read out loud by a text-to-speech engine, so adopt a conversational and slightly comedic tone. "
            "Be concise and keep your summary to a few sentences."
        )

        if len(messages) < 5:
            task_prompt = (
                "The chat is very quiet. Make a brief, funny comment about the silence, "
                "and then just read out the few messages that have appeared."
            )
        else:
            task_prompt = (
                "The chat is active. Do not list every message. Instead, capture the main vibe. "
                "Identify the key topics being discussed, mention any highlights or funny moments, "
                "and give a general sense of the conversation."
            )
        
        full_prompt = f"{base_prompt}\n\n{task_prompt}\n\nHere are the recent messages:\n{formatted_chat}"

        # Call the appropriate AI provider
        return self._summarize_with_openai(full_prompt)

    def _summarize_with_openai(self, prompt_text: str) -> str:
        if not openai:
            return "(OpenAI library not installed. Please run: pip install openai)"
        
        api_key = self.cfg.get("openai_api_key")
        if not api_key:
            return "(OpenAI API key is not set in Options)"

        try:
            client = openai.OpenAI(api_key=api_key)
            response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are a helpful and witty assistant for summarizing YouTube chat."},
                    {"role": "user", "content": prompt_text}
                ],
                temperature=0.7,
                max_tokens=150,
            )
            summary = response.choices[0].message.content
            return summary.strip() if summary else "(OpenAI returned an empty summary)"
        except Exception as e:
            return f"(OpenAI API error: {e})"

# ----------------- Chat Reader Thread -----------------
class ChatReader(QtCore.QThread):
    message_received = QtCore.pyqtSignal(dict)
    error_message = QtCore.pyqtSignal(str)
    stopped = QtCore.pyqtSignal()

    def __init__(self, video_id: str, parent=None):
        super().__init__(parent)
        self.video_id = video_id
        self._stop = threading.Event()

    def run(self):
        # Prefer pytchat, with a new auto-reconnect loop
        if pytchat:
            while not self._stop.is_set():
                chat = None
                try:
                    chat = pytchat.create(video_id=self.video_id, interruptable=False)
                    # Loop while the connection is alive
                    while chat.is_alive() and not self._stop.is_set():
                        for c in chat.get().items:
                            # Pass more metadata for sound events
                            msg = {
                                "author": c.author.name, 
                                "text": c.message,
                                "type": c.type, # e.g., 'textMessage', 'superChat'
                                "is_moderator": c.author.isChatModerator,
                                "is_verified": c.author.isVerified 
                            }
                            self.message_received.emit(msg)
                        time.sleep(0.5)
                    
                    if self._stop.is_set():
                        break 
                    
                    self.error_message.emit("Chat connection lost. Reconnecting in 10 seconds...")

                except Exception as e:
                    self.error_message.emit(f"Chat reader error: {e}. Retrying in 10 seconds...")
                finally:
                    try:
                        if chat:
                            chat.terminate()
                    except Exception:
                        pass
                
                if not self._stop.is_set():
                    time.sleep(10)
            
            self.stopped.emit()
            return

        # Fallback: chat-downloader (provides less metadata)
        if ChatDownloader:
            try:
                url = f"https://www.youtube.com/watch?v={self.video_id}"
                downloader = ChatDownloader().get_chat(url)
                for msg in downloader:
                    if self._stop.is_set():
                        break
                    if msg.get("message_type") == "text_message":
                        m = {
                            "author": msg.get("author", {}).get("name", "?"), 
                            "text": msg.get("message", ""),
                            "type": "textMessage",
                            "is_moderator": False,
                            "is_verified": False
                        }
                        self.message_received.emit(m)
                self.stopped.emit()
            except Exception as e:
                emsg = str(e)
                if "disabled" in emsg.lower():
                    self.error_message.emit("Live chat is disabled for this stream.")
                else:
                    self.error_message.emit(f"Chat downloader error: {emsg}")
                self.stopped.emit()
        else:
            self.error_message.emit("No chat backend is available. Install 'pytchat' or 'chat-downloader'.")
            self.stopped.emit()

    def stop(self):
        self._stop.set()

# ----------------- Options Dialog -----------------
class OptionsDialog(QtWidgets.QDialog):
    def __init__(self, cfg: Dict[str, Any], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Options")
        self.cfg = cfg
        layout = QtWidgets.QFormLayout(self)

        self.openai_key = QtWidgets.QLineEdit(self.cfg.get("openai_api_key", ""))
        self.openai_key.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)
        layout.addRow("OpenAI API key:", self.openai_key)

        self.gcloud_key = QtWidgets.QLineEdit(self.cfg.get("google_cloud_tts_api_key", ""))
        self.gcloud_key.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)
        layout.addRow("Google Cloud TTS API key:", self.gcloud_key)

        self.aws_key = QtWidgets.QLineEdit(self.cfg.get("aws_access_key_id", ""))
        self.aws_secret = QtWidgets.QLineEdit(self.cfg.get("aws_secret_access_key", ""))
        self.aws_secret.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)
        self.aws_region = QtWidgets.QLineEdit(self.cfg.get("aws_region", "us-east-1"))
        layout.addRow("AWS Access Key:", self.aws_key)
        layout.addRow("AWS Secret Key:", self.aws_secret)
        layout.addRow("AWS Region:", self.aws_region)

        self.eleven_key = QtWidgets.QLineEdit(self.cfg.get("elevenlabs_api_key", ""))
        self.eleven_key.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)
        layout.addRow("ElevenLabs API key:", self.eleven_key)

        self.voice = QtWidgets.QLineEdit(self.cfg.get("tts_voice", ""))
        layout.addRow("Preferred voice (optional):", self.voice)

        self.system_tts_mode = QtWidgets.QComboBox()
        self.system_tts_mode.addItems(SYSTEM_TTS_ENGINES)
        self.system_tts_mode.setCurrentText(self.cfg.get("system_tts_backend", SYSTEM_TTS_ENGINES[0]))
        layout.addRow("System TTS engine:", self.system_tts_mode)

        self.sound_pack_combo = QtWidgets.QComboBox()
        self.sound_pack_combo.addItems(list_sound_packs())
        self.sound_pack_combo.setCurrentText(self.cfg.get("sound_pack", "Default"))
        layout.addRow("Sound pack:", self.sound_pack_combo)

        btns = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.StandardButton.Save | QtWidgets.QDialogButtonBox.StandardButton.Cancel
        )
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addRow(btns)

    def get_updated_config(self) -> Dict[str, Any]:
        newcfg = self.cfg.copy()
        newcfg.update({
            "openai_api_key": self.openai_key.text().strip(),
            "google_cloud_tts_api_key": self.gcloud_key.text().strip(),
            "aws_access_key_id": self.aws_key.text().strip(),
            "aws_secret_access_key": self.aws_secret.text().strip(),
            "aws_region": self.aws_region.text().strip() or "us-east-1",
            "elevenlabs_api_key": self.eleven_key.text().strip(),
            "tts_voice": self.voice.text().strip(),
            "sound_pack": self.sound_pack_combo.currentText(),
            "system_tts_backend": self.system_tts_mode.currentText(),
        })
        return newcfg

# ----------------- Main Window -----------------
class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.cfg = load_config()
        self.setWindowTitle(APP_NAME)
        self._restore_window_geometry()

        self.tts = TTSManager(self.cfg)
        self.sound_manager = SoundManager(self.cfg)
        self.summarizer = Summarizer(self.cfg)
        self.pending_messages: List[Dict[str, str]] = []
        self.last_summary_ts = time.time()
        self.reader: Optional[ChatReader] = None

        # Widgets
        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        gl = QtWidgets.QGridLayout(central)

        self.url_edit = QtWidgets.QLineEdit()
        self.url_edit.setPlaceholderText("Paste YouTube live URL or video ID…")
        gl.addWidget(QtWidgets.QLabel("Live stream URL / ID:"), 0, 0)
        gl.addWidget(self.url_edit, 0, 1, 1, 3)

        self.chat_mode = QtWidgets.QComboBox()
        self.chat_mode.addItems(CHAT_MODES)
        self.chat_mode.setCurrentText(self.cfg.get("chat_mode", CHAT_MODES[0]))
        self.chat_mode.currentTextChanged.connect(self._update_cfg)
        gl.addWidget(QtWidgets.QLabel("Chat mode:"), 1, 0)
        gl.addWidget(self.chat_mode, 1, 1, 1, 3)

        self.tts_option = QtWidgets.QComboBox()
        self.tts_option.addItems(TTS_OPTIONS)
        self.tts_option.setCurrentText(self.cfg.get("tts_option", TTS_OPTIONS[0]))
        self.tts_option.currentTextChanged.connect(self._on_tts_option_changed)
        gl.addWidget(QtWidgets.QLabel("TTS:"), 2, 0)
        gl.addWidget(self.tts_option, 2, 1)

        self.system_tts_mode = QtWidgets.QComboBox()
        self.system_tts_mode.addItems(SYSTEM_TTS_ENGINES)
        self.system_tts_mode.setCurrentText(self.cfg.get("system_tts_backend", SYSTEM_TTS_ENGINES[0]))
        self.system_tts_mode.currentTextChanged.connect(self._on_system_tts_changed)
        gl.addWidget(QtWidgets.QLabel("System TTS engine:"), 2, 2)
        gl.addWidget(self.system_tts_mode, 2, 3)

        self.voice_combo = QtWidgets.QComboBox()
        self.voice_combo.setToolTip("Choose a voice for the selected TTS provider")
        self.voice_combo.currentTextChanged.connect(self._on_voice_changed)
        gl.addWidget(QtWidgets.QLabel("Voice:"), 3, 0)
        gl.addWidget(self.voice_combo, 3, 1, 1, 3)
        self._voice_map: Dict[str, str] = {}
        self._reload_voice_list()

        self.summary_interval = QtWidgets.QSpinBox()
        self.summary_interval.setRange(1, 120)
        self.summary_interval.setValue(int(self.cfg.get("summary_interval_minutes", 5)))
        self.summary_interval.valueChanged.connect(self._update_cfg)
        gl.addWidget(QtWidgets.QLabel("Summary interval (min):"), 4, 0)
        gl.addWidget(self.summary_interval, 4, 1)

        self.options_btn = QtWidgets.QPushButton("Options…")
        self.options_btn.clicked.connect(self.open_options)
        gl.addWidget(self.options_btn, 4, 2)

        self.start_btn = QtWidgets.QPushButton("Start")
        self.stop_btn = QtWidgets.QPushButton("Stop")
        self.stop_btn.setEnabled(False)
        gl.addWidget(self.start_btn, 4, 3)

        self.chat_view = QtWidgets.QTextEdit()
        self.chat_view.setReadOnly(True)
        gl.addWidget(self.chat_view, 5, 0, 1, 4)

        self.summary_timer = QtCore.QTimer(self)
        self.summary_timer.timeout.connect(self._maybe_do_summary)
        self.summary_timer.start(5000)

        self.start_btn.clicked.connect(self._on_start)
        self.stop_btn.clicked.connect(self._on_stop)

        # Status bar tools
        self.statusBar().addPermanentWidget(self.stop_btn)
        self.test_tts_btn = QtWidgets.QPushButton("Test TTS")
        self.test_tts_btn.setToolTip("Speak a short sample using the current TTS settings")
        self.test_tts_btn.clicked.connect(self._on_test_tts)
        self.statusBar().addPermanentWidget(self.test_tts_btn)

        self.quick_summary_btn = QtWidgets.QPushButton("Quick Summary")
        self.quick_summary_btn.setToolTip("Summarize the most recent chat messages now")
        self.quick_summary_btn.clicked.connect(self._on_quick_summary)
        self.statusBar().addPermanentWidget(self.quick_summary_btn)

        self.quick_summary_count_sb = QtWidgets.QSpinBox()
        self.quick_summary_count_sb.setRange(5, 500)
        self.quick_summary_count_sb.setValue(int(self.cfg.get("quick_summary_count", 50)))
        self.quick_summary_count_sb.setToolTip("Number of recent messages to include in Quick Summary")
        self.quick_summary_count_sb.valueChanged.connect(self._on_quick_summary_count_changed)
        self.statusBar().addPermanentWidget(QtWidgets.QLabel("Msgs:"))
        self.statusBar().addPermanentWidget(self.quick_summary_count_sb)

        self._update_system_tts_enabled()
        self._update_voice_enabled()

    def _restore_window_geometry(self):
        w = int(self.cfg.get("window_w", 900) or 900)
        h = int(self.cfg.get("window_h", 600) or 600)
        x = self.cfg.get("window_x"); y = self.cfg.get("window_y")
        if x is not None and y is not None:
            try:
                self.setGeometry(int(x), int(y), w, h)
            except Exception:
                self.resize(w, h)
        else:
            self.resize(w, h)
        if self.cfg.get("window_max", False):
            self.showMaximized()

    def _save_window_geometry(self):
        if self.isMaximized():
            self.cfg["window_max"] = True
            rect = self.normalGeometry()
            self.cfg["window_x"] = rect.x(); self.cfg["window_y"] = rect.y()
            self.cfg["window_w"] = rect.width(); self.cfg["window_h"] = rect.height()
        else:
            self.cfg["window_max"] = False
            g = self.geometry()
            self.cfg["window_x"] = g.x(); self.cfg["window_y"] = g.y()
            self.cfg["window_w"] = g.width(); self.cfg["window_h"] = g.height()
        save_config(self.cfg)

    def closeEvent(self, event):
        self._save_window_geometry()
        super().closeEvent(event)

    def _update_cfg(self, *_):
        self.cfg["chat_mode"] = self.chat_mode.currentText()
        self.cfg["summary_provider"] = "OpenAI"
        self.cfg["tts_option"] = self.tts_option.currentText()
        self.cfg["summary_interval_minutes"] = int(self.summary_interval.value())
        save_config(self.cfg)

    def _on_tts_option_changed(self, txt: str):
        self.cfg["tts_option"] = txt
        self._update_cfg()
        self.cfg["tts_voice"] = ""
        save_config(self.cfg)
        self._update_system_tts_enabled()
        self._update_voice_enabled()
        self._reload_voice_list()
        self.tts = TTSManager(self.cfg)

    def _on_system_tts_changed(self, txt: str):
        self.cfg["system_tts_backend"] = txt
        save_config(self.cfg)
        self.tts = TTSManager(self.cfg)
        self._update_voice_enabled()
        self._reload_voice_list()

    def _reload_voice_list(self):
        self.voice_combo.blockSignals(True)
        self.voice_combo.clear()
        self._voice_map = {}
        opt = self.tts_option.currentText()
        if opt == "System (screen reader/SAPI5)":
            if self.system_tts_mode.currentText().startswith("SAPI5") and sys.platform.startswith("win"):
                voices = list_sapi5_voices()
                if voices:
                    self.voice_combo.addItems(voices)
                    wanted = self.cfg.get("tts_voice", "")
                    if wanted and wanted in voices:
                        self.voice_combo.setCurrentIndex(voices.index(wanted))
                    self.voice_combo.setEnabled(True)
                else:
                    self.voice_combo.addItem("(No SAPI5 voices found)")
                    self.voice_combo.setEnabled(False)
            else:
                self.voice_combo.addItem("(Voice controlled by screen reader)")
                self.voice_combo.setEnabled(False)
        elif opt == "OpenAI TTS":
            voices = list_openai_tts_voices()
            if voices:
                for v in voices:
                    self.voice_combo.addItem(v)
                    self._voice_map[v] = v
                sel = self.cfg.get("tts_voice", "")
                if sel and sel in voices:
                    self.voice_combo.setCurrentIndex(voices.index(sel))
                self.voice_combo.setEnabled(True)
            else:
                self.voice_combo.addItem("(No OpenAI TTS voices available)")
                self.voice_combo.setEnabled(False)
        elif opt == "Google Cloud TTS":
            api_key = self.cfg.get("google_cloud_tts_api_key", "")
            if not api_key:
                self.voice_combo.addItem("(Enter Google Cloud TTS API key in Options)")
                self.voice_combo.setEnabled(False)
            else:
                items = list_gcloud_tts_voices(api_key)
                if items:
                    for name, langs in items:
                        disp = f"{name} [{langs}]" if langs else name
                        self.voice_combo.addItem(disp)
                        self._voice_map[disp] = name
                    sel = self.cfg.get("tts_voice", "")
                    if sel:
                        for i in range(self.voice_combo.count()):
                            d = self.voice_combo.itemText(i)
                            if self._voice_map.get(d) == sel:
                                self.voice_combo.setCurrentIndex(i)
                                break
                    self.voice_combo.setEnabled(True)
                else:
                    self.voice_combo.addItem("(No Google Cloud voices found)")
                    self.voice_combo.setEnabled(False)
        elif opt == "Amazon Polly":
            items = list_polly_voices(
                self.cfg.get("aws_access_key_id", ""),
                self.cfg.get("aws_secret_access_key", ""),
                self.cfg.get("aws_region", "us-east-1"),
            )
            if items:
                for vid in items:
                    self.voice_combo.addItem(vid)
                    self._voice_map[vid] = vid
                sel = self.cfg.get("tts_voice", "")
                if sel and sel in items:
                    self.voice_combo.setCurrentIndex(items.index(sel))
                self.voice_combo.setEnabled(True)
            else:
                self.voice_combo.addItem("(No Polly voices found or AWS keys missing)")
                self.voice_combo.setEnabled(False)
        elif opt == "ElevenLabs":
            key = self.cfg.get("elevenlabs_api_key", "")
            if not key:
                self.voice_combo.addItem("(Enter ElevenLabs API key in Options)")
                self.voice_combo.setEnabled(False)
            else:
                pairs = list_elevenlabs_voices(key)
                if pairs:
                    for name, vid in pairs:
                        self.voice_combo.addItem(name)
                        self._voice_map[name] = vid
                    sel = self.cfg.get("tts_voice", "")
                    if sel:
                        for i in range(self.voice_combo.count()):
                            d = self.voice_combo.itemText(i)
                            if self._voice_map.get(d) == sel:
                                self.voice_combo.setCurrentIndex(i)
                                break
                    self.voice_combo.setEnabled(True)
                else:
                    self.voice_combo.addItem("(No ElevenLabs voices found)")
                    self.voice_combo.setEnabled(False)
        else:
            self.voice_combo.addItem("(Select a TTS provider)")
            self.voice_combo.setEnabled(False)
        self.voice_combo.blockSignals(False)

    def _update_voice_enabled(self):
        pass

    def _on_voice_changed(self, display: str):
        if not display or display.startswith("("):
            return
        internal = self._voice_map.get(display, display)
        self.cfg["tts_voice"] = internal
        save_config(self.cfg)

    def open_options(self):
        dlg = OptionsDialog(self.cfg, self)
        if dlg.exec() == QtWidgets.QDialog.DialogCode.Accepted:
            self.cfg = dlg.get_updated_config()
            save_config(self.cfg)
            self.tts = TTSManager(self.cfg)
            self.sound_manager = SoundManager(self.cfg)
            self.summarizer = Summarizer(self.cfg)
            self.system_tts_mode.setCurrentText(self.cfg.get("system_tts_backend", SYSTEM_TTS_ENGINES[0]))
            self._update_system_tts_enabled()
            self._reload_voice_list()
            QtWidgets.QMessageBox.information(self, APP_NAME, "Options saved and applied.")

    def _on_start(self):
        url = self.url_edit.text().strip()
        vid = extract_youtube_video_id(url)
        if not vid:
            QtWidgets.QMessageBox.warning(self, APP_NAME, "Please enter a valid YouTube live URL or video ID.")
            return
        self.sound_manager.play("start")
        self.chat_view.append("Starting chat for video: " + vid)
        self.reader = ChatReader(vid)
        self.reader.message_received.connect(self._on_message)
        self.reader.error_message.connect(self._on_error)
        self.reader.stopped.connect(self._on_stopped)
        self.reader.start()
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.pending_messages.clear()
        self.last_summary_ts = time.time()

    def _on_stop(self):
        self.sound_manager.play("stop")
        if self.reader:
            self.reader.stop()
            self.reader.wait(2000)
            self.reader = None
        self.chat_view.append("[Chat stopped]")
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)

    def _on_stopped(self):
        self.chat_view.append("[Chat thread ended]")
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)

    def _on_message(self, msg: Dict[str, Any]):
        author = msg.get("author", "?")
        text = msg.get("text", "")
        self.pending_messages.append(msg)
        
        if self.chat_mode.currentText() == CHAT_MODES[0]:
            # Play sounds based on message type
            msg_type = msg.get("type")
            if msg_type in ["superChat", "superSticker"]:
                self.sound_manager.play("donation")
            elif msg.get("is_moderator"):
                self.sound_manager.play("moderator")
            elif msg.get("is_verified"):
                self.sound_manager.play("verified")
            else:
                self.sound_manager.play("chat")

            line = f"{author}: {text}"
            self.chat_view.append(line)
            self.tts.speak(line)
        else:
            self.chat_view.append(f"[msg] {author}: {text}")

    def _on_error(self, err: str):
        self.sound_manager.play("error")
        self.chat_view.append(f"[Error] {err}")
        if "reconnecting" in err.lower() or "connection lost" in err.lower():
            self.statusBar().showMessage(err, 5000)
        else:
            QtWidgets.QMessageBox.warning(self, APP_NAME, err)

    def _on_quick_summary(self):
        if not self.pending_messages:
            msg = "No recent messages to summarize."
            self.chat_view.append("[Quick Summary] " + msg)
            try:
                self.tts.speak(msg)
            except Exception:
                pass
            return
        
        self.sound_manager.play("summary")
        count = int(self.quick_summary_count_sb.value()) if hasattr(self, "quick_summary_count_sb") else int(self.cfg.get("quick_summary_count", 50))
        recent = self.pending_messages[-count:]
        summary = self.summarizer.summarize(recent)
        self.chat_view.append("[Quick Summary]\n" + summary + "\n")
        try:
            self.tts.speak(summary)
        except Exception:
            pass

    def _on_quick_summary_count_changed(self, val: int):
        self.cfg["quick_summary_count"] = int(val)
        save_config(self.cfg)

    def _on_test_tts(self):
        provider = self.cfg.get("tts_option", self.tts_option.currentText())
        engine = self.cfg.get("system_tts_backend", self.system_tts_mode.currentText())
        voice = self.cfg.get("tts_voice", "")
        parts = ["LiveChatter TTS test.", f"TTS option: {provider}"]
        if provider == "System (screen reader/SAPI5)":
            parts.append(f"System engine: {engine}")
        if voice:
            parts.append(f"Voice: {voice}")
        text = " ".join(parts)
        self.chat_view.append("[TTS Test] " + text)
        try:
            self.tts.speak(text)
        except Exception as e:
            err = f"TTS test failed: {e}"
            self.chat_view.append(f"[Error] {err}")
            QtWidgets.QMessageBox.warning(self, APP_NAME, err)

    def _maybe_do_summary(self):
        if self.chat_mode.currentText() != CHAT_MODES[1]:
            return
        mins = max(1, int(self.summary_interval.value()))
        if (time.time() - self.last_summary_ts) >= mins * 60 and self.pending_messages:
            self.sound_manager.play("summary")
            msgs = self.pending_messages.copy()
            self.pending_messages.clear()
            self.last_summary_ts = time.time()
            summary = self.summarizer.summarize(msgs)
            self.chat_view.append("[Summary]\n" + summary + "\n")
            self.tts.speak(summary)

    def _update_system_tts_enabled(self):
        pass

# ----------------- Main entry -----------------
def main():
    app = QtWidgets.QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()