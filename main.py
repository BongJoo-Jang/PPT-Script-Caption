"""PPT Script Caption v0.2.0 — accessible bilingual presentation captions.

Windows desktop PowerPoint (COM) is required. No microphone or cloud access.
"""
import json
import re
import sys
import time
from pathlib import Path
import ctypes
import threading
import os
from collections import deque
from ctypes import wintypes

from PySide6.QtCore import Qt, QTimer, QRect, QThread, Signal
from PySide6.QtGui import QFont, QImage, QPixmap, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QApplication, QHBoxLayout, QLabel, QMainWindow, QMessageBox, QPushButton,
    QTextEdit, QVBoxLayout, QWidget, QSplitter, QCheckBox, QComboBox
)

from caption_core import split_captions
from translations import tr

try:
    import pythoncom
    import win32com.client
    import win32gui
    import mss
except ImportError as exc:
    raise SystemExit(f"Missing Windows dependency: {exc}. See README.md")


def extract_note(slide):
    """Extract notes body placeholder, excluding slide number/date/footer placeholders."""
    pieces = []
    for shape in slide.NotesPage.Shapes:
        try:
            if not shape.HasTextFrame or not shape.TextFrame.HasText:
                continue
            # PowerPoint placeholder type ppPlaceholderBody = 2
            if shape.Type == 14 and shape.PlaceholderFormat.Type != 2:
                continue
            if shape.Type != 14:
                continue
            value = shape.TextFrame.TextRange.Text.strip()
            if value:
                pieces.append(value)
        except Exception:
            continue
    return "\n".join(pieces)


FONT_MIN = 18
FONT_MAX = 54
FONT_STEP = 2
FONT_DEFAULT = 24


def config_file():
    """Keep personal display preferences outside the PPT and source folder."""
    config_root = os.getenv("APPDATA")
    folder = Path(config_root) / "PPTScriptCaption" if config_root else Path.home() / ".ppt_script_caption"
    return folder / "settings.json"


def load_settings():
    try:
        doc = json.loads(config_file().read_text(encoding="utf-8"))
        return doc if isinstance(doc, dict) else {}
    except (OSError, ValueError, TypeError, json.JSONDecodeError):
        return {}


def update_settings(**changes):
    """Merge preferences without ever reading/writing presentation notes."""
    try:
        path = config_file()
        path.parent.mkdir(parents=True, exist_ok=True)
        doc = load_settings()
        doc.update(changes)
        path.write_text(json.dumps(doc, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError as exc:
        print(f"[Settings] Could not save preferences: {exc}", flush=True)


def load_font_size():
    try:
        value = int(load_settings().get("audience_font_px", FONT_DEFAULT))
    except (ValueError, TypeError):
        value = FONT_DEFAULT
    return max(FONT_MIN, min(FONT_MAX, value))


def save_font_size(value):
    update_settings(audience_font_px=int(value))


def load_language():
    return "en" if load_settings().get("ui_language") == "en" else "ko"


def load_show_debug():
    return load_settings().get("audience_debug", False) is True


def dxcam_output_table(info_text):
    """Decode DXcam's public output_info() string to device/output metadata."""
    pattern = re.compile(r"Device\[(\d+)\]\s*Output\[(\d+)\].*?Res:\s*\(\s*(\d+)\s*,\s*(\d+)\s*\).*?Primary:\s*(True|False)", re.I)
    results = []
    for line in str(info_text).splitlines():
        found = pattern.search(line)
        if found:
            a, b, w, h, primary = found.groups()
            results.append(dict(device=int(a), output=int(b), width=int(w), height=int(h), primary=(primary.lower()=="true")))
    return results


def dxcam_candidates(outputs, rect, primary, monitor_order):
    """Prefer correct output resolution and primary status over DXcam device 0."""
    w = rect[2] - rect[0]
    h = rect[3] - rect[1]
    ranked = []
    for entry in outputs:
        # Exclude definitely unrelated resolutions.
        if entry['width'] != w or entry['height'] != h:
            continue
        score = (10 if entry['primary'] == primary else 0)
        score += (2 if entry['output'] == monitor_order else 0)
        ranked.append((-score, entry['device'], entry['output']))
    ranked.sort()
    return [(d, o) for _, d, o in ranked]


class PowerPointBridge:
    def __init__(self, language="ko"):
        self.language = language
        self.app = None
        self.presentation = None
        self.slide_captions = []
        self.path = None

    def connect(self):
        # IMPORTANT: Never create a second Office process as a fallback.
        # Dispatch() can attempt to launch PowerPoint and fail with
        # CO_E_SERVER_EXEC_FAILURE (0x80080005), masking the root cause.
        try:
            self.app = win32com.client.GetActiveObject("PowerPoint.Application")
        except pythoncom.com_error as exc:
            hresult = exc.hresult & 0xFFFFFFFF
            raise RuntimeError(tr(self.language, "office_not_found", code=f"0x{hresult:08X}")) from exc
        if self.app.Presentations.Count == 0:
            raise RuntimeError(tr(self.language, "ppt_not_open"))
        self.presentation = self.app.ActivePresentation
        self.path = Path(self.presentation.FullName)
        notes = [extract_note(s) for s in self.presentation.Slides]
        self.slide_captions = [split_captions(n) or [tr(self.language, "missing_notes")] for n in notes]
        edited_path = self.path.with_suffix(".pptcaptions.json")
        if edited_path.exists():
            try:
                doc = json.loads(edited_path.read_text(encoding="utf-8"))
                if doc.get("ppt") == str(self.path) and len(doc.get("captions", [])) == len(self.slide_captions):
                    self.slide_captions = doc["captions"]
            except Exception:
                pass
        return len(self.slide_captions)

    def save(self):
        if self.path is None:
            return
        self.path.with_suffix(".pptcaptions.json").write_text(
            json.dumps({"ppt": str(self.path), "captions": self.slide_captions}, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )

    def slideshow(self):
        if not self.app or self.app.SlideShowWindows.Count < 1:
            raise RuntimeError(tr(self.language, "show_not_started"))
        return self.app.SlideShowWindows(1)

    def current_slide(self):
        return int(self.slideshow().View.Slide.SlideIndex) - 1

    def move_to(self, number):
        self.slideshow().View.GotoSlide(number + 1)

    def rect(self):
        """Find PowerPoint's actual slideshow HWND without accessing Application.HWND.

        Some Office installations don't expose Application.HWND through COM
        (DISP_E_MEMBERNOTFOUND, 0x80020003). Win32 enumeration is safer here.
        """
        self.slideshow()  # Validate that PowerPoint says a slideshow is active.
        import win32process

        # Windows window classes, not Korean/English title alone, identify PPT.
        ppt_pids = set()
        all_windows = []

        def visit(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd) or win32gui.IsIconic(hwnd):
                return
            try:
                cls = win32gui.GetClassName(hwnd).lower()
                title = win32gui.GetWindowText(hwnd).lower()
                bounds = win32gui.GetWindowRect(hwnd)
                area = max(0, bounds[2] - bounds[0]) * max(0, bounds[3] - bounds[1])
                if area < 300 * 200:
                    return
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                all_windows.append((hwnd, pid, cls, title, bounds, area))
                if cls == "pptframeclass":
                    ppt_pids.add(pid)
            except Exception:
                return

        win32gui.EnumWindows(visit, None)

        candidates = []
        for hwnd, pid, cls, title, bounds, area in all_windows:
            # The screenClass window is the PowerPoint slideshow, not editor.
            is_screen = "screenclass" in cls
            is_titled_show = ("slide show" in title or
                              "slideshow" in title or "슬라이드 쇼" in title)
            if not (is_screen or is_titled_show):
                continue
            if ppt_pids and pid not in ppt_pids:
                continue
            # Prefer a native slideshow window over a title-only match.
            candidates.append((2 if is_screen else 1, area, bounds))

        if not candidates:
            classes = ", ".join(sorted({x[2] for x in all_windows if "ppt" in x[2] or "screen" in x[2]}))
            raise RuntimeError(tr(self.language, "show_window_missing") +
                               f" [Window classes: {classes or 'none'}]")
        candidates.sort(reverse=True)
        return candidates[0][2]



def locate_slideshow_rect(language="ko"):
    """Win32-only lookup: safe to call from capture thread; no Office COM."""
    import win32process
    windows = []
    powerpnt_pids = set()

    def visit(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd) or win32gui.IsIconic(hwnd):
            return
        try:
            cls = win32gui.GetClassName(hwnd).lower()
            title = win32gui.GetWindowText(hwnd).lower()
            bounds = win32gui.GetWindowRect(hwnd)
            area = max(0, bounds[2] - bounds[0]) * max(0, bounds[3] - bounds[1])
            if area < 300 * 200:
                return
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            windows.append((pid, cls, title, bounds, area))
            if cls == "pptframeclass":
                powerpnt_pids.add(pid)
        except Exception:
            return

    win32gui.EnumWindows(visit, None)
    candidates = []
    for pid, cls, title, bounds, area in windows:
        is_screen = "screenclass" in cls
        is_title = any(key in title for key in ("slide show", "slideshow", "슬라이드 쇼"))
        if not (is_screen or is_title):
            continue
        if powerpnt_pids and pid not in powerpnt_pids:
            continue
        candidates.append((2 if is_screen else 1, area, bounds))
    if not candidates:
        raise RuntimeError(tr(language, "show_window_missing"))
    candidates.sort(reverse=True)
    return candidates[0][2]



class CaptureWorker(QThread):
    """Background capture. Never uses Office COM, and never blocks Qt hotkeys.

    Try Windows Graphics Capture (WinRT) first for hybrid GPUs, then DXGI,
    and finally MSS. Verify against MSS to avoid capturing the wrong screen.
    """
    frame_ready = Signal(object)
    failed = Signal(str)
    info = Signal(str)
    stats = Signal(object)  # Observed capture rate, backend, source dimensions.

    def __init__(self, parent=None):
        super().__init__(parent)
        self._stop_event = threading.Event()
        self._enabled = threading.Event()
        self._last_error = None
        self.language = "ko"

    def enable(self, value):
        (self._enabled.set if value else self._enabled.clear)()

    def stop(self):
        self._stop_event.set()
        # Ensure background thread exits before its QThread object is destroyed.
        self.wait()

    @staticmethod
    def monitor_for_rect(rect, monitors):
        cx = (rect[0] + rect[2]) / 2
        cy = (rect[1] + rect[3]) / 2
        for idx, m in enumerate(monitors):
            left, top, right, bottom = m['rect']
            if left <= cx < right and top <= cy < bottom:
                return idx, m
        return None, None

    @staticmethod
    def matches_monitor(frame, grab):
        """Avoid showing the wrong monitor when DXGI monitor orders differ."""
        import numpy as np
        h, w = frame.shape[:2]
        if h != grab.height or w != grab.width or frame.shape[-1] != 4:
            return False
        # Compare a sparse sample; allow video/animation frame timing differences.
        ref = np.frombuffer(grab.bgra, dtype=np.uint8).reshape((h, w, 4))
        ys = max(1, h // 28)
        xs = max(1, w // 28)
        a = frame[::ys, ::xs, :3].astype(np.int16)
        b = ref[::ys, ::xs, :3].astype(np.int16)
        return float(np.mean(np.abs(a-b))) < 45.0

    def run(self):
        import mss
        mode = os.getenv('PPT_CAPTION_CAPTURE_MODE', 'auto').lower().strip()
        if mode not in ('auto', 'winrt', 'dxgi', 'mss'):
            mode = 'auto'
        # Prefer Windows Graphics Capture on hybrid Intel/NVIDIA laptops.
        backend_order = (('winrt', 'dxgi') if mode == 'auto'
                         else (mode,) if mode != 'mss' else ())
        dxcam = None
        dx_outputs = []
        device_text = ''
        output_text = ''
        if mode != 'mss':
            try:
                import dxcam as dxcam_module
                dxcam = dxcam_module
                device_text = dxcam.device_info()
                output_text = dxcam.output_info()
                dx_outputs = dxcam_output_table(output_text)
                self.info.emit(tr(self.language, 'log_devices', details=str(device_text).replace('\n', ' / ')))
                self.info.emit(tr(self.language, 'log_outputs', details=str(output_text).replace('\n', ' / ')))
            except Exception as exc:
                dxcam = None
                self.info.emit(tr(self.language, 'log_device_error', error=str(exc)))
        else:
            self.info.emit(tr(self.language, 'log_mss'))
        if dxcam is not None:
            self.info.emit(tr(self.language, 'log_order', order=', '.join(backend_order)))
        try:
            import win32api
            display_monitors = []
            for item in win32api.EnumDisplayMonitors(None, None):
                bounds = tuple(item[2])
                flags = win32api.GetMonitorInfo(item[0]).get('Flags', 0)
                display_monitors.append({'rect': bounds, 'primary': bool(flags & 1)})
        except Exception:
            display_monitors = []
        camera = None
        active_key = None
        failed_monitors = set()
        latest_screen = None
        frames_in_interval = 0
        interval_started = time.perf_counter()
        last_backend = "—"
        last_size = (0, 0)
        with mss.MSS() as sct:  # New API; mss.mss() is deprecated.
            mss_monitors = [
                {'rect': (m['left'], m['top'], m['left']+m['width'], m['top']+m['height']), 'primary': idx == 0}
                for idx, m in enumerate(sct.monitors[1:])
            ]
            monitors = display_monitors or mss_monitors
            while not self._stop_event.is_set():
                if not self._enabled.is_set():
                    now = time.perf_counter()
                    if now - interval_started >= 1.0:
                        self.stats.emit({"mode": "paused", "backend": last_backend,
                                         "capture_fps": 0.0,
                                         "width": last_size[0], "height": last_size[1]})
                        interval_started = now
                        frames_in_interval = 0
                    self._stop_event.wait(0.12)
                    continue
                started = time.perf_counter()
                fps = 20
                try:
                    rect = locate_slideshow_rect(self.language)
                    x1, y1, x2, y2 = rect
                    w, h = x2-x1, y2-y1
                    if w <= 0 or h <= 0:
                        raise RuntimeError(tr(self.language, "show_zero_size"))
                    monitor_idx, mon = self.monitor_for_rect(rect, monitors)
                    mon_key = tuple(mon['rect']) if mon else None
                    if mon_key != latest_screen:
                        if camera is not None:
                            try:
                                camera.release()
                            except Exception:
                                pass
                        camera = None
                        active_key = None
                        latest_screen = mon_key
                    full_on_monitor = mon is not None and (
                        x1 >= mon['rect'][0] and y1 >= mon['rect'][1] and
                        x2 <= mon['rect'][2] and y2 <= mon['rect'][3]
                    )
                    region = {'left': x1, 'top': y1, 'width': w, 'height': h}
                    local = None
                    if full_on_monitor:
                        left, top, *_ = mon['rect']
                        local = (x1-left, y1-top, x2-left, y2-top)
                    if (camera is None and dxcam is not None and full_on_monitor
                            and mon_key not in failed_monitors):
                        try:
                            # Allow explicit overrides for exotic multi-GPU layouts.
                            forced_dev = os.getenv('PPT_CAPTION_GPU_INDEX', '').strip()
                            forced_out = os.getenv('PPT_CAPTION_OUTPUT_INDEX', '').strip()
                            if forced_dev or forced_out:
                                choices = [(int(forced_dev or 0), int(forced_out or monitor_idx))]
                            else:
                                choices = dxcam_candidates(dx_outputs, mon['rect'], mon['primary'], monitor_idx)
                                if not choices:
                                    # Older/newer DXcam information formatting: still test
                                    # discovered adapters rather than assuming device_idx=0.
                                    devs = sorted({int(x) for x in re.findall(r'Device\[(\d+)\]', str(device_text))}) or [0]
                                    outs = sorted({int(x) for x in re.findall(r'Output\[(\d+)\]', str(output_text))}) or [monitor_idx]
                                    choices = [(d, o) for d in devs for o in outs]
                            errors = []
                            for backend_name in backend_order:
                                if camera is not None:
                                    break
                                for dev_idx, out_idx in choices[:12]:
                                    candidate = None
                                    try:
                                        candidate = dxcam.create(
                                            device_idx=dev_idx, output_idx=out_idx,
                                            output_color='BGRA', processor_backend='numpy',
                                            backend=backend_name
                                        )
                                        test_frame = None
                                        # WinRT may not have a frame immediately after startup.
                                        for _ in range(5):
                                            test_frame = candidate.grab(
                                                region=local, new_frame_only=False
                                            )
                                            if test_frame is not None:
                                                break
                                            self._stop_event.wait(0.08)
                                        if test_frame is None:
                                            raise RuntimeError(tr(self.language, 'log_empty_frame'))
                                        ref = sct.grab(region)
                                        if not self.matches_monitor(test_frame, ref):
                                            raise RuntimeError(tr(self.language, 'log_wrong_monitor'))
                                        camera = candidate
                                        active_key = (backend_name, dev_idx, out_idx)
                                        self.info.emit(tr(self.language, 'log_ready',
                                                          backend=backend_name.upper(),
                                                          gpu=dev_idx, output=out_idx))
                                        break
                                    except Exception as exc:
                                        errors.append(
                                            f'{backend_name.upper()} GPU{dev_idx}/출력{out_idx}: '
                                            f'{str(exc)[:180]}'
                                        )
                                        if candidate is not None:
                                            try:
                                                candidate.release()
                                            except Exception:
                                                pass
                            if camera is None:
                                failed_monitors.add(mon_key)
                                self.info.emit(tr(self.language, 'log_fallback', errors=' | '.join(errors[:8])))
                        except Exception as exc:
                            failed_monitors.add(mon_key)
                            self.info.emit(tr(self.language, 'log_init_error', error=str(exc)))
                    qim = None
                    if camera is not None and local is not None:
                        try:
                            frame = camera.grab(region=local, new_frame_only=False)
                            if frame is not None and frame.shape[:2] == (h, w):
                                if not frame.flags['C_CONTIGUOUS']:
                                    frame = frame.copy(order='C')
                                qim = QImage(frame.tobytes(), w, h, w*4, QImage.Format.Format_ARGB32).copy()
                                fps = 30
                        except Exception as exc:
                            backend_name = active_key[0].upper() if active_key else "UNKNOWN"
                            self.info.emit(tr(self.language, 'log_frame_error', backend=backend_name, error=str(exc)))
                            failed_monitors.add(mon_key)
                            try:
                                camera.release()
                            except Exception:
                                pass
                            camera = None
                            active_key = None
                    if qim is None:
                        grab = sct.grab(region)
                        qim = QImage(bytes(grab.bgra), w, h, w*4, QImage.Format.Format_ARGB32).copy()
                        fps = 20
                    self.frame_ready.emit(qim)
                    frames_in_interval += 1
                    last_backend = (f"DXcam {active_key[0].upper()}" if camera is not None and active_key else "MSS")
                    last_size = (w, h)
                    self._last_error = None
                except Exception as exc:
                    error = tr(self.language, "capture_failure",
                               kind=type(exc).__name__, error=str(exc))
                    if error != self._last_error:
                        self.failed.emit(error)
                        self._last_error = error
                    fps = 5
                now = time.perf_counter()
                if now - interval_started >= 1.0:
                    duration = max(0.001, now - interval_started)
                    self.stats.emit({"mode": "error" if self._last_error else "running",
                                     "backend": last_backend,
                                     "capture_fps": frames_in_interval / duration,
                                     "width": last_size[0], "height": last_size[1]})
                    frames_in_interval = 0
                    interval_started = now
                self._stop_event.wait(max(0, 1/fps - (time.perf_counter()-started)))
        if camera is not None:
            try:
                camera.release()
            except Exception:
                pass

class AudienceWindow(QWidget):
    """Only the current caption is shown to attendees; debug is opt-in."""

    def __init__(self, language="ko"):
        super().__init__()
        self.language = language
        self.setWindowTitle(tr(self.language, "audience_title"))
        self.resize(960, 540)
        self.setStyleSheet("background:#101b2c; color:white;")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        self.image = QLabel(tr(self.language, "waiting_show"))
        self.image.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.image.setStyleSheet("background:#101b2c;")
        self.image.setMinimumSize(400, 200)
        layout.addWidget(self.image, 5)
        self.font_px = load_font_size()
        self.caption = QLabel("")
        self.caption.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.caption.setWordWrap(True)
        self.caption.setMinimumHeight(90)
        self.apply_caption_style()
        layout.addWidget(self.caption, 1)
        self.last_image = None
        self._paint_times = deque(maxlen=120)
        # Overlay is a child of the image, never written into the PPT itself.
        self.debug_overlay = QLabel(self.image)
        self.debug_overlay.setStyleSheet(
            "background:rgba(0,0,0,190); color:#ffffff; font-size:12px;"
            "padding:5px 8px; border-radius:3px;"
        )
        self.debug_overlay.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self._debug_visible = False
        self.debug_overlay.hide()

    def set_language(self, language):
        self.language = language
        self.setWindowTitle(tr(language, "audience_title"))
        if self.last_image is None:
            self.image.setText(tr(language, "waiting_show"))

    def apply_caption_style(self):
        self.caption.setStyleSheet(
            f"background:#080c14;color:white;font-size:{self.font_px}px;"
            "font-weight:600;padding:8px;"
        )

    def resize_caption_font(self, delta):
        updated = max(FONT_MIN, min(FONT_MAX, self.font_px + delta))
        if updated != self.font_px:
            self.font_px = updated
            self.apply_caption_style()
            save_font_size(self.font_px)
        return self.font_px

    def set_caption(self, value):
        self.caption.setText(value)

    def set_frame(self, qim):
        self._paint_times.append(time.perf_counter())
        self.last_image = QPixmap.fromImage(qim)
        self.refresh_frame()

    def displayed_fps(self):
        """Rate at which frames actually reached the Qt audience widget."""
        now = time.perf_counter()
        while self._paint_times and now - self._paint_times[0] > 2.0:
            self._paint_times.popleft()
        if len(self._paint_times) < 2:
            return 0.0
        # Count periods, not frames; let the rate fall after a capture stall.
        duration = max(now-self._paint_times[0], 0.001)
        return (len(self._paint_times)-1) / duration

    def set_debug_visible(self, visible):
        self._debug_visible = bool(visible)
        self.debug_overlay.setVisible(self._debug_visible)
        if self._debug_visible:
            self.debug_overlay.raise_()

    def set_debug_text(self, text):
        self.debug_overlay.setText(text)
        self.debug_overlay.adjustSize()
        self.debug_overlay.move(10, 10)
        if self._debug_visible:
            self.debug_overlay.raise_()

    def refresh_frame(self):
        if self.last_image is not None:
            target = self.image.size()
            self.image.setPixmap(self.last_image.scaled(
                target, Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation))
            if self._debug_visible:
                self.debug_overlay.raise_()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.refresh_frame()


class GlobalHotkeys(QThread):
    """RegisterHotKey in a dedicated Win32 message-loop thread.

    Qt's main/UI thread may be busy with PowerPoint COM calls. WM_HOTKEY is
    thread-specific when hwnd=None, so polling on the Qt thread can miss or
    discard inputs. Register and receive on the SAME dedicated thread.
    """
    pressed = Signal(int)
    registration = Signal(object)
    WM_HOTKEY = 0x0312
    WM_QUIT = 0x0012
    MOD_NOREPEAT = 0x4000

    def __init__(self, parent=None):
        super().__init__(parent)
        self._thread_id = 0
        self._lock = threading.Lock()
        self._ready = threading.Event()
        self.registered = []
        self.errors = []

    def run(self):
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        user32 = ctypes.WinDLL("user32", use_last_error=True)
        kernel32.GetCurrentThreadId.restype = wintypes.DWORD
        user32.RegisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int, wintypes.UINT, wintypes.UINT]
        user32.RegisterHotKey.restype = wintypes.BOOL
        user32.UnregisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int]
        user32.UnregisterHotKey.restype = wintypes.BOOL
        user32.GetMessageW.argtypes = [ctypes.POINTER(wintypes.MSG), wintypes.HWND, wintypes.UINT, wintypes.UINT]
        user32.GetMessageW.restype = ctypes.c_int
        msg = wintypes.MSG()
        # Create the thread message queue before posting to it from close().
        user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 0)
        with self._lock:
            self._thread_id = kernel32.GetCurrentThreadId()
        for key_id, vk in [(1, 0x76), (2, 0x77), (3, 0x78), (4, 0x79), (5, 0x7A)]:
            if user32.RegisterHotKey(None, key_id, self.MOD_NOREPEAT, vk):
                self.registered.append(key_id)
            else:
                self.errors.append(f"F{key_id+6}: WinError {ctypes.get_last_error()}")
        self.registration.emit(tuple(self.errors))
        self._ready.set()
        try:
            while True:
                result = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
                if result <= 0:
                    break
                if msg.message == self.WM_HOTKEY:
                    self.pressed.emit(int(msg.wParam))
        finally:
            for key_id in self.registered:
                user32.UnregisterHotKey(None, key_id)
            self.registered.clear()

    def close(self):
        self._ready.wait(2)
        with self._lock:
            tid = self._thread_id
        if tid:
            user32 = ctypes.WinDLL("user32", use_last_error=True)
            user32.PostThreadMessageW.argtypes = [wintypes.DWORD, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
            user32.PostThreadMessageW(tid, self.WM_QUIT, 0, 0)
        self.wait(2000)


class Presenter(QMainWindow):
    """Accessible controller: one-key cue navigation in either UI language."""

    def __init__(self):
        super().__init__()
        self.language = load_language()
        self.resize(880, 760)
        self.bridge = PowerPointBridge(language=self.language)
        self.current = 0
        self.index = 0
        self.visible_caption = True
        self.audience = AudienceWindow(self.language)
        self.capture_data = {"mode": "paused", "backend": "—", "capture_fps": 0.0,
                             "width": 0, "height": 0}

        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        buttons = QHBoxLayout()
        self.connect_btn = QPushButton()
        self.connect_btn.clicked.connect(self.connect)
        buttons.addWidget(self.connect_btn)
        self.output_btn = QPushButton()
        self.output_btn.clicked.connect(self.show_output)
        buttons.addWidget(self.output_btn)
        self.language_label = QLabel()
        buttons.addWidget(self.language_label)
        self.language_combo = QComboBox()
        self.language_combo.addItem("한국어", "ko")
        self.language_combo.addItem("English", "en")
        self.language_combo.setCurrentIndex(0 if self.language == "ko" else 1)
        self.language_combo.currentIndexChanged.connect(self.change_language)
        buttons.addWidget(self.language_combo)
        layout.addLayout(buttons)

        self.status = QLabel()
        layout.addWidget(self.status)
        self.prev = QLabel()
        self.prev.setWordWrap(True)
        self.prev.setStyleSheet("background:#303b50; color:#d8e1ef; padding:16px; border-radius:8px;")
        layout.addWidget(self.prev, 1)
        self.now = QLabel()
        self.now.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.now.setWordWrap(True)
        self.now.setFont(QFont("Malgun Gothic", 16, QFont.Weight.Bold))
        self.now.setStyleSheet("background:#164b49; color:white; padding:24px; border:2px solid #49cdbb; border-radius:8px;")
        layout.addWidget(self.now, 2)
        self.next = QLabel()
        self.next.setWordWrap(True)
        self.next.setStyleSheet("background:#303b50; color:#d8e1ef; padding:16px; border-radius:8px;")
        layout.addWidget(self.next, 1)

        actions = QHBoxLayout()
        self.back_btn = QPushButton()
        self.back_btn.clicked.connect(self.back)
        actions.addWidget(self.back_btn)
        self.go_btn = QPushButton()
        self.go_btn.clicked.connect(self.advance)
        actions.addWidget(self.go_btn)
        self.hide_btn = QPushButton()
        self.hide_btn.clicked.connect(self.toggle_caption)
        actions.addWidget(self.hide_btn)
        layout.addLayout(actions)

        font_bar = QHBoxLayout()
        self.font_down_btn = QPushButton()
        self.font_down_btn.clicked.connect(lambda: self.change_font_size(-FONT_STEP))
        font_bar.addWidget(self.font_down_btn)
        self.font_display = QLabel()
        self.font_display.setAlignment(Qt.AlignmentFlag.AlignCenter)
        font_bar.addWidget(self.font_display)
        self.font_up_btn = QPushButton()
        self.font_up_btn.clicked.connect(lambda: self.change_font_size(FONT_STEP))
        font_bar.addWidget(self.font_up_btn)
        layout.addLayout(font_bar)

        # Show actual observed capture rate and Qt-delivered frame rate separately.
        # This is diagnostics only: capture backend and timing are unchanged.
        self.diagnostics = QLabel()
        self.diagnostics.setStyleSheet("font-size:12px; color:#4b5563; padding:3px;")
        layout.addWidget(self.diagnostics)
        self.show_debug_checkbox = QCheckBox()
        self.show_debug_checkbox.setChecked(load_show_debug())
        self.show_debug_checkbox.toggled.connect(self.change_debug_visibility)
        layout.addWidget(self.show_debug_checkbox)
        self.audience.set_debug_visible(self.show_debug_checkbox.isChecked())

        self.editor_label = QLabel()
        layout.addWidget(self.editor_label)
        self.editor = QTextEdit()
        layout.addWidget(self.editor, 1)
        self.save_btn = QPushButton()
        self.save_btn.clicked.connect(self.save_captions)
        layout.addWidget(self.save_btn)

        # Keep RegisterHotKey on the same Win32 thread as its message loop.
        self.hotkeys = GlobalHotkeys(self)
        self.hotkeys.pressed.connect(self.on_hotkey)
        self.hotkeys.registration.connect(self.on_hotkey_registration)
        self.hotkeys.start()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.tick)
        self.timer.start(1000)

        self.capture_worker = CaptureWorker(self)
        self.capture_worker.language = self.language
        self.capture_worker.frame_ready.connect(self.audience.set_frame)
        self.capture_worker.failed.connect(self.on_capture_error)
        self.capture_worker.info.connect(self.on_capture_info)
        self.capture_worker.stats.connect(self.on_capture_stats)
        self.capture_worker.start()
        self.audience.installEventFilter(self)
        self.capture_worker.enable(False)

        self.diag_timer = QTimer(self)
        self.diag_timer.timeout.connect(self.refresh_diagnostics)
        self.diag_timer.start(1000)
        self.apply_language()

    def change_language(self, _index):
        new_language = self.language_combo.currentData()
        if new_language not in ("ko", "en") or new_language == self.language:
            return
        self.language = new_language
        self.bridge.language = new_language
        self.capture_worker.language = new_language
        self.audience.set_language(new_language)
        update_settings(ui_language=new_language)
        self.apply_language()

    def apply_language(self):
        self.setWindowTitle(tr(self.language, "presenter_title"))
        self.connect_btn.setText(tr(self.language, "connect"))
        self.output_btn.setText(tr(self.language, "show_output"))
        self.language_label.setText(tr(self.language, "language_label"))
        self.back_btn.setText(tr(self.language, "prev_button"))
        self.go_btn.setText(tr(self.language, "next_button"))
        self.hide_btn.setText(tr(self.language, "toggle_button"))
        self.font_down_btn.setText(tr(self.language, "smaller"))
        self.font_up_btn.setText(tr(self.language, "larger"))
        self.font_display.setText(tr(self.language, "font_label", size=self.audience.font_px))
        self.editor_label.setText(tr(self.language, "editor_label"))
        self.save_btn.setText(tr(self.language, "save"))
        self.show_debug_checkbox.setText(tr(self.language, "debug_checkbox"))
        if self.bridge.slide_captions:
            self.refresh()
        else:
            self.status.setText(tr(self.language, "not_connected"))
            self.prev.setText(tr(self.language, "prev_label"))
            self.now.setText(tr(self.language, "now_label"))
            self.next.setText(tr(self.language, "next_label"))
        self.refresh_diagnostics()

    def on_hotkey(self, key_id):
        if key_id == 1:
            self.back()
        elif key_id == 2:
            self.advance()
        elif key_id == 3:
            self.toggle_caption()
        elif key_id == 4:
            self.change_font_size(-FONT_STEP)
        elif key_id == 5:
            self.change_font_size(FONT_STEP)

    def change_font_size(self, delta):
        new_size = self.audience.resize_caption_font(delta)
        self.font_display.setText(tr(self.language, "font_label", size=new_size))

    def on_hotkey_registration(self, errors):
        if errors:
            QMessageBox.warning(self, tr(self.language, "hotkeys_error_title"),
                                tr(self.language, "hotkeys_error", errors=", ".join(errors)))

    def connect(self):
        try:
            count = self.bridge.connect()
            try:
                self.current = self.bridge.current_slide()
            except RuntimeError:
                self.current = 0
            self.index = 0
            self.editor.setPlainText("\n".join(self.bridge.slide_captions[self.current]))
            self.refresh()
            if self.audience.isVisible():
                self.capture_worker.enable(True)
            self.status.setText(tr(self.language, "connected", count=count))
        except Exception as exc:
            QMessageBox.warning(self, tr(self.language, "connect_error_title"),
                                tr(self.language, "connect_error", error=str(exc)))

    def show_output(self):
        self.audience.show()
        self.audience.raise_()
        self.capture_worker.enable(True)

    def eventFilter(self, obj, event):
        if obj is self.audience:
            from PySide6.QtCore import QEvent
            if event.type() == QEvent.Type.Hide:
                self.capture_worker.enable(False)
                self.capture_data["mode"] = "paused"
                self.refresh_diagnostics()
            elif event.type() == QEvent.Type.Show and self.bridge.slide_captions:
                self.capture_worker.enable(True)
        return super().eventFilter(obj, event)

    def change_debug_visibility(self, visible):
        self.audience.set_debug_visible(visible)
        update_settings(audience_debug=bool(visible))
        self.refresh_diagnostics()

    def on_capture_info(self, message):
        print("[Capture]", message, flush=True)

    def on_capture_stats(self, data):
        self.capture_data = dict(data)
        self.refresh_diagnostics()

    def refresh_diagnostics(self):
        data = self.capture_data
        mode = data.get("mode", "paused")
        if mode == "paused" or not self.audience.isVisible():
            message = tr(self.language, "debug_paused")
            self.diagnostics.setText(message)
            self.audience.set_debug_text(message)
            return
        if mode == "error":
            self.diagnostics.setText(tr(self.language, "debug_error"))
            self.audience.set_debug_text(tr(self.language, "debug_error"))
            return
        capture = float(data.get("capture_fps", 0.0))
        display = self.audience.displayed_fps()
        backend = data.get("backend", "—")
        width = data.get("width", 0)
        height = data.get("height", 0)
        self.diagnostics.setText(tr(self.language, "debug_line", capture=capture,
                                    display=display, backend=backend,
                                    width=width, height=height))
        self.audience.set_debug_text(tr(self.language, "debug_audience_line",
                                        capture=capture, display=display, backend=backend))

    def on_capture_error(self, message):
        self.audience.image.setPixmap(QPixmap())
        self.audience.last_image = None
        self.audience.image.setText(message)

    def refresh(self):
        if not self.bridge.slide_captions:
            return
        group = self.bridge.slide_captions[self.current]
        self.index = min(max(self.index, 0), len(group) - 1)
        is_last = self.index == len(group) - 1
        self.prev.setText(tr(self.language, "prev_label") + "\n" +
                          (group[self.index-1] if self.index else "—"))
        self.now.setText(group[self.index])
        self.next.setText(tr(self.language, "next_label") + "\n" +
                          (group[self.index+1] if not is_last else tr(self.language, "next_slide")))
        self.status.setText(tr(self.language, "status_line",
                               slide=self.current+1, total=len(self.bridge.slide_captions),
                               cue=self.index+1, cues=len(group),
                               remaining=len(group)-self.index-1))
        self.audience.set_caption(group[self.index] if self.visible_caption else "")

    def sync_current_editor(self):
        if not self.bridge.slide_captions:
            return
        lines = [line.strip() for line in self.editor.toPlainText().splitlines() if line.strip()]
        if lines:
            self.bridge.slide_captions[self.current] = lines
            self.index = min(self.index, len(lines)-1)

    def swap_slide(self, new_slide, index=0):
        self.sync_current_editor()
        self.current = new_slide
        self.index = index
        self.editor.setPlainText("\n".join(self.bridge.slide_captions[self.current]))
        self.refresh()

    def advance(self):
        if not self.bridge.slide_captions:
            return
        self.sync_current_editor()
        cues = self.bridge.slide_captions[self.current]
        if self.index + 1 < len(cues):
            self.index += 1
            self.refresh()
        elif self.current + 1 < len(self.bridge.slide_captions):
            try:
                self.bridge.move_to(self.current+1)
                self.swap_slide(self.current+1)
            except Exception as exc:
                QMessageBox.warning(self, tr(self.language, "slide_error_title"), str(exc))
        else:
            self.status.setText(tr(self.language, "end_presentation"))

    def back(self):
        if not self.bridge.slide_captions:
            return
        self.sync_current_editor()
        if self.index > 0:
            self.index -= 1
            self.refresh()
        elif self.current > 0:
            target = self.current - 1
            try:
                self.bridge.move_to(target)
                self.swap_slide(target, len(self.bridge.slide_captions[target])-1)
            except Exception as exc:
                QMessageBox.warning(self, tr(self.language, "slide_error_title"), str(exc))

    def toggle_caption(self):
        self.visible_caption = not self.visible_caption
        self.refresh()

    def save_captions(self):
        if not self.bridge.slide_captions:
            return
        try:
            self.sync_current_editor()
            self.bridge.save()
            self.refresh()
            QMessageBox.information(self, tr(self.language, "save_success_title"),
                                    tr(self.language, "save_success"))
        except Exception as exc:
            QMessageBox.warning(self, tr(self.language, "save_error_title"), str(exc))

    def tick(self):
        if not self.bridge.slide_captions:
            return
        try:
            found = self.bridge.current_slide()
            if found != self.current:
                self.swap_slide(found)
        except Exception:
            pass

    def closeEvent(self, event):
        self.hotkeys.close()
        self.capture_worker.stop()
        self.audience.close()
        super().closeEvent(event)


def main():
    pythoncom.CoInitialize()
    app = QApplication(sys.argv)
    ui = Presenter()
    ui.show()
    try:
        sys.exit(app.exec())
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()