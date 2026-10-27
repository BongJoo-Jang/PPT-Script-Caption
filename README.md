# PPT Script Caption

[**한국어 사용설명서**](README.ko.md) | **English README**

**Accessible speaker-note captions for PowerPoint presentations on Windows.**  
**Windows PowerPoint 발표자 노트를 활용하는 접근성 자막 프로그램입니다.**

PPT Script Caption is a free, local desktop presentation aid designed with **speakers with speech disabilities** in mind. It also helps presenters who prefer prepared scripts or who need an accessible way to keep track of their place.

Instead of transcribing your speech, it reads **existing PowerPoint speaker notes** and displays manually controlled caption cues. Audience members see the slideshow and the **current** cue; the presenter sees the **previous, current, and next** cues, along with the number of remaining cues.

> **Version 0.2.0 — source release / Windows field-testing stage.** The PowerPoint/Windows interaction cannot be automated in CI. This is not yet a signed installer or a fully validated public release.

## Features

- **One-key presentation flow:** F8 moves to the next caption; after the last caption on a slide, F8 changes to the next slide. F7 goes back.
- **Presenter preview:** previous, current, and next caption cues (not visible to the audience).
- **Independent audience output:** capture of the real, running slideshow plus its current caption in a separate window. Share this window in **Zoom or Google Meet**.
- **Korean / English interface:** choose **한국어 / English** from the presenter window; the selection is saved across restarts.
- **Live measured diagnostics:** real observed **capture FPS**, frames delivered to the **Qt output widget (display FPS)**, capture backend (WinRT/DXGI/MSS), and source window resolution.
- **Optional audience diagnostics:** a checkbox can show the FPS/backend overlay on audience output; **off by default**.
- **Caption font controls:** F10 smaller, F11 larger; persistent font-size setting.
- **Local-only operation:** speaker notes and caption edits are processed on the PC; the app does not send them to an online service.
- **Optional DXcam acceleration:** WinRT → DXGI → MSS fallback. MSS is a valid choice if it feels just as smooth on your PC.

## What is *not* included yet

- **No automatic speech recognition or voice training** — planned, not implemented.
- **No automatic translation** — changing the **UI language** does **not** translate your speaker notes or captions.
- No direct Zoom/Meet API integration. Share the audience output window via the meeting app's usual window-sharing feature.
- No guarantee of 30/60fps: FPS indicators show **measured values**, not the target capture rate.
- Embedded slide videos may look choppy when re-captured and re-rendered in Python/Qt. Native capture / audio sharing remain subject to OS and meeting-app behavior.

## Requirements

- Windows 10/11 with **installed Microsoft PowerPoint desktop** (not PowerPoint for the web).
- Python 3.12+ (Python 3.14 has been used in field testing).
- A PowerPoint `.pptx` with speaker notes.
- Ideally two monitors or an accessible arrangement where the slideshow window remains **visible and unobstructed**.

### Install (Windows CMD)

```bat
cd /d C:\path\to\PPTScriptCaption
py -3.14 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe main.py
```

If you use Python 3.12, replace `py -3.14` with `py -3.12`.

You may also activate the environment and run:

```bat
.venv\Scripts\activate
python main.py
```

### Presenting in Zoom or Google Meet

1. Open the `.pptx` in **PowerPoint desktop**. Choose **Slide Show → Set Up Slide Show → Browsed by an individual (window)** and start the slide show.
2. Launch PPT Script Caption and select **Connect to PowerPoint**.
3. Select **Show audience window**. Move the separate output window to your desired display.
4. In Zoom / Google Meet, **share only the audience output window**. Keep the presenter controls private.
5. Use the keys below; for easier access, program them onto a USB macro keypad.

| Key | Action |
| --- | --- |
| **F7** | Previous caption; at first caption, previous slide |
| **F8** | Next caption; at last caption, next slide |
| **F9** | Toggle audience caption visibility |
| **F10** | Decrease audience caption font size by 2px |
| **F11** | Increase audience caption font size by 2px |

> These are **Windows global hotkeys**, so avoid apps that already reserve them. Presenter notes are split at `//`, or otherwise by line/sentence boundaries. Edit the cues in the controller and click **Save captions** to write a local companion `.pptcaptions.json` file; the PowerPoint file is **not modified**.

### UI language and diagnostics

Use the **Language / 언어** dropdown to switch between Korean and English. The setting is saved under `%APPDATA%\PPTScriptCaption\settings.json` along with font size and the audience-debug checkbox. It changes **labels, button text, notices, and dialogs**, **not** the text of your speaker notes.

The compact diagnostic label in the **presenter** window shows:

```text
Capture: 19.8fps  |  Display: 19.4fps  |  MSS  |  1600×900
```

- **Capture FPS** = number of successfully captured frames per real elapsed second, measured in the capture worker.
- **Display FPS** = rate at which image frames reached the Qt audience widget, measured independently. It is **not** the final Zoom/Meet viewer frame rate.
- **Backend** = `DXcam WINRT`, `DXcam DXGI`, or `MSS`; backend switches are automatic where available.
- **Resolution** = size of the captured PowerPoint slideshow window.

To include the stats in the shared output temporarily, tick **Show FPS/capture status on audience output**; uncheck it before a formal presentation. This does not change capture performance settings.

### Capture troubleshooting

Use the default automatic mode:

```bat
set PPT_CAPTION_CAPTURE_MODE=auto
python main.py
```

To **skip DXcam entirely** and use the reliable MSS path:

```bat
set PPT_CAPTION_CAPTURE_MODE=mss
python main.py
```

Optionally force `winrt` or `dxgi` instead of `mss`. On hybrid Intel/NVIDIA laptops, DXGI may report `0x887A0004` (unsupported); that is **not** necessarily a VRAM issue. The PowerPoint slideshow needs to remain visible and non-minimized in the capture path.

### Privacy / data handling

**Do not commit personal `.pptx` files, speaker notes, captions, speech recordings, credentials, or conference files.** The included `.gitignore` excludes typical presentation and recording formats and the companion caption JSON files. Check files manually before publishing: `.gitignore` protects only **untracked** files that match it.

App preferences are saved locally at `%APPDATA%\PPTScriptCaption\settings.json`. Speaker-note caption edits are saved alongside the chosen PowerPoint file in a `.pptcaptions.json` companion. **No network service, speech recognition or translation API is called.** When Zoom/Meet shares your audience output, that stream is sent through the meeting provider in the usual way.

## Development and testing

Run dependency-free unit tests:

```bat
python -m unittest discover -s tests -v
python -m py_compile main.py caption_core.py translations.py
```

The GitHub Actions workflow checks syntax and unit tests on Windows with Python 3.12/3.14. It **does not** test Office COM, live capture, live video, multi-monitor setups, or meeting-platform integration. Test those manually before presenting.

Contributions and issue reports from people with speech disabilities, accessibility specialists, educators, and developers are welcome. Please avoid uploading personal or identifiable speaker-note content in issues.

## Roadmap (planned)

Optional **offline text-to-speech (TTS)** is planned so presenters can switch between speaking themselves, synthesized voice, and mixed delivery. Further plans include personalized speech recognition and Korean–English translation. These functions are **not included in v0.2.0**. See [Roadmap / 개발 로드맵](ROADMAP.md).

## License

MIT for this project's source code (see [LICENSE](LICENSE)). Libraries such as PySide6 and DXcam have their **own** licensing terms; respect their requirements for any redistributed binary build. This source package does not include third-party binaries or model weights.