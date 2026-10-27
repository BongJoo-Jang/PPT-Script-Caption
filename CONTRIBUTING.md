# Contributing / 기여 안내

Thank you for helping make presentations more accessible to people with speech disabilities and others who use prepared scripts. / 발표 접근성을 높이는 개선 제안을 환영합니다.

- Reproduce a bug using a **sample or synthetic presentation**, not someone's real speaker notes.
- Report Windows version, Python version, PowerPoint version, GPU, display layout, and capture mode (`auto`, `winrt`, `dxgi`, `mss`) when relevant.
- Never post personal slide decks, voice recordings, captions, personal identifiers, or credentials.
- Keep **F7–F11 global hotkey behavior** predictable and keyboard-friendly. Avoid features that require fine pointer control.
- Keep Korean and English `translations.py` entries in sync and update both READMEs when behavior changes.
- Run `python -m unittest discover -s tests -v` and `python -m py_compile main.py caption_core.py translations.py` before submitting a change.
- Automatic speech recognition and translation proposals are welcome but should remain **opt-in**; local-only manual operation must continue to work without an AI model, internet access, or an API key.