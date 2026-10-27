# Changelog

## 0.2.0 — Public source preview

- Korean / English GUI selection with persisted setting. This translates UI chrome, **not** speaker-note content.
- Presenter window diagnostic indicator: measured capture FPS, measured Qt audience-frame delivery FPS, active capture backend, and source size.
- Optional diagnostic overlay on audience output (off by default).
- Retained global F7–F11 navigation and font controls, COM speaker-note integration, and WinRT/DXGI/MSS capture fallback from the field-tested 0.1.11 baseline.
- Added Korean and English documentation, MIT license, privacy-focused `.gitignore`, dependency-free tests, and Windows CI syntax checks.

### Known limitations

- Tests in this source package do not simulate PowerPoint COM, real capture hardware, or Zoom/Google Meet.
- Capture of embedded video remains dependent on Python/Qt, Windows graphics drivers, and the presentation window staying unobstructed.
- Speech recognition and translation are planned features, **not implemented**.