# Roadmap / 개발 로드맵

This document describes **planned features**, not currently available functionality.
이 문서의 기능은 **개발 계획**이며 현재 버전에 포함되어 있지 않습니다.

## 0.2.x — Accessibility and stability / 접근성·안정화

- Validate the Korean/English interface and measured FPS display on Windows with PowerPoint.
- Verify global F7–F11 hotkeys, mixed-language speaker notes, window sharing, and embedded video playback.
- Add larger controls, configurable caption contrast, and crash/reconnect diagnostics.
- Windows installer / standalone executable after dependency-license review.

## 0.3 — Text-to-speech / 음성 합성(TTS)

- **Offline-first:** use Windows system voices / SAPI (or a separately installed local voice engine) before any online API.
- **Three modes:** manual speaking, TTS speaking, mixed/manual + TTS speaking.
- Read **only the current caption** on demand. Support **play/replay, pause/resume, stop**, speech-rate control, and voice/language selection.
- Keep F7/F8/F9 navigation independent from TTS playback by default. If optional auto-advance is enabled, it must be explicit, reversible, and never advance while a voice is still speaking.
- Show clear playback state and large keyboard/macro-keypad-friendly controls; avoid fine motor precision requirements.
- Avoid overlapping microphone and synthesized audio. Document Zoom/Google Meet **share computer sound** and loopback/virtual-audio options, plus echo prevention.
- **Privacy:** speaker notes remain on the device by default; online TTS requires separate opt-in with transparent data sharing disclosure.

**TTS 목적:** 발화장애인이 원고를 직접 읽거나 필요한 문장만 음성 합성으로 출력하는 **혼합 발표**를 지원합니다. 원치 않을 때는 TTS를 완전히 꺼두고 현재 자막 기능만 사용할 수 있어야 합니다.

## 0.4 — Assisted pacing / 발표 진행 보조

- Optional synchronization of caption advancement to end-of-TTS playback.
- Optional cue-timing indicators and skip/rewind with easy one-key recovery.
- Preserve manual F8 navigation as an always-available fallback.

## Future / 향후 목표

- **Personalized speech recognition (ASR)** tailored to each speaker; a separate local/on-device model if feasible.
- **Caption-aware alignment:** match recognized speech to prepared notes and assist navigation without assuming every pronunciation can be reliably recognized.
- **Automatic translation:** Korean ↔ English caption/audio options with accuracy checks and clear source/translation labeling.
- Make local manual mode usable without an AI model, microphone, network access, or API account.

## Contributions / 참여 안내

Feedback from people with speech disabilities and other accessibility needs is especially welcome. Please report problems using **synthetic sample notes** and avoid sharing personal speech recordings or real conference materials.
발화장애 당사자의 사용성 의견을 환영합니다. 이슈에는 개인 음성 파일이나 실제 발표 대본 대신 가상 예문을 사용해 주세요.