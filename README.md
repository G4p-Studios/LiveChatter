# LiveChatter
## YouTube Live Chat Reader &amp; Summarizer.

LiveChatter is a PyQt6 application that:
- Reads live YouTube chat (pytchat; optional chat-downloader fallback)
- Two chat modes: "Standard (read messages)" and "Summaries (periodic AI summaries)"
- Summary providers: OpenAI / Gemini
- TTS options: System (accessible_output2: NVDA/JAWS/SAPI5), OpenAI TTS, Google Cloud TTS, Amazon Polly, ElevenLabs
- Sound packs via sound_lib (optional)
- Options dialog for API keys and sound settings
- Config persisted as JSON in a .livechatter folder under the current working directory

NOTE: External services require API keys and libraries installed.
