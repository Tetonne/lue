"""Configuration settings for the Lue eBook reader."""

import os
from platformdirs import user_data_dir, user_cache_dir

# Default TTS model
DEFAULT_TTS_MODEL = "edge"

# PUNCTUATION_PAUSE ajout 20260910
PUNCTUATION_PAUSE = {
    ".": 0.5,  # Pause de 500ms sur le point
    ",": 0.2,  # Pause de 200ms sur la virgule
    ";": 0.3,  # Pause de 300ms sur le point-virgule
}

# Default voices for TTS models
TTS_VOICES = {
    "edge": "fr-FR-DeniseNeural",      # Voix Edge TTS naturelle en français
    "kokoro": "ff_siwis",              # Voix Kokoro en français (si disponible, sinon fallback)
    "apple": "Amélie",                 # Voix native macOS Silicon (français)
}

# Language codes for TTS models that require them
TTS_LANGUAGE_CODES = {
    "edge": "fr-FR",
    "kokoro": "f",  # a=English, e=Spanish, j=Japanese, etc.
    "apple": "fr_FR",
}

# TTS model-specific seconds of overlap between sentences (overrides default OVERLAP_SECONDS if specified)
TTS_OVERLAP_SECONDS = {
    "kokoro": 0.6,
}


# Audio processing settings
AUDIO_DATA_DIR = user_cache_dir("lue")
os.makedirs(AUDIO_DATA_DIR, exist_ok=True)
AUDIO_BUFFERS = [os.path.join(AUDIO_DATA_DIR, f"buffer_{i}") for i in range(6)]
MAX_QUEUE_SIZE = 4
OVERLAP_SECONDS = 0.5 # Seconds of overlap between sentences

# Progress tracking settings
PROGRESS_FILE_DIR = user_data_dir("lue")
os.makedirs(PROGRESS_FILE_DIR, exist_ok=True)

# General settings
SHOW_ERRORS_ON_EXIT = True

# PDF parsing settings
PDF_FILTERS_ENABLED = False  # You can also enable this with the --filter or -f command-line option
PDF_FILTER_HEADERS = True  # Filter headers in top margin of pages
PDF_FILTER_FOOTNOTES = True  # Filter page numbers and footnotes in bottom margin of pages

# PDF filtering thresholds (only used when respective filters are enabled)
PDF_HEADER_MARGIN = 0.1  # Top 10% of page considered header area
PDF_FOOTNOTE_MARGIN = 0.1  # Bottom 10% of page considered footnote area

# UI settings
SMOOTH_SCROLLING_ENABLED = True  # Enable smooth scrolling for keyboard navigation
UI_MODE = 2  # 0=minimal (text only), 1=medium (top bar only), 2=full (default), 3=speed reading

# Highlighting settings
SENTENCE_HIGHLIGHTING_ENABLED = True  # Enable sentence-level highlighting
WORD_HIGHLIGHT_MODE = 1  # 0=off, 1=normal highlighting, 2=standout highlighting

# Keyboard settings
# Can be set to "default", "vim", or a path to a custom keyboard shortcuts JSON file
CUSTOM_KEYBOARD_SHORTCUTS = "default"
