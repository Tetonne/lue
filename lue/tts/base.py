"""Abstract base class for TTS models in the Lue eBook reader."""

import asyncio
from abc import ABC, abstractmethod
import hashlib
import logging
import os
import re
import shutil
from typing import Optional, List, Tuple, Any
from rich.console import Console

from .. import audio, config, timing_calculator

# Intégration optionnelle de la prosodie française / anglaise
try:
    from french_prosody import get_french_prosody_engine
    HAS_FRENCH_PROSODY = True
except ImportError:
    HAS_FRENCH_PROSODY = False

try:
    from english_prosody import get_english_prosody_engine
    HAS_ENGLISH_PROSODY = True
except ImportError:
    HAS_ENGLISH_PROSODY = False


def normalize_text(text: str, language: str = "en") -> str:
    """
    Nettoie les artefacts ePub courants (notes de bas de page, liens, 
    multi-espaces, ponctuation aberrante) pour fluidifier la lecture TTS.
    """
    if not text:
        return ""
    
    # 1. Supprime les notes de bas de page entre crochets [1], [2], etc.
    cleaned = re.sub(r'\[\d+\]', '', text)
    
    # 2. Normalise les abréviations selon la langue
    lang = language.lower() if language else "en"
    if lang.startswith("fr"):
        cleaned = re.sub(r'\bDr\.\s*', 'Docteur ', cleaned)
        cleaned = re.sub(r'\bM\.\s*', 'Monsieur ', cleaned)
        cleaned = re.sub(r'\bMme\.\s*', 'Madame ', cleaned)
        cleaned = re.sub(r'\bProf\.\s*', 'Professeur ', cleaned)
    else:
        cleaned = re.sub(r'\bDr\.\s*', 'Doctor ', cleaned)
        cleaned = re.sub(r'\bMr\.\s*', 'Mister ', cleaned)
        cleaned = re.sub(r'\bMrs\.\s*', 'Missus ', cleaned)
        cleaned = re.sub(r'\bProf\.\s*', 'Professor ', cleaned)
    
    # 3. Supprime les URL ou liens
    cleaned = re.sub(r'http[s]?://\S+', '', cleaned)
    
    # 4. Nettoie les espaces multiples
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    
    return cleaned


def prepare_text_chunk(text: str, language: str = "en") -> str:
    """
    Prépare et normalise le bloc de texte pour le TTS de manière multilingue (FR / EN).
    """
    if not text:
        return ""
        
    normalized = normalize_text(text, language)
    if not normalized:
        return ""
        
    lang = language.lower() if language else "en"
    
    if HAS_FRENCH_PROSODY and lang.startswith("fr"):
        return get_french_prosody_engine().prepare_for_tts(normalized)
        
    if HAS_ENGLISH_PROSODY and lang.startswith("en"):
        return get_english_prosody_engine().prepare_for_tts(normalized)
        
    return normalized


class TTSBase(ABC):
    """
    Abstract base class for all TTS models.
    
    This class defines the interface that all TTS models must implement
    to be compatible with the Lue eBook reader.
    """

    def __init__(self, console: Console, voice: Optional[str] = None, lang: Optional[str] = None):
        """
        Initialize the TTS model.
        
        Args:
            console: Rich console instance for user feedback
            voice: Optional voice for the TTS model
            lang: Optional language for the TTS model
        """
        self.console = console
        self.voice = voice
        self.lang = lang
        self.initialized = False

    @property
    @abstractmethod
    def name(self) -> str:
        """Get the unique identifier for this TTS model."""
        pass

    @property
    @abstractmethod
    def output_format(self) -> str:
        """Get the audio format this model produces (e.g. 'mp3', 'wav')."""
        pass

    @abstractmethod
    async def initialize(self) -> bool:
        """Initialize the TTS model asynchronously."""
        pass

    @abstractmethod
    async def generate_audio(self, text: str, output_path: str):
        """Generate audio from text and save to file."""
        pass

    def _get_cache_file_path(self, text: str) -> Optional[str]:
        """
        Génère un chemin de cache unique basé sur le hash SHA-256 du texte, de la voix et de la langue.
        """
        try:
            cache_dir = getattr(config, "AUDIO_CACHE_DIR", os.path.join(getattr(config, "AUDIO_DATA_DIR", "."), "cache"))
            os.makedirs(cache_dir, exist_ok=True)
            
            unique_string = f"{self.name}_{self.voice}_{self.lang}_{text}"
            text_hash = hashlib.sha256(unique_string.encode('utf-8')).hexdigest()
            
            return os.path.join(cache_dir, f"{text_hash}.{self.output_format}")
        except Exception as e:
            logging.warning(f"Impossible de déterminer le chemin du cache : {e}")
            return None

    async def generate_audio_with_timing(self, text: str, output_path: str) -> Any:
        """
        Generate audio from text and save to file, returning processed timing information.
        Intègre la normalisation du texte et le cache audio persistant de manière transparente.
        """
        # 1. Prépare et normalise le texte
        processed_text = prepare_text_chunk(text, self.lang)
        cache_path = self._get_cache_file_path(processed_text)
        
        cache_hit = False
        # Vérifie si le fichier audio existe déjà dans le cache
        if cache_path and await asyncio.to_thread(os.path.exists, cache_path):
            try:
                cache_size = await asyncio.to_thread(os.path.getsize, cache_path)
                if cache_size > 0:
                    await asyncio.to_thread(shutil.copy2, cache_path, output_path)
                    cache_hit = True
            except Exception as e:
                logging.warning(f"Erreur lors de la lecture du cache audio : {e}")
                cache_hit = False

        if not cache_hit:
            # Génère l'audio via l'implémentation spécifique du modèle
            await self.generate_audio(processed_text, output_path)
            
            # Sauvegarde atomique dans le cache via un fichier temporaire
            if cache_path and await asyncio.to_thread(os.path.exists, output_path):
                try:
                    out_size = await asyncio.to_thread(os.path.getsize, output_path)
                    if out_size > 0:
                        temp_cache_path = f"{cache_path}.tmp"
                        await asyncio.to_thread(shutil.copy2, output_path, temp_cache_path)
                        await asyncio.to_thread(os.replace, temp_cache_path, cache_path)
                except Exception as e:
                    logging.warning(f"Erreur lors de la sauvegarde dans le cache audio : {e}")
        
        # 2. Récupère les timings bruts avec le texte préparé
        raw_timings = await self.get_raw_timing_data(processed_text, output_path)
        
        # Récupère la durée réelle de l'audio
        duration = await audio.get_audio_duration(output_path)
        
        # 3. Traite les données de timing via le calculateur centralisé
        return timing_calculator.process_tts_timing_data(text, raw_timings, duration)

    async def get_raw_timing_data(self, text: str, output_path: str) -> List[Tuple[str, float, float]]:
        """Get raw timing data from the TTS engine."""
        return []

    async def warm_up(self):
        """Warm up the model to reduce initial latency."""
        pass

    def get_overlap_seconds(self) -> Optional[float]:
        """Get the TTS-specific overlap seconds for this model."""
        return config.TTS_OVERLAP_SECONDS.get(self.name)