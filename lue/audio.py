"""
Audio management, text cleaning, and playback/producer loops for the Lue eBook reader.
"""

import asyncio
import os
import re
import subprocess
import logging
from functools import lru_cache
from typing import Optional, List, Tuple, Any

from . import config, content_parser

# This pattern is used to both clean text for TTS and detect sentence fragments.
ABBREVIATION_PATTERN = r'\b(Mr|Mrs|Ms|Dr|Prof|Rev|Hon|Jr|Sr|Cpl|Sgt|Gen|Col|Capt|Lt|Pvt|vs|viz|Co|Inc|Ltd|Corp|St|Ave|Blvd)\.'
INITIAL_PATTERN = r'\b([A-Z])\.(?=\s[A-Z])'

# Compile regex patterns once for better performance
COMPILED_ABBR_REGEX = re.compile(ABBREVIATION_PATTERN)
COMPILED_INIT_REGEX = re.compile(INITIAL_PATTERN)
COMPILED_PUNCT_REGEX = re.compile(r'(?:^|\s)[.,:;!?]+(?=\s|$)')
COMPILED_DASH_REGEX = re.compile(r'(?:^|\s)-(?=")')
COMPILED_WHITESPACE_REGEX = re.compile(r'\s+')

# Word mapping functionality moved to timing_calculator.py
# Import it here for backward compatibility
from .timing_calculator import create_word_mapping as _create_word_mapping


@lru_cache(maxsize=1024)
def clean_tts_text(text: str) -> str:
    """
    Removes periods from specific English abbreviations and single initials
    to prevent unnatural pauses in TTS engines. Also removes loose punctuation
    marks that are not connected to any word.
    """
    if not text:
        return ""
        
    # Remove periods from abbreviations and initials
    text = COMPILED_ABBR_REGEX.sub(r'\1', text)
    text = COMPILED_INIT_REGEX.sub(r'\1 ', text)
    
    # Remove loose punctuation marks that are standalone (not connected to words)
    text = COMPILED_PUNCT_REGEX.sub(' ', text)
    
    # Remove standalone dashes that are followed by quotation marks
    text = COMPILED_DASH_REGEX.sub(' ', text)
    
    # Clean up any extra whitespace that might result from removing punctuation
    text = COMPILED_WHITESPACE_REGEX.sub(' ', text).strip()
    
    return text


async def stop_and_clear_audio(reader: Any) -> None:
    """Stop audio playback and clear the audio queue."""
    tasks_to_cancel = []
    for task in [getattr(reader, 'producer_task', None), getattr(reader, 'player_task', None)]:
        if task and not task.done():
            task.cancel()
            tasks_to_cancel.append(task)
            
    if tasks_to_cancel:
        await asyncio.gather(*tasks_to_cancel, return_exceptions=True)
    
    reader.producer_task = None
    reader.player_task = None
    
    processes_to_kill = reader.playback_processes.copy()
    reader.playback_processes.clear()
    
    for process in processes_to_kill:
        try:
            if process.returncode is None:
                process.terminate()
                try: 
                    await asyncio.wait_for(process.wait(), timeout=0.2)
                except asyncio.TimeoutError:
                    process.kill()
                    await asyncio.wait_for(process.wait(), timeout=0.1)
        except (ProcessLookupError, AttributeError, asyncio.TimeoutError): 
            pass
    
    try:
        pkill_proc = await asyncio.create_subprocess_exec(
            'pkill', '-9', '-f', 'ffplay.*buffer_', 
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        await asyncio.wait_for(pkill_proc.wait(), timeout=0.3)
    except (FileNotFoundError, asyncio.TimeoutError): 
        pass
    
    while not reader.audio_queue.empty():
        try:
            reader.audio_queue.get_nowait()
            reader.audio_queue.task_done()
        except asyncio.QueueEmpty: 
            break
    
    await asyncio.sleep(0.1)
    
    # Aggressive cleanup with longer delays for file system operations
    for buf_base in config.AUDIO_BUFFERS:
        for ext in ['.mp3', '.wav']:
            buf = f"{buf_base}{ext}"
            for attempt in range(5):
                try:
                    if os.path.exists(buf): 
                        os.remove(buf)
                    break
                except OSError:
                    if attempt < 4: 
                        await asyncio.sleep(0.1)
            
    await asyncio.sleep(0.2)


async def get_audio_duration(file_path: str) -> Optional[float]:
    """Get the duration of an audio file."""
    command = [
        'ffprobe', '-v', 'error', '-show_entries', 'format=duration', 
        '-of', 'default=noprint_wrappers=1:nokey=1', file_path
    ]
    process = await asyncio.create_subprocess_exec(
        *command, stdout=asyncio.subprocess.PIPE, stderr=subprocess.DEVNULL
    )
    stdout, _ = await process.communicate()
    if process.returncode != 0: 
        return None
    try: 
        return float(stdout.decode().strip())
    except (ValueError, TypeError): 
        return None


async def play_from_current_position(reader: Any) -> None:
    """Start the audio producer and player loops."""
    if not reader.is_paused and reader.running and reader.tts_model:
        for task in [getattr(reader, 'producer_task', None), getattr(reader, 'player_task', None)]:
            if task and not task.done():
                task.cancel()
                try: 
                    await asyncio.wait_for(task, timeout=2.0)
                except (asyncio.CancelledError, asyncio.TimeoutError): 
                    pass
        
        reader.producer_task = None
        reader.player_task = None
        
        await asyncio.sleep(0.05)
        
        reader.producer_task = asyncio.create_task(_producer_loop(reader))
        reader.player_task = asyncio.create_task(_player_loop(reader))


async def _producer_loop(reader: Any) -> None:
    """Producer loop to generate audio files."""
    if not reader.tts_model or not reader.tts_model.initialized:
        try: 
            await asyncio.wait_for(reader.audio_queue.put(None), timeout=0.5)
        except (asyncio.TimeoutError, asyncio.CancelledError): 
            pass
        return

    producer_pos = (reader.chapter_idx, reader.paragraph_idx, reader.sentence_idx)
    buffer_index = 0
    try:
        while reader.running:
            if reader.audio_queue.full():
                await asyncio.sleep(0.1)
                continue
            try:
                c, p, s = producer_pos
                sentences = content_parser.split_into_sentences(reader.chapters[c][p])
                text = sentences[s]
            except IndexError: 
                break
                
            if not text or not text.strip():
                next_pos = reader._advance_position(producer_pos, wrap=False)
                if not next_pos: 
                    break
                producer_pos = next_pos
                continue

            merged = False
            is_abbrev_fragment = bool(COMPILED_ABBR_REGEX.fullmatch(text.strip()))

            if is_abbrev_fragment and s + 1 < len(sentences):
                text += " " + sentences[s+1]
                merged = True

            original_text = text
            output_format = reader.tts_model.output_format
            output_filename = f"{config.AUDIO_BUFFERS[buffer_index]}.{output_format}"
            
            try:
                if not reader.running: 
                    break
                
                for attempt in range(3):
                    try:
                        if os.path.exists(output_filename): 
                            os.remove(output_filename)
                        break
                    except OSError:
                        if attempt < 2: 
                            await asyncio.sleep(0.05)
                
                # Phase 1 (IPA + prosody) is applied only on the synthesis path.
                # The original_text remains the source of timing/navigation data.
                sanitized_text = reader.prepare_tts_text(original_text)
                timing_info = None
                
                if hasattr(reader.tts_model, 'generate_audio_with_timing'):
                    try:
                        timing_info = await reader.tts_model.generate_audio_with_timing(sanitized_text, output_filename)
                    except Exception as e:
                        logging.error(f"TTS timing generation failed for text '{original_text[:50]}...': {e}")
                        await reader.tts_model.generate_audio(sanitized_text, output_filename)
                else:
                    await reader.tts_model.generate_audio(sanitized_text, output_filename)

                duration = await get_audio_duration(output_filename)
                
                if not reader.running: 
                    break
                
                if timing_info is None:
                    from .timing_calculator import process_tts_timing_data
                    timing_info = process_tts_timing_data(original_text, [], duration)
                
                await asyncio.wait_for(
                    reader.audio_queue.put((output_filename, *producer_pos, duration, timing_info)), 
                    timeout=1.0
                )
                
                next_pos = reader._advance_position(producer_pos, wrap=False)
                if merged and next_pos:
                    next_pos = reader._advance_position(next_pos, wrap=False)

                if not next_pos: 
                    break
                producer_pos = next_pos
                buffer_index = (buffer_index + 1) % len(config.AUDIO_BUFFERS)
                
            except asyncio.CancelledError: 
                break
            except Exception as e:
                if reader.running:
                    try:
                        sanitized_for_log = content_parser.sanitize_text_for_tts(original_text) if 'original_text' in locals() else 'N/A'
                        original_for_log = original_text if 'original_text' in locals() else 'N/A'
                        logging.error(f"TTS Error in producer: {e}\nOriginal: '{original_for_log[:100]}...'\nSanitized: '{sanitized_for_log[:100]}...'", exc_info=True)
                    except Exception:
                        logging.error(f"TTS Error in producer: {e}", exc_info=True)
                    await asyncio.sleep(2)
                continue
    except asyncio.CancelledError: 
        pass
    finally:
        try: 
            await asyncio.wait_for(reader.audio_queue.put(None), timeout=0.5)
        except (asyncio.TimeoutError, asyncio.CancelledError): 
            pass


async def _player_loop(reader: Any) -> None:
    """Player loop to play audio files."""
    try:
        while reader.running:
            try:
                item = await asyncio.wait_for(reader.audio_queue.get(), timeout=1.0)
                if item is None:
                    reader.audio_queue.task_done()
                    if reader.active_playback_tasks:
                        await asyncio.gather(*reader.active_playback_tasks, return_exceptions=True)
                    reader.playback_finished_event.set()
                    break
                    
                audio_file, c, p, s, duration, timing_data = item
                if isinstance(timing_data, dict):
                    timing_info = timing_data
                else:
                    timing_info = {"word_timings": timing_data, "speech_duration": duration, "total_duration": duration}

                if not os.path.exists(audio_file) or duration is None or duration <= 0:
                    reader.audio_queue.task_done()
                    continue
                    
                try:
                    reader.loop.call_soon_threadsafe(
                        reader._post_command_sync,
                        ('_new_sentence_started', (c, p, s, duration, timing_data))
                    )
                except RuntimeError:
                    reader.audio_queue.task_done()
                    break
                    
                try:
                    cmd = ['ffplay', '-nodisp', '-autoexit', '-loglevel', 'error']
                    
                    if abs(reader.playback_speed - 1.0) > 0.01:
                        speed = reader.playback_speed
                        filters = []
                        
                        while speed > 2.0:
                            filters.append('atempo=2.0')
                            speed /= 2.0
                        while speed < 0.5:
                            filters.append('atempo=0.5')
                            speed /= 0.5
                        if abs(speed - 1.0) > 0.01:
                            filters.append(f'atempo={speed:.3f}')
                        
                        if filters:
                            cmd.extend(['-af', ','.join(filters)])
                    
                    cmd.append(audio_file)
                    process = await asyncio.create_subprocess_exec(
                        *cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
                    )
                    reader.playback_processes.append(process)
                except Exception:
                    reader.audio_queue.task_done()
                    continue

                async def await_and_remove(proc: asyncio.subprocess.Process, file: str) -> None:
                    task = asyncio.current_task()
                    try:
                        await proc.wait()
                    except Exception: 
                        pass
                    finally:
                        try:
                            if proc in reader.playback_processes: 
                                reader.playback_processes.remove(proc)
                        except ValueError: 
                            pass
                        for attempt in range(3):
                            try:
                                if os.path.exists(file): 
                                    os.remove(file)
                                break
                            except OSError:
                                if attempt < 2: 
                                    await asyncio.sleep(0.05)
                        if task in reader.active_playback_tasks:
                            reader.active_playback_tasks.remove(task)

                playback_task = asyncio.create_task(await_and_remove(process, audio_file))
                reader.active_playback_tasks.append(playback_task)
                
                base_overlap = config.OVERLAP_SECONDS
                if reader.tts_model and hasattr(reader.tts_model, 'get_overlap_seconds'):
                    tts_overlap = reader.tts_model.get_overlap_seconds()
                    if tts_overlap is not None:
                        base_overlap = tts_overlap
                
                speed = reader.playback_speed
                if speed >= 3.0:
                    overlap_seconds = 0.0
                else:
                    overlap_factor = max(0.0, min(1.0, (3.0 - speed) / 2.0))
                    overlap_seconds = base_overlap * overlap_factor
                
                actual_duration = duration / speed
                
                await asyncio.sleep(max(0.1, actual_duration - overlap_seconds))
                reader.audio_queue.task_done()
                
            except asyncio.TimeoutError:
                if not reader.running: 
                    break
                continue
            except asyncio.CancelledError: 
                break
    except asyncio.CancelledError: 
        pass
    finally:
        for process in reader.playback_processes.copy():
            try:
                if process.returncode is None: 
                    process.terminate()
            except (ProcessLookupError, AttributeError): 
                pass
