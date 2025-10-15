"""
Voice-Controlled AI Shell Assistant for Windows with OpenAI Integration
Unlimited offline TTS using pyttsx3 - FIXED VERSION
"""

import os
import sys
import subprocess
import psutil
import threading
import queue
import json
from pathlib import Path
from typing import List, Dict, Tuple, Optional
import re
from dotenv import load_dotenv
import time

# Load environment variables
load_dotenv()

# Try to import speech recognition libraries
try:
    import speech_recognition as sr
    SPEECH_AVAILABLE = True
except ImportError:
    SPEECH_AVAILABLE = False
    print("⚠️ speech_recognition not installed. Run: pip install SpeechRecognition pydub pyaudio")

# TTS will use pyttsx3
try:
    import pyttsx3
    TTS_AVAILABLE = True
except ImportError:
    TTS_AVAILABLE = False
    print("⚠️ pyttsx3 not installed. Run: pip install pyttsx3")

# Import OpenAI
try:
    from openai import OpenAI
    LLM_AVAILABLE = True
except ImportError:
    LLM_AVAILABLE = False
    print("⚠️ openai not installed. Run: pip install openai")


class TextToSpeech:
    """
    Improved TTS using pyttsx3 with proper cleanup and threading
    """
    
    def __init__(self, client=None):
        self.client = client
        self.is_speaking = False
        self.lock = threading.Lock()
        self.engine = None
        self.pyttsx3_available = TTS_AVAILABLE
        
        if self.pyttsx3_available:
            print("✓ Using pyttsx3 offline TTS (unlimited)")
            self._init_engine()
    
    def _init_engine(self):
        """Initialize pyttsx3 engine with error handling"""
        try:
            if self.engine:
                try:
                    self.engine.stop()
                except:
                    pass
                del self.engine
                time.sleep(0.1)  # Brief pause for cleanup
            
            self.engine = pyttsx3.init()
            self.engine.setProperty('rate', 180)
            self.engine.setProperty('volume', 0.9)
            
            # Set better voice if available
            voices = self.engine.getProperty('voices')
            for voice in voices:
                if 'zira' in voice.name.lower() or 'david' in voice.name.lower():
                    self.engine.setProperty('voice', voice.id)
                    break
        except Exception as e:
            print(f"⚠️ TTS engine init error: {e}")
            self.pyttsx3_available = False
    
    def speak(self, text: str, wait: bool = True, force: bool = False):
        """Convert text to speech with proper resource management"""
        if not text or len(text.strip()) == 0 or not self.pyttsx3_available:
            return
        
        # Non-blocking lock check
        if not self.lock.acquire(blocking=False):
            return  # Skip if already speaking
        
        try:
            self.is_speaking = True
            
            # Truncate text if too long
            text_to_speak = text[:500] if len(text) > 500 else text
            
            print(f"🔊 Speaking: {text_to_speak[:50]}...")
            
            # Reinitialize engine for each call (more reliable)
            self._init_engine()
            
            if not self.engine:
                return
            
            # Speak
            self.engine.say(text_to_speak)
            self.engine.runAndWait()
            
            # Cleanup
            try:
                self.engine.stop()
            except:
                pass
            
        except Exception as e:
            print(f"⚠️ TTS Error: {e}")
            # Try to recover
            self.pyttsx3_available = TTS_AVAILABLE
            if self.pyttsx3_available:
                self._init_engine()
        finally:
            self.is_speaking = False
            self.lock.release()
            time.sleep(0.3)  # Pause to release audio device
    
    def stop(self):
        """Stop speaking and cleanup"""
        self.is_speaking = False
        if self.engine:
            try:
                self.engine.stop()
            except:
                pass


class SpeechToText:
    """
    High-accuracy STT using OpenAI Whisper API
    Much better recognition than Google Speech API
    """
    
    def __init__(self, client=None):
        self.recognizer = None
        self.client = client
        self.temp_audio_file = Path("temp_audio.wav")
        
        if SPEECH_AVAILABLE:
            self.recognizer = sr.Recognizer()
            # Optimized settings for clear audio capture
            self.recognizer.energy_threshold = 2000
            self.recognizer.dynamic_energy_threshold = True
            self.recognizer.pause_threshold = 0.8
            self.recognizer.non_speaking_duration = 0.5
            self.recognizer.phrase_threshold = 0.3
            
            if self.client:
                print("✓ Using OpenAI Whisper for STT (high accuracy)")
            else:
                print("⚠️ OpenAI client not available, using Google STT")
    
    def listen(self, timeout: int = 30) -> Optional[str]:
        """Listen to microphone and transcribe using Whisper API"""
        if not self.recognizer:
            return None
        
        try:
            with sr.Microphone(sample_rate=16000) as source:
                print("🎤 Listening... (speak clearly)")
                
                # Quick ambient noise adjustment
                self.recognizer.adjust_for_ambient_noise(source, duration=0.5)
                
                # Listen for speech
                audio = self.recognizer.listen(
                    source,
                    timeout=timeout,
                    phrase_time_limit=20
                )
            
            print("⏳ Processing with Whisper AI...")
            
            # Try OpenAI Whisper first (much more accurate)
            if self.client:
                try:
                    # Save audio to temporary WAV file
                    wav_data = audio.get_wav_data()
                    with open(self.temp_audio_file, "wb") as f:
                        f.write(wav_data)
                    
                    # Transcribe using Whisper
                    with open(self.temp_audio_file, "rb") as audio_file:
                        transcript = self.client.audio.transcriptions.create(
                            model="whisper-1",
                            file=audio_file,
                            language="en"
                        )
                    
                    result = transcript.text.strip()
                    
                    # Cleanup temp file
                    try:
                        self.temp_audio_file.unlink()
                    except:
                        pass
                    
                    # Show transcription with detected PIDs
                    pids = re.findall(r'\b(\d{2,6})\b', result)
                    if pids:
                        print(f"📝 Whisper: \"{result}\" [PIDs: {', '.join(pids)}]")
                    else:
                        print(f"📝 Whisper: \"{result}\"")
                    
                    return result
                
                except Exception as e:
                    print(f"⚠️ Whisper error: {e}, falling back to Google...")
            
            # Fallback to Google Speech Recognition
            try:
                result = self.recognizer.recognize_google(audio, language='en-US')
                
                pids = re.findall(r'\b(\d{2,6})\b', result)
                if pids:
                    print(f"📝 Google: \"{result}\" [PIDs: {', '.join(pids)}]")
                else:
                    print(f"📝 Google: \"{result}\"")
                
                return result
            
            except sr.UnknownValueError:
                print("❌ Could not understand audio - speak louder and clearer")
                return None
            except sr.RequestError as e:
                print(f"❌ Recognition service error: {e}")
                return None
        
        except Exception as e:
            print(f"❌ Microphone error: {e}")
            return None


class ProcessManager:
    """Process Management using psutil"""
    
    @staticmethod
    def get_top_processes(n: int = 5, sort_by: str = 'cpu') -> List[Dict]:
        """Get top N processes by CPU or memory"""
        processes = []
        try:
            for proc in psutil.process_iter(['pid', 'name', 'cpu_percent', 'memory_percent']):
                try:
                    processes.append({
                        'pid': proc.info['pid'],
                        'name': proc.info['name'],
                        'cpu': proc.info['cpu_percent'] or 0,
                        'memory': proc.info['memory_percent'] or 0
                    })
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
            
            if sort_by == 'memory':
                processes.sort(key=lambda x: x['memory'], reverse=True)
            else:
                processes.sort(key=lambda x: x['cpu'], reverse=True)
            
            return processes[:n]
        except Exception as e:
            return [{'error': str(e)}]
    
    @staticmethod
    def kill_process(pid: int) -> str:
        """Kill a specific process by PID"""
        try:
            proc = psutil.Process(pid)
            proc_name = proc.name()
            proc.terminate()
            proc.wait(timeout=3)
            return f"✓ Process {proc_name} (PID {pid}) terminated successfully"
        except psutil.NoSuchProcess:
            return f"❌ Process {pid} not found"
        except psutil.AccessDenied:
            return f"❌ Access denied for PID {pid}. Need admin rights"
        except psutil.TimeoutExpired:
            try:
                proc.kill()
                return f"✓ Process {proc_name} (PID {pid}) force killed"
            except:
                return f"❌ Failed to kill process {pid}"
    
    @staticmethod
    def kill_by_name(name: str, exclude: List[str] = None) -> str:
        """Kill all processes by name"""
        exclude = exclude or []
        killed = []
        try:
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    if name.lower() in proc.info['name'].lower():
                        if not any(ex.lower() in proc.info['name'].lower() for ex in exclude):
                            proc.terminate()
                            killed.append(f"{proc.info['name']} (PID {proc.info['pid']})")
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
            
            if killed:
                return f"✓ Terminated {len(killed)} process(es): {', '.join(killed)}"
            else:
                return f"❌ No processes named '{name}' found"
        except Exception as e:
            return f"❌ Error: {str(e)}"
    
    @staticmethod
    def get_system_info() -> Dict:
        """Get system resource information"""
        try:
            cpu_freq = psutil.cpu_freq()
            mem = psutil.virtual_memory()
            disk = psutil.disk_usage('/')
            return {
                'cpu_percent': psutil.cpu_percent(interval=0.1),
                'cpu_count': psutil.cpu_count(),
                'cpu_freq': f"{cpu_freq.current:.0f}" if cpu_freq else "N/A",
                'memory_percent': mem.percent,
                'memory_used': mem.used / 1024**3,
                'memory_total': mem.total / 1024**3,
                'disk_percent': disk.percent,
                'disk_used': disk.used / 1024**3,
                'disk_total': disk.total / 1024**3,
                'process_count': len(psutil.pids())
            }
        except Exception as e:
            return {'error': str(e)}
    
    @staticmethod
    def get_process_info(pid: int = None) -> Dict:
        """Get detailed info about a specific process"""
        try:
            if pid:
                proc = psutil.Process(pid)
                return {
                    'name': proc.name(),
                    'pid': proc.pid,
                    'status': proc.status(),
                    'cpu_percent': proc.cpu_percent(),
                    'memory_mb': proc.memory_info().rss / 1024 / 1024,
                    'create_time': proc.create_time(),
                    'num_threads': proc.num_threads()
                }
            else:
                return {'error': 'PID not specified'}
        except psutil.NoSuchProcess:
            return {'error': f'Process {pid} not found'}
        except Exception as e:
            return {'error': str(e)}
    
    @staticmethod
    def list_files(path: str = '.') -> str:
        """List files in a directory"""
        try:
            path = Path(path).expanduser()
            if not path.exists():
                return f"❌ Path {path} does not exist"
            
            if path.is_file():
                return f"ℹ️ This is a file: {path.name}"
            
            items = list(path.iterdir())
            if not items:
                return f"ℹ️ Directory {path} is empty"
            
            result = f"Contents of {path}:\n"
            for item in sorted(items)[:15]:
                if item.is_dir():
                    result += f"📁 {item.name}\n"
                else:
                    size_kb = item.stat().st_size / 1024
                    result += f"📄 {item.name} ({size_kb:.1f} KB)\n"
            
            if len(items) > 15:
                result += f"... and {len(items) - 15} more items"
            
            return result
        except Exception as e:
            return f"❌ Error: {str(e)}"


class ConversationalAssistant:
    """LLM-based Command Interpretation"""
    
    def __init__(self, api_key: str = None):
        if not LLM_AVAILABLE:
            raise ImportError("OpenAI library not installed")
        
        self.api_key = api_key or os.getenv('OPENAI_API_KEY')
        if not self.api_key:
            raise ValueError("OPENAI_API_KEY not found")
        
        self.client = OpenAI(api_key=self.api_key)
        self.model = "gpt-3.5-turbo"
        self.conversation_history = []
        self.system_prompt = self._get_system_prompt()
    
    def _get_system_prompt(self) -> str:
        return """You are a helpful AI Shell Assistant for Windows.

Available commands:
1. top_processes - Show top processes
   params: count (1-20), sort_by (cpu/memory)
   
2. kill_process - Kill process by PID
   params: pid (integer)
   
3. kill_by_name - Kill processes by name
   params: name (string)
   
4. system_info - Show system stats
   
5. process_info - Get process details
   params: pid (integer)
   
6. list_files - List directory contents
   params: path (string)

CRITICAL PID EXTRACTION RULES:
- Extract ALL digits from commands like: "kill 1234", "terminate 5678", "stop 9012"
- Look for [PID detected: NUMBER] hints in user messages
- If you see a number after "kill/terminate/stop", use that exact PID
- NEVER use fake PIDs - only real numbers from user input
- If unclear, show top_processes first

Response format: "Short response. <ACTION>{"command": "...", "params": {...}}</ACTION>"

Examples:
- "kill 21808" → "Terminating. <ACTION>{"command": "kill_process", "params": {"pid": 21808}}</ACTION>"
- "top processes" → "Checking. <ACTION>{"command": "top_processes", "params": {"count": 5}}</ACTION>"
- "21808" → "Stopping PID 21808. <ACTION>{"command": "kill_process", "params": {"pid": 21808}}</ACTION>"

Keep responses under 15 words. Execute commands immediately."""
    
    def chat(self, user_message: str) -> Tuple[str, Optional[str], Optional[Dict]]:
        """Process user message"""
        try:
            # Enhanced PID detection with explicit hints
            kill_keywords = ['kill', 'terminate', 'stop', 'end', 'close', 'quit']
            numbers = re.findall(r'\b(\d{2,6})\b', user_message)
            
            # Add PID hint if command looks like a kill command
            has_kill_word = any(word in user_message.lower() for word in kill_keywords)
            is_just_number = len(user_message.split()) <= 2 and numbers
            
            if numbers and (has_kill_word or is_just_number):
                # Add explicit hint for LLM
                user_message = f"{user_message} [PID detected: {numbers[0]}]"
            
            self.conversation_history.append({
                "role": "user",
                "content": user_message
            })
            
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "system", "content": self.system_prompt}] + self.conversation_history,
                temperature=0.5,  # Lower for more consistent parsing
                max_tokens=150
            )
            
            assistant_message = response.choices[0].message.content
            
            self.conversation_history.append({
                "role": "assistant",
                "content": assistant_message
            })
            
            # Keep history manageable
            if len(self.conversation_history) > 20:
                self.conversation_history = self.conversation_history[-20:]
            
            # Parse response
            command = None
            params = None
            response_text = assistant_message
            
            if "<ACTION>" in assistant_message and "</ACTION>" in assistant_message:
                action_start = assistant_message.find("<ACTION>") + len("<ACTION>")
                action_end = assistant_message.find("</ACTION>")
                action_json = assistant_message[action_start:action_end].strip()
                response_text = assistant_message[:assistant_message.find("<ACTION>")].strip()
                
                try:
                    action_data = json.loads(action_json)
                    command = action_data.get('command')
                    params = action_data.get('params', {})
                except json.JSONDecodeError as e:
                    print(f"⚠️ JSON parse error: {e}")
            
            return response_text, command, params
        
        except Exception as e:
            return f"Error: {str(e)}", None, None


class ShellAssistant:
    """Main AI Shell Assistant"""
    
    def __init__(self):
        self.proc_manager = ProcessManager()
        self.assistant = None
        self.running = True
        self.voice_mode = True
        self.recent_pids = []
        
        try:
            self.assistant = ConversationalAssistant()
            print("✓ OpenAI API connected")
        except Exception as e:
            print(f"✗ Error: {e}")
            sys.exit(1)
        
        self.tts = TextToSpeech(client=self.assistant.client if self.assistant else None)
        self.stt = SpeechToText(client=self.assistant.client if self.assistant else None)
        
        if TTS_AVAILABLE:
            print("✓ Text-to-Speech ready")
        if SPEECH_AVAILABLE:
            print("✓ Speech-to-Text ready")
    
    def speak(self, text: str, wait: bool = True, force: bool = False):
        """Speak text with proper timing"""
        if not text or len(text.strip()) == 0:
            return
        
        if TTS_AVAILABLE:
            self.tts.speak(text, wait=wait, force=force)
        else:
            print(f"🤖 {text}")
    
    def listen(self) -> Optional[str]:
        """Get user input"""
        if self.voice_mode and SPEECH_AVAILABLE:
            return self.stt.listen()
        else:
            return input("\n🔹 > ").strip()
    
    def execute_command(self, cmd_type: str, params: Dict) -> str:
        """Execute command"""
        try:
            if cmd_type == 'top_processes':
                count = int(params.get('count', 5))
                count = max(1, min(count, 20))
                sort_by = params.get('sort_by', 'cpu')
                processes = self.proc_manager.get_top_processes(count, sort_by)
                
                if processes and 'error' in processes[0]:
                    return f"Error: {processes[0]['error']}"
                
                self.recent_pids = [proc['pid'] for proc in processes]
                
                result = f"Top {count} processes by {sort_by.upper()}:\n"
                for i, proc in enumerate(processes, 1):
                    result += f"{i}. {proc['name']} (PID {proc['pid']}) - CPU {proc['cpu']:.1f}%, Mem {proc['memory']:.1f}%\n"
                return result
            
            elif cmd_type == 'kill_process':
                pid = params.get('pid')
                if not pid:
                    return "❌ No PID specified"
                
                pid_int = int(pid)
                result = self.proc_manager.kill_process(pid_int)
                
                # Try fuzzy matching if not found
                if "not found" in result.lower() and self.recent_pids:
                    pid_str = str(pid_int)
                    matches = [p for p in self.recent_pids if str(p).startswith(pid_str)]
                    
                    if len(matches) == 1:
                        matched_pid = matches[0]
                        result = self.proc_manager.kill_process(matched_pid)
                        if "terminated" in result.lower():
                            result = f"✓ Matched {pid_int}* → {matched_pid}. {result}"
                    elif len(matches) > 1:
                        result = f"❌ PID {pid_int} ambiguous. Matches: {matches[:5]}"
                
                return result
            
            elif cmd_type == 'kill_by_name':
                name = params.get('name')
                if not name:
                    return "❌ No process name specified"
                return self.proc_manager.kill_by_name(name)
            
            elif cmd_type == 'system_info':
                info = self.proc_manager.get_system_info()
                if 'error' in info:
                    return f"Error: {info['error']}"
                
                return f"""System Status:
CPU: {info['cpu_percent']}% ({info['cpu_count']} cores @ {info['cpu_freq']} MHz)
Memory: {info['memory_percent']}% ({info['memory_used']:.1f}/{info['memory_total']:.1f} GB)
Disk: {info['disk_percent']}% ({info['disk_used']:.1f}/{info['disk_total']:.1f} GB)
Processes: {info['process_count']}"""
            
            elif cmd_type == 'process_info':
                pid = params.get('pid')
                if not pid:
                    return "❌ No PID specified"
                
                info = self.proc_manager.get_process_info(int(pid))
                if 'error' in info:
                    return f"Error: {info['error']}"
                
                return f"""Process: {info['name']} (PID {pid})
Status: {info['status']}
CPU: {info['cpu_percent']}%
Memory: {info['memory_mb']:.1f} MB
Threads: {info['num_threads']}"""
            
            elif cmd_type == 'list_files':
                path = params.get('path', '.')
                return self.proc_manager.list_files(path)
            
            else:
                return f"❌ Unknown command: {cmd_type}"
        
        except Exception as e:
            return f"❌ Execution error: {str(e)}"
    
    def get_help(self) -> str:
        """Display help"""
        return """╔════════════════════════════════════════╗
║        AI SHELL ASSISTANT HELP         ║
╚════════════════════════════════════════╝

Voice Commands:
  • "show top processes"
  • "kill process [PID]" or just say the PID
  • "kill [process name]"
  • "system info"
  • "list files in [path]"

Keyboard:
  • 'help' - Show this help
  • 'v' - Toggle voice mode
  • 'exit' - Quit assistant

Tips:
  • Speak clearly and wait for response
  • Say PIDs slowly for accuracy
  • Use admin rights for protected processes
"""
    
    def run_interactive(self):
        """Main loop with improved flow"""
        print("=" * 60)
        print("    AI SHELL ASSISTANT - High Accuracy Voice")
        print("=" * 60)
        print("\n✓ OpenAI Whisper STT (95%+ accuracy)")
        print("✓ Unlimited offline speech (pyttsx3)")
        print("✓ No API rate limits")
        print("✓ Intelligent PID matching")
        print("✓ Natural conversation")
        print("\nType 'help' for commands\n")
        
        self.speak("AI Shell Assistant ready.", force=True)
        
        while self.running:
            try:
                user_input = self.listen()
                
                if not user_input:
                    continue
                
                user_lower = user_input.lower()
                
                # Handle exit/goodbye
                if any(word in user_lower for word in ['exit', 'quit', 'bye', 'goodbye']):
                    self.speak("Goodbye!", force=True)
                    self.running = False
                    break
                
                # Handle special commands
                if user_lower == 'help':
                    print(self.get_help())
                    continue
                
                if user_lower in ['voice', 'v']:
                    self.voice_mode = not self.voice_mode
                    mode = "enabled" if self.voice_mode else "disabled"
                    msg = f"Voice mode {mode}"
                    print(msg)
                    self.speak(msg, force=True)
                    continue
                
                # Process with LLM
                print("\n⏳ Processing...")
                response_text, command, params = self.assistant.chat(user_input)
                
                # Speak and show response
                if response_text:
                    print(f"\n🤖 {response_text}")
                    self.speak(response_text, wait=True, force=True)
                
                # Execute command
                if command:
                    print(f"🔧 Executing: {command}")
                    result = self.execute_command(command, params)
                    print(f"\n{result}\n")
                    
                    # Speak brief result for certain commands
                    if command in ['kill_process', 'kill_by_name']:
                        if "✓" in result and len(result) < 80:
                            self.speak(result.split('\n')[0], wait=True, force=True)
            
            except KeyboardInterrupt:
                print("\n\n🛑 Interrupted")
                self.running = False
                break
            except Exception as e:
                print(f"⚠️ Error: {str(e)}")
                import traceback
                traceback.print_exc()


def main():
    """Entry point"""
    print("\n" + "=" * 60)
    print("  AI SHELL ASSISTANT - OpenAI Whisper STT")
    print("=" * 60)
    print("\nRequired packages:")
    print("  pip install psutil SpeechRecognition pyaudio")
    print("  pip install python-dotenv openai pyttsx3")
    print("\nFeatures:")
    print("  • OpenAI Whisper for 95%+ accurate voice recognition")
    print("  • Unlimited offline TTS (pyttsx3)")
    print("  • GPT-3.5 for natural language processing")
    print("\nSetup: Set OPENAI_API_KEY in environment\n")
    
    if not os.getenv('OPENAI_API_KEY'):
        print("❌ ERROR: OPENAI_API_KEY not set!")
        print("\nSet it with:")
        print("  Windows: setx OPENAI_API_KEY \"your-key-here\"")
        print("  Or create .env file with: OPENAI_API_KEY=your-key-here")
        sys.exit(1)
    
    try:
        assistant = ShellAssistant()
        assistant.run_interactive()
    except Exception as e:
        print(f"❌ Fatal error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()