"""
Voice-Controlled AI Shell Assistant for Windows with OpenAI Integration
With Animated GUI - FIXED VERSION
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
import tkinter as tk
from tkinter import scrolledtext, ttk
from PIL import Image, ImageTk, ImageSequence

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
                time.sleep(0.1)
            
            self.engine = pyttsx3.init()
            self.engine.setProperty('rate', 180)
            self.engine.setProperty('volume', 0.9)
            
            voices = self.engine.getProperty('voices')
            for voice in voices:
                if 'ravi' in voice.name.lower() or 'ravi' in voice.name.lower():
                    self.engine.setProperty('voice', voice.id)
                    break
        except Exception as e:
            print(f"⚠️ TTS engine init error: {e}")
            self.pyttsx3_available = False
    
    def speak(self, text: str, wait: bool = True, force: bool = False):
        """Convert text to speech with proper resource management"""
        if not text or len(text.strip()) == 0 or not self.pyttsx3_available:
            return
        
        if not self.lock.acquire(blocking=False):
            return
        
        try:
            self.is_speaking = True
            text_to_speak = text[:500] if len(text) > 500 else text
            print(f"🔊 Speaking: {text_to_speak[:50]}...")
            self._init_engine()
            
            if not self.engine:
                return
            
            self.engine.say(text_to_speak)
            self.engine.runAndWait()
            
            try:
                self.engine.stop()
            except:
                pass
            
        except Exception as e:
            print(f"⚠️ TTS Error: {e}")
            self.pyttsx3_available = TTS_AVAILABLE
            if self.pyttsx3_available:
                self._init_engine()
        finally:
            self.is_speaking = False
            self.lock.release()
            time.sleep(0.3)
    
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
    """
    
    def __init__(self, client=None):
        self.recognizer = None
        self.client = client
        self.temp_audio_file = Path("temp_audio.wav")
        
        if SPEECH_AVAILABLE:
            self.recognizer = sr.Recognizer()
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
                self.recognizer.adjust_for_ambient_noise(source, duration=0.5)
                audio = self.recognizer.listen(source, timeout=timeout, phrase_time_limit=20)
            
            print("⏳ Processing with Whisper AI...")
            
            if self.client:
                try:
                    wav_data = audio.get_wav_data()
                    with open(self.temp_audio_file, "wb") as f:
                        f.write(wav_data)
                    
                    with open(self.temp_audio_file, "rb") as audio_file:
                        transcript = self.client.audio.transcriptions.create(
                            model="whisper-1",
                            file=audio_file,
                            language="en"
                        )
                    
                    result = transcript.text.strip()
                    
                    try:
                        self.temp_audio_file.unlink()
                    except:
                        pass
                    
                    pids = re.findall(r'\b(\d{2,6})\b', result)
                    if pids:
                        print(f"📝 Whisper: \"{result}\" [PIDs: {', '.join(pids)}]")
                    else:
                        print(f"📝 Whisper: \"{result}\"")
                    
                    return result
                
                except Exception as e:
                    print(f"⚠️ Whisper error: {e}, falling back to Google...")
            
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
    def open_application(app_name: str) -> str:
        """Open an application and return parent/child PID info"""
        # Common application mappings
        app_paths = {
            'camera': 'microsoft.windows.camera:',
            'notepad': 'notepad.exe',
            'calculator': 'calc.exe',
            'paint': 'mspaint.exe',
            'chrome': r'C:\Program Files\Google\Chrome\Application\chrome.exe',
            'brave': r'C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe',
            'firefox': r'C:\Program Files\Mozilla Firefox\firefox.exe',
            'edge': r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
            'spotify': os.path.join(os.getenv('APPDATA', ''), r'Spotify\Spotify.exe'),
            'discord': os.path.join(os.getenv('LOCALAPPDATA', ''), r'Discord\app-1.0.9003\Discord.exe'),
            'vscode': r'C:\Users\{}\AppData\Local\Programs\Microsoft VS Code\Code.exe'.format(os.getenv('USERNAME')),
            'code': r'C:\Users\{}\AppData\Local\Programs\Microsoft VS Code\Code.exe'.format(os.getenv('USERNAME')),
            'word': r'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE',
            'excel': r'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE',
            'powerpoint': r'C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE',
            'teams': os.path.join(os.getenv('LOCALAPPDATA', ''), r'Microsoft\Teams\current\Teams.exe'),
            'outlook': r'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE',
        }
        
        app_name_lower = app_name.lower()
        
        try:
            # Get current process count before opening
            initial_pids = set(psutil.pids())
            
            # Try to find the app path
            app_path = None
            for key, path in app_paths.items():
                if key in app_name_lower:
                    app_path = path
                    break
            
            # If not found in mappings, try to use it as-is
            if not app_path:
                app_path = app_name
            
            # Open the application
            if app_path.startswith('microsoft.windows.'):
                # Windows Store apps
                process = subprocess.Popen(['explorer.exe', app_path], 
                                          shell=True,
                                          stdout=subprocess.PIPE,
                                          stderr=subprocess.PIPE)
                parent_pid = process.pid
            else:
                # Regular desktop applications
                if os.path.exists(app_path):
                    process = subprocess.Popen([app_path],
                                              stdout=subprocess.PIPE,
                                              stderr=subprocess.PIPE)
                    parent_pid = process.pid
                else:
                    # Try using start command for apps in PATH
                    process = subprocess.Popen(['start', '', app_path],
                                              shell=True,
                                              stdout=subprocess.PIPE,
                                              stderr=subprocess.PIPE)
                    parent_pid = process.pid
            
            # Wait a moment for the app to start
            time.sleep(2)
            
            # Get new PIDs after opening
            new_pids = set(psutil.pids()) - initial_pids
            
            # Find the actual parent process
            actual_parent_pid = None
            child_pids = []
            
            # Try to find the main process by name
            for pid in new_pids:
                try:
                    proc = psutil.Process(pid)
                    proc_name = proc.name().lower()
                    
                    # Check if this is the main application process
                    if any(key in proc_name for key in app_paths.keys() if key in app_name_lower):
                        actual_parent_pid = pid
                        
                        # Get children of this process
                        children = proc.children(recursive=True)
                        child_pids = [child.pid for child in children]
                        break
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
            
            # If we found the process
            if actual_parent_pid:
                parent_proc = psutil.Process(actual_parent_pid)
                result = f"✓ Opened {app_name}\n"
                result += f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
                result += f"📌 PARENT PROCESS:\n"
                result += f"   • Name: {parent_proc.name()}\n"
                result += f"   • PID: {actual_parent_pid}\n"
                result += f"   • Status: {parent_proc.status()}\n"
                result += f"   • CPU: {parent_proc.cpu_percent()}%\n"
                result += f"   • Memory: {parent_proc.memory_info().rss / 1024 / 1024:.1f} MB\n"
                
                if child_pids:
                    result += f"\n👶 CHILD PROCESSES ({len(child_pids)}):\n"
                    for idx, child_pid in enumerate(child_pids[:5], 1):  # Show first 5
                        try:
                            child_proc = psutil.Process(child_pid)
                            result += f"   {idx}. {child_proc.name()} - PID {child_pid}\n"
                        except:
                            result += f"   {idx}. PID {child_pid} (terminated)\n"
                    
                    if len(child_pids) > 5:
                        result += f"   ... and {len(child_pids) - 5} more child processes\n"
                else:
                    result += f"\n👶 CHILD PROCESSES: None\n"
                
                result += f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
                return result
            else:
                # Fallback: just show that we tried to open it
                result = f"✓ Attempted to open {app_name}\n"
                result += f"Process started but detailed info not available.\n"
                result += f"The application may be starting in the background."
                return result
        
        except FileNotFoundError:
            return f"❌ Application '{app_name}' not found.\nPlease check if it's installed or provide the full path."
        except Exception as e:
            return f"❌ Error opening '{app_name}': {str(e)}"
    
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

7. open_app - Open an application
   params: app_name (string)
   Examples: "open chrome", "open spotify", "open calculator", "open camera"

CRITICAL PID EXTRACTION RULES:
- Extract ALL digits from commands like: "kill 1234", "terminate 5678", "stop 9012"
- Look for [PID detected: NUMBER] hints in user messages
- If you see a number after "kill/terminate/stop", use that exact PID
- NEVER use fake PIDs - only real numbers from user input

OPEN APP RULES:
- When user says "open X" or "launch X" or "start X", use open_app command
- Extract the app name after "open/launch/start"
- Examples: 
  * "open chrome" → <ACTION>{"command": "open_app", "params": {"app_name": "chrome"}}</ACTION>
  * "launch spotify" → <ACTION>{"command": "open_app", "params": {"app_name": "spotify"}}</ACTION>
  * "start calculator" → <ACTION>{"command": "open_app", "params": {"app_name": "calculator"}}</ACTION>
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
            kill_keywords = ['kill', 'terminate', 'stop', 'end', 'close', 'quit']
            numbers = re.findall(r'\b(\d{2,6})\b', user_message)
            
            has_kill_word = any(word in user_message.lower() for word in kill_keywords)
            is_just_number = len(user_message.split()) <= 2 and numbers
            
            if numbers and (has_kill_word or is_just_number):
                user_message = f"{user_message} [PID detected: {numbers[0]}]"
            
            self.conversation_history.append({
                "role": "user",
                "content": user_message
            })
            
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "system", "content": self.system_prompt}] + self.conversation_history,
                temperature=0.5,
                max_tokens=150
            )
            
            assistant_message = response.choices[0].message.content
            
            self.conversation_history.append({
                "role": "assistant",
                "content": assistant_message
            })
            
            if len(self.conversation_history) > 20:
                self.conversation_history = self.conversation_history[-20:]
            
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


class AssistantGUI:
    """GUI for the AI Assistant"""
    
    def __init__(self, assistant):
        self.assistant = assistant
        self.root = tk.Tk()
        self.root.title("AI Shell Assistant")
        self.root.geometry("900x600")
        
        # ===== COLOR SCHEME =====
        # Change these values to customize the entire UI color scheme
        # Theme: FUTURISTIC BLACK (Active)
        self.colors = {
            'bg_main': '#000000',        # Pure black main background
            'bg_frame': '#0a0a0a',       # Very dark gray panels
            'bg_text': '#000000',        # Black text area background
            'accent': '#00ffff',         # Bright cyan accent (futuristic)
            'success': '#00ff00',        # Bright green (Matrix-style)
            'warning': '#ffff00',        # Bright yellow
            'error': '#ff0000',          # Bright red
            'text': '#00ff00',           # Green text (Matrix/terminal style)
            'text_user': '#00ffff',      # Cyan for user input
            'text_assistant': '#00ff00', # Green for assistant
            'text_command': '#ffff00'    # Yellow for commands
        }
        
        # UNCOMMENT ONE OF THESE PRESET THEMES:
        
        # Theme 2: Cyberpunk Purple
        # self.colors = {
        #     'bg_main': '#1a0933',
        #     'bg_frame': '#2d1b4e',
        #     'bg_text': '#1a0933',
        #     'accent': '#ff00ff',
        #     'success': '#00ff41',
        #     'warning': '#ffff00',
        #     'error': '#ff073a',
        #     'text': '#e0e0e0',
        #     'text_user': '#00ff41',
        #     'text_assistant': '#ff00ff',
        #     'text_command': '#ffff00'
        # }
        
        # Theme 3: Matrix Green
        # self.colors = {
        #     'bg_main': '#0d0d0d',
        #     'bg_frame': '#1a1a1a',
        #     'bg_text': '#000000',
        #     'accent': '#00ff00',
        #     'success': '#00ff00',
        #     'warning': '#ffff00',
        #     'error': '#ff0000',
        #     'text': '#00ff00',
        #     'text_user': '#00ff00',
        #     'text_assistant': '#33ff33',
        #     'text_command': '#66ff66'
        # }
        
        # Theme 4: Ocean Blue
        # self.colors = {
        #     'bg_main': '#001f3f',
        #     'bg_frame': '#003d5c',
        #     'bg_text': '#001529',
        #     'accent': '#39cccc',
        #     'success': '#2ecc40',
        #     'warning': '#ffdc00',
        #     'error': '#ff4136',
        #     'text': '#dddddd',
        #     'text_user': '#7fdbff',
        #     'text_assistant': '#39cccc',
        #     'text_command': '#ffdc00'
        # }
        
        # Theme 5: Sunset Orange
        # self.colors = {
        #     'bg_main': '#2d1b1b',
        #     'bg_frame': '#4a2f2f',
        #     'bg_text': '#1a0f0f',
        #     'accent': '#ff6b35',
        #     'success': '#00d9a3',
        #     'warning': '#ffc857',
        #     'error': '#e63946',
        #     'text': '#f5f5f5',
        #     'text_user': '#00d9a3',
        #     'text_assistant': '#ff6b35',
        #     'text_command': '#ffc857'
        # }
        
        # Theme 6: Midnight Purple
        # self.colors = {
        #     'bg_main': '#1b1b2f',
        #     'bg_frame': '#2d2d44',
        #     'bg_text': '#16162a',
        #     'accent': '#bb86fc',
        #     'success': '#03dac6',
        #     'warning': '#ffb300',
        #     'error': '#cf6679',
        #     'text': '#e0e0e0',
        #     'text_user': '#03dac6',
        #     'text_assistant': '#bb86fc',
        #     'text_command': '#ffb300'
        # }
        
        self.root.configure(bg=self.colors['bg_main'])
        
        # Animation variables
        self.gif_frames = []
        self.current_frame = 0
        self.animation_running = True
        
        # Thread-safe queues for GUI updates
        self.output_queue = queue.Queue()
        self.status_queue = queue.Queue()
        
        self.setup_ui()
        self.load_animation()
        self.process_queues()  # Start processing queue updates
        
    def setup_ui(self):
        """Create the UI layout"""
        # Main container
        main_frame = tk.Frame(self.root, bg=self.colors['bg_main'])
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Left side - Animation (with cyan border for futuristic look)
        left_frame = tk.Frame(main_frame, bg=self.colors['bg_frame'], 
                             relief=tk.SOLID, borderwidth=2, 
                             highlightbackground=self.colors['accent'], 
                             highlightthickness=2)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # Title
        title_label = tk.Label(left_frame, text="🤖 AI ASSISTANT", 
                               font=('Consolas', 16, 'bold'), 
                               bg=self.colors['bg_frame'], fg=self.colors['accent'])
        title_label.pack(pady=10)
        
        # Animation label
        self.animation_label = tk.Label(left_frame, bg=self.colors['bg_frame'])
        self.animation_label.pack(expand=True, pady=20)
        
        # Status label
        self.status_label = tk.Label(left_frame, text="● READY", 
                                     font=('Consolas', 12, 'bold'), 
                                     bg=self.colors['bg_frame'], fg=self.colors['success'])
        self.status_label.pack(pady=10)
        
        # Right side - Output (with cyan border)
        right_frame = tk.Frame(main_frame, bg=self.colors['bg_frame'], 
                              relief=tk.SOLID, borderwidth=2,
                              highlightbackground=self.colors['accent'], 
                              highlightthickness=2)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # Output title
        output_title = tk.Label(right_frame, text="📋 SYSTEM OUTPUT", 
                               font=('Consolas', 14, 'bold'), 
                               bg=self.colors['bg_frame'], fg=self.colors['accent'])
        output_title.pack(pady=10)
        
        # Output text area
        self.output_text = scrolledtext.ScrolledText(
            right_frame,
            wrap=tk.WORD,
            font=('Consolas', 10),
            bg=self.colors['bg_text'],
            fg=self.colors['text'],
            insertbackground=self.colors['accent'],
            relief=tk.FLAT,
            padx=10,
            pady=10,
            selectbackground=self.colors['accent'],
            selectforeground='#000000'
        )
        self.output_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        # Configure tags for colored output
        self.output_text.tag_config('user', foreground=self.colors['text_user'], font=('Consolas', 10, 'bold'))
        self.output_text.tag_config('assistant', foreground=self.colors['text_assistant'], font=('Consolas', 10, 'bold'))
        self.output_text.tag_config('command', foreground=self.colors['text_command'], font=('Consolas', 10, 'italic'))
        self.output_text.tag_config('result', foreground=self.colors['text'])
        self.output_text.tag_config('error', foreground=self.colors['error'])
        
        # Initial message
        self.append_output(">>> AI SHELL ASSISTANT INITIALIZED <<<\n", 'assistant')
        self.append_output("SYSTEM READY FOR VOICE COMMANDS\n\n", 'result')
        
    def load_animation(self):
        """Load the GIF animation"""
        try:
            # Try to load sparkle-erio.gif from current directory
            gif_path = "sparkle-erio.gif"
            if not os.path.exists(gif_path):
                # If not found, create a placeholder
                self.create_placeholder()
                return
            
            gif = Image.open(gif_path)
            
            # Extract all frames
            for frame in ImageSequence.Iterator(gif):
                frame = frame.convert('RGBA')
                frame = frame.resize((300, 300), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(frame)
                self.gif_frames.append(photo)
            
            # Start animation
            self.animate()
            
        except Exception as e:
            print(f"⚠️ Could not load GIF: {e}")
            self.create_placeholder()
    
    def create_placeholder(self):
        """Create a placeholder if GIF is not found"""
        placeholder = tk.Label(
            self.animation_label.master,
            text="🌟\n\nPlace sparkle-erio.gif\nin the same folder\nas this script",
            font=('Arial', 14),
            bg=self.colors['bg_frame'],
            fg=self.colors['accent'],
            justify=tk.CENTER
        )
        self.animation_label.pack_forget()
        placeholder.pack(expand=True, pady=20)
    
    def animate(self):
        """Animate the GIF"""
        if self.animation_running and self.gif_frames:
            self.animation_label.configure(image=self.gif_frames[self.current_frame])
            self.current_frame = (self.current_frame + 1) % len(self.gif_frames)
            self.root.after(50, self.animate)  # ~20 FPS
    
    def append_output(self, text, tag='result'):
        """Thread-safe append text to output area"""
        self.output_queue.put((text, tag))
    
    def update_status(self, status, color='#00ff88'):
        """Thread-safe update status label"""
        self.status_queue.put((status, color))
    
    def process_queues(self):
        """Process queued GUI updates from worker threads"""
        # Process output queue
        try:
            while not self.output_queue.empty():
                text, tag = self.output_queue.get_nowait()
                self.output_text.insert(tk.END, text, tag)
                self.output_text.see(tk.END)
        except queue.Empty:
            pass
        
        # Process status queue
        try:
            while not self.status_queue.empty():
                status, color = self.status_queue.get_nowait()
                self.status_label.config(text=f"● {status}", fg=color)
        except queue.Empty:
            pass
        
        # Schedule next queue processing
        if self.animation_running:
            self.root.after(100, self.process_queues)  # Check every 100ms
    
    def run(self):
        """Start the GUI"""
        self.root.mainloop()
    
    def destroy(self):
        """Clean up and close"""
        self.animation_running = False
        self.root.destroy()


class ShellAssistant:
    """Main AI Shell Assistant with GUI"""
    
    def __init__(self):
        self.proc_manager = ProcessManager()
        self.assistant = None
        self.running = True
        self.voice_mode = True
        self.recent_pids = []
        self.gui = None
        
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
            
            elif cmd_type == 'open_app':
                app_name = params.get('app_name')
                if not app_name:
                    return "❌ No application name specified"
                return self.proc_manager.open_application(app_name)
            
            else:
                return f"❌ Unknown command: {cmd_type}"
        
        except Exception as e:
            return f"❌ Execution error: {str(e)}"
    
    def get_help(self) -> str:
        """Display help"""
        return """Voice Commands:
  • "show top processes"
  • "kill process [PID]" or just say the PID
  • "kill [process name]"
  • "open [app name]" - Launch applications
  • "system info"
  • "list files in [path]"

Supported Apps to Open:
  • Browsers: chrome, brave, firefox, edge
  • Apps: spotify, discord, camera, calculator
  • Office: word, excel, powerpoint, outlook
  • Dev: vscode, code, teams
  • Basic: notepad, paint

Keyboard:
  • 'help' - Show this help
  • 'v' - Toggle voice mode
  • 'exit' - Quit assistant

Tips:
  • Speak clearly and wait for response
  • Say PIDs slowly for accuracy
  • Use admin rights for protected processes
  • Opening apps will show parent PID and child PIDs
"""
    
    def run_interactive(self):
        """Main loop with GUI"""
        # Initialize GUI
        self.gui = AssistantGUI(self)
        
        # Show welcome message
        self.gui.append_output("=" * 50 + "\n", 'result')
        self.gui.append_output("🎙️ AI SHELL ASSISTANT - Voice Enabled\n", 'assistant')
        self.gui.append_output("=" * 50 + "\n\n", 'result')
        self.gui.append_output("✓ OpenAI Whisper STT (95%+ accuracy)\n", 'result')
        self.gui.append_output("✓ Unlimited offline speech (pyttsx3)\n", 'result')
        self.gui.append_output("✓ No API rate limits\n", 'result')
        self.gui.append_output("✓ Intelligent PID matching\n\n", 'result')
        
        # Start listening in separate thread
        listen_thread = threading.Thread(target=self._listen_loop, daemon=True)
        listen_thread.start()
        
        self.speak("AI Shell Assistant ready.", force=True)
        
        # Run GUI main loop
        self.gui.run()
    
    def _listen_loop(self):
        """Continuous listening loop"""
        while self.running:
            try:
                self.gui.update_status("🎤 Listening...", '#00d9ff')
                user_input = self.listen()
                
                if not user_input:
                    continue
                
                # Display transcription
                self.gui.append_output(f"\n👤 YOU: ", 'user')
                self.gui.append_output(f"{user_input}\n", 'result')
                
                user_lower = user_input.lower()
                
                # Handle exit/goodbye
                if any(word in user_lower for word in ['exit', 'quit', 'bye', 'goodbye']):
                    self.gui.append_output("\n🤖 ASSISTANT: ", 'assistant')
                    self.gui.append_output("Goodbye!\n", 'result')
                    self.speak("Goodbye!", force=True)
                    self.running = False
                    time.sleep(1)
                    self.gui.destroy()
                    break
                
                # Handle help
                if user_lower == 'help':
                    help_text = self.get_help()
                    self.gui.append_output(f"\n{help_text}\n", 'result')
                    continue
                
                # Handle voice toggle
                if user_lower in ['voice', 'v']:
                    self.voice_mode = not self.voice_mode
                    mode = "enabled" if self.voice_mode else "disabled"
                    msg = f"Voice mode {mode}"
                    self.gui.append_output(f"\n🤖 ASSISTANT: ", 'assistant')
                    self.gui.append_output(f"{msg}\n", 'result')
                    self.speak(msg, force=True)
                    continue
                
                # Process with LLM
                self.gui.update_status("⏳ Processing...", '#ffd700')
                response_text, command, params = self.assistant.chat(user_input)
                
                # Display assistant response
                if response_text:
                    self.gui.append_output(f"\n🤖 ASSISTANT: ", 'assistant')
                    self.gui.append_output(f"{response_text}\n", 'result')
                    self.speak(response_text, wait=True, force=True)
                
                # Execute command
                if command:
                    self.gui.append_output(f"\n⚙️  EXECUTING: ", 'command')
                    self.gui.append_output(f"{command}\n", 'command')
                    self.gui.update_status("⚙️ Executing...", '#ff6b6b')
                    
                    result = self.execute_command(command, params)
                    
                    self.gui.append_output(f"\n📊 RESULT:\n", 'command')
                    
                    # Color-code results
                    if "✓" in result:
                        self.gui.append_output(f"{result}\n", 'result')
                    elif "❌" in result:
                        self.gui.append_output(f"{result}\n", 'error')
                    else:
                        self.gui.append_output(f"{result}\n", 'result')
                    
                    # Speak brief result for certain commands
                    if command in ['kill_process', 'kill_by_name']:
                        if "✓" in result and len(result) < 80:
                            self.speak(result.split('\n')[0], wait=True, force=True)
                
                self.gui.update_status("✓ Ready", '#00ff88')
            
            except Exception as e:
                error_msg = f"⚠️ Error: {str(e)}"
                self.gui.append_output(f"\n{error_msg}\n", 'error')
                print(error_msg)
                import traceback
                traceback.print_exc()
                self.gui.update_status("⚠️ Error", '#ff6b6b')
                time.sleep(2)
                self.gui.update_status("✓ Ready", '#00ff88')


def main():
    """Entry point"""
    print("\n" + "=" * 60)
    print("  AI SHELL ASSISTANT - GUI with Whisper STT")
    print("=" * 60)
    print("\nRequired packages:")
    print("  pip install psutil SpeechRecognition pyaudio")
    print("  pip install python-dotenv openai pyttsx3 pillow")
    print("\nFeatures:")
    print("  • Animated GUI with real-time transcription")
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