"""
Voice-Controlled AI Shell Assistant for Windows with OpenAI Integration
Optimized for rate limiting and reduced TTS calls
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

# TTS will use OpenAI's built-in API (no extra installation needed)
TTS_AVAILABLE = True

# Import OpenAI
try:
    from openai import OpenAI
    LLM_AVAILABLE = True
except ImportError:
    LLM_AVAILABLE = False
    TTS_AVAILABLE = False
    print("⚠️ openai not installed. Run: pip install openai")


class TextToSpeech:
    """
    Optimized TTS with rate limiting and caching
    - Implements request throttling
    - Caches common responses
    - Handles rate limit errors gracefully
    """
    
    def __init__(self, client=None):
        self.client = client
        self.temp_dir = Path("./tts_cache")
        self.temp_dir.mkdir(exist_ok=True)
        self.tts_thread = None
        self.last_request_time = 0
        self.min_interval = 0.5  # Minimum 0.5 seconds between TTS requests
        self.cache = {}
    
    def speak(self, text: str, wait: bool = True, force: bool = False):
        """Convert text to speech using OpenAI API with rate limiting"""
        if not text or len(text.strip()) == 0 or not self.client:
            return
        
        # Skip short acknowledgments if not forced
        if not force and len(text) < 10:
            return
        
        def _speak_async():
            try:
                # Rate limiting: wait if needed
                time_since_last = time.time() - self.last_request_time
                if time_since_last < self.min_interval:
                    time.sleep(self.min_interval - time_since_last)
                
                # Truncate text if too long
                text_to_speak = text[:1000] if len(text) > 1000 else text
                
                # Check cache
                cache_key = hash(text_to_speak)
                if cache_key in self.cache:
                    audio_file = self.cache[cache_key]
                    if audio_file.exists():
                        self._play_audio(audio_file)
                        return
                
                print(f"🔊 Speaking...")
                
                # Generate audio using OpenAI TTS
                response = self.client.audio.speech.create(
                    model="tts-1",
                    voice="alloy",
                    input=text_to_speak
                )
                
                # Save audio with timestamp
                audio_file = self.temp_dir / f"speech_{int(time.time() * 1000)}.mp3"
                
                # Write response content to file
                with open(audio_file, 'wb') as f:
                    f.write(response.content)
                
                # Cache it
                self.cache[cache_key] = audio_file
                
                # Play audio
                self._play_audio(audio_file)
                
                # Update last request time
                self.last_request_time = time.time()
                
                # Clean up old files (keep only last 10)
                try:
                    files = sorted(self.temp_dir.glob("speech_*.mp3"))
                    for old_file in files[:-10]:
                        old_file.unlink()
                except:
                    pass
            
            except Exception as e:
                if "429" in str(e) or "rate_limit" in str(e).lower():
                    print(f"⚠️ Rate limit reached. Waiting before retry...")
                    time.sleep(25)  # Wait 25 seconds for rate limit to reset
                else:
                    print(f"⚠️ TTS Error: {e}")
        
        # Run TTS in background thread
        self.tts_thread = threading.Thread(target=_speak_async, daemon=True)
        self.tts_thread.start()
        
        # Optionally wait for completion
        if wait and self.tts_thread:
            self.tts_thread.join(timeout=10)
    
    def _play_audio(self, audio_file: Path):
        """Play audio file"""
        try:
            if sys.platform == "win32":
                os.startfile(str(audio_file))
            elif sys.platform == "darwin":  # macOS
                subprocess.run(["afplay", str(audio_file)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            else:  # Linux
                subprocess.run(["paplay", str(audio_file)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            
            # Estimate duration and wait
            estimated_duration = 2.0
            time.sleep(estimated_duration)
        except Exception as e:
            print(f"⚠️ Audio playback error: {e}")
    
    def stop(self):
        """Stop speaking"""
        pass


class SpeechToText:
    """
    Optimized STT with continuous long-duration listening
    """
    
    def __init__(self):
        self.recognizer = None
        if SPEECH_AVAILABLE:
            self.recognizer = sr.Recognizer()
            self.recognizer.energy_threshold = 3000
            self.recognizer.dynamic_energy_threshold = True
            self.recognizer.pause_threshold = 1.0
            self.recognizer.non_speaking_duration = 0.3
    
    def listen(self, timeout: int = 30) -> Optional[str]:
        """Listen to microphone and convert speech to text"""
        if not self.recognizer:
            return None
        
        try:
            with sr.Microphone() as source:
                print("🎤 Listening... (speak now)")
                
                self.recognizer.adjust_for_ambient_noise(source, duration=0.3)
                
                audio = self.recognizer.listen(
                    source,
                    timeout=timeout,
                    phrase_time_limit=30
                )
            
            print("⏳ Processing audio...")
            
            text = self.recognizer.recognize_google(
                audio,
                language='en-US',
                show_all=False
            )
            
            return text.strip()
        
        except sr.UnknownValueError:
            print("⚠️ Could not understand audio")
            return None
        except sr.RequestError as e:
            print(f"⚠️ Speech Recognition API error: {e}")
            return None
        except Exception as e:
            print(f"⚠️ STT Error: {e}")
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
            return f"Process {proc_name} with ID {pid} has been terminated successfully"
        except psutil.NoSuchProcess:
            return f"Process with ID {pid} was not found on the system"
        except psutil.AccessDenied:
            return f"Access denied. Cannot terminate process {pid}. You may need administrator rights"
        except psutil.TimeoutExpired:
            try:
                proc.kill()
                return f"Process {proc_name} with ID {pid} was force killed"
            except:
                return f"Failed to kill process {pid}"
    
    @staticmethod
    def kill_by_name(name: str, exclude: List[str] = None) -> str:
        """Kill all processes by name, with exclusions"""
        exclude = exclude or []
        killed = []
        try:
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    if name.lower() in proc.info['name'].lower():
                        if not any(ex.lower() in proc.info['name'].lower() for ex in exclude):
                            proc.terminate()
                            killed.append(f"{proc.info['name']} with PID {proc.info['pid']}")
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
            
            if killed:
                killed_text = ", ".join(killed)
                return f"Successfully terminated {len(killed)} process(es): {killed_text}"
            else:
                return f"No processes named {name} were found on the system"
        except Exception as e:
            return f"Error occurred while killing processes: {str(e)}"
    
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
                return f"Path {path} does not exist"
            
            if path.is_file():
                return f"This is a file: {path.name}"
            
            items = list(path.iterdir())
            if not items:
                return f"The directory {path} is empty"
            
            result = f"Contents of {path}:\n"
            for item in sorted(items)[:15]:
                if item.is_dir():
                    result += f"[Folder] {item.name}\n"
                else:
                    size_kb = item.stat().st_size / 1024
                    result += f"[File] {item.name} - {size_kb:.1f} KB\n"
            
            if len(items) > 15:
                result += f"And {len(items) - 15} more items"
            
            return result
        except Exception as e:
            return f"Error: {str(e)}"


class ConversationalAssistant:
    """LLM-based Command Interpretation"""
    
    def __init__(self, api_key: str = None):
        if not LLM_AVAILABLE:
            raise ImportError("OpenAI library not installed")
        
        self.api_key = api_key or os.getenv('OPENAI_API_KEY')
        if not self.api_key:
            raise ValueError("OPENAI_API_KEY not found. Set it in environment or .env file")
        
        self.client = OpenAI(api_key=self.api_key)
        self.model = "gpt-3.5-turbo"
        self.conversation_history = []
        self.system_prompt = self._get_system_prompt()
    
    def _get_system_prompt(self) -> str:
        return """You are a friendly and helpful AI Shell Assistant for Windows. You have the ability to control processes and monitor system resources.

You can perform these actions:
1. top_processes - Show top processes by CPU or memory
   params: count (1-20), sort_by (cpu/memory)
   
2. kill_process - Kill process by PID
   params: pid (integer)
   
3. kill_by_name - Kill processes by name
   params: name (process name), exclude (list of names to exclude)
   
4. system_info - Show system resource usage
   params: none
   
5. process_info - Get detailed info about a process
   params: pid (integer)
   
6. list_files - List files in a directory
   params: path (directory path)

CRITICAL RULES:
- ALWAYS respond with BOTH a conversational message AND a command in <ACTION> tags
- NEVER ask clarifying questions - make reasonable assumptions and execute immediately
- ALWAYS extract numbers (PIDs, counts) from speech even if partially heard
- Extract any number mentioned as a PID for kill commands
- If user says a number alone or with "kill", treat as PID to terminate
- Keep responses SHORT (under 10 words)

Examples:
- User: "top processes" → Response: "Checking now. <ACTION>{"command": "top_processes", "params": {"count": 5, "sort_by": "cpu"}}</ACTION>"
- User: "show memory" → Response: "Checking memory. <ACTION>{"command": "top_processes", "params": {"count": 5, "sort_by": "memory"}}</ACTION>"
- User: "kill 712" → Response: "Terminating process 712. <ACTION>{"command": "kill_process", "params": {"pid": 712}}</ACTION>"
- User: "system info" → Response: "Getting info. <ACTION>{"command": "system_info", "params": {}}</ACTION>"

Always execute commands - no exceptions, no clarifications needed."""
    
    
    def chat(self, user_message: str) -> Tuple[str, Optional[str], Optional[Dict]]:
        """Process user message and generate response"""
        try:
            self.conversation_history.append({
                "role": "user",
                "content": user_message
            })
            
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "system", "content": self.system_prompt}] + self.conversation_history,
                temperature=0.7,
                max_tokens=200  # Reduced from 500 to encourage brevity
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
                except json.JSONDecodeError:
                    pass
            
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
        self.speaking = False  # Track if currently speaking
        self.last_voice_output = ""  # Store what to speak
        
        try:
            self.assistant = ConversationalAssistant()
            print("✓ OpenAI API connected successfully")
        except Exception as e:
            print(f"✗ Error initializing OpenAI: {e}")
            sys.exit(1)
        
        self.tts = TextToSpeech(client=self.assistant.client)
        self.stt = SpeechToText()
        
        if LLM_AVAILABLE and self.assistant.client:
            print("✓ Text-to-Speech available")
        else:
            print("ℹ️ Text-to-Speech not available")
        
        if SPEECH_AVAILABLE:
            print("✓ Speech-to-Text available")
        else:
            print("ℹ️ Speech-to-Text not available (text mode only)")
    
    def speak(self, text: str, wait: bool = True, force: bool = False):
        """Speak text using TTS - blocks until complete by default"""
        if not text or len(text.strip()) == 0:
            return
        
        if TTS_AVAILABLE and (force or len(text) > 15):
            print(f"🔊 Speaking...")
            self.tts.speak(text, wait=wait, force=force)
            time.sleep(2)  # Ensure speech completes before listening
        else:
            print(f"📝 {text}")
            time.sleep(1)
    
    def listen(self) -> Optional[str]:
        """Listen and get user input"""
        if self.voice_mode and SPEECH_AVAILABLE:
            return self.stt.listen()
        else:
            return input("\n🔹 > ").strip()
    
    def execute_command(self, cmd_type: str, params: Dict) -> str:
        """Execute the interpreted command"""
        try:
            if cmd_type == 'top_processes':
                count = int(params.get('count', 5))
                count = max(1, min(count, 20))
                sort_by = params.get('sort_by', 'cpu')
                processes = self.proc_manager.get_top_processes(count, sort_by)
                
                if processes and 'error' in processes[0]:
                    return f"Error: {processes[0]['error']}"
                
                result = f"Top {count} processes by {sort_by.upper()}:\n"
                for i, proc in enumerate(processes, 1):
                    result += f"{i}. {proc['name']} - PID {proc['pid']}: CPU {proc['cpu']:.1f}%, Memory {proc['memory']:.1f}%\n"
                return result
            
            elif cmd_type == 'kill_process':
                pid = params.get('pid')
                if pid:
                    return self.proc_manager.kill_process(int(pid))
                return "Please specify a PID"
            
            elif cmd_type == 'kill_by_name':
                name = params.get('name')
                exclude = params.get('exclude', [])
                if name:
                    return self.proc_manager.kill_by_name(name, exclude)
                return "Please specify a process name"
            
            elif cmd_type == 'system_info':
                info = self.proc_manager.get_system_info()
                if 'error' in info:
                    return f"Error: {info['error']}"
                
                result = f"""System Information:
CPU Usage: {info['cpu_percent']}%
Memory Usage: {info['memory_percent']}% ({info['memory_used']:.2f}/{info['memory_total']:.2f} GB)
Disk Usage: {info['disk_percent']}%
Active Processes: {info['process_count']}"""
                return result
            
            elif cmd_type == 'process_info':
                pid = params.get('pid')
                if pid:
                    info = self.proc_manager.get_process_info(int(pid))
                    if 'error' in info:
                        return f"Error: {info['error']}"
                    
                    result = f"Process: {info['name']} (PID {pid})\nCPU: {info['cpu_percent']}% | Memory: {info['memory_mb']:.1f} MB"
                    return result
                return "Please specify a PID"
            
            elif cmd_type == 'list_files':
                path = params.get('path', '.')
                return self.proc_manager.list_files(path)
            
            else:
                return f"Unknown command: {cmd_type}"
        
        except Exception as e:
            return f"Error executing command: {str(e)}"
    
    def get_help(self) -> str:
        """Display available commands"""
        return """Available Commands:
[*] Show top processes / Show top processes by memory
[!] Kill process [PID] / Kill [process name]
[*] System info / What's my CPU usage
[*] List files / List C:\\Users
[*] 'v' to toggle voice, 'help' for help, 'exit' to quit"""
    
    def run_interactive(self):
        """Main event loop"""
        print("=" * 70)
        print("AI SHELL ASSISTANT - Optimized for Rate Limits")
        print("=" * 70)
        print("\nOptimizations:")
        print("✓ Rate limit handling (25s wait on 429)")
        print("✓ TTS request throttling (0.5s minimum between calls)")
        print("✓ Response caching")
        print("✓ Brief responses to reduce API costs")
        print("✓ Only TTS for command results (not confirmations)")
        print("=" * 70)
        print("\nType 'help' for commands\n")
        
        self.speak("Hello! AI Shell Assistant ready.", force=True)
        
        while self.running:
            try:
                user_input = self.listen()
                
                if not user_input:
                    if self.voice_mode:
                        self.speak("Sorry, didn't catch that. Try again.", force=True)
                    continue
                
                if user_input.lower() in ['exit', 'quit']:
                    self.speak("Goodbye!", force=True)
                    self.running = False
                    break
                
                elif user_input.lower() == 'help':
                    help_text = self.get_help()
                    print(help_text)
                    continue
                
                elif user_input.lower() in ['voice', 'v']:
                    self.voice_mode = not self.voice_mode
                    mode = "enabled" if self.voice_mode else "disabled"
                    print(f"Voice mode {mode}")
                    continue
                
                print("\n⏳ Processing...")
                response_text, command, params = self.assistant.chat(user_input)
                
                if response_text:
                    print(f"Assistant: {response_text}")
                
                if command:
                    print(f"🔧 Executing: {command}")
                    result = self.execute_command(command, params)
                    print(f"\n{result}\n")
                    
                    # Speak summarized version for process commands
                    if command == 'top_processes' and self.last_voice_output:
                        self.speak(self.last_voice_output, wait=True, force=True)
                        self.last_voice_output = ""
                    # Speak brief results for other commands
                    elif command in ['kill_process', 'kill_by_name', 'system_info']:
                        # Only speak if result is short enough and not an error
                        if "Error" not in result and len(result) < 100:
                            self.speak(result, wait=True, force=True)
                        elif "Error" in result:
                            print(f"⚠️ {result}")
            
            except KeyboardInterrupt:
                print("\n")
                self.running = False
            except Exception as e:
                print(f"Error: {str(e)}")


def main():
    """Main entry point"""
    print("=" * 70)
    print("AI SHELL ASSISTANT (Rate-Limited & Optimized)")
    print("=" * 70)
    print("\nRequired: pip install psutil SpeechRecognition pydub pyaudio python-dotenv openai")
    print("Setup: Set OPENAI_API_KEY environment variable\n")
    
    if not os.getenv('OPENAI_API_KEY'):
        print("ERROR: OPENAI_API_KEY not set!")
        sys.exit(1)
    
    try:
        assistant = ShellAssistant()
        assistant.run_interactive()
    except Exception as e:
        print(f"Fatal error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()