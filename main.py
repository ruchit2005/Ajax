"""
Voice-Controlled AI Shell Assistant for Windows with OpenAI Integration
Full TTS/STT with conversational AI assistant capabilities
Optimized for Windows OS with process control, I/O management, and system calls
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

# Try to import TTS
try:
    from pyttsx3 import init as tts_init
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
    OS Concept: Audio I/O and Device Management
    - Uses pyttsx3 for text-to-speech synthesis
    - Cross-platform audio output
    - Real-time speech generation
    """
    
    def __init__(self):
        self.engine = None
        if TTS_AVAILABLE:
            try:
                self.engine = tts_init()
                self.engine.setProperty('rate', 150)  # Speed
                self.engine.setProperty('volume', 0.9)  # Volume
            except Exception as e:
                print(f"⚠️ TTS initialization warning: {e}")
    
    def speak(self, text: str, wait: bool = True):
        """Convert text to speech"""
        if not self.engine:
            return
        
        try:
            self.engine.say(text)
            if wait:
                self.engine.runAndWait()
            else:
                # Non-blocking speech
                threading.Thread(target=self.engine.runAndWait, daemon=True).start()
        except Exception as e:
            print(f"⚠️ TTS Error: {e}")
    
    def stop(self):
        """Stop speaking"""
        if self.engine:
            try:
                self.engine.stop()
            except:
                pass


class SpeechToText:
    """
    OS Concept: Audio I/O and Real-time Processing
    - Uses Google Speech Recognition API
    - Microphone input stream management
    - Real-time audio processing
    """
    
    def __init__(self):
        self.recognizer = None
        if SPEECH_AVAILABLE:
            self.recognizer = sr.Recognizer()
            self.recognizer.energy_threshold = 4000
    
    def listen(self, timeout: int = 10) -> Optional[str]:
        """Listen to microphone and convert speech to text"""
        if not self.recognizer:
            return None
        
        try:
            with sr.Microphone() as source:
                print("🎤 Listening... (speak now)")
                self.recognizer.adjust_for_ambient_noise(source, duration=0.5)
                audio = self.recognizer.listen(source, timeout=timeout, phrase_time_limit=15)
            
            print("⏳ Processing audio...")
            text = self.recognizer.recognize_google(audio)
            return text
        except sr.UnknownValueError:
            return None
        except sr.RequestError:
            return None
        except Exception as e:
            print(f"⚠️ STT Error: {e}")
            return None


class ProcessManager:
    """
    OS Concept: Process Management
    - Uses psutil to interact with process table
    - Manages process lifecycle (creation, termination, monitoring)
    - Demonstrates IPC and process control
    """
    
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
            
            # Sort by specified metric
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
            proc.terminate()  # Graceful termination
            proc.wait(timeout=3)
            return f"Process {proc_name} with ID {pid} has been terminated successfully"
        except psutil.NoSuchProcess:
            return f"Process with ID {pid} was not found on the system"
        except psutil.AccessDenied:
            return f"Access denied. Cannot terminate process {pid}. You may need administrator rights"
        except psutil.TimeoutExpired:
            try:
                proc.kill()  # Forced termination
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
                        # Check exclusions
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
                'cpu_percent': psutil.cpu_percent(interval=1),
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
            for item in sorted(items)[:15]:  # Limit to 15 items
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
    """
    OS Concept: LLM-based Command Interpretation and Conversation
    - Uses OpenAI to understand natural language
    - Maintains conversation history
    - Provides conversational responses
    """
    
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

When a user asks you to perform an action:
1. First, respond conversationally to acknowledge their request
2. If you need to execute a command, respond ONLY with valid JSON in this format at the end of your message:

<ACTION>
{
    "command": "command_name",
    "params": {"key": "value"}
}
</ACTION>

Examples:
- User: "Show me the top 5 processes"
- Response: "Sure! Let me check the top 5 processes consuming the most CPU for you." + <ACTION>{"command": "top_processes", "params": {"count": 5, "sort_by": "cpu"}}</ACTION>

- User: "What's my system status?"
- Response: "I'll get your current system information." + <ACTION>{"command": "system_info", "params": {}}</ACTION>

Be conversational, friendly, and helpful. Always explain what you're doing in natural language."""
    
    def chat(self, user_message: str) -> Tuple[str, Optional[str], Optional[Dict]]:
        """
        Process user message and generate response
        Returns: (response_text, command, params)
        """
        try:
            # Add user message to history
            self.conversation_history.append({
                "role": "user",
                "content": user_message
            })
            
            # Get response from OpenAI
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "system", "content": self.system_prompt}] + self.conversation_history,
                temperature=0.7,
                max_tokens=500
            )
            
            assistant_message = response.choices[0].message.content
            
            # Add assistant response to history
            self.conversation_history.append({
                "role": "assistant",
                "content": assistant_message
            })
            
            # Keep conversation history manageable
            if len(self.conversation_history) > 20:
                self.conversation_history = self.conversation_history[-20:]
            
            # Extract action if present
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
            return f"I encountered an error: {str(e)}", None, None


class ShellAssistant:
    """
    Main AI Shell Assistant with TTS/STT and Conversational AI
    OS Concepts Used:
    - Process Management (psutil)
    - I/O Streams (stdin/stdout/stderr, audio I/O)
    - System Calls (subprocess, psutil)
    - Signal Handling (process termination)
    - Resource Management (file handles, memory)
    - Threading for async operations
    - Audio device management
    - LLM-based command interpretation
    """
    
    def __init__(self):
        self.proc_manager = ProcessManager()
        self.tts = TextToSpeech()
        self.stt = SpeechToText()
        self.assistant = None
        self.running = True
        self.voice_mode = False
        
        # Initialize conversational assistant
        try:
            self.assistant = ConversationalAssistant()
            print("✓ OpenAI API connected successfully")
        except Exception as e:
            print(f"✗ Error initializing OpenAI: {e}")
            sys.exit(1)
        
        # Check TTS availability
        if TTS_AVAILABLE:
            print("✓ Text-to-Speech available")
        else:
            print("ℹ️ Text-to-Speech not available")
        
        # Check STT availability
        if SPEECH_AVAILABLE:
            print("✓ Speech-to-Text available")
        else:
            print("ℹ️ Speech-to-Text not available (text mode only)")
    
    def speak(self, text: str):
        """Speak text using TTS"""
        if TTS_AVAILABLE:
            print(f"🔊 {text}")
            self.tts.speak(text)
        else:
            print(f"📝 {text}")
    
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
CPU Cores: {info['cpu_count']}
CPU Frequency: {info['cpu_freq']} MHz
Memory Usage: {info['memory_percent']}% ({info['memory_used']:.2f} GB of {info['memory_total']:.2f} GB)
Disk Usage: {info['disk_percent']}% ({info['disk_used']:.2f} GB of {info['disk_total']:.2f} GB)
Active Processes: {info['process_count']}"""
                return result
            
            elif cmd_type == 'process_info':
                pid = params.get('pid')
                if pid:
                    info = self.proc_manager.get_process_info(int(pid))
                    if 'error' in info:
                        return f"Error: {info['error']}"
                    
                    result = f"""Process Information for PID {pid}:
Name: {info['name']}
Status: {info['status']}
CPU Usage: {info['cpu_percent']}%
Memory: {info['memory_mb']:.2f} MB
Threads: {info['num_threads']}"""
                    return result
                return "Please specify a PID"
            
            elif cmd_type == 'list_files':
                path = params.get('path', '.')
                return self.proc_manager.list_files(path)
            
            elif cmd_type == 'help':
                return self.get_help()
            
            else:
                return f"Unknown command: {cmd_type}"
        
        except Exception as e:
            return f"Error executing command: {str(e)}"
    
    def get_help(self) -> str:
        """Display available commands"""
        return r"""Available Commands:
[*] PROCESS MANAGEMENT: Show me top processes, Top 10 by memory, Process info for PID
[!] PROCESS CONTROL: Kill process [PID], Kill all Python tasks, Kill Chrome except localhost
[*] SYSTEM MONITORING: System info, CPU usage, RAM usage, Disk space
[*] FILE OPERATIONS: List files, Show desktop, List C:\Users
[*] VOICE MODE: Press 'v' to enable voice control
[*] SHORTCUTS: 'help' for help, 'voice' to toggle voice, 'exit' to quit"""
    
    def run_interactive(self):
        """
        OS Concept: Main Event Loop & Signal Handling
        - Maintains process state
        - Handles user input (stdin/audio)
        - Manages graceful shutdown
        """
        print("=" * 70)
        print("🤖 AI SHELL ASSISTANT - Conversational Voice-Controlled System")
        print("=" * 70)
        print("\nOS Concepts Demonstrated:")
        print("✓ Process Management (psutil)")
        print("✓ I/O Streams (stdin/stdout, audio I/O)")
        print("✓ System Calls (subprocess, psutil)")
        print("✓ Signal Handling (process termination)")
        print("✓ Resource Management (file handles, memory)")
        print("✓ Threading (async operations)")
        print("✓ Audio Device Management (TTS/STT)")
        print("✓ LLM-based Conversational AI (OpenAI GPT-3.5)")
        print("=" * 70)
        print("\nType 'help' for commands or 'v' for voice mode\n")
        
        self.speak("Hello! I'm your AI Shell Assistant. How can I help you today?")
        
        while self.running:
            try:
                user_input = self.listen()
                
                if not user_input:
                    if self.voice_mode:
                        self.speak("I didn't catch that. Could you please repeat?")
                    continue
                
                # Special commands
                if user_input.lower() in ['exit', 'quit']:
                    self.speak("Goodbye! Thank you for using the AI Shell Assistant.")
                    self.running = False
                    break
                
                elif user_input.lower() == 'help':
                    help_text = self.get_help()
                    print(help_text)
                    self.speak("Displayed help menu. What would you like to do?")
                    continue
                
                elif user_input.lower() in ['voice', 'v']:
                    self.voice_mode = not self.voice_mode
                    if self.voice_mode:
                        self.speak("Voice mode enabled. I'm listening!")
                    else:
                        print("Voice mode disabled. Using text input.")
                    continue
                
                # Process with AI assistant
                print("⏳ Processing your request...")
                response_text, command, params = self.assistant.chat(user_input)
                
                # Speak the response
                self.speak(response_text)
                
                # Execute command if present
                if command:
                    print(f"\n🔧 Executing: {command}")
                    result = self.execute_command(command, params)
                    print(f"\n{result}\n")
                    self.speak(result)
            
            except KeyboardInterrupt:
                print("\n\n")
                self.speak("Interrupted. Exiting...")
                self.running = False
            except Exception as e:
                error_msg = f"An error occurred: {str(e)}"
                print(f"❌ {error_msg}")
                self.speak(error_msg)


def main():
    """Main entry point"""
    print("=" * 70)
    print("AI SHELL ASSISTANT WITH TTS/STT (Windows Optimized + OpenAI)")
    print("=" * 70)
    
    print("\n📦 Required Installations:")
    print("pip install psutil SpeechRecognition pydub pyaudio python-dotenv openai pyttsx3")
    
    print("\n🔑 Setup:")
    print("1. Set OPENAI_API_KEY environment variable or create .env file")
    print("   .env file content: OPENAI_API_KEY=sk-...")
    print("2. Install all requirements above")
    print("3. Run this script")
    
    print("\n🎯 Features:")
    print("✓ Conversational AI Assistant")
    print("✓ Text-to-Speech (TTS)")
    print("✓ Speech-to-Text (STT)")
    print("✓ Process Management & Control")
    print("✓ Real-time System Monitoring")
    print("✓ Voice & Text Input Modes")
    
    print("=" * 70 + "\n")
    
    # Check if OpenAI API key is set
    if not os.getenv('OPENAI_API_KEY'):
        print("❌ ERROR: OPENAI_API_KEY not found!")
        print("\nSet it using one of these methods:")
        print("1. Environment variable: set OPENAI_API_KEY=sk-...")
        print("2. Create .env file with: OPENAI_API_KEY=sk-...")
        sys.exit(1)
    
    try:
        assistant = ShellAssistant()
        assistant.run_interactive()
    except Exception as e:
        print(f"❌ Fatal error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()