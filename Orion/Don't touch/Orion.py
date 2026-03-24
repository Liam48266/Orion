import faster_whisper
import sounddevice as sd
import numpy as np
import os
import glob
import threading
import tkinter as tk
from tkinter import messagebox, scrolledtext, simpledialog
import ctypes
import sys
import datetime
from playsound import playsound
import keyboard  # ← For global hotkey

# ================== MUTE OCEAN BLUE THEME ==================
BG_MAIN        = "#0a1f3d"
BG_PANEL       = "#132e54"
ACCENT_CYAN    = "#00b4d8"
TITLE_CYAN     = "#00ccff"
WHITE          = "#ffffff"
RED_STOP       = "#c62828"

# ================== PORTABLE FOLDERS ==================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_DIR  = os.path.join(BASE_DIR, "Conversation Logs")
APP_DIR  = os.path.join(BASE_DIR, "App directories")
DIALOGUES_FOLDER = os.path.join(BASE_DIR, "Dialogues")

os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(APP_DIR, exist_ok=True)

current_log_path = None
current_hotkey = 'ctrl+\\'   # Default: rare hotkey (Ctrl + Backslash)

# ================== SMART DIALOGUES FINDER ==================
PATH_FILE = os.path.join(APP_DIR, "dialogues_path.txt")

def find_dialogues_folder():
    if os.path.exists(PATH_FILE):
        try:
            with open(PATH_FILE, "r", encoding="utf-8") as f:
                saved = f.read().strip()
            if os.path.exists(saved) and any(f.lower().endswith('.mp3') for f in os.listdir(saved)):
                return saved
        except:
            pass

    current = BASE_DIR
    for _ in range(8):
        candidate = os.path.join(current, "Dialogues")
        if os.path.exists(candidate) and any(f.lower().endswith('.mp3') for f in os.listdir(candidate)):
            with open(PATH_FILE, "w", encoding="utf-8") as f:
                f.write(candidate)
            return candidate
        parent = os.path.dirname(current)
        if parent == current: break
        current = parent

    messagebox.showerror("Dialogues Missing", 
                         "Please create a folder named 'Dialogues' next to this .exe\n"
                         "and put your .mp3 voice files inside it.")
    sys.exit(1)

DIALOGUES_FOLDER = find_dialogues_folder()

# ================== LOAD WHISPER MODEL ==================
model = faster_whisper.WhisperModel("small.en", device="cpu", compute_type="int8")

# ================== MEDIA KEYS ==================
VK_MEDIA_NEXT_TRACK = 0xB0
VK_MEDIA_PREV_TRACK = 0xB1
VK_MEDIA_PLAY_PAUSE = 0xB3
VK_VOLUME_UP = 0xAF
VK_VOLUME_DOWN = 0xAE

def send_media_key(vk):
    ctypes.windll.user32.keybd_event(vk, 0, 0, 0)
    ctypes.windll.user32.keybd_event(vk, 0, 2, 0)

# ================== DEFAULT APPS & WEBSITES ==================
websites = {
    "youtube": "https://www.youtube.com",
    "google": "https://www.google.com",
    "gmail": "https://mail.google.com",
    "calendar": "https://calendar.google.com",
    "wikipedia": "https://en.wikipedia.org",
    "netflix": "https://www.netflix.com",
    "reddit": "https://www.reddit.com",
    "facebook": "https://www.facebook.com",
    "twitter": "https://twitter.com",
    "x": "https://x.com",
    "news": "https://cnn.com",
    "weather": "https://weather.com",
    "github": "https://github.com",
    "amazon": "https://amazon.com",
    "twitch": "https://www.twitch.tv",
    "spotify": "https://open.spotify.com",
    "roblox": "https://www.roblox.com",
    "fortnite": "https://www.epicgames.com/fortnite",
    "valorant": "https://playvalorant.com",
    "genshin": "https://genshin.hoyoverse.com",
}

apps = {
    "notepad": "notepad",
    "calculator": "calc",
    "paint": "mspaint",
    "chrome": "chrome",
    "edge": "msedge",
    "firefox": "firefox",
    "task manager": "taskmgr",
    "discord": "discord",
    "steam": "steam",
    "spotify": "spotify",
    "vlc": "vlc",
    "obs": "obs64",
    "blender": "blender",
    "vscode": "code",
    "word": "winword",
    "excel": "excel",
    "powerpoint": "powerpnt",
    "teams": "teams",
    "zoom": "zoom",
    "skype": "skype",
    "slack": "slack",
}

# ================== LOAD CUSTOM APPS/WEBSITES FROM .py FILES ==================
def load_custom_apps():
    for filename in os.listdir(APP_DIR):
        if filename.endswith(".py"):
            name = filename[:-3].replace("_", " ").lower()
            path = os.path.join(APP_DIR, filename)
            try:
                with open(path, "r", encoding="utf-8") as f:
                    content = f.read().strip()
                if "path =" in content:
                    value = content.split("path =", 1)[1].strip().strip('"\'')
                    apps[name] = value
                elif "url =" in content:
                    value = content.split("url =", 1)[1].strip().strip('"\'')
                    websites[name] = value
            except:
                pass

load_custom_apps()

# ================== DIALOGUE MAP ==================
dialogue_map = {
    "Orion online and ready. Say 'Orion' followed by your command.": os.path.join(DIALOGUES_FOLDER, "opening dialogue.mp3"),
    "Shutting down. Goodbye, sir.": os.path.join(DIALOGUES_FOLDER, "shutting down dialogue.mp3"),
    "Opening, sir.": os.path.join(DIALOGUES_FOLDER, "open app.mp3"),
    "Apologies, sir. i am currently unable to open that": os.path.join(DIALOGUES_FOLDER, "cant open.mp3"),
    "Command not recognized.": os.path.join(DIALOGUES_FOLDER, "not recognized.mp3"),
    "Playing all MP3 files from your music directory.": os.path.join(DIALOGUES_FOLDER, "play music.mp3"),
    "No MP3 files found in your music directory.": os.path.join(DIALOGUES_FOLDER, "no music found.mp3"),
    "Pausing music.": os.path.join(DIALOGUES_FOLDER, "pause music.mp3"),
    "Skipping to the next song, sir.": os.path.join(DIALOGUES_FOLDER, "skip song.mp3"),
    "Playing previous song, sir.": os.path.join(DIALOGUES_FOLDER, "previous song.mp3"),
    "Increasing volume, sir.": os.path.join(DIALOGUES_FOLDER, "increase volume.mp3"),
    "Reducing volume, sir.": os.path.join(DIALOGUES_FOLDER, "reduce volume.mp3"),
    "Emergency shutdown in 30 seconds, sir.": os.path.join(DIALOGUES_FOLDER, "emergency shuttdown.mp3"),
    "Locking the computer.": os.path.join(DIALOGUES_FOLDER, "lock computer.mp3"),
    "Opening the weather forecast.": os.path.join(DIALOGUES_FOLDER, "weather.mp3"),
    "Shutdown canceled.": os.path.join(DIALOGUES_FOLDER, "cancel shuttdown.mp3"),
}

# ================== GUI GLOBALS ==================
root = None
main_frame = None
logs_frame = None
log_listbox = None
log_viewer = None
running = False
listen_thread = None
hotkey_entry = None

# ================== SPEAK ==================
def speak(text):
    print("Orion:", text)
    if current_log_path:
        try:
            with open(current_log_path, "a", encoding="utf-8") as f:
                f.write(f"Orion: {text}\n")
        except:
            pass
    
    def play_worker():
        try:
            file_path = dialogue_map.get(text.strip())
            if file_path and os.path.exists(file_path):
                playsound(file_path, block=False)
        except:
            pass
    threading.Thread(target=play_worker, daemon=True).start()

# ================== LISTEN ==================
def listen():
    RATE = 16000
    CHANNELS = 1
    RECORD_SECONDS = 6
    print("🎤 Listening...")
    recording = sd.rec(int(RECORD_SECONDS * RATE), samplerate=RATE, channels=CHANNELS, dtype=np.float32)
    sd.wait()
    audio_np = recording.flatten()
    segments, _ = model.transcribe(audio_np, beam_size=5, language="en", vad_filter=True)
    text = " ".join(segment.text for segment in segments).strip().lower()
    if text:
        print("You said:", text)
        return text
    return None

# ================== PROCESS COMMAND ==================
def process_command(command):
    cmd = command.lower()

    if any(phrase in cmd for phrase in ["add new app", "add app"]):
        add_new_app()
        return
    elif any(phrase in cmd for phrase in ["add new website", "add website"]):
        add_new_website()
        return

    if any(phrase in cmd for phrase in ["open conversation logs folder", "open logs folder", "show logs"]):
        speak("Opening conversation logs folder.")
        os.startfile(LOG_DIR)
        return

    elif any(phrase in cmd for phrase in ["open latest conversation log", "open last log", "latest log"]):
        files = [f for f in os.listdir(LOG_DIR) if f.endswith(".txt")]
        if files:
            latest = max(files, key=lambda x: os.path.getmtime(os.path.join(LOG_DIR, x)))
            speak("Opening latest conversation log.")
            os.startfile(os.path.join(LOG_DIR, latest))
        else:
            speak("No conversation logs found yet.")
        return

    if any(phrase in cmd for phrase in ["what time is it", "whats the time", "what's the time", "time", "what day is it", "whats the date", "what's the date"]):
        speak("Opening, sir.")
        os.system('start ms-clock:')

    elif any(phrase in cmd for phrase in ["play music", "play some music"]):
        speak("Playing all MP3 files from your music directory.")
        music_dir = os.path.expanduser("~/Music")
        mp3_files = sorted(glob.glob(os.path.join(music_dir, "*.mp3")))
        if mp3_files:
            playlist_path = os.path.join(music_dir, "orion_playlist.m3u")
            with open(playlist_path, "w", encoding="utf-8") as f:
                f.write("#EXTM3U\n")
                for mp3 in mp3_files:
                    f.write(mp3 + "\n")
            os.startfile(playlist_path)
        else:
            speak("No MP3 files found in your music directory.")

    elif any(phrase in cmd for phrase in ["stop music", "pause music"]):
        speak("Pausing music.")
        send_media_key(VK_MEDIA_PLAY_PAUSE)

    elif any(phrase in cmd for phrase in ["next song", "skip song"]):
        speak("Skipping to the next song, sir.")
        send_media_key(VK_MEDIA_NEXT_TRACK)

    elif any(phrase in cmd for phrase in ["previous song", "go back"]):
        speak("Playing previous song, sir.")
        send_media_key(VK_MEDIA_PREV_TRACK)

    elif "volume up" in cmd or "louder" in cmd:
        speak("Increasing volume, sir.")
        for _ in range(5): send_media_key(VK_VOLUME_UP)

    elif "volume down" in cmd or "quieter" in cmd:
        speak("Reducing volume, sir.")
        for _ in range(5): send_media_key(VK_VOLUME_DOWN)

    elif any(phrase in cmd for phrase in ["emergency shutdown", "emergency computer shutdown"]):
        speak("Emergency shutdown in 30 seconds, sir.")
        os.system("shutdown /s /t 30")

    elif any(phrase in cmd for phrase in ["emergency computer restart", "emergency restart"]):
        speak("Emergency shutdown in 30 seconds, sir.")
        os.system("shutdown /r /t 30")

    elif "cancel shutdown" in cmd:
        os.system("shutdown /a")
        speak("Shutdown canceled.")

    elif any(phrase in cmd for phrase in ["lock computer", "lock screen"]):
        speak("Locking the computer.")
        os.system("rundll32.exe user32.dll,LockWorkStation")

    elif any(phrase in cmd for phrase in ["what's the weather", "weather"]):
        speak("Opening the weather forecast.")
        os.startfile("https://weather.com")

    elif any(verb in cmd for verb in ["open ", "start ", "launch "]):
        for verb in ["open ", "start ", "launch "]:
            if verb in cmd:
                to_open = cmd.split(verb, 1)[1].strip()
                break
        else:
            to_open = ""

        if to_open:
            opened = False
            for key in sorted(apps.keys(), key=len, reverse=True):
                if key in to_open:
                    speak("Opening, sir.")
                    os.startfile(apps[key])
                    opened = True
                    break
            if not opened:
                for key in sorted(websites.keys(), key=len, reverse=True):
                    if key in to_open:
                        speak("Opening, sir.")
                        os.startfile(websites[key])
                        opened = True
                        break
            if not opened:
                speak("Apologies, sir. i am currently unable to open that")

    elif any(phrase in cmd for phrase in ["exit", "stop", "shutdown", "quit"]):
        speak("Shutting down. Goodbye, sir.")
        stop_jarvis()
        return

    else:
        speak("Command not recognized.")

# ================== DYNAMIC ADD FUNCTIONS ==================
def add_new_app():
    name = simpledialog.askstring("Add New App", "Enter command name:")
    if not name: return
    path = simpledialog.askstring("Add New App", "Enter full path to .exe or .lnk:")
    if not path: return
    safe_name = name.lower().replace(" ", "_") + ".py"
    with open(os.path.join(APP_DIR, safe_name), "w", encoding="utf-8") as f:
        f.write(f'path = "{path}"\n')
    messagebox.showinfo("Success", f"'{name}' added!\nRestart Orion to use it.")

def add_new_website():
    name = simpledialog.askstring("Add New Website", "Enter command name:")
    if not name: return
    url = simpledialog.askstring("Add New Website", "Enter full URL:")
    if not url: return
    safe_name = name.lower().replace(" ", "_") + ".py"
    with open(os.path.join(APP_DIR, safe_name), "w", encoding="utf-8") as f:
        f.write(f'url = "{url}"\n')
    messagebox.showinfo("Success", f"'{name}' website added!\nRestart Orion to use it.")

# ================== TOGGLE LISTENING WITH HOTKEY ==================
def toggle_listening():
    global running
    if running:
        stop_jarvis()
    else:
        start_jarvis()

def apply_hotkey():
    global current_hotkey
    new_key = hotkey_entry.get().strip().lower()
    if not new_key:
        return
    try:
        keyboard.remove_hotkey(current_hotkey)
    except:
        pass
    keyboard.add_hotkey(new_key, toggle_listening)
    current_hotkey = new_key
    messagebox.showinfo("Hotkey Updated", f"Toggle hotkey is now: {new_key}\n(works globally even when minimized)")

# ================== LISTEN LOOP ==================
def listen_loop():
    global running
    while running:
        command = listen()
        if command and "orion" in command.lower():
            if current_log_path:
                try:
                    with open(current_log_path, "a", encoding="utf-8") as f:
                        f.write(f"You said: {command}\n")
                except:
                    pass
            process_command(command)

def start_jarvis():
    global running, listen_thread, current_log_path
    if not running:
        running = True
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        current_log_path = os.path.join(LOG_DIR, f"Conversation_{timestamp}.txt")
        with open(current_log_path, "w", encoding="utf-8") as f:
            f.write(f"Orion Conversation Log - {datetime.datetime.now().strftime('%Y-%m-%d %I:%M %p')}\n")
            f.write("="*60 + "\n\n")
        speak("Orion online and ready. Say 'Orion' followed by your command.")
        listen_thread = threading.Thread(target=listen_loop, daemon=True)
        listen_thread.start()

def stop_jarvis():
    global running
    running = False
    speak("Shutting down. Goodbye, sir.")

# ================== SHOW COMMANDS ==================
def show_commands():
    cmd_list = """Orion Commands List:

• What time is it / What day is it → opens clock
• Play music / Pause music / Skip song / Previous song
• Volume up / Volume down
• Emergency shutdown / Emergency computer restart
• Lock computer
• Cancel shutdown
• What's the weather
• Open any app or website
• Add new app
• Add new website
• Open conversation logs folder
• Open latest conversation log

Press your hotkey (default Ctrl + \\) to toggle listening on/off
Say 'Orion' + command!"""
    messagebox.showinfo("All Commands", cmd_list)

def show_instruction_manual():
    manual = """Orion Instruction Manual

1. Click 'Start Orion' or press your hotkey to begin listening
2. Say 'Orion' followed by any command
3. To add your own apps or websites:
   - Click 'Add New App' or 'Add New Website'
   - Follow the prompts
4. All custom apps/websites are saved as .py files in 'App directories'
5. Conversation logs are saved automatically in 'Conversation Logs'
6. Change the toggle hotkey anytime in the main screen

Enjoy your personal voice assistant!"""
    messagebox.showinfo("Instruction Manual", manual)

# ================== SWITCH SCREENS ==================
def show_main():
    logs_frame.pack_forget()
    main_frame.pack(fill="both", expand=True)

def show_logs():
    main_frame.pack_forget()
    logs_frame.pack(fill="both", expand=True)
    load_log_list()

def load_log_list():
    log_listbox.delete(0, tk.END)
    files = sorted([f for f in os.listdir(LOG_DIR) if f.endswith(".txt")],
                   key=lambda x: os.path.getmtime(os.path.join(LOG_DIR, x)), reverse=True)
    for f in files:
        log_listbox.insert(tk.END, f)

def show_selected_log(event=None):
    selection = log_listbox.curselection()
    if not selection:
        return
    filename = log_listbox.get(selection[0])
    path = os.path.join(LOG_DIR, filename)
    try:
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        log_viewer.delete("1.0", tk.END)
        log_viewer.insert(tk.END, content)
    except:
        pass

# ================== BUILD GUI ==================
def create_gui():
    global root, main_frame, logs_frame, log_listbox, log_viewer, hotkey_entry
    root = tk.Tk()
    root.title("Orion")
    root.geometry("920x680")
    root.configure(bg=BG_MAIN)

    title_bar = tk.Frame(root, bg=BG_MAIN, height=80)
    title_bar.pack(fill="x")
    tk.Label(title_bar, text="ORION", font=("Arial", 28, "bold"), fg=TITLE_CYAN, bg=BG_MAIN).pack(pady=15)

    nav_frame = tk.Frame(root, bg=BG_MAIN)
    nav_frame.pack(pady=8)
    tk.Button(nav_frame, text="Main", font=("Arial", 12, "bold"), bg=ACCENT_CYAN, fg=WHITE, width=12, command=show_main).pack(side="left", padx=5)
    tk.Button(nav_frame, text="Conversation Logs", font=("Arial", 12, "bold"), bg=ACCENT_CYAN, fg=WHITE, width=18, command=show_logs).pack(side="left", padx=5)
    tk.Button(nav_frame, text="Instruction Manual", font=("Arial", 12, "bold"), bg=ACCENT_CYAN, fg=WHITE, width=18, command=show_instruction_manual).pack(side="left", padx=5)

    main_frame = tk.Frame(root, bg=BG_MAIN)
    main_frame.pack(fill="both", expand=True)

    btn_frame = tk.Frame(main_frame, bg=BG_MAIN)
    btn_frame.pack(pady=30)

    tk.Button(btn_frame, text="▶ Start Orion", font=("Arial", 13, "bold"), bg=ACCENT_CYAN, fg=WHITE, width=16, height=2, command=start_jarvis).pack(side="left", padx=12)
    tk.Button(btn_frame, text="⛔ Stop Orion", font=("Arial", 13, "bold"), bg=RED_STOP, fg=WHITE, width=16, height=2, command=stop_jarvis).pack(side="left", padx=12)
    tk.Button(btn_frame, text="📋 Commands List", font=("Arial", 13, "bold"), bg="#0277bd", fg=WHITE, width=16, height=2, command=show_commands).pack(side="left", padx=12)

    add_frame = tk.Frame(main_frame, bg=BG_MAIN)
    add_frame.pack(pady=10)
    tk.Button(add_frame, text="Add New App", font=("Arial", 11, "bold"), bg=ACCENT_CYAN, fg=WHITE, width=18, command=add_new_app).pack(side="left", padx=8)
    tk.Button(add_frame, text="Add New Website", font=("Arial", 11, "bold"), bg=ACCENT_CYAN, fg=WHITE, width=18, command=add_new_website).pack(side="left", padx=8)

    # ================== HOTKEY CONFIG ==================
    hotkey_frame = tk.Frame(main_frame, bg=BG_MAIN)
    hotkey_frame.pack(pady=20)
    tk.Label(hotkey_frame, text="Toggle Listening Hotkey (global):", fg=WHITE, bg=BG_MAIN, font=("Arial", 11, "bold")).pack()
    hotkey_entry = tk.Entry(hotkey_frame, font=("Arial", 12), width=25, justify="center")
    hotkey_entry.insert(0, current_hotkey)
    hotkey_entry.pack(pady=5)
    tk.Button(hotkey_frame, text="Apply Hotkey", font=("Arial", 10, "bold"), bg=ACCENT_CYAN, fg=WHITE, command=apply_hotkey).pack()

    tk.Label(main_frame, text="Status: Ready (click Start or press hotkey)", fg=ACCENT_CYAN, bg=BG_MAIN, font=("Arial", 11)).pack(pady=10)

    logs_frame = tk.Frame(root, bg=BG_MAIN)
    tk.Label(logs_frame, text="Saved Conversations", font=("Arial", 12, "bold"), fg=WHITE, bg=BG_MAIN).pack(anchor="w", padx=20, pady=10)
    
    global log_listbox
    log_listbox = tk.Listbox(logs_frame, height=12, bg=BG_PANEL, fg=WHITE, font=("Consolas", 10))
    log_listbox.pack(fill="both", expand=True, padx=20, pady=5)
    log_listbox.bind("<<ListboxSelect>>", show_selected_log)

    tk.Label(logs_frame, text="Log Content", font=("Arial", 12, "bold"), fg=WHITE, bg=BG_MAIN).pack(anchor="w", padx=20, pady=10)
    global log_viewer
    log_viewer = scrolledtext.ScrolledText(logs_frame, bg=BG_PANEL, fg=WHITE, font=("Consolas", 10))
    log_viewer.pack(fill="both", expand=True, padx=20, pady=5)

    btn_logs = tk.Frame(logs_frame, bg=BG_MAIN)
    btn_logs.pack(pady=10)
    tk.Button(btn_logs, text="Open Logs Folder", bg=ACCENT_CYAN, fg=WHITE, command=lambda: os.startfile(LOG_DIR)).pack(side="left", padx=8)
    tk.Button(btn_logs, text="Refresh List", bg=ACCENT_CYAN, fg=WHITE, command=load_log_list).pack(side="left", padx=8)

    show_main()

    # Register initial hotkey
    keyboard.add_hotkey(current_hotkey, toggle_listening)

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()

def on_close():
    try:
        keyboard.remove_hotkey(current_hotkey)
    except:
        pass
    stop_jarvis()
    if root:
        root.destroy()

# ================== START ==================
if __name__ == "__main__":
    try:
        create_gui()
    except KeyboardInterrupt:
        sys.exit(0)