"""
Wallpaper AI Slideshow
Author: Anton Rosh
Version: 1.0.0
License: MIT
Copyright (c) 2024 Anton Rosh
"""

# Standard library imports
import os
import sys
import subprocess
import time
import threading
import random
import ctypes
import logging
import json
from datetime import datetime
import shutil
import tempfile

# Third-party imports
import win32api
import win32con
import win32gui
import win32process
import win32event
from win32api import GetLastError
from winerror import ERROR_ALREADY_EXISTS
import psutil
import requests
from cryptography.fernet import Fernet
import pystray
from PIL import Image, ImageDraw, ImageEnhance, ImageFilter

# Tkinter imports
import tkinter as tk
from tkinter import messagebox, Toplevel, IntVar, StringVar
from ttkbootstrap import Style
from ttkbootstrap.widgets import Frame, Label, Button, Entry, Checkbutton, Notebook, OptionMenu
from tkinter.ttk import Progressbar

# Constants - Move these to the top
API_KEY_FILE = "api_key.enc"
ENCRYPTION_KEY_FILE = "encryption_key.key"
DEFAULT_PROMPTS = {
    "Random": "Random",
    "Nature Landscapes": "A high-resolution photo of a serene mountain valley with a clear blue lake surrounded by pine forests and snow-capped peaks.",
    "Space and Galaxy": "A stunning photo of the Milky Way arching across the night sky over a remote desert with no light pollution.",
    "Anime and Manga": "A high-definition digital artwork featuring a vibrant anime-style cityscape with detailed characters and intricate backgrounds.",
    "Cityscapes": "A breathtaking photo of a metropolitan skyline at sunset, with glowing skyscrapers reflected in a nearby river.",
    "Underwater Scenes": "A real underwater photograph of a coral reef teeming with colorful fish, sea turtles, and rays of sunlight piercing the water.",
    "Vintage and Retro": "A nostalgic photo of a 1950s-style diner with classic cars parked outside and neon signs glowing under the twilight sky.",
    "Fantasy Worlds": "A detailed digital illustration of a mythical forest with glowing fireflies, enchanted waterfalls, and ancient ruins.",
    "Floral Designs": "A vivid macro photograph of fresh tulips and roses in full bloom, with soft natural light accentuating their colors.",
    "Animal Portraits": "A professional wildlife photo of a lion lounging in the golden grasslands, its majestic mane blowing gently in the breeze.",
    "Seasonal Themes": "A cozy winter photo of a snow-covered cabin with smoke rising from the chimney, surrounded by frosty evergreen trees.",
    "Inspirational Quotes": "A motivational design featuring an elegant font overlaying a photo of a calm sunrise over a tranquil ocean.",
    "Gaming Scenes": "A high-quality screenshot or promotional art from a popular video game, showcasing a dynamic action sequence or vivid landscapes.",
    "Artistic Illustrations": "A vibrant digital artwork of a bustling fantasy market filled with colorful stalls and imaginative characters.",
    "Monochrome Designs": "A dramatic black-and-white photo of a solitary lighthouse on a rocky coastline under a cloudy sky.",
    "Sports Highlights": "A dynamic photo of a basketball player mid-dunk, with an excited crowd visible in the background.",
    "Technological Themes": "A photo of an ultramodern data center, with glowing blue servers and neatly organized cables creating a futuristic feel.",
    "Minimalist Designs": "A crisp photo of a single pebble resting on a smooth sandy surface, captured with perfect symmetry and soft lighting.",
    "Custom Creations": "A personalized collage of real-life travel photos and memorabilia, neatly arranged against a corkboard background.",
    "Woodland Scenes": "A tranquil photo of a forest pathway covered with autumn leaves, bathed in warm golden sunlight.",
    "Beach Vistas": "A real photo of a tropical beach with crystal-clear waters, white sand, and gently swaying palm trees."
}
WALLPAPERS_DIR = "generated_wallpapers"
METADATA_FILE = os.path.join(WALLPAPERS_DIR, "metadata.json")
TEMP_DIR = os.path.join(tempfile.gettempdir(), "wallpaper_ai_slideshow")

# Initialize logger at module level
logger = logging.getLogger(__name__)

class EncryptionManager:
    def __init__(self):
        self.fernet = self._init_encryption()
    
    def _init_encryption(self):
        if not os.path.exists(ENCRYPTION_KEY_FILE):
            encryption_key = Fernet.generate_key()
            with open(ENCRYPTION_KEY_FILE, "wb") as key_file:
                key_file.write(encryption_key)
        
        with open(ENCRYPTION_KEY_FILE, "rb") as key_file:
            encryption_key = key_file.read()
        return Fernet(encryption_key)
    
    def encrypt(self, data):
        return self.fernet.encrypt(data)
    
    def decrypt(self, encrypted_data):
        return self.fernet.decrypt(encrypted_data)

# Create global encryption manager
encryption_manager = EncryptionManager()

# Add debug logging
def setup_logging():
    import logging
    log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'app.log')
    logging.basicConfig(
        filename=log_file,
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    # Also log to console when running from exe
    console = logging.StreamHandler()
    console.setLevel(logging.DEBUG)
    logging.getLogger('').addHandler(console)
    return logging.getLogger()

def kill_existing_instances():
    try:
        current_pid = os.getpid()
        current_process = psutil.Process(current_pid)
        
        # Get the name of our executable
        if getattr(sys, 'frozen', False):
            our_name = os.path.basename(sys.executable).lower()
        else:
            our_name = "wallpaper_ai_slideshow"
        
        # Find and terminate other instances
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                # Skip our own process and parent processes
                if proc.info['pid'] == current_pid or proc.info['pid'] in [p.pid for p in current_process.parents()]:
                    continue
                
                if our_name in proc.info['name'].lower():
                    proc.terminate()
                    proc.wait(timeout=3)
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.TimeoutExpired):
                continue
    except Exception as e:
        logger.error(f"Error killing existing instances: {e}")

def create_mutex():
    try:
        mutex_name = "Global\\WallpaperAISlideshow_Mutex"  # Use Global namespace
        mutex = win32event.CreateMutex(None, False, mutex_name)  # Don't initially own
        last_error = GetLastError()
        
        if (last_error == ERROR_ALREADY_EXISTS):
            if mutex:
                win32api.CloseHandle(mutex)
            logger.warning("Application already running")
            messagebox.showwarning("Warning", "Application is already running!")
            sys.exit(0)
        
        return mutex
    except Exception as e:
        logger.error(f"Failed to create mutex: {e}")
        messagebox.showerror("Error", f"Failed to create mutex: {str(e)}")
        sys.exit(1)

# Function to save API key
def save_api_key(api_key):
    encrypted_key = encryption_manager.encrypt(api_key.encode())
    with open(API_KEY_FILE, "wb") as key_file:
        key_file.write(encrypted_key)
    return True

# Function to load API key
def load_api_key():
    if not os.path.exists(API_KEY_FILE):
        return None
    with open(API_KEY_FILE, "rb") as key_file:
        encrypted_key = key_file.read()
    return encryption_manager.decrypt(encrypted_key).decode()

# Modify get_temp_path function (new)
def get_temp_path(filename):
    """Get path for temporary file and ensure temp directory exists"""
    os.makedirs(TEMP_DIR, exist_ok=True)
    return os.path.join(TEMP_DIR, filename)

# Replace upscale_to_4k function
def upscale_to_4k(image_path, save_path=None):
    """Upscale image to 4K (3840x2160) with improved quality and crop to 16:9"""
    target_width = 3840
    target_height = 2160

    with Image.open(image_path) as img:
        # If image is square (from DALL-E), crop it to 16:9
        if img.width == img.height:
            crop_height = img.width * 9 // 16
            top_margin = (img.height - crop_height) // 2
            img = img.crop((0, top_margin, img.width, top_margin + crop_height))
            logger.debug(f"Cropped square image to 16:9 ratio")
        
        # High quality resize to 4K using Lanczos
        img_resized = img.resize((target_width, target_height), Image.Resampling.LANCZOS)
        
        # Apply image enhancements
        try:
            # Apply subtle sharpening
            sharpener = ImageEnhance.Sharpness(img_resized)
            img_enhanced = sharpener.enhance(1.3)  # Slight sharpness boost
            
            # Apply noise reduction
            img_enhanced = img_enhanced.filter(ImageFilter.UnsharpMask(radius=2, percent=150, threshold=3))
            
            # Enhance color
            color_enhancer = ImageEnhance.Color(img_enhanced)
            img_enhanced = color_enhancer.enhance(1.1)  # Subtle color boost
            
            # Adjust contrast
            contrast = ImageEnhance.Contrast(img_enhanced)
            img_enhanced = contrast.enhance(1.1)  # Subtle contrast boost
            
            # Adjust brightness
            brightness = ImageEnhance.Brightness(img_enhanced)
            img_enhanced = brightness.enhance(1.05)  # Very subtle brightness boost
            
            logger.debug("Applied image enhancements successfully")
            
        except Exception as e:
            logger.warning(f"Image enhancement failed, using original resized image: {e}")
            img_enhanced = img_resized
        
        # Save with high quality
        if save_path:
            img_enhanced.save(save_path, "JPEG", quality=95, optimize=True)
            return save_path
        else:
            upscale_path = get_temp_path("upscaled_wallpaper.jpg")
            img_enhanced.save(upscale_path, "JPEG", quality=95, optimize=True)
            return upscale_path

# Function to resize and set wallpaper
def resize_and_set_wallpaper(image_path):
    upscaled_path = upscale_to_4k(image_path)
    abs_path = os.path.abspath(upscaled_path)
    ctypes.windll.user32.SystemParametersInfoW(20, 0, abs_path, 0)
    return "Wallpaper has been updated successfully!"

def ensure_wallpapers_dir():
    """Ensure wallpapers directory exists and load metadata"""
    os.makedirs(WALLPAPERS_DIR, exist_ok=True)
    if not os.path.exists(METADATA_FILE):
        with open(METADATA_FILE, 'w') as f:
            json.dump({}, f)

# Modify save_generated_image function to save only final version
def save_generated_image(image_path, prompt):
    """Save generated image with metadata"""
    ensure_wallpapers_dir()
    
    # Create unique filename based on timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"wallpaper_{timestamp}.jpg"
    filepath = os.path.join(WALLPAPERS_DIR, filename)
    
    # Process image to 4K and save directly to library
    upscale_to_4k(image_path, filepath)
    
    # Update metadata
    with open(METADATA_FILE, 'r') as f:
        metadata = json.load(f)
    
    metadata[filename] = {
        'prompt': prompt,
        'date': timestamp,
        'path': filepath
    }
    
    with open(METADATA_FILE, 'w') as f:
        json.dump(metadata, f, indent=2)
    
    return filepath

class LoadingDialog:
    def __init__(self, parent):
        self.top = Toplevel(parent)
        self.parent = parent
        self.top.title("Generating Wallpaper")
        
        # Make it modal and set size/position
        self.top.transient(parent)
        self.top.grab_set()
        w, h = 400, 200
        x = parent.winfo_x() + (parent.winfo_width() - w) // 2
        y = parent.winfo_y() + (parent.winfo_height() - h) // 2
        self.top.geometry(f"{w}x{h}+{x}+{y}")
        
        self.top.attributes('-topmost', True)
        self.top.protocol("WM_DELETE_WINDOW", lambda: None)
        
        # Main frame
        frame = Frame(self.top, padding=20)
        frame.pack(fill='both', expand=True)
        
        # Step indicators
        self.steps = [
            "🎨 Crafting your custom wallpaper...",
            "🌟 AI is bringing your vision to life...",
            "📥 Downloading the masterpiece...",
            "✨ Perfecting the image quality...",
            "🖼️ Setting as your wallpaper..."
        ]
        
        # Create step labels
        self.step_labels = []
        for step in self.steps:
            label = Label(
                frame,
                text=f"⭐ {step}",
                font=("Arial", 10),
                foreground="gray"
            )
            label.pack(pady=2, anchor='w')
            self.step_labels.append(label)
        
        # Progress bar at bottom
        self.progress = Progressbar(
            frame,
            mode='determinate',
            length=360
        )
        self.progress.pack(fill='x', pady=(15, 0))
        
        self.current_step = -1
        self.progress['maximum'] = len(self.steps)
        self.advance()
        
    def advance(self):
        self.current_step += 1
        self.progress['value'] = self.current_step
        
        for i, label in enumerate(self.step_labels):
            if i < self.current_step:
                label.config(text=f"✅ {self.steps[i]}", foreground="green")
            elif i == self.current_step:
                label.config(text=f"⭐ {self.steps[i]}", foreground="blue")
            else:
                label.config(text=f"⭐ {self.steps[i]}", foreground="gray")
        
        self.top.update()
    
    def destroy(self):
        try:
            if self.top:
                self.top.grab_release()
                self.top.destroy()
            if self.parent:
                self.parent.focus_force()
        except Exception as e:
            logger.error(f"Error destroying loading dialog: {e}")

def generate_wallpaper(prompt, status_label, root, save_to_library=True):
    loading = None
    temp_path = None
    try:
        loading = LoadingDialog(root)
        root.update()
        
        # Step 1: Check API Key
        api_key = load_api_key()
        if not api_key:
            loading.destroy()
            messagebox.showerror("Error", "API key not found. Please enter and save your API key.")
            return
        
        # Step 2: Generate Image
        loading.advance()
        url = "https://api.openai.com/v1/images/generations"
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }
        data = {
            "prompt": prompt,
            "n": 1,
            "size": "1024x1024",
            "response_format": "url",
            "quality": "hd"
        }
        
        response = requests.post(url, headers=headers, json=data)
        if response.status_code != 200:
            raise Exception(f"API Error: {response.json().get('error', {}).get('message', 'Unknown error')}")
        
        # Step 3: Download Image
        loading.advance()
        image_url = response.json()["data"][0]["url"]
        img_data = requests.get(image_url).content
        temp_path = get_temp_path("wallpaper.png")
        with open(temp_path, "wb") as handler:
            handler.write(img_data)
        
        # Step 4: Process Image
        loading.advance()
        if save_to_library:
            saved_path = save_generated_image(temp_path, prompt)
            logger.info(f"Saved generated image to library: {saved_path}")
            # Refresh library list after saving
            refresh_library_list()
        
        # Step 5: Set Wallpaper
        loading.advance()
        status = resize_and_set_wallpaper(temp_path)
        status_label.config(text=status, foreground="green")
        
        # Complete
        time.sleep(0.5)  # Short pause to show completion
        loading.advance()
        time.sleep(0.5)  # Show completed state briefly
        
    except Exception as e:
        logger.exception("Error generating wallpaper")
        messagebox.showerror("Error", f"An error occurred: {e}")
    finally:
        if loading:
            loading.destroy()
        # Clean up temp files
        try:
            shutil.rmtree(TEMP_DIR, ignore_errors=True)
        except Exception as e:
            logger.warning(f"Failed to clean temporary files: {e}")
        root.update()

# Add global variable for system tray icon
_system_tray_icon = None

# Update minimize_to_tray function
def minimize_to_tray(root):
    global _system_tray_icon
    
    def show_app(icon, item):
        icon.stop()
        root.after(0, root.deiconify)

    def quit_app(icon, item):
        icon.stop()
        root.after(0, root.quit)

    try:
        # Create a system tray icon
        image = Image.new("RGB", (64, 64), (255, 255, 255))
        draw = ImageDraw.Draw(image)
        draw.rectangle([16, 16, 48, 48], fill="black")

        menu = pystray.Menu(
            pystray.MenuItem("Show", show_app),
            pystray.MenuItem("Quit", quit_app),
        )
        
        _system_tray_icon = pystray.Icon("Wallpaper AI Slideshow", image, "Wallpaper AI Slideshow", menu)
        root.withdraw()
        
        # Run in separate thread
        threading.Thread(target=_system_tray_icon.run, daemon=True).start()
    except Exception as e:
        logger.error(f"Failed to create system tray icon: {e}")
        messagebox.showerror("Error", "Failed to minimize to system tray")
        root.deiconify()  # Show window again if minimizing fails

# Add cleanup on exit
def cleanup():
    try:
        # Kill any running instances first
        kill_existing_instances()
        
        # Clean up temp directory
        if os.path.exists(TEMP_DIR):
            shutil.rmtree(TEMP_DIR, ignore_errors=True)
            
    except Exception as e:
        logger.error(f"Error during cleanup: {e}")

def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)

# GUI
global_library_listbox = None

# Modify refresh_library_list function to show simpler entries
def refresh_library_list():
    """Global function to refresh the library listbox"""
    if global_library_listbox:
        global_library_listbox.delete(0, tk.END)
        try:
            ensure_wallpapers_dir()
            with open(METADATA_FILE, 'r') as f:
                metadata = json.load(f)
            for filename, info in metadata.items():
                global_library_listbox.insert(tk.END, f"{info['date']} - {info['prompt'][:50]}...")
        except Exception as e:
            logger.error(f"Error refreshing library list: {e}")

# Add these functions before create_gui()
def open_file_location(event):
    """Open the folder containing the selected wallpaper"""
    selection = global_library_listbox.curselection()
    if not selection:
        return
    
    try:
        ensure_wallpapers_dir()
        with open(METADATA_FILE, 'r') as f:
            metadata = json.load(f)
            
        filename = list(metadata.keys())[selection[0]]
        filepath = metadata[filename]['path']
        
        if os.path.exists(filepath):
            # Use explorer to open and select the file
            subprocess.run(['explorer', '/select,', os.path.normpath(filepath)])
        else:
            messagebox.showerror("Error", "Wallpaper file not found")
    except Exception as e:
        logger.error(f"Error opening file location: {e}")
        messagebox.showerror("Error", f"Could not open file location: {e}")

def use_selected_wallpaper():
    """Set the selected wallpaper as current wallpaper"""
    selection = global_library_listbox.curselection()
    if not selection:
        return
    
    try:
        with open(METADATA_FILE, 'r') as f:
            metadata = json.load(f)
        
        filename = list(metadata.keys())[selection[0]]
        filepath = metadata[filename]['path']
        
        if os.path.exists(filepath):
            status = resize_and_set_wallpaper(filepath)
            return status
        else:
            messagebox.showerror("Error", "Wallpaper file not found")
    except Exception as e:
        logger.error(f"Error using selected wallpaper: {e}")
        messagebox.showerror("Error", f"Could not set wallpaper: {e}")

def create_gui():
    try:
        style = Style(theme="flatly")  # Fluent Design Theme
        root = style.master
        
        # Force window to be visible and focused
        root.lift()
        root.attributes('-topmost', True)
        root.after_idle(root.attributes, '-topmost', False)
        
        # Set window properties
        root.title("Wallpaper AI Slideshow")
        root.geometry("600x600")
        root.state('normal')  # Ensure window is not minimized
        
        # Center window on screen
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - 600) // 2
        y = (screen_height - 600) // 2
        root.geometry(f"600x600+{x}+{y}")
        
        # Set window icon
        icon_paths = [
            get_resource_path("app_icon.ico"),  # Try bundled path first
            "app_icon.ico",  # Try current directory
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_icon.ico")  # Try script directory
        ]
        
        icon_set = False
        for icon_path in icon_paths:
            if (os.path.exists(icon_path)):
                try:
                    root.iconbitmap(icon_path)
                    icon_set = True
                    logger.info(f"Successfully set icon from: {icon_path}")
                    break
                except Exception as e:
                    logger.warning(f"Failed to set icon from {icon_path}: {e}")
        
        if not icon_set:
            logger.warning("Could not set app icon from any location")
        
        # Tabs
        notebook = Notebook(root)
        notebook.pack(expand=True, fill="both", padx=10, pady=10)

        # API Key Tab (Add this new tab)
        api_key_tab = Frame(notebook)
        notebook.add(api_key_tab, text="API Key Settings")
        
        Label(api_key_tab, text="OpenAI API Key:").pack(pady=5)
        api_key_var = StringVar()
        api_key_entry = Entry(api_key_tab, textvariable=api_key_var, width=40)
        api_key_entry.pack(pady=5)
        
        # Try to load existing API key
        existing_key = load_api_key()
        if existing_key:
            api_key_var.set(existing_key)
            Label(api_key_tab, text="✓ API key is saved", foreground="green").pack(pady=5)
        
        def save_key():
            key = api_key_var.get().strip()
            if key:
                try:
                    if save_api_key(key):
                        messagebox.showinfo("Success", "API key saved successfully!")
                        # Add success indicator label if not exists
                        for widget in api_key_tab.winfo_children():
                            if isinstance(widget, Label) and widget.cget("text") == "✓ API key is saved":
                                break
                        else:
                            Label(api_key_tab, text="✓ API key is saved", foreground="green").pack(pady=5)
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to save API key: {e}")
            else:
                messagebox.showwarning("Warning", "Please enter an API key")
        
        Button(api_key_tab, text="Save API Key", command=save_key).pack(pady=10)
        
        # Add help text
        help_text = """
        To get your OpenAI API key:
        1. Go to platform.openai.com
        2. Sign in or create an account
        3. Go to API keys section
        4. Create a new API key
        5. Copy and paste it here
        """
        Label(api_key_tab, text=help_text, justify=tk.LEFT, wraplength=500).pack(pady=20)

        # Wallpaper Tab
        wallpaper_tab = Frame(notebook)
        notebook.add(wallpaper_tab, text="Wallpaper Settings")

        Label(wallpaper_tab, text="Custom Prompt (optional):").pack(pady=5)
        custom_prompt = StringVar()
        prompt_entry = Entry(wallpaper_tab, textvariable=custom_prompt, width=40, state="disabled")
        prompt_entry.pack(pady=5)

        use_custom_prompt = IntVar(value=0)

        def toggle_custom_prompt():
            if use_custom_prompt.get():
                prompt_entry.config(state="normal")
            else:
                prompt_entry.config(state="disabled")

        Checkbutton(
            wallpaper_tab, text="Use Custom Prompt", variable=use_custom_prompt, command=toggle_custom_prompt
        ).pack(pady=5)

        # Default Prompts Dropdown
        Label(wallpaper_tab, text="Select Default Prompt:").pack(pady=5)
        selected_prompt = StringVar(value="Random")
        OptionMenu(wallpaper_tab, selected_prompt, *DEFAULT_PROMPTS.keys()).pack(pady=5)

        # Interval Dropdown
        Label(wallpaper_tab, text="Auto-Change Interval:").pack(pady=5)
        interval_var = StringVar(value="Never")
        intervals = ["Never", "1 hour", "2 hours", "3 hours", "6 hours", 
                    "12 hours", "24 hours", "5 minutes", "15 minutes", 
                    "30 minutes", "45 minutes", "60 minutes"]
        OptionMenu(wallpaper_tab, interval_var, *intervals).pack(pady=5)

        # Add estimated cost label
        def update_cost_estimate(*args):
            interval = interval_var.get()
            if interval == "Never":
                cost_text = "Cost: $0/month (Manual only)"
            else:
                # Parse interval to hours
                if "hour" in interval:
                    hours = float(interval.split()[0])
                    images_per_month = (30 * 24) / hours
                elif "minute" in interval:
                    minutes = float(interval.split()[0])
                    images_per_month = (30 * 24 * 60) / minutes
                
                cost_per_month = (images_per_month * 0.040)  # $0.040 per image
                cost_text = f"Est. Cost: ${cost_per_month:.2f}/month"
            
            cost_label.config(text=cost_text)

        cost_label = Label(wallpaper_tab, text="Cost: $0/month (Manual only)", 
                          font=("Arial", 9), foreground="gray")
        cost_label.pack(pady=2)
        
        # Bind interval change to cost update
        interval_var.trace('w', update_cost_estimate)

        # Status Label
        status_label = Label(root, text="Welcome to Wallpaper AI Slideshow", font=("Arial", 10), anchor="w")
        status_label.pack(fill="x", side="bottom", pady=5)

        # Generate Now Button
        def generate_now():
            # Disable the generate button
            generate_button.config(state='disabled')
            
            prompt = (
                custom_prompt.get()
                if use_custom_prompt.get()
                else DEFAULT_PROMPTS[selected_prompt.get()]
                if selected_prompt.get() != "Random"
                else random.choice(list(DEFAULT_PROMPTS.values()))
            )
            
            def generation_complete():
                generate_button.config(state='normal')
                root.update()
            
            def run_generation():
                try:
                    generate_wallpaper(prompt, status_label, root)
                finally:
                    root.after(0, generation_complete)
            
            threading.Thread(target=run_generation, daemon=True).start()

        # Store generate button as global
        global generate_button
        generate_button = Button(wallpaper_tab, text="Generate Wallpaper Now", command=generate_now)
        generate_button.pack(pady=10)
        Button(wallpaper_tab, text="Hide App", command=lambda: minimize_to_tray(root)).pack(pady=10)

        # Add Library Tab
        library_tab = Frame(notebook)
        notebook.add(library_tab, text="Wallpaper Library")
        
        # Make library_listbox global
        global global_library_listbox
        global_library_listbox = tk.Listbox(library_tab, width=70, height=15)
        global_library_listbox.pack(pady=5, padx=5)
        
        # Rest of the library tab code remains the same
        global_library_listbox.bind('<Double-Button-1>', open_file_location)
        
        Button(library_tab, text="Use Selected Wallpaper", 
               command=use_selected_wallpaper).pack(pady=5)
        Button(library_tab, text="Refresh Library", 
               command=refresh_library_list).pack(pady=5)
        
        # Add tooltip label
        Label(library_tab, text="Tip: Double-click to open file location", 
              foreground="gray").pack(pady=5)

        # Initial library load
        refresh_library_list()

        # Add cleanup on window close
        def on_closing():
            global _system_tray_icon
            try:
                # Stop system tray icon if it exists
                if _system_tray_icon is not None:
                    _system_tray_icon.stop()
                    _system_tray_icon = None
                
                # Run cleanup
                cleanup()
                
                # Schedule quit after a short delay
                root.after(100, root.quit)
                
            except Exception as e:
                logger.error(f"Error during cleanup: {e}")
                root.quit
        
        root.protocol("WM_DELETE_WINDOW", on_closing)
        
        # Force window to front after creation
        root.after(1000, lambda: root.focus_force())
        
        root.mainloop()
        
    except Exception as e:
        logger.exception("Failed to create GUI")
        messagebox.showerror("Error", f"Failed to create GUI: {e}")
        sys.exit(1)

# Startup check
def check_startup():
    try:
        # Basic file system check
        test_file = "test_write.tmp"
        try:
            with open(test_file, "w") as f:
                f.write("test")
            os.remove(test_file)
        except Exception as e:
            print(f"File system check failed: {e}")
            return False

        # Check critical files
        if not os.path.exists("app_icon.ico"):
            print("Warning: App icon not found")
            # Continue anyway, not critical

        # Test Windows API
        try:
            ctypes.windll.user32.SystemParametersInfoW(0, 0, None, 0)
        except Exception as e:
            print(f"Windows API check failed: {e}")
            return False

        return True
    except Exception as e:
        print(f"Startup check failed: {e}")
        return False

if __name__ == "__main__":
    logger = setup_logging()
    mutex = None
    
    try:
        logger.info("Application starting...")
        logger.debug(f"Python version: {sys.version}")
        logger.debug(f"Current directory: {os.getcwd()}")
        
        # Kill existing instances first
        kill_existing_instances()
        
        # Create mutex
        mutex = create_mutex()
        if not mutex:
            sys.exit(1)
            
        # Set working directory
        if getattr(sys, 'frozen', False):
            os.chdir(os.path.dirname(sys.executable))
        
        # Run startup checks
        if not check_startup():
            logger.error("Startup checks failed")
            sys.exit(1)

        # Start GUI
        create_gui()

    except Exception as e:
        logger.exception("Fatal error occurred")
        try:
            messagebox.showerror("Fatal Error", str(e))
        except:
            print(f"Fatal error: {e}")
        sys.exit(1)
    finally:
        if mutex:
            try:
                win32api.CloseHandle(mutex)
            except:
                pass
        cleanup()
        logger.info("Application shutting down")
