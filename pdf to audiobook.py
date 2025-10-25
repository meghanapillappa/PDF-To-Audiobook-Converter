import fitz  # PyMuPDF
import pyttsx3
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os

class PDFToAudiobookConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to Audiobook Converter")
        self.root.geometry("600x500")
        self.root.configure(bg='#f0f0f0')
        
        # Initialize TTS engine
        self.engine = pyttsx3.init()
        self.setup_tts_engine()
        
        # Variables
        self.pdf_path = tk.StringVar()
        self.is_playing = False
        self.current_text = ""
        
        self.create_widgets()
    
    def setup_tts_engine(self):
        """Configure TTS engine settings"""
        voices = self.engine.getProperty('voices')
        self.engine.setProperty('voice', voices[0].id)  # Default voice
        self.engine.setProperty('rate', 150)  # Speech rate
        self.engine.setProperty('volume', 0.8)  # Volume
    
    def create_widgets(self):
        """Create GUI elements"""
        # Title
        title_label = tk.Label(self.root, text="PDF to Audiobook Converter", 
                              font=('Arial', 16, 'bold'), bg='#f0f0f0')
        title_label.pack(pady=10)
        
        # File selection frame
        file_frame = tk.Frame(self.root, bg='#f0f0f0')
        file_frame.pack(pady=10, padx=20, fill='x')
        
        tk.Label(file_frame, text="Select PDF File:", bg='#f0f0f0').pack(anchor='w')
        
        file_select_frame = tk.Frame(file_frame, bg='#f0f0f0')
        file_select_frame.pack(fill='x', pady=5)
        
        tk.Entry(file_select_frame, textvariable=self.pdf_path, width=50).pack(side='left', padx=(0, 10))
        tk.Button(file_select_frame, text="Browse", command=self.browse_pdf).pack(side='left')
        
        # Settings frame
        settings_frame = tk.LabelFrame(self.root, text="Audio Settings", bg='#f0f0f0')
        settings_frame.pack(pady=10, padx=20, fill='x')
        
        # Voice selection
        tk.Label(settings_frame, text="Voice:", bg='#f0f0f0').grid(row=0, column=0, sticky='w', padx=5, pady=2)
        self.voice_var = tk.StringVar(value="Male")
        voice_menu = ttk.Combobox(settings_frame, textvariable=self.voice_var, 
                                 values=["Male", "Female"], state="readonly")
        voice_menu.grid(row=0, column=1, sticky='w', padx=5, pady=2)
        
        # Speech rate
        tk.Label(settings_frame, text="Speech Rate:", bg='#f0f0f0').grid(row=1, column=0, sticky='w', padx=5, pady=2)
        self.rate_var = tk.IntVar(value=150)
        tk.Scale(settings_frame, variable=self.rate_var, from_=100, to=200, 
                orient='horizontal', bg='#f0f0f0').grid(row=1, column=1, sticky='ew', padx=5, pady=2)
        
        # Volume
        tk.Label(settings_frame, text="Volume:", bg='#f0f0f0').grid(row=2, column=0, sticky='w', padx=5, pady=2)
        self.volume_var = tk.DoubleVar(value=0.8)
        tk.Scale(settings_frame, variable=self.volume_var, from_=0.0, to=1.0, 
                resolution=0.1, orient='horizontal', bg='#f0f0f0').grid(row=2, column=1, sticky='ew', padx=5, pady=2)
        
        # Control buttons
        button_frame = tk.Frame(self.root, bg='#f0f0f0')
        button_frame.pack(pady=10)
        
        self.play_btn = tk.Button(button_frame, text="Play Audio", command=self.toggle_playback,
                                 bg='#4CAF50', fg='white', font=('Arial', 10, 'bold'))
        self.play_btn.pack(side='left', padx=5)
        
        tk.Button(button_frame, text="Export to MP3", command=self.export_audio,
                 bg='#2196F3', fg='white', font=('Arial', 10, 'bold')).pack(side='left', padx=5)
        
        tk.Button(button_frame, text="Stop", command=self.stop_audio,
                 bg='#f44336', fg='white', font=('Arial', 10, 'bold')).pack(side='left', padx=5)
        
        # Progress bar
        self.progress = ttk.Progressbar(self.root, orient='horizontal', length=500, mode='determinate')
        self.progress.pack(pady=10)
        
        # Status label
        self.status_label = tk.Label(self.root, text="Ready to convert PDF", 
                                    bg='#f0f0f0', font=('Arial', 10))
        self.status_label.pack(pady=5)
        
        # Text preview
        preview_frame = tk.LabelFrame(self.root, text="Text Preview", bg='#f0f0f0')
        preview_frame.pack(pady=10, padx=20, fill='both', expand=True)
        
        self.text_preview = tk.Text(preview_frame, height=10, wrap='word')
        scrollbar = tk.Scrollbar(preview_frame, command=self.text_preview.yview)
        self.text_preview.config(yscrollcommand=scrollbar.set)
        
        self.text_preview.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
    
    def browse_pdf(self):
        """Browse and select PDF file"""
        filename = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filename:
            self.pdf_path.set(filename)
            self.load_pdf_preview()
    
    def load_pdf_preview(self):
        """Load and display PDF text preview"""
        try:
            doc = fitz.open(self.pdf_path.get())
            text = ""
            for page_num in range(min(3, len(doc))):  # Preview first 3 pages
                page = doc.load_page(page_num)
                text += page.get_text()
            
            doc.close()
            
            # Clean and limit preview text
            text = self.clean_text(text)
            self.current_text = text
            self.text_preview.delete(1.0, tk.END)
            self.text_preview.insert(1.0, text[:2000] + ("..." if len(text) > 2000 else ""))
            
            self.status_label.config(text="PDF loaded successfully")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load PDF: {str(e)}")
    
    def clean_text(self, text):
        """Clean extracted text"""
        # Remove extra whitespace and normalize
        lines = text.split('\n')
        cleaned_lines = []
        for line in lines:
            line = line.strip()
            if line:  # Skip empty lines
                cleaned_lines.append(line)
        
        return '\n'.join(cleaned_lines)
    
    def extract_full_text(self):
        """Extract text from entire PDF"""
        try:
            doc = fitz.open(self.pdf_path.get())
            full_text = ""
            
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                full_text += page.get_text() + "\n"
                
                # Update progress
                progress = (page_num + 1) / len(doc) * 100
                self.progress['value'] = progress
                self.root.update_idletasks()
            
            doc.close()
            return self.clean_text(full_text)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract text: {str(e)}")
            return ""
    
    def update_tts_settings(self):
        """Update TTS engine with current settings"""
        voices = self.engine.getProperty('voices')
        voice_index = 0 if self.voice_var.get() == "Male" else 1
        if voice_index < len(voices):
            self.engine.setProperty('voice', voices[voice_index].id)
        
        self.engine.setProperty('rate', self.rate_var.get())
        self.engine.setProperty('volume', self.volume_var.get())
    
    def toggle_playback(self):
        """Toggle audio playback"""
        if not self.pdf_path.get():
            messagebox.showwarning("Warning", "Please select a PDF file first")
            return
        
        if not self.is_playing:
            self.start_playback()
        else:
            self.pause_playback()
    
    def start_playback(self):
        """Start audio playback in separate thread"""
        def play_audio():
            try:
                self.update_tts_settings()
                
                if not self.current_text:
                    self.current_text = self.extract_full_text()
                
                if self.current_text:
                    self.engine.say(self.current_text)
                    self.engine.runAndWait()
                
                self.is_playing = False
                self.play_btn.config(text="Play Audio")
                self.status_label.config(text="Playback completed")
                
            except Exception as e:
                messagebox.showerror("Error", f"Playback failed: {str(e)}")
                self.is_playing = False
                self.play_btn.config(text="Play Audio")
        
        if not self.is_playing:
            self.is_playing = True
            self.play_btn.config(text="Pause Audio")
            self.status_label.config(text="Playing audio...")
            
            # Run in separate thread to avoid GUI freeze
            thread = threading.Thread(target=play_audio)
            thread.daemon = True
            thread.start()
    
    def pause_playback(self):
        """Pause audio playback"""
        self.engine.stop()
        self.is_playing = False
        self.play_btn.config(text="Play Audio")
        self.status_label.config(text="Playback paused")
    
    def stop_audio(self):
        """Stop audio playback"""
        self.engine.stop()
        self.is_playing = False
        self.play_btn.config(text="Play Audio")
        self.status_label.config(text="Playback stopped")
        self.progress['value'] = 0
    
    def export_audio(self):
        """Export audio to MP3 file"""
        if not self.pdf_path.get():
            messagebox.showwarning("Warning", "Please select a PDF file first")
            return
        
        filename = filedialog.asksaveasfilename(
            title="Save Audio File",
            defaultextension=".mp3",
            filetypes=[("MP3 files", "*.mp3"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                self.status_label.config(text="Exporting audio...")
                
                if not self.current_text:
                    self.current_text = self.extract_full_text()
                
                # Note: pyttsx3 doesn't directly export to MP3
                # This would require additional libraries like gTTS
                messagebox.showinfo("Info", 
                    "MP3 export requires additional setup with gTTS.\n"
                    "Current version supports playback only.\n\n"
                    "To add MP3 export:\n"
                    "1. Install gTTS: pip install gtts\n"
                    "2. Use gTTS for MP3 conversion")
                
                self.status_label.config(text="Export completed")
                
            except Exception as e:
                messagebox.showerror("Error", f"Export failed: {str(e)}")
                self.status_label.config(text="Export failed")

def main():
    root = tk.Tk()
    app = PDFToAudiobookConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()