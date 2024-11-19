import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pptx import Presentation
from googletrans import Translator
from pptx.util import Pt
from pptx.dml.color import RGBColor
from PIL import Image, ImageTk



# Extract text from PPTX file
def extract_text_from_pptx(file_path):
    prs = Presentation(file_path)
    text_runs = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if run is not None:
                        text_runs.append(run.text)


    return text_runs



# Translate text
def translate_text(texts, dest='zh-cn', progress_callback=None):
    translator = Translator()
    translations = []
    for index, text in enumerate(texts):
        try:
            translation= translator.translate(text, dest=dest)
            translations.append(translation.text)

        except Exception as e:
            translations.append('Translation error')
            print(f"Error: {e}")

        if progress_callback:
            progress_callback(index + 1, len(texts))
    return translations


# Append or replace text in PPTX
def modify_pptx(file_path, original_texts, translations, method='append', progress_callback=None):
    prs = Presentation(file_path)
    i = 0
    num_texts = len(translations)
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue

            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if i < len(translations) and run.text in original_texts:
                        if method == 'replace':
                            run.text = translations[i]
                        elif method == 'append':
                            run.text += " " + translations[i]
                        i += 1

                        if progress_callback:
                            progress_callback(i, num_texts)

    prs.save('translated_presentation.pptx')


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("BabbelFish")
        self.geometry("600x400")


       

        # Open image with Pillow
        self.original_image = Image.open('image.jpeg')
        self.logo_image = ImageTk.PhotoImage(self.original_image)


        # Set the window size to the image size
        width, height = self.original_image.size
        self.geometry(f"{width}x{height}")

        # Logo
        
        self.logo_label = ttk.Label(self, image=self.logo_image)
        self.logo_label.place(x=0, y=0, relwidth=1, relheight=1)


         # Navigation Frame
        nav_frame = tk.Frame(self, height=100, bg='#ADADAD')
        nav_frame.pack(fill='x')





        # Styling for ttk Button
        style = ttk.Style()
        style.configure("LightBlue.TButton", foreground='black', background='#ADD8E6')


        style.configure("Status.TLabel", foreground='black', background='#D3FDAD', font=('Hevetica', 10))
       


        # Widgets
        self.file_label = ttk.Label(self, text="No file selected")
        self.file_label.pack(pady=20)


        # Progress Bar

        self.progress = ttk.Progressbar(self, orient='horizontal', length=300, mode='determinate')
        self.progress.pack(pady=20)


        ttk.Button(self, text="Open PPTX File", command = self.load_file, style="LightBlue.TButton").pack()


        self.language_label = ttk.Label(nav_frame, text="Select target language:")
        self.language_label.pack(pady=10)

        self.language_var = tk.StringVar()
        self.language_entry = ttk.Entry(nav_frame, textvariable=self.language_var)

        self.language_var.set("zh-cn")
        self.language_entry.pack()


        self.translate_button = ttk.Button(nav_frame, text="Translate Text", command=self.translate, style="LightBlue.TButton")
        self.translate_button.pack(pady=20)


        self.save_button = ttk.Button(nav_frame, text="Save Translated PPTX", command=self.save, style="LightBlue.TButton")

        self.save_button.pack(pady=20)


        self.status_label = ttk.Label(nav_frame, text="", style="Status.TLabel")
        self.status_label.pack(pady=20)


        # Data
        self.file_path = None
        self.original_texts = []
        self.translations = []




    def update_progress(self, current, total):
        self.progress['value'] = (current / total) * 100
        self.update_idletasks()



    def load_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("PowerPoint files", "*.pptx")])

        if self.file_path:
            self.file_label.config(text=f"Selected File: {self.file_path}")
            self.original_texts = extract_text_from_pptx(self.file_path)
            self.status_label.config(text="Text extracted. Ready to translate.")



    def translate(self):
        if not self.original_texts:
            messagebox.showerror("Error", "No text extracted or file not loaded.")
            return
        
        self.progress['value'] = 0
        self.progress['maximum'] = 100
        self.translations = translate_text(self.original_texts, self.language_var.get(), self.update_progress)
        self.status_label.config(text="Translation complete. Ready to save.")

    

    def save(self):
        if not self.translations:
            messagebox.showerror("Error", "No translations available.")
            return
        self.progress['value'] = 0
        modify_pptx(self.file_path, self.original_texts, self.translations, method='append', progress_callback=self.update_progress)

        messagebox.showinfo("Success", "Presentation saved as 'translated_presentation.pptx")


if __name__ == "__main__":
    app = App()
    app.mainloop()