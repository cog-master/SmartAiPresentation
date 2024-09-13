import tkinter as tk
from tkinter import ttk, filedialog
from slide_control import slide_control
from generate_ppt import generate_ppt

class ImageSelectFrame(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)

        self.image_path = None
        self.image_name_label = ttk.Label(self, text="No PPT selected")
        self.image_name_label.pack(pady=10)

        self.select_button = ttk.Button(self, text="Select PPT", command=self.select_image)
        self.select_button.pack()

        self.submit_button = ttk.Button(self, text="Submit", state="disabled", command=self.submit)
        self.submit_button.pack(pady=10)

    def select_image(self):
        import platform
        self.image_path = filedialog.askopenfilename(defaultextension=".pptx", filetypes=[("PowerPoint Presentations", "*.pptx")])

        if platform.system() == "Windows":
            self.image_path = self.image_path.replace('/', '\\')
            file_name = self.image_path.split("\\")[-1]
        else:
            file_name = self.image_path.split("/")[-1]
        
        if self.image_path:
            self.image_name_label.config(text=f'Selected file: {file_name}')
            self.submit_button.config(state="normal")
    
    def submit(self):
        slide_control(self.image_path)


class PromptFrame(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)

        self.prompt_entry = ttk.Entry(self)
        self.prompt_entry.pack(pady=10)

        category_options = ['Select Category for the PPT', 'professional', 'minimalist', 'modern']

        self.category_value = ''
        self.category = ttk.Combobox(self, values=category_options, textvariable=self.category_value, state='readonly')
        self.category.pack(pady=10)
        self.category.current(0)
        
        self.submit_button = ttk.Button(self, text="Submit", command=self.submit)
        self.submit_button.pack(pady=10)

        self.prompt_info_label = ttk.Label(self, text="")
        self.prompt_info_label.pack(pady=10)

    def submit(self):

        if self.prompt_entry.get() and 'Select' not in self.category_value:
            self.prompt_info_label.config(text='Processing......')
            self.submit_button.config(state='disabled')
            self.update_idletasks()
            ppt = generate_ppt(self.prompt_entry.get())
            ppt = ppt.update_ppxt()
            self.prompt_info_label.config(text=ppt)
            self.submit_button.config(state='enabled')
        else:
            self.prompt_info_label.config(text='Fill all the Required Field ...')



class ModernUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Modern Tkinter UI")
        self.geometry('300x300')
        self.resizable(False,False)
        menubar = tk.Menu(self)
        menubar.add_command(label="Select PPT", command=self.show_image_select_frame)
        menubar.add_command(label="Prompt", command=self.show_prompt_frame)
        self.config(menu=menubar)

        self.current_frame = None
        self.image_select_frame = ImageSelectFrame(self)
        self.prompt_frame = PromptFrame(self)

        # Show image selection frame by default
        self.show_image_select_frame()

    def show_image_select_frame(self):
        if self.current_frame is not self.image_select_frame:
            if self.current_frame:
                self.current_frame.pack_forget()
            self.current_frame = self.image_select_frame
            self.current_frame.pack(fill=tk.BOTH, expand=True)

    def show_prompt_frame(self):
        if self.current_frame is not self.prompt_frame:
            if self.current_frame:
                self.current_frame.pack_forget()
            self.current_frame = self.prompt_frame
            self.current_frame.pack(fill=tk.BOTH, expand=True)

if __name__ == "__main__":
    ui = ModernUI()
    ui.mainloop()
