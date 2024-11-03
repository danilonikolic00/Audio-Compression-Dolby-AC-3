import subprocess
import sys
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox, Menu, Toplevel, Label, StringVar
import os
import ffmpeg

settings = {
    'bitrate': '192',
    'sample_rate': '44.1',
    'channels': '2'
}

def get_ffmpeg_path():
    if getattr(sys, 'frozen', False):
        #exe
        ffmpeg_path = os.path.join(sys._MEIPASS, "ffmpeg.exe")
    else:
        #script
        ffmpeg_path = os.path.join(os.getcwd(), "ffmpeg.exe") 
    return ffmpeg_path

def get_ffmpeg_probe():
    if getattr(sys, 'frozen', False):
        # exe
        ffprobe_path = os.path.join(sys._MEIPASS, "ffprobe.exe")
    else:
        # script
        ffprobe_path = os.path.join(os.getcwd(), "ffprobe.exe") 
    return ffprobe_path

def select_file():
    file_paths = filedialog.askopenfilenames(filetypes=[("Audio Files", "*.mp3;*.wav;*.aac;*.flac;*.ogg;*.m4a")])
    if file_paths:
        input_entry.config(state='normal')
        input_entry.delete(0, ttk.END)
        input_entry.insert(0, ';'.join(file_paths))
        input_entry.config(state='readonly')

def convert_to_ac3():
    file_paths = input_entry.get().split(';')
    if not file_paths[0]:
        messagebox.showwarning("Input Error", "Please select audio files.")
        return

    ffmpeg_path = get_ffmpeg_path()

    total_files = len(file_paths)
    progress_bar['maximum'] = total_files
    
    make_progress_bar_visible()
    
    for index, input_file in enumerate(file_paths):
        if not input_file:
            continue
        
        output_file = filedialog.asksaveasfilename(defaultextension=".ac3", filetypes=[("AC3 File", "*.ac3")])

        if not output_file:
            messagebox.showinfo("Conversion Aborted", "Conversion aborted.")
            input_entry.config(state='normal')
            input_entry.delete(0, ttk.END)
            input_entry.config(state='readonly')
            reset_progress()
            make_progress_bar_transparent()
            return

        # Sample rate to integer
        sample_rate_int = int(float(settings['sample_rate']) * 1000)
        
        if output_file:
            try:
                ffmpeg_command = (
                    ffmpeg.input(input_file)
                    .output(output_file, acodec='ac3', ab=f"{settings['bitrate']}k", ar=sample_rate_int, ac=settings['channels'])
                    .compile()
                )

                full_command = [ffmpeg_path] + ffmpeg_command[1:]  # Skip the initial 'ffmpeg' added by compile()

                # print("Running command:", full_command)

                subprocess.run(full_command, creationflags=subprocess.CREATE_NO_WINDOW, check=True)

                progress_bar['value'] = index + 1
                progress_label.config(text=f"Converting {index + 1} of {total_files} files...")
                root.update_idletasks()

                original_size = get_file_size(input_file)
                converted_size = get_file_size(output_file)
                percentage = 100 - (converted_size / original_size * 100) if original_size > 0 else 0

                original_size_mb = original_size / (1024 * 1024)
                converted_size_mb = converted_size / (1024 * 1024)

                # Insert only the file name into the tree, store absolute path in iid
                input_file_name = os.path.basename(input_file)
                output_file_name = os.path.basename(output_file)
                tree.insert("", ttk.END, iid=output_file, values=(
                    input_file_name, output_file_name, 
                    f"{original_size_mb:.2f} MB", f"{converted_size_mb:.2f} MB", 
                    f"{percentage:.2f}"))

            except Exception as e:
                messagebox.showerror("Conversion Error", f"An error occurred with {input_file}: {e}")
            
    if output_file:
        progress_label.config(text="Conversion Complete!")

    input_entry.config(state='normal')
    input_entry.delete(0, ttk.END)
    input_entry.config(state='readonly')

    root.after(500, lambda: (reset_progress(), make_progress_bar_transparent()))

def get_file_size(file_path):
    try:
        return os.path.getsize(file_path)
    except FileNotFoundError:
        return 0

def get_additional_info(file_path):
    try:
        ffprobe = get_ffmpeg_probe()
        probe = ffmpeg.probe(file_path,cmd=ffprobe)
        stream = next((s for s in probe['streams'] if s['codec_type'] == 'audio'), None)
        
        if stream:
            codec = stream.get('codec_name', 'Unknown')
            sample_rate = stream.get('sample_rate', 'Unknown')
            channels = stream.get('channels', 'Unknown')
            bitrate = stream.get('bit_rate', 'Unknown')
            
            duration = stream.get('duration', 'Unknown')
            duration = float(duration) if duration != 'Unknown' else 'Unknown'
            if duration != 'Unknown':
                duration = f"{int(duration // 60)}:{int(duration % 60):02d}"

            if bitrate != 'Unknown':
                bitrate = f"{int(bitrate) / 1000:.2f} kbps"
            else:
                bitrate = 'Unknown'

            return {
                'Codec': codec,
                'Sample Rate': f"{sample_rate} Hz",
                'Channels': channels,
                'Bitrate': bitrate,
                'Duration': duration
            }
        else:
            return {'Error': 'No audio stream found'}
    except Exception as e:
        return {'Error': str(e)}

def make_progress_bar_transparent():
    progress_bar.configure(style="Transparent.Horizontal.TProgressbar")

def make_progress_bar_visible():
    progress_bar.configure(style="Visible.Horizontal.TProgressbar")

def reset_progress():
    progress_bar['value'] = 0
    progress_label.config(text="")

def clear_table():
    if not tree.get_children():
        messagebox.showwarning("Clear Error", "No files to clear.")
    else:
        for item in tree.get_children():
            tree.delete(item)

def clear_selected_rows():
    selected_items = tree.selection()
    if not selected_items:
        messagebox.showwarning("Clear Error", "No rows selected.")
    else:
        for item in selected_items:
            tree.delete(item)

def play_file(event):
    selected_item = tree.selection()
    if selected_item:
        # Retrieve the absolute path stored in the iid
        output_file = selected_item[0]
        if os.path.exists(output_file):
            os.startfile(output_file)  # Windows-specific
        else:
            messagebox.showerror("Error", "Cannot play file on this system.")

def show_additional_info():
    selected_item = tree.selection()
    if selected_item:
        output_file = selected_item[0]
        
        if os.path.exists(output_file):
            info = get_additional_info(output_file)
            
            info_window = Toplevel(root)
            info_window.title(f"Details: {tree.item(selected_item)['values'][1]}")
            info_window.geometry("350x350")
            info_window.resizable(False,False)
            info_window.attributes('-topmost', True)

            info_text = (
                f"Duration: {info.get('Duration', 'N/A')}\n"
                f"Codec: {info.get('Codec', 'N/A')}\n"
                f"Sample Rate: {info.get('Sample Rate', 'N/A')}\n"
                f"Channels: {info.get('Channels', 'N/A')}\n"
                f"Bitrate: {info.get('Bitrate', 'N/A')}\n"
            )

            info_label = Label(info_window, text=info_text, font=('Helvetica', 12), anchor='center', justify='left', wraplength=350)
            info_label.pack(expand=True,pady=20, padx=20)

def on_right_click(event):
    context_menu.post(event.x_root, event.y_root)

def update_progress_bar_width(event):
    window_width = root.winfo_width()
    new_width = int(window_width * 0.6)
    progress_bar.config(length=new_width)

def open_settings_window():
    settings_window = Toplevel(root)
    settings_window.title("Settings")
    settings_window.geometry("350x350")
    settings_window.resizable(False, False)

    center_frame = ttk.Frame(settings_window)
    center_frame.pack(expand=True, fill='both', pady=(30, 0))

    def apply_settings():
        settings['bitrate'] = bitrate_var.get()
        settings['sample_rate'] = sample_rate_var.get()
        settings['channels'] = channels_var.get()
        settings_window.destroy()

    ttk.Label(center_frame, text="Bitrate (kbps):").pack(pady=5)
    bitrate_var = StringVar(value=settings['bitrate'])
    bitrate_menu = ttk.Combobox(center_frame, textvariable=bitrate_var, values=['96', '128', '192', '256', '320', '448','640'], state='readonly')
    bitrate_menu.pack(pady=5)
    bitrate_menu.set(settings['bitrate']) 

    ttk.Label(center_frame, text="Sample Rate (kHz):").pack(pady=5)
    sample_rate_var = StringVar(value=settings['sample_rate'])
    sample_rate_menu = ttk.Combobox(center_frame, textvariable=sample_rate_var, values=['32', '44.1', '48'], state='readonly')
    sample_rate_menu.pack(pady=5)
    sample_rate_menu.set(settings['sample_rate'])  

    ttk.Label(center_frame, text="Channels:").pack(pady=5)
    channels_var = StringVar(value=settings['channels'])
    channels_menu = ttk.Combobox(center_frame, textvariable=channels_var, values=['1', '2', '3', '4', '5'], state='readonly')
    channels_menu.pack(pady=5)
    channels_menu.set(settings['channels']) 

    apply_button = ttk.Button(center_frame, text="Apply", command=apply_settings)
    apply_button.pack(pady=10)

root = ttk.Window(themename="superhero")
root.title("Audio to AC3 Converter")
root.geometry('1050x650')
root.minsize(650,650)

input_label = ttk.Label(root, text="Select audio files:", font=('Helvetica', 14))
input_label.pack(pady=10)

input_frame = ttk.Frame(root)
input_frame.pack(pady=5, padx=20, fill='x')

input_entry = ttk.Entry(input_frame, width=0, font=('Helvetica', 12), state='readonly')
input_entry.pack(side='left', expand=True, fill='x', padx=(0, 10))

select_button = ttk.Button(input_frame, text="Browse", command=select_file, bootstyle=INFO, width=15)
select_button.pack(side='right')

button_frame = ttk.Frame(root)
button_frame.pack(pady=20)

convert_button = ttk.Button(button_frame, text="Convert to AC3", command=convert_to_ac3, bootstyle=SUCCESS, width=20)  # Adjust width here
convert_button.pack(side='left', padx=10)

settings_button = ttk.Button(button_frame, text="Settings", command=open_settings_window, bootstyle=PRIMARY, width=20)  # Same width as convert_button
settings_button.pack(side='left', padx=10)

input_label = ttk.Label(root, text="Converted files:", font=('Helvetica', 14))
input_label.pack(pady=10)

style = ttk.Style()
style.configure('Treeview', font=("Helvetica", 10))

tree_frame = ttk.Frame(root)
tree_frame.pack(pady=10, fill=ttk.BOTH, expand=True)

tree = ttk.Treeview(tree_frame, columns=("Input File", "Output File", "Original Size", "Compressed Size", "Compression (%)"), show='headings', bootstyle=PRIMARY, style='Treeview')
tree.heading("Input File", text="Input File")
tree.column("Input File", anchor=ttk.CENTER)
tree.heading("Output File", text="Output File")
tree.column("Output File", anchor=ttk.CENTER)
tree.heading("Original Size", text="Original Size")
tree.column("Original Size", anchor=ttk.CENTER)
tree.heading("Compressed Size", text="Compressed Size")
tree.column("Compressed Size", anchor=ttk.CENTER)
tree.heading("Compression (%)", text="Compression (%)")
tree.column("Compression (%)", anchor=ttk.CENTER)
tree.pack(fill=ttk.BOTH, expand=True, padx=5)

#play file
tree.bind("<Double-1>", play_file)

#show additional info menu
tree.bind("<Button-3>", on_right_click)

clear_button = ttk.Button(root, text="Clear History", command=clear_table, bootstyle=DANGER, width=15)
clear_button.pack(pady=10)

clear_selected_button = ttk.Button(root, text="Clear Selected", command=clear_selected_rows, bootstyle=WARNING, width=15)
clear_selected_button.pack(pady=10)

progress_label = ttk.Label(root, text="", font=('Helvetica', 12))
progress_label.pack(pady=(10, 0))

style = ttk.Style()
style.configure("Transparent.Horizontal.TProgressbar", troughcolor=root.cget('bg'), background=root.cget('bg'), bordercolor=root.cget('bg'))
style.configure("Visible.Horizontal.TProgressbar", troughcolor="lightgray", background="#007bff")

progress_bar = ttk.Progressbar(root, style="Transparent.Horizontal.TProgressbar")
progress_bar.pack(pady=(0, 20))

context_menu = Menu(root, tearoff=0)
context_menu.add_command(label="Show Additional Info", command=show_additional_info)

root.bind("<Configure>", update_progress_bar_width)

root.mainloop()