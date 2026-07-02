import customtkinter as ctk
from tkinter import filedialog, messagebox
from pathlib import Path
import threading
from converter import PPTToDocConverter, ConversionError


class MainWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("PPT to DOC Converter")
        self.root.geometry("700x600")

        self.input_file = None
        self.output_file = None
        self.is_converting = False

        self._build_ui()

    def _build_ui(self):
        # Main container
        main_frame = ctk.CTkFrame(self.root)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Title
        title = ctk.CTkLabel(
            main_frame,
            text="PowerPoint to Word Converter",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title.pack(pady=(0, 20))

        # Instructions
        instructions = ctk.CTkLabel(
            main_frame,
            text="Convert PowerPoint presentations (.pptx) to Word documents (.docx)",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        instructions.pack(pady=(0, 30))

        # Input file section
        input_frame = ctk.CTkFrame(main_frame)
        input_frame.pack(fill="x", pady=(0, 15))

        input_label = ctk.CTkLabel(
            input_frame,
            text="Input PowerPoint File:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        input_label.pack(anchor="w", padx=15, pady=(15, 5))

        self.input_path_label = ctk.CTkLabel(
            input_frame,
            text="No file selected",
            font=ctk.CTkFont(size=12),
            text_color="gray",
            anchor="w"
        )
        self.input_path_label.pack(fill="x", padx=15, pady=(0, 10))

        input_btn = ctk.CTkButton(
            input_frame,
            text="Browse Input File",
            command=self._select_input_file,
            width=200
        )
        input_btn.pack(pady=(0, 15))

        # Output file section
        output_frame = ctk.CTkFrame(main_frame)
        output_frame.pack(fill="x", pady=(0, 15))

        output_label = ctk.CTkLabel(
            output_frame,
            text="Output Word Document:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        output_label.pack(anchor="w", padx=15, pady=(15, 5))

        self.output_path_label = ctk.CTkLabel(
            output_frame,
            text="Will be generated automatically",
            font=ctk.CTkFont(size=12),
            text_color="gray",
            anchor="w"
        )
        self.output_path_label.pack(fill="x", padx=15, pady=(0, 10))

        output_btn = ctk.CTkButton(
            output_frame,
            text="Browse Output Location (Optional)",
            command=self._select_output_file,
            width=200
        )
        output_btn.pack(pady=(0, 15))

        # Progress section
        progress_frame = ctk.CTkFrame(main_frame)
        progress_frame.pack(fill="x", pady=(0, 15))

        progress_label = ctk.CTkLabel(
            progress_frame,
            text="Progress:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        progress_label.pack(anchor="w", padx=15, pady=(15, 5))

        self.progress_bar = ctk.CTkProgressBar(progress_frame)
        self.progress_bar.pack(fill="x", padx=15, pady=(0, 10))
        self.progress_bar.set(0)

        self.progress_text = ctk.CTkLabel(
            progress_frame,
            text="Ready",
            font=ctk.CTkFont(size=11),
            text_color="gray"
        )
        self.progress_text.pack(anchor="w", padx=15, pady=(0, 15))

        # Log section
        log_label = ctk.CTkLabel(
            main_frame,
            text="Conversion Log:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        log_label.pack(anchor="w", pady=(0, 5))

        self.log_text = ctk.CTkTextbox(main_frame, height=150)
        self.log_text.pack(fill="both", expand=True, pady=(0, 15))

        # Convert button
        self.convert_btn = ctk.CTkButton(
            main_frame,
            text="Convert to Word",
            command=self._start_conversion,
            height=40,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color="#2ecc71",
            hover_color="#27ae60"
        )
        self.convert_btn.pack(fill="x")

    def _select_input_file(self):
        file_path = filedialog.askopenfilename(
            title="Select PowerPoint File",
            filetypes=[("PowerPoint Files", "*.pptx"), ("All Files", "*.*")]
        )
        if file_path:
            self.input_file = Path(file_path)
            self.input_path_label.configure(
                text=str(self.input_file),
                text_color="white"
            )

            # Auto-generate output path
            if not self.output_file:
                output_path = self.input_file.parent / f"{self.input_file.stem}_converted.docx"
                self.output_file = output_path
                self.output_path_label.configure(
                    text=str(self.output_file),
                    text_color="white"
                )

            self._log("Input file selected")

    def _select_output_file(self):
        initial_name = ""
        if self.input_file:
            initial_name = f"{self.input_file.stem}_converted.docx"

        file_path = filedialog.asksaveasfilename(
            title="Save Word Document As",
            defaultextension=".docx",
            initialfile=initial_name,
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        if file_path:
            self.output_file = Path(file_path)
            self.output_path_label.configure(
                text=str(self.output_file),
                text_color="white"
            )
            self._log("Output location selected")

    def _log(self, message):
        """Add message to log"""
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")

    def _update_progress(self, percent, message):
        """Update progress bar and text"""
        self.progress_bar.set(percent / 100)
        self.progress_text.configure(text=message)

    def _start_conversion(self):
        if self.is_converting:
            return

        if not self.input_file:
            messagebox.showerror("Error", "Please select an input PowerPoint file")
            return

        if not self.output_file:
            messagebox.showerror("Error", "Please select an output location")
            return

        # Clear log
        self.log_text.delete("1.0", "end")

        # Disable convert button
        self.is_converting = True
        self.convert_btn.configure(
            state="disabled",
            text="Converting...",
            fg_color="gray"
        )

        # Run conversion in a separate thread
        thread = threading.Thread(target=self._run_conversion, daemon=True)
        thread.start()

    def _run_conversion(self):
        try:
            converter = PPTToDocConverter(
                progress_cb=self._update_progress_safe,
                log_cb=self._log_safe
            )

            converter.convert(self.input_file, self.output_file)

            # Success
            self.root.after(0, lambda: messagebox.showinfo(
                "Success",
                f"Conversion completed successfully!\n\nOutput saved to:\n{self.output_file}"
            ))

        except ConversionError as e:
            self.root.after(0, lambda: messagebox.showerror("Conversion Error", str(e)))
            self._log_safe(f"ERROR: {e}")

        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Unexpected Error", str(e)))
            self._log_safe(f"UNEXPECTED ERROR: {e}")

        finally:
            # Re-enable convert button
            self.root.after(0, self._reset_ui)

    def _update_progress_safe(self, percent, message):
        """Thread-safe progress update"""
        self.root.after(0, lambda: self._update_progress(percent, message))

    def _log_safe(self, message):
        """Thread-safe log update"""
        self.root.after(0, lambda: self._log(message))

    def _reset_ui(self):
        """Reset UI after conversion"""
        self.is_converting = False
        self.convert_btn.configure(
            state="normal",
            text="Convert to Word",
            fg_color="#2ecc71"
        )
