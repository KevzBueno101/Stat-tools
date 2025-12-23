"""
GUI Components Module
Contains reusable CustomTkinter widgets and components.
"""

import customtkinter as ctk
from typing import Callable


class LabeledEntry(ctk.CTkFrame):
    """
    A reusable frame containing a label and entry widget.
    """
    
    def __init__(self, master, label_text: str, default_value: str = "", 
                 **kwargs):
        """
        Initialize a labeled entry widget.
        
        Parameters:
        -----------
        master : CTk widget
            Parent widget
        label_text : str
            Text for the label
        default_value : str
            Default value for the entry field
        """
        super().__init__(master, **kwargs)
        
        # Configure grid
        self.grid_columnconfigure(0, weight=1)
        
        # Create label
        self.label = ctk.CTkLabel(
            self, 
            text=label_text,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.label.grid(row=0, column=0, sticky="w", padx=5, pady=(0, 5))
        
        # Create entry
        self.entry = ctk.CTkEntry(
            self,
            placeholder_text=f"Enter {label_text.lower()}",
            font=ctk.CTkFont(size=13)
        )
        self.entry.grid(row=1, column=0, sticky="ew", padx=5, pady=(0, 10))
        
        # Set default value if provided
        if default_value:
            self.entry.insert(0, default_value)
    
    def get(self) -> str:
        """Get the current value of the entry."""
        return self.entry.get()
    
    def set(self, value: str):
        """Set the value of the entry."""
        self.entry.delete(0, 'end')
        self.entry.insert(0, value)
    
    def clear(self):
        """Clear the entry field."""
        self.entry.delete(0, 'end')


class ExcelAnalysisFrame(ctk.CTkFrame):
    """
    A frame for Excel file analysis controls.
    """
    
    def __init__(self, master, on_analyze_callback, **kwargs):
        """
        Initialize Excel analysis frame.
        
        Parameters:
        -----------
        master : CTk widget
            Parent widget
        on_analyze_callback : callable
            Callback function when analyze button is clicked
        """
        super().__init__(master, **kwargs)
        
        self.on_analyze_callback = on_analyze_callback
        
        # Configure grid
        self.grid_columnconfigure(0, weight=1)
        
        # Title
        title_label = ctk.CTkLabel(
            self,
            text="📊 Excel File Analysis",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        title_label.grid(row=0, column=0, padx=15, pady=(15, 10), sticky="w")
        
        # File path display
        self.filepath_label = ctk.CTkLabel(
            self,
            text="No file selected",
            font=ctk.CTkFont(size=11),
            text_color="gray50",
            anchor="w"
        )
        self.filepath_label.grid(row=1, column=0, padx=15, pady=(0, 10), sticky="ew")
        
        # Button frame
        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=2, column=0, padx=15, pady=(0, 10), sticky="ew")
        button_frame.grid_columnconfigure((0, 1), weight=1)
        
        # Browse button
        self.browse_button = ctk.CTkButton(
            button_frame,
            text="Browse Excel File",
            command=self.browse_file,
            font=ctk.CTkFont(size=13),
            height=35
        )
        self.browse_button.grid(row=0, column=0, padx=(0, 5), sticky="ew")
        
        # Analyze button
        self.analyze_button = ctk.CTkButton(
            button_frame,
            text="Analyze File",
            command=self.analyze_file,
            font=ctk.CTkFont(size=13),
            height=35,
            state="disabled",
            fg_color="green",
            hover_color="darkgreen"
        )
        self.analyze_button.grid(row=0, column=1, padx=(5, 0), sticky="ew")
        
        # Column selection frame (initially hidden)
        self.column_frame = ctk.CTkFrame(self)
        self.column_frame.grid(row=3, column=0, padx=15, pady=(0, 15), sticky="ew")
        self.column_frame.grid_columnconfigure(0, weight=1)
        self.column_frame.grid_remove()  # Hide initially
        
        # Column 1 selection
        col1_label = ctk.CTkLabel(
            self.column_frame,
            text="Column 1 (X):",
            font=ctk.CTkFont(size=12, weight="bold")
        )
        col1_label.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="w")
        
        self.col1_var = ctk.StringVar(value="")
        self.col1_menu = ctk.CTkOptionMenu(
            self.column_frame,
            variable=self.col1_var,
            values=["No columns available"],
            font=ctk.CTkFont(size=12)
        )
        self.col1_menu.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="ew")
        
        # Column 2 selection
        col2_label = ctk.CTkLabel(
            self.column_frame,
            text="Column 2 (Y):",
            font=ctk.CTkFont(size=12, weight="bold")
        )
        col2_label.grid(row=2, column=0, padx=10, pady=(0, 5), sticky="w")
        
        self.col2_var = ctk.StringVar(value="")
        self.col2_menu = ctk.CTkOptionMenu(
            self.column_frame,
            variable=self.col2_var,
            values=["No columns available"],
            font=ctk.CTkFont(size=12)
        )
        self.col2_menu.grid(row=3, column=0, padx=10, pady=(0, 10), sticky="ew")
        
        # Store current filepath
        self.current_filepath = None
        self.numeric_columns = []
    
    def browse_file(self):
        """Open file dialog to select Excel file."""
        from tkinter import filedialog
        
        filepath = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[
                ("Excel Files", "*.xlsx *.xls"),
                ("All Files", "*.*")
            ]
        )
        
        if filepath:
            self.load_file(filepath)
    
    def load_file(self, filepath):
        """Load Excel file and extract numeric columns."""
        from excel_utils import read_excel_file, get_numeric_columns
        
        success, df, error = read_excel_file(filepath)
        
        if not success:
            from tkinter import messagebox
            messagebox.showerror("File Error", error)
            return
        
        # Get numeric columns
        self.numeric_columns = get_numeric_columns(df)
        
        if len(self.numeric_columns) < 2:
            from tkinter import messagebox
            messagebox.showerror(
                "Invalid File",
                "The Excel file must contain at least 2 numeric columns for correlation analysis."
            )
            return
        
        # Update UI
        self.current_filepath = filepath
        filename = filepath.split('/')[-1].split('\\')[-1]
        self.filepath_label.configure(text=f"File: {filename}")
        
        # Update column dropdowns
        self.col1_menu.configure(values=self.numeric_columns)
        self.col2_menu.configure(values=self.numeric_columns)
        self.col1_var.set(self.numeric_columns[0])
        self.col2_var.set(self.numeric_columns[1] if len(self.numeric_columns) > 1 else self.numeric_columns[0])
        
        # Show column selection and enable analyze button
        self.column_frame.grid()
        self.analyze_button.configure(state="normal")
    
    def analyze_file(self):
        """Trigger analysis callback."""
        if self.current_filepath and self.on_analyze_callback:
            col1 = self.col1_var.get()
            col2 = self.col2_var.get()
            
            if col1 == col2:
                from tkinter import messagebox
                messagebox.showerror(
                    "Invalid Selection",
                    "Please select two different columns for correlation analysis."
                )
                return
            
            self.on_analyze_callback(self.current_filepath, col1, col2)


class ResultsDisplay(ctk.CTkFrame):
    """
    A reusable frame for displaying calculation results.
    """
    
    def __init__(self, master, **kwargs):
        """
        Initialize a results display widget.
        
        Parameters:
        -----------
        master : CTk widget
            Parent widget
        """
        super().__init__(master, **kwargs)
        
        # Configure appearance
        self.configure(fg_color=("gray90", "gray20"))
        
        # Configure grid
        self.grid_columnconfigure(0, weight=1)
        
        # Title
        self.title_label = ctk.CTkLabel(
            self,
            text="Results",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        self.title_label.grid(row=0, column=0, padx=20, pady=(15, 10))
        
        # Results text widget (using CTkTextbox for better formatting)
        self.results_text = ctk.CTkTextbox(
            self,
            font=ctk.CTkFont(size=13),
            height=450,
            wrap="word"
        )
        self.results_text.grid(row=1, column=0, padx=20, pady=(0, 15), sticky="nsew")
        
        # Initially disable editing
        self.results_text.configure(state="disabled")
    
    def display_results(self, results_dict: dict):
        """
        Display formatted results in the text widget.
        
        Parameters:
        -----------
        results_dict : dict
            Dictionary containing computation results
        """
        # Enable editing to update text
        self.results_text.configure(state="normal")
        self.results_text.delete("1.0", "end")
        
        # Format and insert results
        output = []
        output.append("═" * 45)
        
        # Check if this is an Excel analysis (has column names and r_value)
        if 'column_1' in results_dict and 'r_value' in results_dict:
            output.append("EXCEL FILE CORRELATION ANALYSIS")
            output.append("═" * 45)
            output.append(f"\n📊 Variables Analyzed:")
            output.append(f"   X: {results_dict['column_1']}")
            output.append(f"   Y: {results_dict['column_2']}")
            output.append(f"\nSample Size (n): {results_dict['sample_size']}")
            if results_dict.get('rows_with_missing', 0) > 0:
                output.append(f"   (Original rows: {results_dict['original_rows']}, "
                            f"Excluded {results_dict['rows_with_missing']} with missing data)")
            output.append(f"Significance Level (α): {results_dict['alpha']}")
            output.append(f"Test Type: {results_dict['test_type'].title()}")
            output.append("\n" + "═" * 45)
            output.append("\n📈 Correlation Results:")
            output.append(f"\nPearson's r: {results_dict['r_value']:.6f}")
            output.append(f"P-value: {results_dict['p_value']:.6f}")
            output.append(f"\n" + "─" * 45)
            output.append("\n🎯 Critical Value Analysis:")
            output.append(f"\nDegrees of Freedom (df): {results_dict['degrees_of_freedom']}")
            output.append(f"t Critical: {results_dict['t_critical']:.6f}")
            output.append(f"r Critical: {results_dict['r_critical']:.6f}")
            output.append(f"\n|r| = {abs(results_dict['r_value']):.6f}")
            output.append(f"Required: {results_dict['r_critical']:.6f}")
            output.append("\n" + "═" * 45)
            output.append("\n📊 Interpretation:")
            output.append(f"\n{results_dict['significance_interpretation']}")
            
            # Add strength interpretation
            r_abs = abs(results_dict['r_value'])
            if r_abs >= 0.7:
                strength = "Strong"
            elif r_abs >= 0.4:
                strength = "Moderate"
            elif r_abs >= 0.2:
                strength = "Weak"
            else:
                strength = "Very Weak"
            
            direction = "positive" if results_dict['r_value'] > 0 else "negative"
            output.append(f"\nCorrelation Strength: {strength} {direction} correlation")
        
        else:
            # Standard critical value display
            output.append(f"Test Type: {results_dict['test_type'].title()}")
            output.append(f"Sample Size (n): {results_dict['sample_size']}")
            output.append(f"Significance Level (α): {results_dict['alpha']}")
            output.append("═" * 45)
            output.append(f"\nDegrees of Freedom (df): {results_dict['degrees_of_freedom']}")
            output.append(f"\nt Critical: {results_dict['t_critical']:.6f}")
            output.append(f"\nr Critical: {results_dict['r_critical']:.6f}")
            output.append("\n" + "═" * 45)
            output.append("\n📊 Interpretation:")
            output.append(f"\nFor a correlation to be statistically significant")
            output.append(f"at α = {results_dict['alpha']}, the absolute value of r")
            output.append(f"must be greater than {results_dict['r_critical']:.4f}")
        
        self.results_text.insert("1.0", "\n".join(output))
        
        # Disable editing again
        self.results_text.configure(state="disabled")
    
    def clear(self):
        """Clear the results display."""
        self.results_text.configure(state="normal")
        self.results_text.delete("1.0", "end")
        self.results_text.configure(state="disabled")