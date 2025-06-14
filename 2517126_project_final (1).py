import math
import re
import tkinter as tk
from tkinter import ttk, messagebox
import ttkbootstrap as tb

import locale

# try to set locale properly
for _candidate in ("en_US.utf8", "English_United_States.1252", "C"):
    try:
        locale.setlocale(locale.LC_ALL, _candidate)
        break
    except locale.Error:
        continue

# setup theme
style = tb.Style(theme="superhero")
root = style.master
root.title("Smart Cable Selector (Auto-Filter)")
root.geometry("900x700")

# theme list
_themes = [
    "minty", "journal", "simplex", "sandstone", "lumen", "united", "cosmo",
    "superhero", "cyborg", "morph", "darkly"
]
_theme_idx = 0


def toggle_theme():
    global _theme_idx
    _theme_idx = (_theme_idx + 1) % len(_themes)
    style.theme_use(_themes[_theme_idx])


def show_help():
    help_text = """
SMART CABLE SELECTOR - USER GUIDE

WHAT THIS APP DOES:
This application helps you select the right electrical cable for your project and calculates important electrical parameters like voltage drop, power losses, and costs.

HOW TO USE:

STEP 1 - ENTER YOUR LOAD REQUIREMENTS:
   - Load Type: Choose your application type (Industrial, Residential, etc.)
   - Active Power: Enter power consumption in MW (Megawatts)
   - Reactive Power: Enter reactive power in MVar (Megavolt-amperes reactive)
   - System Voltage: Enter your system voltage in kV (Kilovolts)

STEP 2 - SPECIFY INSTALLATION CONDITIONS:
   - Cable Type: Choose Single-core (individual cables) or Three-core (bundled)
   - Arrangement: For single-core, choose Flat (side-by-side) or Trefoil (triangular)
   - Parallel Circuits: Number of parallel cable runs (1-6 for three-core, 1-2 for single-core)
   - Cable Length: Total cable length in kilometers
   - Ambient Temperature: Installation environment temperature (5-40°C)

STEP 3 - AUTOMATIC FILTERING:
   The app automatically shows only suitable cables as you type. Cables are filtered by:
   - Voltage rating (must meet or exceed your system voltage)
   - Current capacity (after temperature and trench corrections)
   - Cable type compatibility

STEP 4 - SELECT A CABLE:
   - Click on any cable from the filtered list
   - View cable parameters and capacity check
   - Green "Valid" status means the cable can handle your load safely

STEP 5 - CALCULATE PERFORMANCE:
   - Click "Calculate Losses & Regulation" for detailed analysis (or press Enter)
   - Review voltage regulation, power losses, and 10-year costs
   - Lower total cost usually indicates better cable choice

TIPS:
   - Start with Load Type and System Voltage
   - The app shows the cheapest suitable cable highlighted in green
   - Safety Margin shows how much extra capacity you have
   - Consider both cable cost and energy loss cost for best value

COMMON VALUES:
   - Residential: 0.01-0.1 MW, 400V-1kV
   - Commercial: 0.1-1 MW, 1kV-10kV  
   - Industrial: 1-10 MW, 3kV-35kV

KEYBOARD SHORTCUTS:
   - Press Enter: Calculate losses and regulation (when cable selected)
   - Press F1: Show this help dialog
   
   ** Smart Cable Selector - Ata Turk - 2025**
   """

    help_window = tk.Toplevel(root)
    help_window.title("How to Use - Smart Cable Selector")
    help_window.geometry("700x600")
    help_window.transient(root)
    help_window.grab_set()

    help_window.update_idletasks()
    x = (help_window.winfo_screenwidth() // 2) - (700 // 2)
    y = (help_window.winfo_screenheight() // 2) - (600 // 2)
    help_window.geometry(f"700x600+{x}+{y}")

    help_frame = ttk.Frame(help_window, padding=20)
    help_frame.pack(fill=tk.BOTH, expand=True)

    text_frame = ttk.Frame(help_frame)
    text_frame.pack(fill=tk.BOTH, expand=True)

    help_text_widget = tk.Text(text_frame, wrap=tk.WORD, font=("Segoe UI", 10),
                               background="#2a2a2a", foreground="#ffffff",
                               insertbackground="#ffffff", padx=15, pady=15)
    help_scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=help_text_widget.yview)
    help_text_widget.configure(yscrollcommand=help_scrollbar.set)

    help_text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    help_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    help_text_widget.insert("1.0", help_text)
    help_text_widget.configure(state="disabled")

    close_button = ttk.Button(help_frame, text="Close", command=help_window.destroy,
                              style="success.TButton")
    close_button.pack(pady=(10, 0))


# validators
def _is_positive_float(txt: str) -> bool:
    if txt == "":
        return True
    if txt in (".", "-"):
        return False
    if txt.startswith("0") and len(txt) > 1 and txt[1] != ".":
        return txt[1] == "."
    try:
        val = float(txt)
        return val >= 0
    except ValueError:
        return False


def _is_positive_int(txt: str) -> bool:
    if txt == "":
        return True
    try:
        return int(txt) > 0
    except ValueError:
        return False


# cable database
cable_list = [
    # 0.6/1kV single-core
    {"id": 1, "code": "1x10 mm2", "voltage": "0.6/1 kV", "fcc": 81, "tcc": 69, "resistance": 1.83,
     "inductance_flat": 0.34, "inductance_trefoil": 0.41, "capacitance": None, "price": 102300},
    {"id": 2, "code": "1x16 mm2", "voltage": "0.6/1 kV", "fcc": 108, "tcc": 92, "resistance": 1.15,
     "inductance_flat": 0.317, "inductance_trefoil": 0.387, "capacitance": None, "price": 156600},
    {"id": 3, "code": "1x25 mm2", "voltage": "0.6/1 kV", "fcc": 146, "tcc": 124, "resistance": 0.727,
     "inductance_flat": 0.304, "inductance_trefoil": 0.374, "capacitance": None, "price": 241400},
    {"id": 4, "code": "1x35 mm2", "voltage": "0.6/1 kV", "fcc": 180, "tcc": 153, "resistance": 0.524,
     "inductance_flat": 0.291, "inductance_trefoil": 0.36, "capacitance": None, "price": 328300},
    {"id": 5, "code": "1x50 mm2", "voltage": "0.6/1 kV", "fcc": 220, "tcc": 187, "resistance": 0.387,
     "inductance_flat": 0.281, "inductance_trefoil": 0.351, "capacitance": None, "price": 434000},
    {"id": 6, "code": "1x70 mm2", "voltage": "0.6/1 kV", "fcc": 279, "tcc": 237, "resistance": 0.268,
     "inductance_flat": 0.272, "inductance_trefoil": 0.341, "capacitance": None, "price": 623000},
    {"id": 7, "code": "1x95 mm2", "voltage": "0.6/1 kV", "fcc": 347, "tcc": 294, "resistance": 0.193,
     "inductance_flat": 0.264, "inductance_trefoil": 0.333, "capacitance": None, "price": 847000},
    {"id": 8, "code": "1x120 mm2", "voltage": "0.6/1 kV", "fcc": 405, "tcc": 343, "resistance": 0.153,
     "inductance_flat": 0.259, "inductance_trefoil": 0.329, "capacitance": None, "price": 1080000},

    # 0.6/1kV three-core
    {"id": 9, "code": "3x16+10 mm2", "voltage": "0.6/1 kV", "fcc": None, "tcc": 89, "resistance": 1.15,
     "inductance_flat": None, "inductance_trefoil": 0.264, "capacitance": None, "price": 640000},
    {"id": 10, "code": "3x25+16 mm2", "voltage": "0.6/1 kV", "fcc": None, "tcc": 120, "resistance": 0.727,
     "inductance_flat": None, "inductance_trefoil": 0.265, "capacitance": None, "price": 990000},
    {"id": 11, "code": "3x35+16 mm2", "voltage": "0.6/1 kV", "fcc": None, "tcc": 147, "resistance": 0.524,
     "inductance_flat": None, "inductance_trefoil": 0.258, "capacitance": None, "price": 1300000},
    {"id": 12, "code": "3x50+25 mm2", "voltage": "0.6/1 kV", "fcc": None, "tcc": 179, "resistance": 0.387,
     "inductance_flat": None, "inductance_trefoil": 0.256, "capacitance": None, "price": 1750000},
    {"id": 13, "code": "3x70+35 mm2", "voltage": "0.6/1 kV", "fcc": None, "tcc": 224, "resistance": 0.268,
     "inductance_flat": None, "inductance_trefoil": 0.253, "capacitance": None, "price": 2520000},
    {"id": 14, "code": "3x95+50 mm2", "voltage": "0.6/1 kV", "fcc": None, "tcc": 277, "resistance": 0.193,
     "inductance_flat": None, "inductance_trefoil": 0.247, "capacitance": None, "price": 3400000},
    {"id": 15, "code": "3x120+70 mm2", "voltage": "0.6/1 kV", "fcc": None, "tcc": 323, "resistance": 0.153,
     "inductance_flat": None, "inductance_trefoil": 0.246, "capacitance": None, "price": 4400000},
    {"id": 16, "code": "3x150+70 mm2", "voltage": "0.6/1 kV", "fcc": None, "tcc": 368, "resistance": 0.124,
     "inductance_flat": None, "inductance_trefoil": 0.248, "capacitance": None, "price": 5200000},

    # 3.6/6kV single-core
    {"id": 17, "code": "1x25 mm2", "voltage": "3.6/6 kV", "fcc": 196, "tcc": 163, "resistance": 0.727,
     "inductance_flat": 0.77, "inductance_trefoil": 0.43, "capacitance": 0.25, "price": 351000},
    {"id": 18, "code": "1x35 mm2", "voltage": "3.6/6 kV", "fcc": 238, "tcc": 198, "resistance": 0.524,
     "inductance_flat": 0.75, "inductance_trefoil": 0.41, "capacitance": 0.28, "price": 532000},
    {"id": 19, "code": "1x50 mm2", "voltage": "3.6/6 kV", "fcc": 286, "tcc": 238, "resistance": 0.387,
     "inductance_flat": 0.72, "inductance_trefoil": 0.39, "capacitance": 0.31, "price": 638000},
    {"id": 20, "code": "1x70 mm2", "voltage": "3.6/6 kV", "fcc": 356, "tcc": 296, "resistance": 0.268,
     "inductance_flat": 0.68, "inductance_trefoil": 0.37, "capacitance": 0.36, "price": 825000},
    {"id": 21, "code": "1x95 mm2", "voltage": "3.6/6 kV", "fcc": 434, "tcc": 361, "resistance": 0.193,
     "inductance_flat": 0.65, "inductance_trefoil": 0.36, "capacitance": 0.4, "price": 1070000},
    {"id": 22, "code": "1x120 mm2", "voltage": "3.6/6 kV", "fcc": 600, "tcc": 417, "resistance": 0.153,
     "inductance_flat": 0.63, "inductance_trefoil": 0.34, "capacitance": 0.44, "price": 1300000},
    {"id": 23, "code": "1x150 mm2", "voltage": "3.6/6 kV", "fcc": 559, "tcc": 473, "resistance": 0.124,
     "inductance_flat": 0.62, "inductance_trefoil": 0.33, "capacitance": 0.48, "price": 1640000},
    {"id": 24, "code": "1x185 mm2", "voltage": "3.6/6 kV", "fcc": 637, "tcc": 543, "resistance": 0.0991,
     "inductance_flat": 0.6, "inductance_trefoil": 0.32, "capacitance": 0.52, "price": 1950000},

    # 3.6/6kV three-core
    {"id": 25, "code": "3x25+16 mm2", "voltage": "3.6/6 kV", "fcc": None, "tcc": 143, "resistance": 0.727,
     "inductance_flat": None, "inductance_trefoil": 0.37, "capacitance": 0.25, "price": 1647000},
    {"id": 26, "code": "3x35+16 mm2", "voltage": "3.6/6 kV", "fcc": None, "tcc": 172, "resistance": 0.524,
     "inductance_flat": None, "inductance_trefoil": 0.35, "capacitance": 0.28, "price": 2002000},
    {"id": 27, "code": "3x50+16 mm2", "voltage": "3.6/6 kV", "fcc": None, "tcc": 205, "resistance": 0.387,
     "inductance_flat": None, "inductance_trefoil": 0.34, "capacitance": 0.3, "price": 2594000},
    {"id": 28, "code": "3x70+16 mm2", "voltage": "3.6/6 kV", "fcc": None, "tcc": 253, "resistance": 0.268,
     "inductance_flat": None, "inductance_trefoil": 0.32, "capacitance": 0.35, "price": 3450000},
    {"id": 29, "code": "3x95+16 mm2", "voltage": "3.6/6 kV", "fcc": None, "tcc": 307, "resistance": 0.193,
     "inductance_flat": None, "inductance_trefoil": 0.31, "capacitance": 0.39, "price": 4727000},
    {"id": 30, "code": "3x120+16 mm2", "voltage": "3.6/6 kV", "fcc": None, "tcc": 352, "resistance": 0.153,
     "inductance_flat": None, "inductance_trefoil": 0.3, "capacitance": 0.43, "price": 5784000},
    {"id": 31, "code": "3x150+25 mm2", "voltage": "3.6/6 kV", "fcc": None, "tcc": 397, "resistance": 0.124,
     "inductance_flat": None, "inductance_trefoil": 0.29, "capacitance": 0.47, "price": 6963000},
    {"id": 32, "code": "3x185+25 mm2", "voltage": "3.6/6 kV", "fcc": None, "tcc": 453, "resistance": 0.0991,
     "inductance_flat": None, "inductance_trefoil": 0.28, "capacitance": 0.5, "price": 8481000},

    # 6/10kV single-core
    {"id": 33, "code": "1x35 mm2", "voltage": "6/10 kV", "fcc": 231, "tcc": 195, "resistance": 0.524,
     "inductance_flat": 0.661, "inductance_trefoil": 0.383, "capacitance": 0.223, "price": 785100},
    {"id": 34, "code": "1x50 mm2", "voltage": "6/10 kV", "fcc": 277, "tcc": 234, "resistance": 0.387,
     "inductance_flat": 0.636, "inductance_trefoil": 0.366, "capacitance": 0.248, "price": 944900},
    {"id": 35, "code": "1x70 mm2", "voltage": "6/10 kV", "fcc": 345, "tcc": 292, "resistance": 0.268,
     "inductance_flat": 0.606, "inductance_trefoil": 0.349, "capacitance": 0.285, "price": 1226000},
    {"id": 36, "code": "1x95 mm2", "voltage": "6/10 kV", "fcc": 418, "tcc": 354, "resistance": 0.193,
     "inductance_flat": 0.582, "inductance_trefoil": 0.334, "capacitance": 0.32, "price": 1533000},
    {"id": 37, "code": "1x120 mm2", "voltage": "6/10 kV", "fcc": 481, "tcc": 407, "resistance": 0.153,
     "inductance_flat": 0.563, "inductance_trefoil": 0.323, "capacitance": 0.35, "price": 1872000},
    {"id": 38, "code": "1x150 mm2", "voltage": "6/10 kV", "fcc": 537, "tcc": 460, "resistance": 0.124,
     "inductance_flat": 0.546, "inductance_trefoil": 0.313, "capacitance": 0.382, "price": 2362000},
    {"id": 39, "code": "1x185 mm2", "voltage": "6/10 kV", "fcc": 612, "tcc": 527, "resistance": 0.0991,
     "inductance_flat": 0.529, "inductance_trefoil": 0.304, "capacitance": 0.415, "price": 2838000},

    # 6/10kV three-core
    {"id": 40, "code": "3x35 mm2", "voltage": "6/10 kV", "fcc": None, "tcc": 173, "resistance": 0.524,
     "inductance_flat": None, "inductance_trefoil": 0.374, "capacitance": 0.189, "price": 2127000},
    {"id": 41, "code": "3x50 mm2", "voltage": "6/10 kV", "fcc": None, "tcc": 206, "resistance": 0.387,
     "inductance_flat": None, "inductance_trefoil": 0.355, "capacitance": 0.209, "price": 2716000},
    {"id": 42, "code": "3x70 mm2", "voltage": "6/10 kV", "fcc": None, "tcc": 257, "resistance": 0.268,
     "inductance_flat": None, "inductance_trefoil": 0.336, "capacitance": 0.236, "price": 3603000},
    {"id": 43, "code": "3x95 mm2", "voltage": "6/10 kV", "fcc": None, "tcc": 313, "resistance": 0.193,
     "inductance_flat": None, "inductance_trefoil": 0.32, "capacitance": 0.263, "price": 4901000},
    {"id": 44, "code": "3x120 mm2", "voltage": "6/10 kV", "fcc": None, "tcc": 360, "resistance": 0.153,
     "inductance_flat": None, "inductance_trefoil": 0.308, "capacitance": 0.291, "price": 5934000},
    {"id": 45, "code": "3x150 mm2", "voltage": "6/10 kV", "fcc": None, "tcc": 410, "resistance": 0.124,
     "inductance_flat": None, "inductance_trefoil": 0.299, "capacitance": 0.314, "price": 7125000},
    {"id": 46, "code": "3x185 mm2", "voltage": "6/10 kV", "fcc": None, "tcc": 469, "resistance": 0.0991,
     "inductance_flat": None, "inductance_trefoil": 0.29, "capacitance": 0.341, "price": 8659000},

    # 12/20kV single-core
    {"id": 47, "code": "1x95 mm2", "voltage": "12/20 kV", "fcc": 420, "tcc": 358, "resistance": 0.193,
     "inductance_flat": 0.59, "inductance_trefoil": 0.36, "capacitance": 0.218, "price": 1560000},
    {"id": 48, "code": "1x120 mm2", "voltage": "12/20 kV", "fcc": 483, "tcc": 412, "resistance": 0.153,
     "inductance_flat": 0.571, "inductance_trefoil": 0.349, "capacitance": 0.238, "price": 1906000},
    {"id": 49, "code": "1x150 mm2", "voltage": "12/20 kV", "fcc": 540, "tcc": 466, "resistance": 0.124,
     "inductance_flat": 0.554, "inductance_trefoil": 0.338, "capacitance": 0.258, "price": 2393000},
    {"id": 50, "code": "1x185 mm2", "voltage": "12/20 kV", "fcc": 614, "tcc": 534, "resistance": 0.0991,
     "inductance_flat": 0.538, "inductance_trefoil": 0.329, "capacitance": 0.278, "price": 2877000},
    {"id": 51, "code": "1x240 mm2", "voltage": "12/20 kV", "fcc": 718, "tcc": 627, "resistance": 0.0754,
     "inductance_flat": 0.518, "inductance_trefoil": 0.317, "capacitance": 0.308, "price": 3543000},
    {"id": 52, "code": "1x300 mm2", "voltage": "12/20 kV", "fcc": 813, "tcc": 715, "resistance": 0.0601,
     "inductance_flat": 0.501, "inductance_trefoil": 0.308, "capacitance": 0.336, "price": 4455000},
    {"id": 53, "code": "1x400 mm2", "voltage": "12/20 kV", "fcc": 904, "tcc": 819, "resistance": 0.047,
     "inductance_flat": 0.48, "inductance_trefoil": 0.298, "capacitance": 0.377, "price": 5669000},

    # 20.3/35kV single-core
    {"id": 54, "code": "1x150 mm2", "voltage": "20.3/35 kV", "fcc": 559, "tcc": 473, "resistance": 0.124,
     "inductance_flat": 0.64, "inductance_trefoil": 0.41, "capacitance": 0.17, "price": 1720000},
    {"id": 55, "code": "1x185 mm2", "voltage": "20.3/35 kV", "fcc": 637, "tcc": 543, "resistance": 0.0991,
     "inductance_flat": 0.63, "inductance_trefoil": 0.39, "capacitance": 0.18, "price": 2020000},
    {"id": 56, "code": "1x240 mm2", "voltage": "20.3/35 kV", "fcc": 745, "tcc": 641, "resistance": 0.0754,
     "inductance_flat": 0.6, "inductance_trefoil": 0.38, "capacitance": 0.2, "price": 2530000},
    {"id": 57, "code": "1x300 mm2", "voltage": "20.3/35 kV", "fcc": 846, "tcc": 735, "resistance": 0.0601,
     "inductance_flat": 0.59, "inductance_trefoil": 0.37, "capacitance": 0.21, "price": 3150000},
    {"id": 58, "code": "1x400 mm2", "voltage": "20.3/35 kV", "fcc": 938, "tcc": 845, "resistance": 0.047,
     "inductance_flat": 0.57, "inductance_trefoil": 0.35, "capacitance": 0.23, "price": 4100000},
    {"id": 59, "code": "1x500 mm2", "voltage": "20.3/35 kV", "fcc": 1010, "tcc": 950, "resistance": 0.0366,
     "inductance_flat": 0.55, "inductance_trefoil": 0.34, "capacitance": 0.26, "price": 5150000},
    {"id": 60, "code": "1x630 mm2", "voltage": "20.3/35 kV", "fcc": 1120, "tcc": 1040, "resistance": 0.0283,
     "inductance_flat": 0.52, "inductance_trefoil": 0.33, "capacitance": 0.29, "price": 6550000},

    # 20.3/35kV three-core
    {"id": 61, "code": "3x95 mm2", "voltage": "20.3/35 kV", "fcc": None, "tcc": 307, "resistance": 0.193,
     "inductance_flat": None, "inductance_trefoil": 0.4, "capacitance": 0.15, "price": 4600000},
    {"id": 62, "code": "3x120 mm2", "voltage": "20.3/35 kV", "fcc": None, "tcc": 352, "resistance": 0.153,
     "inductance_flat": None, "inductance_trefoil": 0.39, "capacitance": 0.16, "price": 5500000},
    {"id": 63, "code": "3x150 mm2", "voltage": "20.3/35 kV", "fcc": None, "tcc": 397, "resistance": 0.124,
     "inductance_flat": None, "inductance_trefoil": 0.37, "capacitance": 0.17, "price": 6400000},
    {"id": 64, "code": "3x185 mm2", "voltage": "20.3/35 kV", "fcc": None, "tcc": 453, "resistance": 0.0991,
     "inductance_flat": None, "inductance_trefoil": 0.36, "capacitance": 0.18, "price": 7500000},
    {"id": 65, "code": "3x240 mm2", "voltage": "20.3/35 kV", "fcc": None, "tcc": 529, "resistance": 0.0754,
     "inductance_flat": None, "inductance_trefoil": 0.35, "capacitance": 0.2, "price": 9400000},
    {"id": 66, "code": "3x300 mm2", "voltage": "20.3/35 kV", "fcc": None, "tcc": 626, "resistance": 0.0601,
     "inductance_flat": None, "inductance_trefoil": 0.29, "capacitance": 0.22, "price": 11300000},
    {"id": 67, "code": "3x400 mm2", "voltage": "20.3/35 kV", "fcc": None, "tcc": 720, "resistance": 0.047,
     "inductance_flat": None, "inductance_trefoil": 0.28, "capacitance": 0.24, "price": 14300000},
]


def parse_voltage_kV(voltage_str):
    # get max voltage from string like "6/10 kV"
    if not voltage_str:
        return 0.0
    parts = voltage_str.replace('kV', '').strip().split('/')
    try:
        nums = [float(p) for p in parts]
        return max(nums) if nums else 0.0
    except:
        return float(parts[-1]) if parts else 0.0


def get_temp_factor(temp_c):
    temp_table = {5: 1.15, 10: 1.10, 15: 1.05, 20: 1.00, 25: 0.95, 30: 0.90, 35: 0.85, 40: 0.80}
    if temp_c <= 5:
        return 1.15
    if temp_c >= 40:
        return 0.80
    rounded_up = min(k for k in temp_table if k >= temp_c)
    return temp_table[rounded_up]


def get_trench_factor(num_circuits):
    trench_table = {1: 1.00, 2: 0.90, 3: 0.85, 4: 0.80, 5: 0.75, 6: 0.70}
    num_circuits = max(1, min(6, num_circuits))
    return trench_table[num_circuits]


def auto_filter_cables():
    # clear table
    for item in tree.get_children():
        tree.delete(item)

    selected_cable["data"] = None
    if 'update_cable_display' in globals():
        update_cable_display()

    try:
        P = float(active_power_var.get().replace(',', '.')) or 0.0
        Q = float(reactive_power_var.get().replace(',', '.')) or 0.0
        V = float(voltage_var.get().replace(',', '.')) or 0.0
        N = int(circuits_spin.get()) or 1
        ambient = float(ambient_spin.get()) or 20.0
    except:
        # show all if invalid input
        for idx, cable in enumerate(cable_list):
            tags = ["oddrow" if idx % 2 == 0 else "evenrow"]
            tree.insert("", tk.END, values=(cable["id"], cable["code"], cable["voltage"], f"{cable['price']}"),
                        tags=tags)
        return

    max_c = 2 if cable_type_combo.get() == "Single-core" else 6
    if not (1 <= N <= max_c):
        N = min(max_c, max(1, N))

    if not (5 <= ambient <= 40):
        ambient = max(5, min(40, ambient))

    if P <= 0 or V <= 0:
        # show all if no valid load
        for idx, cable in enumerate(cable_list):
            tags = ["oddrow" if idx % 2 == 0 else "evenrow"]
            tree.insert("", tk.END, values=(cable["id"], cable["code"], cable["voltage"], f"{cable['price']}"),
                        tags=tags)
        return

    # calc required current
    S_MVA = math.sqrt(P ** 2 + Q ** 2)
    I_total = S_MVA * 1e6 / (math.sqrt(3) * V * 1e3) if V > 0 else 0
    I_per_circuit = I_total / N

    ctype = cable_type_combo.get()
    arrangement = arrangement_combo.get()
    results = []

    for cable in cable_list:
        code = cable["code"]
        is_single_core = code.startswith("1x")

        # filter by type
        if (ctype == "Single-core" and not is_single_core) or (ctype == "Three-core" and is_single_core):
            continue

        # filter by voltage
        if parse_voltage_kV(cable["voltage"]) < V:
            continue

        # check capacity
        base_capacity = cable["fcc"] if is_single_core and arrangement == "Flat" else cable["tcc"]
        if base_capacity is None:
            continue

        # apply derating
        temp_factor = get_temp_factor(ambient)
        cables_in_trench = N * 3 if is_single_core else N
        trench_factor = get_trench_factor(min(cables_in_trench, 6))
        derated_capacity = base_capacity * temp_factor * trench_factor

        if derated_capacity >= I_per_circuit:
            results.append((cable["id"], code, cable["voltage"], base_capacity, cable["price"]))

    # show results
    if results:
        min_rating = min(parse_voltage_kV(r[2]) for r in results)
        results = [r for r in results if parse_voltage_kV(r[2]) == min_rating]
        min_price = min(r[4] for r in results)

        for idx, (cid, code, volt, cap, price) in enumerate(results):
            tags = ["best"] if price == min_price else []
            tags.append("oddrow" if idx % 2 == 0 else "evenrow")
            tree.insert("", tk.END, values=(cid, code, volt, f"{price}"), tags=tags)


# GUI setup
input_frame = ttk.Labelframe(root, text="Filter Criteria", padding=15)
input_frame.grid(row=0, column=0, padx=15, pady=15, sticky="ew")

# left column
load_type_label = ttk.Label(input_frame, text="Load Type:")
load_type_combo = ttk.Combobox(input_frame,
                               values=["Industrial", "Residential", "Commercial", "Municipal"],
                               state="readonly", width=15)
load_type_combo.current(0)

cable_type_label = ttk.Label(input_frame, text="Cable Type:")
cable_type_combo = ttk.Combobox(input_frame,
                                values=["Single-core", "Three-core"],
                                state="readonly", width=15)
cable_type_combo.current(0)

arrangement_label = ttk.Label(input_frame, text="Arrangement:")
arrangement_combo = ttk.Combobox(input_frame, values=["Flat", "Trefoil"],
                                 state="readonly", width=15)
arrangement_combo.current(0)

# middle column
active_power_label = ttk.Label(input_frame, text="Active Power (MW):")
active_power_var = tk.StringVar(value="0.4")
active_power_spin = ttk.Entry(
    input_frame,
    textvariable=active_power_var,
    validate="key",
    validatecommand=(root.register(_is_positive_float), "%P"),
    width=15
)

reactive_power_label = ttk.Label(input_frame, text="Reactive Power (MVar):")
reactive_power_var = tk.StringVar(value="0.3")
reactive_power_spin = ttk.Entry(
    input_frame,
    textvariable=reactive_power_var,
    validate="key",
    validatecommand=(root.register(_is_positive_float), "%P"),
    width=15
)

voltage_label = ttk.Label(input_frame, text="System Voltage (kV):")
voltage_var = tk.StringVar(value="0.8")
voltage_spin = ttk.Entry(
    input_frame,
    textvariable=voltage_var,
    validate="key",
    validatecommand=(root.register(_is_positive_float), "%P"),
    width=15
)

# right column
circuits_label = ttk.Label(input_frame, text="Parallel Circuits:")
circuits_spin = ttk.Spinbox(
    input_frame, from_=1, to=6, increment=1,
    validate="key",
    validatecommand=(root.register(_is_positive_int), "%P"),
    width=13
)
circuits_spin.set("2")

cable_length_label = ttk.Label(input_frame, text="Cable Length (km):")
cable_length_var = tk.StringVar(value="0.4")
cable_length_spin = ttk.Entry(
    input_frame,
    textvariable=cable_length_var,
    validate="key",
    validatecommand=(root.register(_is_positive_float), "%P"),
    width=15
)

ambient_label = ttk.Label(input_frame, text="Ambient Temp (°C):")
ambient_spin = ttk.Spinbox(
    input_frame, from_=5, to=40, increment=1,
    validate="key",
    validatecommand=(root.register(_is_positive_float), "%P"),
    width=13
)
ambient_spin.set("20")

# layout grid
load_type_label.grid(row=0, column=0, sticky=tk.W, pady=5, padx=(0, 10))
load_type_combo.grid(row=0, column=1, sticky=tk.W, pady=5, padx=(0, 20))
cable_type_label.grid(row=1, column=0, sticky=tk.W, pady=5, padx=(0, 10))
cable_type_combo.grid(row=1, column=1, sticky=tk.W, pady=5, padx=(0, 20))
arrangement_label.grid(row=2, column=0, sticky=tk.W, pady=5, padx=(0, 10))
arrangement_combo.grid(row=2, column=1, sticky=tk.W, pady=5, padx=(0, 20))

active_power_label.grid(row=0, column=2, sticky=tk.W, pady=5, padx=(0, 10))
active_power_spin.grid(row=0, column=3, sticky=tk.W, pady=5, padx=(0, 20))
reactive_power_label.grid(row=1, column=2, sticky=tk.W, pady=5, padx=(0, 10))
reactive_power_spin.grid(row=1, column=3, sticky=tk.W, pady=5, padx=(0, 20))
voltage_label.grid(row=2, column=2, sticky=tk.W, pady=5, padx=(0, 10))
voltage_spin.grid(row=2, column=3, sticky=tk.W, pady=5, padx=(0, 20))

circuits_label.grid(row=0, column=4, sticky=tk.W, pady=5, padx=(0, 10))
circuits_spin.grid(row=0, column=5, sticky=tk.W, pady=5)
cable_length_label.grid(row=1, column=4, sticky=tk.W, pady=5, padx=(0, 10))
cable_length_spin.grid(row=1, column=5, sticky=tk.W, pady=5)
ambient_label.grid(row=2, column=4, sticky=tk.W, pady=5, padx=(0, 10))
ambient_spin.grid(row=2, column=5, sticky=tk.W, pady=5)

input_frame.columnconfigure(1, weight=1, minsize=150)
input_frame.columnconfigure(3, weight=1, minsize=150)
input_frame.columnconfigure(5, weight=1, minsize=150)

# tip
tip_frame = ttk.Frame(root)
tip_frame.grid(row=1, column=0, padx=15, pady=0, sticky="ew")

tip_label = ttk.Label(tip_frame,
                      text="Quick Tip: Select a cable from the list below and press Enter to calculate instantly",
                      font=("TkDefaultFont", 9), foreground="#4a90e2")
tip_label.pack()


def on_cable_type_change(event=None):
    ctype = cable_type_combo.get()
    if ctype == "Single-core":
        arrangement_combo.config(state="readonly")
        circuits_spin.config(from_=1, to=2)
        if int(circuits_spin.get()) > 2:
            circuits_spin.set("2")
    else:
        arrangement_combo.config(state="disabled")
        circuits_spin.config(from_=1, to=6)
        if int(circuits_spin.get()) > 6:
            circuits_spin.set("6")
    root.after(100, auto_filter_cables)


cable_type_combo.bind("<<ComboboxSelected>>", on_cable_type_change)

# buttons
button_frame = ttk.Frame(root, padding=(15, 0))
button_frame.grid(row=5, column=0, pady=(0, 15), sticky="ew")

help_button = ttk.Button(
    button_frame,
    text="? Help",
    command=show_help,
    style="info.TButton"
)
help_button.grid(row=0, column=0, pady=5, padx=(0, 10))

theme_button = ttk.Button(
    button_frame,
    text="Change Theme",
    command=toggle_theme,
    style="secondary.TButton"
)
theme_button.grid(row=0, column=3, sticky="e", padx=5, pady=5)

button_frame.columnconfigure(2, weight=1)

# results tree
result_frame = ttk.Frame(root, padding=15)
result_frame.grid(row=2, column=0, padx=15, pady=(0, 15), sticky="nsew")
root.rowconfigure(2, weight=1)

columns = ("ID", "Code", "Voltage", "Price (TL/km)")
tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=6)

normal_fg = style.colors.fg
tree.tag_configure("oddrow", background="#2a2a2a", foreground="#ffffff")
tree.tag_configure("evenrow", background="#3a3a3a", foreground="#ffffff")
tree.tag_configure("best", background=style.colors.success, foreground="white")

for col in columns:
    tree.heading(col, text=col,
                 command=lambda c=col: sort_tree(c, False))

tree.column("ID", anchor="center", width=80)
tree.column("Code", anchor="center", width=150)
tree.column("Voltage", anchor="center", width=120)
tree.column("Price (TL/km)", anchor="center", width=150)

style.configure("Treeview", rowheight=20)

tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=tree.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
tree.configure(yscrollcommand=scrollbar.set)

# cable selection frames
cable_capacity_container = ttk.Frame(root)
cable_capacity_container.grid(row=3, column=0, padx=15, pady=(0, 15), sticky="ew")
root.columnconfigure(0, weight=1)

# left side
selection_frame = ttk.Labelframe(cable_capacity_container, text="Selected Cable: None", padding=10)
selection_frame.grid(row=0, column=0, sticky="ew", padx=(0, 8))

selected_cable = {"data": None}
last_calculation_log = {"content": None}


def show_detailed_log():
    if last_calculation_log["content"] is None:
        messagebox.showinfo("No Log", "Please calculate losses and regulation first.")
        return

    show_log_window(last_calculation_log["content"])


resistance_label = ttk.Label(selection_frame, text="Resistance: - Ω/km")
resistance_label.grid(row=0, column=0, sticky=tk.W, pady=1)

inductance_label = ttk.Label(selection_frame, text="Inductance: - mH/km")
inductance_label.grid(row=1, column=0, sticky=tk.W, pady=1)

capacitance_label = ttk.Label(selection_frame, text="Capacitance: - μF/km")
capacitance_label.grid(row=2, column=0, sticky=tk.W, pady=1)

# right side
capacity_main_frame = ttk.Labelframe(cable_capacity_container, text="Capacity Check", padding=10)
capacity_main_frame.grid(row=0, column=1, sticky="ew", padx=(8, 0))

base_capacity_label = ttk.Label(capacity_main_frame, text="Base Capacity: - A")
base_capacity_label.grid(row=0, column=0, sticky=tk.W, pady=1, padx=(0, 15))

current_per_circuit_label = ttk.Label(capacity_main_frame, text="Current per Circuit: - A")
current_per_circuit_label.grid(row=1, column=0, sticky=tk.W, pady=1, padx=(0, 15))

derated_capacity_label = ttk.Label(capacity_main_frame, text="Derated Capacity: - A")
derated_capacity_label.grid(row=0, column=1, sticky=tk.W, pady=1)

safety_margin_label = ttk.Label(capacity_main_frame, text="Safety Margin: - %")
safety_margin_label.grid(row=1, column=1, sticky=tk.W, pady=1)

capacity_status_label = ttk.Label(capacity_main_frame, text="Status: -", font=("TkDefaultFont", 10, "bold"))
capacity_status_label.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=1)

cable_capacity_container.columnconfigure(0, weight=1)
cable_capacity_container.columnconfigure(1, weight=1)

capacity_main_frame.columnconfigure(0, weight=1)
capacity_main_frame.columnconfigure(1, weight=1)

capacity_main_frame.grid_remove()

# more buttons
calc_button = ttk.Button(button_frame, text="Calculate Losses & Regulation", state="disabled", style="success.TButton")
calc_button.grid(row=0, column=1, pady=5, padx=(0, 5))

view_log_button = ttk.Button(button_frame, text="View Log", command=show_detailed_log, style="info.TButton")
view_log_button.grid(row=0, column=2, pady=5, padx=(0, 10))

# summary frame
summary_frame = ttk.Labelframe(root, text="    Key Results Summary", padding=10)
summary_frame.grid(row=4, column=0, padx=15, pady=(0, 15), sticky="ew")

voltage_reg_head = ttk.Label(summary_frame, text="Voltage Regulation:", font=("TkDefaultFont", 10, "bold"),
                             foreground="#ff8c00")
voltage_reg_head.grid(row=0, column=0, sticky=tk.W, padx=(10, 5), pady=5)
voltage_reg_value = ttk.Label(summary_frame, text="-")
voltage_reg_value.grid(row=0, column=1, sticky=tk.W, padx=(0, 5), pady=5)

energy_cost_head = ttk.Label(summary_frame, text="10-Year Energy Loss Cost:", font=("TkDefaultFont", 10, "bold"),
                             foreground="#ff8c00")
energy_cost_head.grid(row=0, column=2, sticky=tk.W, padx=(20, 5), pady=5)
energy_cost_value = ttk.Label(summary_frame, text="-")
energy_cost_value.grid(row=0, column=3, sticky=tk.W, padx=(0, 5), pady=5)

cable_cost_head = ttk.Label(summary_frame, text="Cable Installation Cost:", font=("TkDefaultFont", 10, "bold"),
                            foreground="#ff8c00")
cable_cost_head.grid(row=1, column=0, sticky=tk.W, padx=(10, 5), pady=5)
cable_cost_value = ttk.Label(summary_frame, text="-")
cable_cost_value.grid(row=1, column=1, sticky=tk.W, padx=(0, 5), pady=5)

total_cost_head = ttk.Label(summary_frame, text="Total 10-Year Cost:", font=("TkDefaultFont", 11, "bold"),
                            foreground="#ff8c00")
total_cost_head.grid(row=1, column=2, sticky=tk.W, padx=(20, 5), pady=5)
total_cost_value = ttk.Label(summary_frame, text="-", font=("TkDefaultFont", 11))
total_cost_value.grid(row=1, column=3, sticky=tk.W, padx=(0, 5), pady=5)

summary_frame.columnconfigure(0, weight=0, minsize=150)
summary_frame.columnconfigure(1, weight=1, minsize=120)
summary_frame.columnconfigure(2, weight=0, minsize=150)
summary_frame.columnconfigure(3, weight=1, minsize=120)

summary_frame.grid_remove()


def show_log_window(log_content):
    log_window = tk.Toplevel(root)
    log_window.title("Detailed Calculation Log")
    log_window.geometry("650x600")
    log_window.transient(root)
    log_window.grab_set()

    log_window.update_idletasks()
    x = (log_window.winfo_screenwidth() // 2) - (650 // 2)
    y = (log_window.winfo_screenheight() // 2) - (600 // 2)
    log_window.geometry(f"650x600+{x}+{y}")

    log_main_frame = ttk.Frame(log_window, padding=20)
    log_main_frame.pack(fill=tk.BOTH, expand=True)

    log_text_frame = ttk.Frame(log_main_frame)
    log_text_frame.pack(fill=tk.BOTH, expand=True)

    log_text_widget = tk.Text(log_text_frame, wrap=tk.WORD, font=("Consolas", 9),
                              background="#1a1a1a", foreground="#ffffff",
                              insertbackground="#ffffff", padx=15, pady=15)
    log_scrollbar = ttk.Scrollbar(log_text_frame, orient=tk.VERTICAL, command=log_text_widget.yview)
    log_text_widget.configure(yscrollcommand=log_scrollbar.set)

    log_text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    log_text_widget.insert("1.0", log_content)
    log_text_widget.configure(state="disabled")

    button_frame = ttk.Frame(log_main_frame)
    button_frame.pack(fill=tk.X, pady=(10, 0))

    def copy_log():
        try:
            root.clipboard_clear()
            root.clipboard_append(log_content)
            root.update()

            copy_btn.config(text="✓ Copied!", state="disabled")
            log_window.after(2000, lambda: copy_btn.config(text="Copy", state="normal"))

        except Exception as e:
            messagebox.showerror("Copy Error", f"Failed to copy log: {str(e)}")

    copy_btn = ttk.Button(button_frame, text="Copy", command=copy_log, style="success.TButton")
    copy_btn.pack(side=tk.LEFT, padx=(0, 10))

    close_btn = ttk.Button(button_frame, text="Close", command=log_window.destroy, style="secondary.TButton")
    close_btn.pack(side=tk.LEFT)


def sort_tree(col, descending=False):
    rows = [(tree.set(iid, col), iid) for iid in tree.get_children('')]
    if col in ("Capacity (A)", "Price (TL/km)"):
        rows = [(float(txt.replace(',', '')), iid) for txt, iid in rows]
    elif col == "Voltage":
        parsed = []
        for txt, iid in rows:
            try:
                parsed.append((float(txt.split('/')[1].replace('kV', '').strip()), iid))
            except:
                parsed.append((0.0, iid))
        rows = parsed
    rows.sort(reverse=descending)
    for index, (_, iid) in enumerate(rows):
        tree.move(iid, '', index)

    arrow = "▼" if descending else "▲"
    for c in tree["columns"]:
        lbl = c + (f" {arrow}" if c == col else "")
        tree.heading(c, text=lbl,
                     command=lambda c=c, d=(not descending if c == col else False): sort_tree(c, d))


def on_cable_select(event):
    selection = tree.selection()
    if not selection:
        selected_cable["data"] = None
        update_cable_display()
        return

    try:
        cable_id = int(tree.set(selection[0], "ID"))
        cable = next((c for c in cable_list if c["id"] == cable_id), None)
        if cable:
            selected_cable["data"] = cable
            update_cable_display()
        else:
            selected_cable["data"] = None
            update_cable_display()
    except (ValueError, tk.TclError):
        selected_cable["data"] = None
        update_cable_display()


def update_cable_display():
    cable = selected_cable["data"]
    if not cable:
        selection_frame.config(text="Selected Cable: None")
        resistance_label.config(text="Resistance: - Ω/km")
        inductance_label.config(text="Inductance: - mH/km")
        capacitance_label.config(text="Capacitance: - μF/km")
        base_capacity_label.config(text="Base Capacity: - A")
        derated_capacity_label.config(text="Derated Capacity: - A")
        current_per_circuit_label.config(text="Current per Circuit: - A")
        safety_margin_label.config(text="Safety Margin: - %")
        capacity_status_label.config(text="Status: -", foreground="white")
        capacity_main_frame.config(text="Capacity Check")
        try:
            capacity_main_frame.configure(foreground="white")
        except:
            pass
        capacity_main_frame.grid_remove()
        calc_button.config(state="disabled")
        return

    selection_frame.config(text=f"Selected Cable: {cable['code']} - {cable['voltage']}")
    try:
        selection_frame.configure(foreground="#ff8c00")
    except:
        pass

    capacity_main_frame.grid()

    resistance_label.config(text=f"Resistance: {cable['resistance']} Ω/km")

    # get inductance
    arrangement = arrangement_combo.get()
    if cable['code'].startswith('1x'):  # single core
        if arrangement == "Flat" and cable['inductance_flat'] is not None:
            inductance_value = cable['inductance_flat']
        elif cable['inductance_trefoil'] is not None:
            inductance_value = cable['inductance_trefoil']
        else:
            inductance_value = "N/A"
    else:  # three core
        inductance_value = cable['inductance_trefoil'] if cable['inductance_trefoil'] is not None else "N/A"

    inductance_label.config(text=f"Inductance: {inductance_value} mH/km")

    cap_value = cable['capacitance'] if cable['capacitance'] is not None else "N/A"
    capacitance_label.config(text=f"Capacitance: {cap_value} μF/km")

    # capacity calculations
    if cable['code'].startswith('1x'):
        base_capacity = cable['fcc'] if arrangement == "Flat" else cable['tcc']
    else:
        base_capacity = cable['tcc']

    base_capacity_label.config(text=f"Base Capacity: {base_capacity} A")

    try:
        P = float(active_power_var.get()) or 0.0
        Q = float(reactive_power_var.get()) or 0.0
        V = float(voltage_var.get()) or 0.0
        N = int(circuits_spin.get()) or 1
        ambient = float(ambient_spin.get()) or 20.0

        # calc current per circuit
        S_MVA = math.sqrt(P ** 2 + Q ** 2)
        I_total = S_MVA * 1e6 / (math.sqrt(3) * V * 1e3) if V > 0 else 0
        I_per_circuit = I_total / N

        # derating
        temp_factor = get_temp_factor(ambient)
        cables_in_trench = N * 3 if cable['code'].startswith('1x') else N
        trench_factor = get_trench_factor(min(cables_in_trench, 6))
        derated_capacity = base_capacity * temp_factor * trench_factor

        # safety margin
        safety_margin = ((derated_capacity - I_per_circuit) / derated_capacity * 100) if derated_capacity > 0 else 0

        derated_capacity_label.config(text=f"Derated Capacity: {derated_capacity:.1f} A")
        current_per_circuit_label.config(text=f"Current per Circuit: {I_per_circuit:.1f} A")
        safety_margin_label.config(text=f"Safety Margin: {safety_margin:.1f}%")

        if I_per_circuit <= derated_capacity:
            capacity_status_label.config(text="Status: Valid", foreground="#00ff00")
        else:
            capacity_status_label.config(text="Status: Invalid", foreground="#ff0000")

    except:
        derated_capacity_label.config(text="Derated Capacity: - A")
        current_per_circuit_label.config(text="Current per Circuit: - A")
        safety_margin_label.config(text="Safety Margin: - %")
        capacity_status_label.config(text="Status: -", foreground="white")

    calc_button.config(state="normal")


def on_arrangement_change(event=None):
    if selected_cable["data"]:
        update_cable_display()
    root.after(100, auto_filter_cables)


def on_input_change(*args):
    if selected_cable["data"]:
        update_cable_display()
    root.after(500, auto_filter_cables)


arrangement_combo.bind("<<ComboboxSelected>>", on_arrangement_change)
tree.bind("<<TreeviewSelect>>", on_cable_select)

# bind events
active_power_var.trace_add('write', on_input_change)
reactive_power_var.trace_add('write', on_input_change)
voltage_var.trace_add('write', on_input_change)
cable_length_var.trace_add('write', on_input_change)


def on_spinbox_change(event=None):
    on_input_change()


circuits_spin.bind('<KeyRelease>', on_spinbox_change)
circuits_spin.bind('<<Increment>>', on_spinbox_change)
circuits_spin.bind('<<Decrement>>', on_spinbox_change)
ambient_spin.bind('<KeyRelease>', on_spinbox_change)
ambient_spin.bind('<<Increment>>', on_spinbox_change)
ambient_spin.bind('<<Decrement>>', on_spinbox_change)


def calculate_losses_and_regulation():
    if not selected_cable["data"]:
        messagebox.showwarning("No Selection", "Please select a cable first.")
        return

    try:
        length = float(cable_length_var.get()) or 1.0
        P = float(active_power_var.get()) or 0.0  # MW
        Q = float(reactive_power_var.get()) or 0.0  # MVar
        V = float(voltage_var.get()) or 0.0  # kV
        N = int(circuits_spin.get()) or 1
        ambient = float(ambient_spin.get()) or 20.0
    except:
        messagebox.showerror("Invalid Input", "Please check your input values.")
        return

    cable = selected_cable["data"]
    load_type = load_type_combo.get()
    arrangement = arrangement_combo.get()

    # main calcs
    S_MVA = math.sqrt(P ** 2 + Q ** 2)
    I_total = S_MVA * 1e6 / (math.sqrt(3) * V * 1e3) if V > 0 else 0
    I_per_circuit = I_total / N

    cos_phi = P / S_MVA if S_MVA > 0 else 1.0
    sin_phi = Q / S_MVA if S_MVA > 0 else 0.0

    # cable params
    R = cable['resistance']  # ohm/km
    if cable['code'].startswith('1x'):
        L = (cable['inductance_flat'] if arrangement == "Flat" and cable['inductance_flat']
             else cable['inductance_trefoil']) or 0.3
    else:
        L = cable['inductance_trefoil'] or 0.3

    X = 2 * math.pi * 50 * L / 1000  # reactance at 50Hz

    # capacity check
    if cable['code'].startswith('1x'):
        base_capacity = cable['fcc'] if arrangement == "Flat" else cable['tcc']
    else:
        base_capacity = cable['tcc']

    temp_factor = get_temp_factor(ambient)
    cables_in_trench = N * 3 if cable['code'].startswith('1x') else N
    trench_factor = get_trench_factor(min(cables_in_trench, 6))
    derated_capacity = base_capacity * temp_factor * trench_factor

    capacity_check = "PASS" if I_per_circuit <= derated_capacity else "FAIL"
    capacity_margin = ((derated_capacity - I_per_circuit) / derated_capacity * 100) if derated_capacity > 0 else 0

    # power losses
    I_per_circuit_A = I_per_circuit

    if cable['code'].startswith('1x'):  # single core
        P_loss_per_circuit_kW = 3 * (I_per_circuit_A ** 2) * R * length / 1000
        Q_loss_per_circuit_kVar = 3 * (I_per_circuit_A ** 2) * X * length / 1000
    else:  # three core
        P_loss_per_circuit_kW = 3 * (I_per_circuit_A ** 2) * R * length / 1000
        Q_loss_per_circuit_kVar = 3 * (I_per_circuit_A ** 2) * X * length / 1000

    # total losses
    P_loss_total_kW = P_loss_per_circuit_kW * N
    Q_loss_total_kVar = Q_loss_per_circuit_kVar * N

    P_loss_MW = P_loss_total_kW / 1000
    Q_loss_MVar = Q_loss_total_kVar / 1000

    # voltage regulation
    VLN_volts = V * 1000 / math.sqrt(3)

    voltage_regulation_percent = (I_per_circuit * R * cos_phi + I_per_circuit * X * sin_phi) * length / VLN_volts * 100

    voltage_drop_volts = (I_per_circuit * R * cos_phi + I_per_circuit * X * sin_phi) * length

    terminal_voltage_ll = V * 1000 - voltage_drop_volts * math.sqrt(3)

    theta = math.acos(cos_phi)

    # economics
    electricity_price = 2500  # TL/MWh
    hours_per_day = {"Industrial": 10, "Residential": 5, "Municipal": 12, "Commercial": 8}
    daily_hours = hours_per_day.get(load_type, 8)
    annual_hours = daily_hours * 365

    annual_energy_loss_MWh = P_loss_MW * annual_hours
    energy_loss_cost_10yr = annual_energy_loss_MWh * 10 * electricity_price

    # cable cost
    cable_cost_per_km = cable['price']
    if cable['code'].startswith('1x'):  # single core: need 3 cables per circuit
        total_cable_length = length * N * 3
    else:  # three core: 1 cable per circuit
        total_cable_length = length * N

    cable_installation_cost = cable_cost_per_km * total_cable_length
    total_cost_10yr = energy_loss_cost_10yr + cable_installation_cost

    # update display
    voltage_reg_value.config(text=f"{voltage_regulation_percent:.3f}% ({voltage_drop_volts:.1f}V drop)")
    energy_cost_value.config(text=f"{energy_loss_cost_10yr:,.0f} TL")
    cable_cost_value.config(text=f"{cable_installation_cost:,.0f} TL")
    total_cost_value.config(text=f"{total_cost_10yr:,.0f} TL")

    summary_frame.grid()

    # cost breakdown
    cable_percent = cable_installation_cost / total_cost_10yr * 100 if total_cost_10yr > 0 else 0
    energy_percent = energy_loss_cost_10yr / total_cost_10yr * 100 if total_cost_10yr > 0 else 0

    # detailed log
    result_output = f"""
════════════════════════════
                        CABLE ANALYSIS RESULTS                              
════════════════════════════

SELECTED CABLE: {cable['code']} - {cable['voltage']}
Cable Length: {length} km | Parallel Circuits: {N} | Arrangement: {arrangement} | Load Type: {load_type}

ELECTRICAL PARAMETERS:
├─ Resistance (R): {R} Ω/km
├─ Inductance (L): {L} mH/km  
└─ Reactance (X): {X:.4f} Ω/km (at 50Hz)

LOAD CONDITIONS:
├─ Active Power (P): {P} MW
├─ Reactive Power (Q): {Q} MVar
├─ Apparent Power: {S_MVA:.3f} MVA
├─ Power Factor: {cos_phi:.3f}
├─ Total Current: {I_total:.1f} A
└─ Current per Circuit: {I_per_circuit:.1f} A

LINE LOSSES:
├─ Active Power Loss: {P_loss_total_kW:.3f} kW ({P_loss_MW / P * 100:.2f}% of load)
└─ Reactive Power Loss: {Q_loss_total_kVar:.3f} kVar ({Q_loss_MVar / Q * 100:.2f}% of load)

CAPACITY CHECK:
├─ Base Capacity: {base_capacity} A
├─ Temperature Factor: {temp_factor:.2f}
├─ Trench Factor: {trench_factor:.2f}
├─ Derated Capacity: {derated_capacity:.1f} A
├─ Current per Circuit: {I_per_circuit:.1f} A
├─ Status: {capacity_check}
└─ Safety Margin: {capacity_margin:.1f}%

VOLTAGE REGULATION (Formula: VR = (I×R×cos φ + I×X×sin φ)×L / VLN×100%):
├─ Power Factor Angle (θ): {math.degrees(theta):.2f}°
├─ Voltage Drop (L-N): {voltage_drop_volts:.1f} V
├─ Voltage Regulation: {voltage_regulation_percent:.3f}%
└─ Terminal Voltage (L-L): {terminal_voltage_ll:.1f} V

ECONOMIC ANALYSIS (10-Year Period):
├─ Operating Hours: {daily_hours} hours/day ({annual_hours} hours/year)
├─ Annual Energy Loss: {annual_energy_loss_MWh:.2f} MWh
├─ Electricity Price: {electricity_price} TL/MWh
├─ 10-Year Energy Loss Cost: {energy_loss_cost_10yr:,.0f} TL
├─ Cable Installation Cost: {cable_installation_cost:,.0f} TL
└─ TOTAL 10-YEAR COST: {total_cost_10yr:,.0f} TL

COST BREAKDOWN:
├─ Cable Cost: {cable_percent:.1f}%
└─ Energy Loss Cost: {energy_percent:.1f}%

════════════════════════════

**Generated by Smart Cable Selector - Ata Turk - 2025**
"""

    last_calculation_log["content"] = result_output


calc_button.config(command=calculate_losses_and_regulation)


def on_enter_key(event=None):
    if selected_cable["data"] is not None:
        calculate_losses_and_regulation()
    return "break"


def on_f1_key(event=None):
    show_help()
    return "break"


# key bindings
root.bind('<Return>', on_enter_key)
root.bind('<KP_Enter>', on_enter_key)
root.bind('<F1>', on_f1_key)

root.focus_set()


def initialize_app():
    on_cable_type_change()
    auto_filter_cables()


root.after(1000, initialize_app)

root.mainloop()