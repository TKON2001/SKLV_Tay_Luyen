import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import threading
import time
import pyautogui
import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
import pygetwindow as gw
import keyboard
import json
import re
import os
import shutil
import sys
import unicodedata
import colorsys
import math

try:
    from pyautogui import FailSafeException  # type: ignore
except Exception:  # pragma: no cover - fallback khi không có pyautogui thật
    class FailSafeException(Exception):
        """Stub FailSafeException khi pyautogui không sẵn có."""

try:
    import numpy as np
except Exception:
    np = None

# --- CẤU HÌNH QUAN TRỌNG ---
# Nếu bạn không thêm Tesseract vào PATH khi cài đặt, hãy đảm bảo thiết lập đúng đường dẫn.

_DEFAULT_TESSERACT_PATHS = [
    r'F:\\Tesseract-OCR\\tesseract.exe',
    r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe',
    r'C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe',
]


def _try_show_messagebox(title: str, message: str) -> None:
    """Hiển thị messagebox an toàn ngay cả khi chưa tạo Tk root."""

    try:
        tmp_root = tk.Tk()
        tmp_root.withdraw()
        messagebox.showerror(title, message)
        tmp_root.destroy()
    except Exception:
        # Nếu môi trường không hỗ trợ GUI (ví dụ chạy test), in ra stderr.
        print(f"{title}: {message}", file=sys.stderr)


def _ensure_tesseract_available() -> bool:
    """Cố gắng tìm và cấu hình đường dẫn đến Tesseract-OCR."""

    candidates: list[str] = []

    # 1. Ưu tiên các biến môi trường do người dùng chỉ định.
    for env_key in ("TESSERACT_CMD", "TESSERACT_PATH"):
        env_path = os.environ.get(env_key)
        if env_path:
            candidates.append(env_path)

    # 2. Thử tìm trong PATH hiện tại.
    detected_in_path = shutil.which("tesseract")
    if detected_in_path:
        candidates.append(detected_in_path)

    # 3. Thêm các đường dẫn mặc định phổ biến trên Windows.
    candidates.extend(_DEFAULT_TESSERACT_PATHS)

    checked_paths: list[str] = []

    for path in candidates:
        normalized = os.path.normpath(os.path.expandvars(path))
        if not normalized:
            continue
        if not os.path.exists(normalized):
            checked_paths.append(normalized)
            continue

        pytesseract.pytesseract.tesseract_cmd = normalized
        try:
            pytesseract.get_tesseract_version()
            return True
        except pytesseract.TesseractNotFoundError:
            checked_paths.append(normalized)
        except Exception:
            checked_paths.append(normalized)

    # Cuối cùng thử phiên bản mặc định (nếu người dùng đã cấu hình trước đó).
    try:
        pytesseract.get_tesseract_version()
        return True
    except pytesseract.TesseractNotFoundError:
        pass

    # Ghi log debug nếu cần.
    if checked_paths:
        print("Đã kiểm tra các đường dẫn Tesseract nhưng không hợp lệ:", file=sys.stderr)
        for path in checked_paths:
            print(f"  - {path}", file=sys.stderr)

    return False


if not _ensure_tesseract_available():
    _try_show_messagebox(
        "Lỗi",
        "Không tìm thấy Tesseract-OCR. Vui lòng cài đặt và thêm vào PATH hoặc đặt biến môi trường TESSERACT_CMD.",
    )
    sys.exit(1)

# --- Lớp ứng dụng chính ---
class AutoRefineApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Auto Tẩy Luyện Tool v1.0")
        self.root.geometry("900x700")
        self.root.resizable(True, True)

        self.is_running = False
        self.automation_thread = None
        self.game_window = None
        self.config = {
            "refine_button": [0, 0],
            "stats": [
                {
                    "name": "Chỉ số 1",
                    "area": [0, 0, 0, 0],
                    "lock_button": [0, 0],
                    "desired_value": 0,
                    "lock_ocr_area": [0, 0, 0, 0],
                    "lock_unchecked_keyword": "",
                    "lock_checked_keyword": "",
                },
                {
                    "name": "Chỉ số 2",
                    "area": [0, 0, 0, 0],
                    "lock_button": [0, 0],
                    "desired_value": 0,
                    "lock_ocr_area": [0, 0, 0, 0],
                    "lock_unchecked_keyword": "",
                    "lock_checked_keyword": "",
                },
                {
                    "name": "Chỉ số 3",
                    "area": [0, 0, 0, 0],
                    "lock_button": [0, 0],
                    "desired_value": 0,
                    "lock_ocr_area": [0, 0, 0, 0],
                    "lock_unchecked_keyword": "",
                    "lock_checked_keyword": "",
                },
                {
                    "name": "Chỉ số 4",
                    "area": [0, 0, 0, 0],
                    "lock_button": [0, 0],
                    "desired_value": 0,
                    "lock_ocr_area": [0, 0, 0, 0],
                    "lock_unchecked_keyword": "",
                    "lock_checked_keyword": "",
                },
            ],
            "upgrade_area": [0, 0, 0, 0],
            "upgrade_button": [0, 0],
            "require_red": False,
            "lock_templates": {
                "checked": "lock_checked.png",
                "unchecked": "lock_unchecked.png"
            }
        }
        self.locked_stats = [False] * 4
        self.pending_upgrade = False
        self.config_file = "config_tay_luyen.json"
        self.require_red_var = tk.BooleanVar(value=False)
        self._tpl_checked = None
        self._tpl_unchecked = None

        # --- Tạo giao diện ---
        self.create_widgets()
        self.load_config()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        # Tải template nếu có
        try:
            self._load_lock_templates()
        except Exception as _e:
            self.log(f"⚠️ Không thể tải template lock: {_e}")

    def create_widgets(self):
        # Tạo scrollable frame
        canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        main_frame = ttk.Frame(scrollable_frame, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. Khung chọn cửa sổ
        window_frame = ttk.LabelFrame(main_frame, text="1. Chọn Cửa Sổ Game", padding="10")
        window_frame.pack(fill=tk.X, pady=5)
        
        self.window_label = ttk.Label(window_frame, text="Chưa chọn cửa sổ nào", width=50)
        self.window_label.pack(side=tk.LEFT, padx=5)
        ttk.Button(window_frame, text="Chọn Cửa Sổ", command=self.select_game_window).pack(side=tk.LEFT)

        # 2. Khung thiết lập tọa độ
        coords_frame = ttk.LabelFrame(main_frame, text="2. Thiết Lập Tọa Độ và Chỉ Số Mong Muốn", padding="10")
        coords_frame.pack(fill=tk.X, pady=5)
        coords_frame.columnconfigure(1, weight=1)
        coords_frame.columnconfigure(2, weight=1)

        # Nút Tẩy Luyện
        ttk.Label(coords_frame, text="Nút Tẩy Luyện:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.refine_btn_label = ttk.Label(coords_frame, text="Chưa thiết lập")
        self.refine_btn_label.grid(row=0, column=1, sticky=tk.W)
        ttk.Button(coords_frame, text="Thiết lập", command=lambda: self.setup_coord("refine_button")).grid(row=0, column=2, sticky=tk.W, padx=6)

        row_idx = 1

        # Các chỉ số
        stats_container = ttk.Frame(coords_frame)
        stats_container.grid(row=row_idx, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        stats_container.columnconfigure(0, weight=1)
        row_idx += 1

        self.stat_entries = []
        rows_per_stat = 5
        for i in range(4):
            stat_frame = ttk.LabelFrame(stats_container, text=f"Chỉ số {i+1}")
            stat_frame.pack(fill=tk.X, pady=4)
            for col in range(6):
                stat_frame.columnconfigure(col, weight=1 if col in (1, 3, 5) else 0)

            ttk.Label(stat_frame, text="Mục tiêu:").grid(row=0, column=0, sticky=tk.W, padx=4, pady=2)
            desired_val_entry = ttk.Entry(stat_frame, width=10)
            desired_val_entry.grid(row=0, column=1, sticky=tk.W, padx=2, pady=2)

            current_label = ttk.Label(stat_frame, text="Giá trị hiện tại: --")
            current_label.grid(row=0, column=2, columnspan=2, sticky=tk.W, padx=4, pady=2)

            lock_status_label = ttk.Label(stat_frame, text="Trạng thái khóa: --")
            lock_status_label.grid(row=0, column=4, columnspan=2, sticky=tk.W, padx=4, pady=2)

            ttk.Label(stat_frame, text="Vùng đọc:").grid(row=1, column=0, sticky=tk.W, padx=4, pady=2)
            area_label = ttk.Label(stat_frame, text="Chưa đặt")
            area_label.grid(row=1, column=1, columnspan=3, sticky=tk.W, padx=2, pady=2)
            ttk.Button(stat_frame, text="Đặt vùng", command=lambda idx=i: self.setup_coord("stat_area", idx)).grid(row=1, column=4, sticky=tk.W, padx=4, pady=2)

            ttk.Label(stat_frame, text="Nút khóa:").grid(row=2, column=0, sticky=tk.W, padx=4, pady=2)
            lock_label = ttk.Label(stat_frame, text="Chưa đặt")
            lock_label.grid(row=2, column=1, columnspan=3, sticky=tk.W, padx=2, pady=2)
            ttk.Button(stat_frame, text="Đặt nút", command=lambda idx=i: self.setup_coord("stat_lock", idx)).grid(row=2, column=4, sticky=tk.W, padx=4, pady=2)

            ttk.Label(stat_frame, text="Vùng xác nhận:").grid(row=3, column=0, sticky=tk.W, padx=4, pady=2)
            lock_ocr_label = ttk.Label(stat_frame, text="Chưa đặt")
            lock_ocr_label.grid(row=3, column=1, columnspan=3, sticky=tk.W, padx=2, pady=2)
            ttk.Button(stat_frame, text="Đặt vùng", command=lambda idx=i: self.setup_coord("stat_lock_ocr", idx)).grid(row=3, column=4, sticky=tk.W, padx=4, pady=2)

            keyword_frame = ttk.Frame(stat_frame)
            keyword_frame.grid(row=4, column=0, columnspan=6, sticky="ew", padx=2, pady=(4, 2))
            for col in range(6):
                weight = 1 if col in (1, 4) else 0
                keyword_frame.columnconfigure(col, weight=weight)

            ttk.Label(keyword_frame, text="Bỏ tích:").grid(row=0, column=0, sticky=tk.W, padx=4, pady=2)
            lock_unchecked_entry = ttk.Entry(keyword_frame, width=18)
            lock_unchecked_entry.grid(row=0, column=1, sticky=tk.EW, padx=2, pady=2)
            ttk.Button(keyword_frame, text="Chụp", command=lambda idx=i: self.capture_lock_keyword(idx, checked=False)).grid(row=0, column=2, sticky=tk.W, padx=4, pady=2)

            ttk.Label(keyword_frame, text="Đã khóa:").grid(row=0, column=3, sticky=tk.W, padx=(12, 4), pady=2)
            lock_checked_entry = ttk.Entry(keyword_frame, width=18)
            lock_checked_entry.grid(row=0, column=4, sticky=tk.EW, padx=2, pady=2)
            ttk.Button(keyword_frame, text="Chụp", command=lambda idx=i: self.capture_lock_keyword(idx, checked=True)).grid(row=0, column=5, sticky=tk.W, padx=4, pady=2)

            self.stat_entries.append({
                "desired_value": desired_val_entry,
                "area_label": area_label,
                "lock_label": lock_label,
                "lock_ocr_label": lock_ocr_label,
                "lock_unchecked_entry": lock_unchecked_entry,
                "lock_checked_entry": lock_checked_entry,
                "current_label": current_label,
                "lock_status_label": lock_status_label,
            })

        ttk.Separator(coords_frame, orient=tk.HORIZONTAL).grid(row=row_idx, column=0, columnspan=3, sticky="ew", pady=8)
        row_idx += 1

        ttk.Label(coords_frame, text="Vùng nút Thăng Cấp:").grid(row=row_idx, column=0, sticky=tk.W, pady=2)
        self.upgrade_area_label = ttk.Label(coords_frame, text="Chưa đặt")
        self.upgrade_area_label.grid(row=row_idx, column=1, sticky=tk.W)
        ttk.Button(coords_frame, text="Đặt vùng", command=lambda: self.setup_coord("upgrade_area")).grid(row=row_idx, column=2, sticky=tk.W, padx=6)
        row_idx += 1

        ttk.Label(coords_frame, text="Nút Thăng Cấp:").grid(row=row_idx, column=0, sticky=tk.W, pady=2)
        self.upgrade_btn_label = ttk.Label(coords_frame, text="Chưa đặt")
        self.upgrade_btn_label.grid(row=row_idx, column=1, sticky=tk.W)
        ttk.Button(coords_frame, text="Thiết lập", command=lambda: self.setup_coord("upgrade_button")).grid(row=row_idx, column=2, sticky=tk.W, padx=6)
        row_idx += 1

        # 3. Khung điều khiển
        control_frame = ttk.LabelFrame(main_frame, text="3. Điều Khiển", padding="10")
        control_frame.pack(fill=tk.X, pady=5)
        
        self.start_button = ttk.Button(control_frame, text="Bắt Đầu (F5)", command=self.start_automation)
        self.start_button.pack(side=tk.LEFT, padx=10)
        self.stop_button = ttk.Button(control_frame, text="Dừng Lại (F6)", command=self.stop_automation, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=10)
        
        # Nút test OCR
        ttk.Button(control_frame, text="Test OCR", command=self.test_ocr).pack(side=tk.LEFT, padx=10)
        
        # Checkbox: Bắt buộc chữ đỏ
        ttk.Checkbutton(control_frame, text="Bắt buộc chữ đỏ", variable=self.require_red_var, command=self.save_config).pack(side=tk.LEFT, padx=10)

        # 4. Khung log
        log_frame = ttk.LabelFrame(main_frame, text="Log", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Tạo frame cho log với scrollbar
        log_container = ttk.Frame(log_frame)
        log_container.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = tk.Text(log_container, height=10, state=tk.DISABLED, wrap=tk.WORD)
        log_scrollbar = ttk.Scrollbar(log_container, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Pack canvas và scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Hotkeys
        try:
            keyboard.add_hotkey('f5', self.start_automation)
            keyboard.add_hotkey('f6', self.stop_automation)
        except Exception as e:
            self.log(f"Không thể đăng ký hotkey: {e}")

    def log(self, message):
        def _log():
            self.log_text.config(state=tk.NORMAL)
            timestamp = time.strftime("%H:%M:%S")
            self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)
        self.root.after(0, _log)

    def select_game_window(self):
        windows = [w for w in gw.getAllTitles() if w]
        selected_title = simpledialog.askstring("Chọn cửa sổ", "Nhập một phần tên của cửa sổ game (ví dụ: LDPlayer, BlueStacks, Nox):")
        if not selected_title:
            return

        found_windows = [w for w in windows if selected_title.lower() in w.lower()]
        if not found_windows:
            messagebox.showwarning("Không tìm thấy", "Không tìm thấy cửa sổ nào có tên chứa '{}'".format(selected_title))
            return
        
        try:
            self.game_window = gw.getWindowsWithTitle(found_windows[0])[0]
            self.window_label.config(text=self.game_window.title)
            self.log(f"Đã chọn cửa sổ: {self.game_window.title}")
            self.game_window.activate()
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể chọn cửa sổ: {e}")

    def setup_coord(self, coord_type, index=None):
        if not self.game_window:
            messagebox.showerror("Lỗi", "Vui lòng chọn cửa sổ game trước!")
            return

        self.game_window.activate()
        time.sleep(0.5)

        msg = ""
        if coord_type == "refine_button":
            msg = "Di chuyển chuột đến NÚT TẨY LUYỆN và nhấn F8"
        elif coord_type == "stat_area":
            msg = f"Thiết lập vùng cho Chỉ số {index+1}:\n1. Di chuyển chuột đến GÓC TRÊN-TRÁI của vùng chỉ số và nhấn F8.\n2. Di chuyển chuột đến GÓC DƯỚI-PHẢI và nhấn F8 lần nữa."
        elif coord_type == "stat_lock":
            msg = f"Di chuyển chuột đến NÚT KHÓA của Chỉ số {index+1} và nhấn F8"
        elif coord_type == "stat_lock_ocr":
            msg = (
                f"Thiết lập vùng xác nhận khóa của Chỉ số {index+1}:\n"
                "1. Di chuyển chuột đến GÓC TRÊN-TRÁI của dòng chữ xác nhận và nhấn F8.\n"
                "2. Di chuyển chuột đến GÓC DƯỚI-PHẢI và nhấn F8 lần nữa."
            )
        elif coord_type == "upgrade_area":
            msg = "Thiết lập vùng nhận diện chữ 'Thăng cấp':\n1. Di chuyển chuột đến GÓC TRÊN-TRÁI của vùng nút/chữ và nhấn F8.\n2. Di chuyển chuột đến GÓC DƯỚI-PHẢI và nhấn F8 lần nữa."
        elif coord_type == "upgrade_button":
            msg = "Di chuyển chuột đến NÚT THĂNG CẤP và nhấn F8"

        info_window = tk.Toplevel(self.root)
        info_window.title("Hướng dẫn")
        info_window.geometry("400x150")
        info_window.transient(self.root)
        info_window.grab_set()
        
        tk.Label(info_window, text=msg, padx=20, pady=20, justify=tk.LEFT).pack()
        
        positions = []
        def on_f8(event):
            if event.name == 'f8':
                pos = pyautogui.position()
                positions.append(pos)
                self.log(f"Đã ghi nhận tọa độ: {pos}")
                needs_box = coord_type in {"stat_area", "stat_lock_ocr", "upgrade_area"}
                if (needs_box and len(positions) == 2) or (not needs_box and len(positions) == 1):
                    keyboard.unhook_all()
                    info_window.destroy()

        keyboard.on_press(on_f8)
        self.root.wait_window(info_window)

        if coord_type == "refine_button" and positions:
            self.config["refine_button"] = list(positions[0])
            self.refine_btn_label.config(text=f"X={positions[0][0]}, Y={positions[0][1]}")
            self.save_config()
        elif coord_type == "stat_area" and len(positions) == 2:
            x1, y1 = positions[0]
            x2, y2 = positions[1]
            area = [min(x1, x2), min(y1, y2), abs(x2-x1), abs(y2-y1)]
            self.config["stats"][index]["area"] = area
            self.stat_entries[index]["area_label"].config(text=f"Đã đặt ({area[2]}x{area[3]})")
            self.save_config()
        elif coord_type == "stat_lock" and positions:
            self.config["stats"][index]["lock_button"] = list(positions[0])
            self.stat_entries[index]["lock_label"].config(text=f"X={positions[0][0]}, Y={positions[0][1]}")
            self.save_config()
        elif coord_type == "stat_lock_ocr" and len(positions) == 2:
            x1, y1 = positions[0]
            x2, y2 = positions[1]
            area = [min(x1, x2), min(y1, y2), abs(x2-x1), abs(y2-y1)]
            self.config["stats"][index]["lock_ocr_area"] = area
            self.stat_entries[index]["lock_ocr_label"].config(text=f"Đã đặt ({area[2]}x{area[3]})")
            self.save_config()
        elif coord_type == "upgrade_area" and len(positions) == 2:
            x1, y1 = positions[0]
            x2, y2 = positions[1]
            area = [min(x1, x2), min(y1, y2), abs(x2-x1), abs(y2-y1)]
            self.config["upgrade_area"] = area
            self.upgrade_area_label.config(text=f"Đã đặt ({area[2]}x{area[3]})")
            self.save_config()
        elif coord_type == "upgrade_button" and positions:
            self.config["upgrade_button"] = list(positions[0])
            self.upgrade_btn_label.config(text=f"X={positions[0][0]}, Y={positions[0][1]}")
            self.save_config()

    def _sync_stat_entries_to_config(self, *, strict: bool = False) -> bool:
        """Đồng bộ các ô nhập liệu của chỉ số vào cấu hình nội bộ.

        Khi ``strict`` được bật, nếu người dùng nhập giá trị mong muốn không hợp lệ
        thì hàm sẽ hiển thị thông báo lỗi và trả về ``False`` để caller xử lý.
        """

        for i, stat in enumerate(self.config.get("stats", [])):
            entries = self.stat_entries[i]

            # Đồng bộ desired value
            desired_text = entries["desired_value"].get().strip()
            if desired_text:
                try:
                    stat["desired_value"] = int(desired_text)
                except ValueError:
                    if strict:
                        messagebox.showerror("Lỗi", f"Giá trị mong muốn của Chỉ số {i+1} không hợp lệ!")
                        return False
            else:
                stat["desired_value"] = 0

            # Đồng bộ từ khóa OCR
            stat["lock_unchecked_keyword"] = entries["lock_unchecked_entry"].get().strip()
            stat["lock_checked_keyword"] = entries["lock_checked_entry"].get().strip()

        return True

    @staticmethod
    def _normalize_text(text: str) -> str:
        if not text:
            return ""
        normalized = unicodedata.normalize('NFD', text)
        normalized = ''.join(ch for ch in normalized if unicodedata.category(ch) != 'Mn')
        normalized = normalized.upper()
        normalized = re.sub(r'[^0-9A-Z%+\-]+', ' ', normalized)
        return re.sub(r'\s+', ' ', normalized).strip()

    def capture_lock_keyword(self, stat_index: int, *, checked: bool) -> None:
        if not (0 <= stat_index < len(self.config.get("stats", []))):
            return

        stat_cfg = self.config["stats"][stat_index]
        area = stat_cfg.get("lock_ocr_area", [0, 0, 0, 0])
        if sum(area) == 0:
            messagebox.showerror(
                "Lỗi",
                f"Vui lòng đặt vùng xác nhận khóa cho Chỉ số {stat_index + 1} trước khi lấy OCR!",
            )
            return

        try:
            ax, ay, aw, ah = map(int, area)
            if aw <= 0 or ah <= 0:
                raise ValueError("Kích thước vùng không hợp lệ")
        except Exception:
            messagebox.showerror(
                "Lỗi",
                f"Vùng xác nhận khóa của Chỉ số {stat_index + 1} không hợp lệ, vui lòng đặt lại!",
            )
            return

        try:
            snap = pyautogui.screenshot(region=(ax, ay, aw, ah))
        except Exception as exc:
            self.log(f"❌ Không thể chụp vùng OCR khóa: {exc}")
            messagebox.showerror("Lỗi", f"Không thể chụp vùng OCR khóa: {exc}")
            return

        processed = self.process_image_for_ocr(snap)
        debug_tag = f"lock_{stat_index + 1}_{'checked' if checked else 'unchecked'}"
        try:
            processed.save(f"debug_{debug_tag}.png")
        except Exception:
            pass

        text = self.ocr_read_text(processed, debug_tag=debug_tag)
        cleaned = text.strip()
        normalized = self._normalize_text(cleaned)

        entry_key = "lock_checked_entry" if checked else "lock_unchecked_entry"
        entry = self.stat_entries[stat_index][entry_key]

        if checked:
            if 'V' not in normalized:
                self.log(
                    f"⚠️ Không phát hiện chữ V trong vùng khóa của Chỉ số {stat_index + 1}."
                )
                messagebox.showwarning(
                    "Cảnh báo",
                    "Không phát hiện được chữ V trong vùng đã khóa. Vui lòng đảm bảo ô khóa có dấu V và vùng chụp đủ lớn.",
                )
                return
            entry_value = 'V'
        else:
            if 'V' in normalized:
                self.log(
                    f"⚠️ Vẫn còn chữ V trong vùng bỏ tích của Chỉ số {stat_index + 1}."
                )
                messagebox.showwarning(
                    "Cảnh báo",
                    "Ô khóa vẫn còn chữ V. Vui lòng bỏ tích trước khi chụp mẫu bỏ tích.",
                )
                return
            # Cho phép OCR trống đối với trạng thái bỏ tích
            entry_value = ''
            if not cleaned:
                self.log(
                    f"ℹ️ Vùng bỏ tích của Chỉ số {stat_index + 1} không có chữ – dùng mẫu trống."
                )

        entry.delete(0, tk.END)
        entry.insert(0, entry_value)

        self.save_config()

        state_label = "đã khóa" if checked else "bỏ tích"
        display_text = entry_value or '(trống)'
        self.log(
            f"✅ Đã ghi nhận mẫu OCR {state_label} cho Chỉ số {stat_index + 1}: '{display_text}'"
        )

    def _update_lock_status_label(self, stat_index: int | None, status: bool | None, source: str) -> None:
        if stat_index is None or not (0 <= stat_index < len(self.stat_entries)):
            return

        label = self.stat_entries[stat_index]["lock_status_label"]
        if status is None:
            text = f"Trạng thái khóa: {source}"
        else:
            state_txt = "ĐÃ TÍCH" if status else "CHƯA TÍCH"
            if source:
                text = f"Trạng thái khóa: {state_txt} ({source})"
            else:
                text = f"Trạng thái khóa: {state_txt}"

        try:
            self.root.after(0, lambda txt=text, lbl=label: lbl.config(text=txt))
        except Exception:
            try:
                label.config(text=text)
            except Exception:
                pass

    def test_ocr(self):
        if not self.game_window:
            messagebox.showerror("Lỗi", "Vui lòng chọn cửa sổ game trước!")
            return

        self.log("=== TEST OCR ===")
        for i, stat in enumerate(self.config["stats"]):
            if sum(stat["area"]) == 0:
                self.log(f"Chỉ số {i+1}: Chưa thiết lập vùng đọc")
                continue
                
            try:
                x, y, w, h = stat["area"]
                screenshot = pyautogui.screenshot(region=(x, y, w, h))
                processed_img = self.process_image_for_ocr(screenshot)
                
                # Đọc OCR với cơ chế dự phòng (phóng to/threshold + nhiều cấu hình)
                text = self.ocr_read_text(processed_img, debug_tag=f"stat_{i+1}")
                current_value, range_max, is_percent = self.parse_ocr_result(text)
                if range_max is not None:
                    if is_percent:
                        fmt_current = self.format_percent_value(current_value)
                        fmt_range = self.format_percent_value(range_max)
                        self.log(f"Chỉ số {i+1}: '{text.strip()}' -> {fmt_current}% / MAX {fmt_range}%")
                        self.stat_entries[i]["current_label"].config(text=f"Hiện tại: {fmt_current}% / Max: {fmt_range}%")
                    else:
                        self.log(f"Chỉ số {i+1}: '{text.strip()}' -> {current_value} / MAX {range_max}")
                        self.stat_entries[i]["current_label"].config(text=f"Hiện tại: {current_value} / Max: {range_max}")
                else:
                    if is_percent:
                        fmt_current = self.format_percent_value(current_value)
                        self.log(f"Chỉ số {i+1}: '{text.strip()}' -> {fmt_current}%")
                        self.stat_entries[i]["current_label"].config(text=f"Giá trị hiện tại: {fmt_current}%")
                    else:
                        self.log(f"Chỉ số {i+1}: '{text.strip()}' -> {current_value}")
                        self.stat_entries[i]["current_label"].config(text=f"Giá trị hiện tại: {current_value}")
                
                # Lưu ảnh để debug
                processed_img.save(f"debug_stat_{i+1}.png")
                self.log(f"Đã lưu ảnh debug: debug_stat_{i+1}.png")

            except Exception as e:
                self.log(f"Lỗi khi đọc chỉ số {i+1}: {e}")

            lock_pos = stat.get("lock_button", [0, 0])
            if sum(lock_pos) > 0:
                try:
                    checked = self.is_lock_checked(lock_pos, stat_index=i)
                    state_txt = "ĐÃ TÍCH" if checked else "CHƯA TÍCH"
                    self.log(f"   → Trạng thái khóa {i+1}: {state_txt}")
                except Exception as exc:
                    self.log(f"   ⚠️ Không thể xác định trạng thái khóa {i+1}: {exc}")
                    self._update_lock_status_label(i, None, "Lỗi kiểm tra")
            else:
                self._update_lock_status_label(i, None, "Chưa đặt nút khóa")

    def process_image_for_ocr(self, img):
        # Chuyển sang ảnh xám
        gray = img.convert('L')
        # Tăng độ tương phản
        enhancer = ImageEnhance.Contrast(gray)
        contrast_img = enhancer.enhance(2.0)
        # Tăng độ sáng
        brightness_enhancer = ImageEnhance.Brightness(contrast_img)
        bright_img = brightness_enhancer.enhance(1.2)
        # Áp dụng bộ lọc để làm nét
        sharpened_img = bright_img.filter(ImageFilter.SHARPEN)
        return sharpened_img

    def ocr_read_text(self, base_img, debug_tag: str | None = None) -> str:
        # Tạo các biến thể ảnh để tăng tỷ lệ đọc
        variants = []
        try:
            variants.append(("proc", base_img))
        except Exception:
            pass
        try:
            # Phóng to và nhị phân đơn giản
            upscale = base_img.resize((max(1, base_img.width * 2), max(1, base_img.height * 2)), Image.LANCZOS)
            gray = upscale.convert('L')
            bin_img = gray.point(lambda p: 255 if p > 180 else 0)
            variants.append(("bin2x", bin_img))
        except Exception:
            pass

        configs = [
            r'--oem 1 --psm 6 -c tessedit_char_whitelist=0123456789%+().-',
            r'--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789%+().-'
        ]

        best_text = ""
        for vname, vimg in variants:
            for cfg in configs:
                try:
                    txt = pytesseract.image_to_string(vimg, config=cfg)
                    if txt and txt.strip():
                        if debug_tag and vname.startswith("bin"):
                            try:
                                vimg.save(f"debug_{debug_tag}_{vname}.png")
                            except Exception:
                                pass
                        return txt
                    if len(txt.strip()) > len(best_text.strip()):
                        best_text = txt
                except Exception:
                    continue

        if debug_tag:
            try:
                for vname, vimg in variants:
                    if vname != "proc":
                        vimg.save(f"debug_{debug_tag}_{vname}.png")
            except Exception:
                pass
        return best_text

    def has_red_text(self, img):
        # Kiểm tra pixel đỏ nổi bật, ưu tiên vùng bên trái (nơi chuỗi "+giá trị" hiển thị)
        rgb = img.convert('RGB')
        width, height = rgb.size
        x_start = 0
        x_end = int(width * 0.65)
        pixels = rgb.load()
        red_count = 0
        total_samples = 0
        step_y = max(1, height // 20)
        step_x = max(1, max(1, (x_end - x_start) // 40))
        for y in range(0, height, step_y):
            for x in range(x_start, x_end, step_x):
                r, g, b = pixels[x, y]
                if r > 150 and r - max(g, b) > 55:
                    red_count += 1
                total_samples += 1
        ratio = red_count / max(1, total_samples)
        return ratio > 0.06

    def analyze_upgrade_area(self, *, log: bool = True) -> dict | None:
        """Trả về thống kê màu sắc của vùng Thăng Cấp và vị trí click gợi ý."""

        try:
            if sum(self.config.get("upgrade_area", [0, 0, 0, 0])) == 0:
                return None

            ux, uy, uw, uh = self.config["upgrade_area"]
            shot = pyautogui.screenshot(region=(ux, uy, uw, uh))
            rgb = shot.convert('RGB')
            px = rgb.load()
            width, height = rgb.size

            golden_pixels = 0
            bright_gold_pixels = 0
            red_pixels = 0
            total_pixels = max(1, width * height)

            golden_coords: list[tuple[int, int]] = []
            step_x = max(1, width // 80)
            step_y = max(1, height // 60)

            for y in range(0, height, step_y):
                for x in range(0, width, step_x):
                    r, g, b = px[x, y]

                    if r > 180 and g > 160 and b < 120:
                        golden_pixels += 1
                        if r > 200 and g > 180 and b < 100:
                            bright_gold_pixels += 1
                        golden_coords.append((x, y))

                    if r > 200 and r - max(g, b) > 80:
                        red_pixels += 1

            golden_ratio = golden_pixels / total_pixels
            bright_gold_ratio = bright_gold_pixels / total_pixels
            red_ratio = red_pixels / total_pixels

            has_golden_button = golden_ratio > 0.10 or bright_gold_ratio > 0.05
            has_red_badge = red_ratio > 0.02
            is_active = has_golden_button or has_red_badge

            hotspot = None
            if golden_coords:
                avg_x = sum(x for x, _ in golden_coords) / len(golden_coords)
                avg_y = sum(y for _, y in golden_coords) / len(golden_coords)
                hotspot = (int(ux + avg_x), int(uy + avg_y))
            else:
                hotspot = (int(ux + uw // 2), int(uy + uh // 2))

            if log:
                self.log(
                    f"   DEBUG Upgrade: golden={golden_ratio:.3f}, bright_gold={bright_gold_ratio:.3f}, red={red_ratio:.3f}"
                )
                if is_active:
                    if has_golden_button and has_red_badge:
                        self.log("   ✅ Nút Thăng Cấp: ACTIVE (vàng + badge đỏ)")
                    elif has_golden_button:
                        self.log("   ✅ Nút Thăng Cấp: ACTIVE (nút vàng)")
                    else:
                        self.log("   ✅ Nút Thăng Cấp: ACTIVE (badge đỏ)")
                else:
                    self.log(
                        f"   ❌ Nút Thăng Cấp: INACTIVE (golden={golden_ratio:.3f}, red={red_ratio:.3f})"
                    )

            return {
                "active": is_active,
                "hotspot": hotspot,
                "has_golden": has_golden_button,
                "has_red": has_red_badge,
                "golden_ratio": golden_ratio,
                "red_ratio": red_ratio,
            }

        except Exception as e:
            if log:
                self.log(f"   ❌ Lỗi kiểm tra nút Thăng Cấp: {e}")
            return None

    def is_upgrade_available(self) -> bool:
        info = self.analyze_upgrade_area(log=True)
        return bool(info and info.get("active"))

    def click_upgrade_button(self) -> tuple[bool, tuple[int, int] | None, str]:
        """Cố gắng click nút Thăng Cấp. Trả về (success, vị trí, phương thức)."""

        if sum(self.config.get("upgrade_button", [0, 0])) > 0:
            bx, by = self.config["upgrade_button"]
            try:
                pyautogui.moveTo(bx, by)
                pyautogui.click(bx, by)
                return True, (bx, by), "preset"
            except FailSafeException:
                raise
            except Exception as exc:
                self.log(f"   ⚠️ Lỗi click nút Thăng Cấp preset: {exc}")

        info = self.analyze_upgrade_area(log=False)
        if info and info.get("hotspot"):
            hx, hy = info["hotspot"]
            try:
                pyautogui.moveTo(hx, hy)
                pyautogui.click(hx, hy)
                method = "hotspot" if info.get("active") else "center"
                return True, (hx, hy), method
            except FailSafeException:
                raise
            except Exception as exc:
                self.log(f"   ⚠️ Lỗi click hotspot Thăng Cấp: {exc}")

        return False, None, "none"

    def perform_upgrade_sequence(self) -> bool:
        """Thực hiện chuỗi thao tác thăng cấp và bỏ tích các dòng đã khóa."""

        if sum(self.config.get("upgrade_button", [0, 0])) == 0 and \
           sum(self.config.get("upgrade_area", [0, 0, 0, 0])) == 0:
            self.log("⚠️ Chưa cấu hình nút/vùng Thăng Cấp. Không thể thăng cấp tự động.")
            return False

        max_click_attempts = 4
        for attempt in range(max_click_attempts):
            if not self.is_running:
                return False

            self.log(f"▶️ Thử thăng cấp lần {attempt + 1}/{max_click_attempts}...")
            clicked, pos, method = self.click_upgrade_button()

            if not clicked:
                self.log("   ⚠️ Không xác định được vị trí nút Thăng Cấp. Sẽ thử lại sau 0.7s.")
                time.sleep(0.7)
                continue

            if pos:
                self.log(f"   ✅ Đã click nút Thăng Cấp tại ({pos[0]}, {pos[1]}) [{method}]")

            time.sleep(1.6)

            settle_checks = 0
            info = None
            while settle_checks < 3:
                info = self.analyze_upgrade_area(log=False)
                if not info or not info.get("active"):
                    break
                settle_checks += 1
                self.log("   ⏳ Nút vẫn đang sáng, chờ thêm 0.6s để xác nhận...")
                time.sleep(0.6)

            if settle_checks >= 3 and info and info.get("active"):
                self.log("   ⚠️ Có vẻ thăng cấp chưa thành công, thử click lại.")
                time.sleep(0.6)
                continue

            time.sleep(0.8)

            success_unlock = self.unlock_all_locks(max_attempts=6, force_click=True)
            if not success_unlock:
                if self.brute_force_unlock_locks(cycles=3):
                    time.sleep(0.4)
                    success_unlock = self.unlock_all_locks(max_attempts=4, force_click=True)
            if success_unlock:
                self.locked_stats = [False] * 4
                self.log("✅ Đã thăng cấp thành công và bỏ tích các dòng!")
                return True

            self.log("⚠️ Đã thăng cấp nhưng không bỏ tích hết các dòng, sẽ thử lại.")
            time.sleep(0.8)

        self.log("❌ Thử thăng cấp nhiều lần nhưng chưa thành công hoàn toàn.")
        return False

    def is_lock_checked(self, lock_pos: list[int] | tuple[int, int], *, stat_index: int | None = None) -> bool:
        # Phân tích hình ảnh của ô khóa để xác định trạng thái: tìm dấu tích vàng
        try:
            lx, ly = int(lock_pos[0]), int(lock_pos[1])
        except Exception:
            self._update_lock_status_label(stat_index, None, "Chưa đặt nút khóa")
            return False

        stat_cfg = None
        if stat_index is not None and 0 <= stat_index < len(self.config.get("stats", [])):
            stat_cfg = self.config["stats"][stat_index]
        else:
            for idx, cfg in enumerate(self.config.get("stats", [])):
                pos = cfg.get("lock_button", [0, 0])
                try:
                    if int(pos[0]) == lx and int(pos[1]) == ly:
                        stat_cfg = cfg
                        stat_index = idx
                        break
                except Exception:
                    continue

        # Ưu tiên OCR theo vùng xác nhận nếu người dùng cấu hình
        if stat_cfg:
            area = stat_cfg.get("lock_ocr_area", [0, 0, 0, 0])
            unchecked_kw = self._normalize_text(stat_cfg.get("lock_unchecked_keyword", ""))
            checked_kw = self._normalize_text(stat_cfg.get("lock_checked_keyword", ""))
            if sum(area) > 0 and (unchecked_kw or checked_kw):
                try:
                    ax, ay, aw, ah = map(int, area)
                    if aw > 0 and ah > 0:
                        snap_area = pyautogui.screenshot(region=(ax, ay, aw, ah))
                        processed = self.process_image_for_ocr(snap_area)
                        debug_tag = None
                        if stat_index is not None:
                            debug_tag = f"lock_{stat_index + 1}"
                        raw_text = self.ocr_read_text(processed, debug_tag=debug_tag)
                        norm_text = self._normalize_text(raw_text)
                        if norm_text:
                            label = f"Lock {stat_index + 1}" if stat_index is not None else f"Lock {lock_pos}"
                            self.log(f"   OCR {label}: '{raw_text.strip()}' -> {norm_text}")
                        label_idx = f"khóa {stat_index + 1}" if stat_index is not None else f"khóa {lock_pos}"
                        if unchecked_kw and unchecked_kw in norm_text:
                            self.log(f"   ✅ OCR xác nhận {label_idx}: phát hiện từ khóa bỏ tích '{stat_cfg.get('lock_unchecked_keyword', '')}'")
                            self._update_lock_status_label(stat_index, False, "OCR")
                            return False
                        if checked_kw and checked_kw in norm_text:
                            self.log(f"   🔒 OCR xác nhận {label_idx}: phát hiện từ khóa đã khóa '{stat_cfg.get('lock_checked_keyword', '')}'")
                            self._update_lock_status_label(stat_index, True, "OCR")
                            return True
                        if norm_text:
                            self._update_lock_status_label(stat_index, None, "OCR không khớp")
                except Exception as exc:
                    label_idx = f"khóa {stat_index + 1}" if stat_index is not None else f"khóa {lock_pos}"
                    self.log(f"   ⚠️ OCR {label_idx}: lỗi nhận diện - {exc}")
                    self._update_lock_status_label(stat_index, None, "Lỗi OCR")
            elif sum(area) > 0:
                self._update_lock_status_label(stat_index, None, "Chưa có từ khóa OCR")

        # Vùng chụp đủ lớn để bao phủ hoàn toàn dấu tích vàng
        box_size = 32
        half = box_size // 2
        left = max(0, lx - half)
        top = max(0, ly - half)
        snap = pyautogui.screenshot(region=(left, top, box_size, box_size))

        # Thử nhận diện trực tiếp dấu chữ V
        v_state = self._detect_lock_by_checkmark(snap, stat_index=stat_index, lock_pos=lock_pos)
        if v_state is True:
            label_idx = f"khóa {stat_index + 1}" if stat_index is not None else f"khóa {lock_pos}"
            self.log(f"   🔒 Chữ V xác nhận {label_idx} đang ĐÃ KHÓA")
            self._update_lock_status_label(stat_index, True, "Chữ V")
            return True
        if v_state is False:
            label_idx = f"khóa {stat_index + 1}" if stat_index is not None else f"khóa {lock_pos}"
            self.log(f"   ✅ Chữ V xác nhận {label_idx} đang BỎ TÍCH")
            self._update_lock_status_label(stat_index, False, "Chữ V")
            return False

        # Nếu có template, ưu tiên so khớp mẫu
        try:
            if self._tpl_checked is not None or self._tpl_unchecked is not None:
                sim_checked = self._template_similarity(snap, self._tpl_checked)
                sim_unchecked = self._template_similarity(snap, self._tpl_unchecked)
                # Ngưỡng quyết định bằng similarity
                # checked ~0.65 trở lên và chênh lệch > 0.10 so với unchecked
                if sim_checked >= 0.65 and (sim_checked - max(-1.0, sim_unchecked)) >= 0.10:
                    self.log(f"   TEMPLATE Lock {lock_pos}: sim_checked={sim_checked:.3f}, sim_unchecked={sim_unchecked:.3f} => TÍCH")
                    self._update_lock_status_label(stat_index, True, "Template")
                    return True
                if sim_unchecked >= 0.65 and (sim_unchecked - max(-1.0, sim_checked)) >= 0.10:
                    self.log(f"   TEMPLATE Lock {lock_pos}: sim_checked={sim_checked:.3f}, sim_unchecked={sim_unchecked:.3f} => TRỐNG")
                    self._update_lock_status_label(stat_index, False, "Template")
                    return False
                # Nếu mơ hồ, fallback sang phân tích màu
        except Exception:
            pass

        # Chuyển sang RGB để phân tích màu sắc
        rgb_img = snap.convert('RGB')
        width, height = rgb_img.size
        pixels = rgb_img.load()

        # Đếm pixel vàng (dấu tích)
        yellow_pixels = 0
        bright_yellow_pixels = 0
        hsv_yellow_pixels = 0
        total_pixels = width * height

        for y in range(height):
            for x in range(width):
                r, g, b = pixels[x, y]

                # Kiểm tra màu vàng: R cao, G cao, B thấp
                if r > 185 and g > 175 and b < 110:
                    yellow_pixels += 1
                    # Vàng sáng (dấu tích)
                    if r > 225 and g > 215 and b < 85:
                        bright_yellow_pixels += 1

                # Kiểm tra theo HSV để bao phủ trường hợp màu vàng đậm/nhạt
                h, s, v = colorsys.rgb_to_hsv(r / 255.0, g / 255.0, b / 255.0)
                if 0.12 <= h <= 0.18 and s >= 0.42 and v >= 0.55:
                    hsv_yellow_pixels += 1

        # Tính tỉ lệ pixel vàng
        yellow_ratio = yellow_pixels / total_pixels
        bright_yellow_ratio = bright_yellow_pixels / total_pixels
        hsv_yellow_ratio = hsv_yellow_pixels / total_pixels

        # Debug log để kiểm tra
        self.log(
            "   DEBUG Lock {}: yellow_ratio={:.3f}, bright_yellow_ratio={:.3f}, hsv_yellow_ratio={:.3f}".format(
                lock_pos, yellow_ratio, bright_yellow_ratio, hsv_yellow_ratio
            )
        )

        # Nhận diện dấu tích vàng: yêu cầu chặt chẽ hơn để tránh nhầm nền
        has_checkmark = (
            (bright_yellow_ratio > 0.018 and yellow_ratio > 0.085)
            or (yellow_ratio > 0.12 and hsv_yellow_ratio > 0.060)
        )
        
        status = "TÍCH" if has_checkmark else "TRỐNG"
        self.log(f"   Kết quả Lock {lock_pos}: {status}")
        self._update_lock_status_label(stat_index, has_checkmark, "Màu sắc")

        return has_checkmark

    def _detect_lock_by_checkmark(
        self,
        snap: Image.Image,
        *,
        stat_index: int | None = None,
        lock_pos: tuple[int, int] | list[int] | None = None,
    ) -> bool | None:
        """Cố gắng phát hiện chữ V vàng trong ô khóa.

        Trả về ``True`` nếu chắc chắn có chữ V, ``False`` nếu chắc chắn không có,
        hoặc ``None`` nếu không thể kết luận (để dùng fallback khác).
        """

        try:
            w, h = snap.size
            if w <= 0 or h <= 0:
                return None

            scale = 3 if max(w, h) < 48 else 2
            upscaled = snap.resize((w * scale, h * scale), Image.LANCZOS)
            gray = upscaled.convert("L")
            enhanced = ImageEnhance.Contrast(gray).enhance(3.0)
            sharpened = enhanced.filter(ImageFilter.UnsharpMask(radius=2, percent=170, threshold=3))
            bw = sharpened.point(lambda p: 255 if p > 160 else 0)

            config = "--psm 8 --oem 3 -c tessedit_char_whitelist=Vv"
            raw_text = pytesseract.image_to_string(bw, lang="eng", config=config)
            normalized = self._normalize_text(raw_text)

            label_idx = f"khóa {stat_index + 1}" if stat_index is not None else f"lock {lock_pos}"

            tokens = [tok for tok in normalized.split() if tok]
            if any(tok == "V" for tok in tokens) or normalized == "VV":
                self.log(f"   DEBUG chữ V {label_idx}: phát hiện trực tiếp '{raw_text.strip()}'")
                return True

            total_pixels = bw.width * bw.height
            white_pixels = sum(1 for px in bw.getdata() if px == 255)
            bright_ratio = white_pixels / total_pixels if total_pixels else 0.0

            diag_hits = 0
            diag_total = 0
            tolerance = max(1, int(bw.width * 0.06))
            for y in range(int(bw.height * 0.25), bw.height):
                left_expected = int((y / bw.height) * (bw.width / 2))
                right_expected = bw.width - 1 - left_expected
                for offset in range(-tolerance, tolerance + 1):
                    diag_total += 2
                    lx = left_expected + offset
                    rx = right_expected + offset
                    if 0 <= lx < bw.width and bw.getpixel((lx, y)) == 255:
                        diag_hits += 1
                    if 0 <= rx < bw.width and bw.getpixel((rx, y)) == 255:
                        diag_hits += 1

            diag_score = diag_hits / diag_total if diag_total else 0.0
            self.log(
                f"   DEBUG chữ V {label_idx}: bright_ratio={bright_ratio:.3f}, diag_score={diag_score:.3f}, raw='{raw_text.strip()}'"
            )

            if bright_ratio <= 0.010 and diag_score <= 0.060:
                return False
            if bright_ratio >= 0.045 or diag_score >= 0.180:
                return True
        except Exception as exc:
            self.log(f"   ⚠️ Nhận diện chữ V lỗi: {exc}")

        return None

    def ensure_unchecked(self, lock_pos: list[int] | tuple[int, int], *, force: bool = False, stat_index: int | None = None) -> bool:
        """Đảm bảo ô khóa được bỏ tích.

        Khi ``force`` được bật, hàm sẽ cố gắng click bỏ tích ngay cả khi hệ thống
        nhận diện rằng ô đã bỏ tích (dùng cho trường hợp nhận diện bị sai màu).
        """
        try:
            # Đảm bảo cửa sổ game đang active để click có tác dụng
            try:
                if self.game_window:
                    self.game_window.activate()
                    time.sleep(0.2)
            except Exception:
                pass

            x, y = int(lock_pos[0]), int(lock_pos[1])

            # Kiểm tra trạng thái ban đầu
            if not force and not self.is_lock_checked(lock_pos, stat_index=stat_index):
                self.log(f"   ✅ Lock {lock_pos} đã ở trạng thái bỏ tích")
                return True
            elif force:
                self.log(f"   🔁 Force bỏ tích Lock {lock_pos} bất kể trạng thái nhận diện")

            # Thử click với vài vị trí lân cận để tăng độ chính xác
            click_positions = [
                (x, y),           # Vị trí chính xác
                (x+1, y),         # Lệch phải 1px
                (x, y+1),         # Lệch xuống 1px
                (x-1, y-1),       # Lệch chéo
            ]
            
            for attempt in range(3):  # Rút ngắn số lần thử để thao tác nhanh hơn
                self.log(f"   Thử bỏ tích lần {attempt + 1}/3...")
                
                for offset_x, offset_y in click_positions:
                    try:
                        # Click với vị trí offset
                        pyautogui.moveTo(offset_x, offset_y)
                        time.sleep(0.12) # Chờ trước khi click
                        pyautogui.click(offset_x, offset_y)
                        time.sleep(0.35)  # Chờ UI cập nhật
                        
                        # Kiểm tra kết quả (đọc hai lần để chống nhiễu)
                        unchecked_1 = not self.is_lock_checked(lock_pos, stat_index=stat_index)
                        time.sleep(0.12)
                        unchecked_2 = not self.is_lock_checked(lock_pos, stat_index=stat_index)
                        if unchecked_1 and unchecked_2:
                            self.log(f"   ✅ Đã bỏ tích thành công Lock {lock_pos}")
                            return True
                            
                    except FailSafeException:
                        raise
                    except Exception as e:
                        self.log(f"   ⚠️ Lỗi khi click Lock {lock_pos}: {e}")
                        continue
                
                # Nếu vẫn chưa bỏ tích được, thử click mạnh hơn
                if attempt < 2:
                    time.sleep(0.25)
                    try:
                        # Double click để chắc chắn
                        pyautogui.doubleClick(x, y)
                        time.sleep(0.25)
                        unchecked_1 = not self.is_lock_checked(lock_pos, stat_index=stat_index)
                        time.sleep(0.1)
                        unchecked_2 = not self.is_lock_checked(lock_pos, stat_index=stat_index)
                        if unchecked_1 and unchecked_2:
                            self.log(f"   ✅ Đã bỏ tích bằng double click Lock {lock_pos}")
                            return True
                    except FailSafeException:
                        raise
                    except Exception:
                        pass
            
            # Kiểm tra lần cuối
            final_check_1 = not self.is_lock_checked(lock_pos, stat_index=stat_index)
            time.sleep(0.12)
            final_check_2 = not self.is_lock_checked(lock_pos, stat_index=stat_index)
            final_check = final_check_1 and final_check_2
            if final_check:
                self.log(f"   ✅ Cuối cùng đã bỏ tích Lock {lock_pos}")
                return True
            else:
                self.log(f"   ❌ Không thể bỏ tích Lock {lock_pos} sau 3 lần thử")
                return False

        except FailSafeException:
            raise
        except Exception as e:
            self.log(f"   ❌ Lỗi trong ensure_unchecked: {e}")
            return False

    def unlock_all_locks(
        self,
        max_attempts: int = 5,
        *,
        force_click: bool = False,
        target_indices: list[int] | None = None,
    ) -> bool:
        """Bỏ tích các ô khóa được chỉ định.

        ``force_click`` cho phép bỏ qua nhận diện ban đầu và click bắt buộc để
        xử lý các trường hợp OCR màu bị sai. ``target_indices`` cho phép giới
        hạn danh sách chỉ số cần thao tác (mặc định là tất cả các chỉ số có cấu
        hình nút khóa).
        """

        if target_indices is None:
            indices = list(range(len(self.config["stats"])))
        else:
            indices = [idx for idx in target_indices if 0 <= idx < len(self.config["stats"])]

        pending: list[tuple[int, list[int] | tuple[int, int]]] = []

        for idx in indices:
            stat_cfg = self.config["stats"][idx]
            lock_pos = stat_cfg.get("lock_button", [0, 0])
            if sum(lock_pos) == 0:
                continue

            if force_click:
                pending.append((idx, lock_pos))
            else:
                if self.is_lock_checked(lock_pos, stat_index=idx):
                    pending.append((idx, lock_pos))

        if not pending:
            # Không có ô nào cần bỏ tích
            return True

        self.log("🔄 Đang bỏ tích các ô khóa...")

        for attempt in range(max_attempts):
            self.log(f"   Lần thử bỏ tích: {attempt + 1}/{max_attempts}")
            next_pending: list[tuple[int, list[int] | tuple[int, int]]] = []

            for idx, lock_pos in pending:
                if self.ensure_unchecked(lock_pos, force=force_click, stat_index=idx):
                    self.locked_stats[idx] = False
                else:
                    next_pending.append((idx, lock_pos))

            if not next_pending:
                self.log("✅ Đã bỏ tích thành công các dòng!")
                return True

            if attempt < max_attempts - 1:
                self.log(f"   ↻ Còn {len(next_pending)} dòng chưa bỏ tích, thử lại sau 0.6s...")
                time.sleep(0.6)

            pending = next_pending

        self.log("⚠️ Không thể bỏ tích hết các dòng sau nhiều lần thử.")
        return False

    def brute_force_unlock_locks(self, cycles: int = 2, jitter: int = 2) -> bool:
        """Nhấp mạnh vào toàn bộ các ô khóa mà không cần nhận diện trạng thái.

        Hàm này được dùng khi thao tác thường xuyên ``unlock_all_locks`` thất bại
        vì OCR hoặc template nhận diện sai. Logic: di chuyển chuột tới mỗi vị trí
        đã cấu hình, click với nhiều offset nhỏ và double-click nhằm đảm bảo
        checkbox được bỏ tích.
        """

        coords: list[tuple[int, int]] = []
        for stat_cfg in self.config.get("stats", []):
            lock_pos = stat_cfg.get("lock_button", [0, 0])
            if sum(lock_pos) == 0:
                continue
            try:
                coords.append((int(lock_pos[0]), int(lock_pos[1])))
            except Exception:
                continue

        if not coords:
            return False

        try:
            if self.game_window:
                self.game_window.activate()
                time.sleep(0.2)
        except Exception:
            pass

        self.log("   🔁 Bruteforce: thử nhấp mạnh để bỏ tích tất cả ô khóa...")

        for cycle in range(max(1, cycles)):
            for idx, (x, y) in enumerate(coords):
                offsets = [
                    (0, 0),
                    (1, 0),
                    (0, 1),
                    (-1, 0),
                    (0, -1),
                ]
                if jitter > 0:
                    offsets.extend([
                        (jitter, jitter),
                        (-jitter, jitter),
                        (jitter, -jitter),
                        (-jitter, -jitter),
                    ])

                for ox, oy in offsets:
                    px, py = x + ox, y + oy
                    try:
                        pyautogui.moveTo(px, py)
                        time.sleep(0.05)
                        pyautogui.click(px, py)
                        time.sleep(0.08)
                    except FailSafeException:
                        raise
                    except Exception as exc:
                        self.log(f"      ⚠️ Bruteforce: lỗi click ô khóa {idx+1}: {exc}")
                try:
                    pyautogui.doubleClick(x, y)
                except FailSafeException:
                    raise
                except Exception:
                    pass
                time.sleep(0.12)

            time.sleep(0.25)

        return True

    def normalize_vi(self, s: str) -> str:
        # Bỏ dấu tiếng Việt để so khớp văn bản đơn giản
        nfkd = unicodedata.normalize('NFKD', s)
        ascii_str = ''.join([c for c in nfkd if not unicodedata.combining(c)])
        return ascii_str.lower()

    def clean_ocr_text(self, text: str) -> str:
        # Làm sạch một số lỗi OCR phổ biến
        s = unicodedata.normalize('NFKC', text)
        s = s.replace(',', '.')
        s = s.replace(' ', '')
        s = s.replace('–', '-')
        # Giữ lại ký tự hợp lệ cho parse
        s = re.sub(r"[^0-9%+().\-]", "", s)
        # Chuyển các ký tự dễ nhầm thành số
        trans = {
            'O': '0', 'o': '0', 'D': '0',
            'l': '1', 'I': '1', 'í': '1',
            'S': '5'
        }
        s = ''.join(trans.get(ch, ch) for ch in s)
        return s

    @staticmethod
    def format_percent_value(value) -> str:
        """Trả về chuỗi hiển thị phần trăm đúng như giá trị OCR thu được."""

        if isinstance(value, (int, float)):
            text = f"{value}"
            if isinstance(value, float) and '.' in text:
                text = text.rstrip('0').rstrip('.')
            return text
        return str(value)

    @staticmethod
    def _count_integer_digits_from_token(token: str | None) -> int | None:
        """Đếm số chữ số phần nguyên trong chuỗi OCR ban đầu."""

        if not token:
            return None

        normalized = unicodedata.normalize('NFKC', token)
        normalized = normalized.replace(',', '.')
        match = re.search(r'-?\d+(?:\.\d+)?', normalized)
        if not match:
            return None

        number_part = match.group(0)
        if '.' in number_part:
            integer_part = number_part.split('.')[0]
        else:
            integer_part = number_part

        integer_part = integer_part.lstrip('0')
        if not integer_part:
            integer_part = '0'
        return len(integer_part)

    def fix_percent_current_with_max(self, current_value: float, range_max: float | None) -> float:
        """Thử khôi phục dấu chấm bị mất dựa trên range_max."""

        if range_max is None:
            return current_value

        candidates = [current_value]
        for div in (10.0, 100.0, 1000.0):
            candidates.append(current_value / div)

        best = current_value
        best_delta = abs(current_value - range_max)
        for c in candidates:
            delta = abs(c - range_max)
            if delta < best_delta:
                best_delta = delta
                best = c
        if best > 400:
            best = 400.0
        elif best < -400:
            best = -400.0

        return best

    def normalize_percent_value(
        self,
        value: float,
        reference: float | None = None,
        *legacy_tokens: str,
        raw_token: str | None = None,
        reference_token: str | None = None,
        **legacy_kwargs,
    ) -> float:
        """Chuẩn hoá giá trị % mà không làm mất 3 chữ số như 224%.

        Nếu ``reference`` được cung cấp (thường là giá trị MAX hoặc CURRENT tương ứng),
        ưu tiên chọn ứng viên gần ``reference`` nhất. Nếu không có ``reference``, chọn
        ứng viên nằm trong khoảng [0, 400] với độ lớn lớn nhất để tránh rơi xuống 2 chữ số.
        """

        # --- Tương thích ngược ---
        # Các phiên bản cũ có thể truyền đối số vị trí hoặc keyword lạ. Gom các giá trị này
        # về ``raw_token``/``reference_token`` và bỏ qua phần còn lại để tránh lỗi runtime.
        if legacy_tokens:
            if raw_token is None:
                raw_token = legacy_tokens[0]
            if len(legacy_tokens) > 1 and reference_token is None:
                reference_token = legacy_tokens[1]

        if "raw_token" in legacy_kwargs and raw_token is None:
            raw_token = legacy_kwargs.pop("raw_token")
        if "reference_token" in legacy_kwargs and reference_token is None:
            reference_token = legacy_kwargs.pop("reference_token")
        legacy_kwargs.clear()

        candidates = [value]
        for div in (10.0, 100.0, 1000.0, 10000.0):
            candidates.append(value / div)

        if reference is not None:
            best = min(candidates, key=lambda cand: abs(cand - reference))
        else:
            best = None

            # Không có reference: chọn ứng viên trong khoảng hợp lý nhất (0..400)
            plausible = [cand for cand in candidates if 0 <= cand <= 400]
            if plausible:
                best = max(plausible)

        if best is None:
            best = value

        # Nếu OCR gốc có >=3 chữ số phần nguyên nhưng giá trị hiện tại <100, khôi phục bằng cách nhân 10.
        integer_digits = self._count_integer_digits_from_token(raw_token)
        if integer_digits and integer_digits >= 3 and abs(best) < 100:
            adjusted = best
            digits = integer_digits
            while digits >= 3 and abs(adjusted) < 100:
                adjusted *= 10.0
                digits -= 1
            best = adjusted

        ref_digits = self._count_integer_digits_from_token(reference_token)
        if reference is not None and ref_digits and ref_digits >= 3 and abs(reference) < 100:
            adjusted_ref = reference
            digits = ref_digits
            while digits >= 3 and abs(adjusted_ref) < 100:
                adjusted_ref *= 10.0
                digits -= 1
            # Giữ best gần reference đã điều chỉnh nếu cần
            scale = adjusted_ref / reference if reference else 1.0
            if scale not in (0.0, 1.0):
                best *= scale

        if best > 400:
            best = 400.0
        elif best < -400:
            best = -400.0

        return best

    def normalize_percent_value(self, value: float, reference: float | None = None) -> float:
        """Chuẩn hoá giá trị % mà không làm mất 3 chữ số như 224%.

        Nếu ``reference`` được cung cấp (thường là giá trị MAX hoặc CURRENT tương ứng),
        ưu tiên chọn ứng viên gần ``reference`` nhất. Nếu không có ``reference``, chọn
        ứng viên nằm trong khoảng [0, 400] với độ lớn lớn nhất để tránh rơi xuống 2 chữ số.
        """

        candidates = [value]
        for div in (10.0, 100.0, 1000.0, 10000.0):
            candidates.append(value / div)

        if reference is not None:
            best = min(candidates, key=lambda cand: abs(cand - reference))
            return best

        # Không có reference: chọn ứng viên trong khoảng hợp lý nhất (0..400)
        plausible = [cand for cand in candidates if 0 <= cand <= 400]
        if plausible:
            # Ưu tiên giá trị lớn nhất trong khoảng hợp lý để giữ đủ chữ số
            return max(plausible)
        return value

    def is_read_valid(self, current_value, range_max, is_percent: bool) -> bool:
        if is_percent:
            # Cho phép chỉ số % lên tới 400 để không làm mất 3 chữ số như 224%
            if current_value < 0 or current_value > 400:
                return False
            if range_max is not None and current_value > range_max * 1.5 + 1:
                return False
            return True
        else:
            # Số nguyên không vượt quá 10 lần max
            if current_value < 0:
                return False
            if range_max is not None and current_value > range_max * 10:
                return False
            return True

    def parse_ocr_result(self, text):
        # Phân tích chắc chắn theo cấu trúc: "+GIÁ_TRỊ [%(tuỳ chọn)] (MIN-MAX[%(tuỳ chọn)])"
        # Bỏ dấu phẩy, chuẩn hoá khoảng trắng
        raw = text.strip()
        cleaned = self.clean_ocr_text(raw)

        # Chuẩn hoá ký tự gạch ngang (OCR có thể thành '–' hoặc '-')
        cleaned = cleaned.replace('–', '-')

        range_max = None
        range_raw_token = None
        is_percent = '%' in cleaned

        # 1) Tìm cặp min-max dạng phần trăm trong ngoặc: (a%-b%)
        pm = re.search(r'\((\d+(?:\.\d+)?)%\s*-\s*(\d+(?:\.\d+)?)%\)?', cleaned)
        if pm:
            try:
                a = float(pm.group(1))
                b = float(pm.group(2))
                range_max = max(a, b)
                is_percent = True
                range_raw_token = pm.group(2) if b >= a else pm.group(1)
            except:
                range_max = None
        else:
            # 1b) Nếu không phải phần trăm, thử bắt cặp số nguyên trong ngoặc: (min-max)
            nm = re.search(r'\((\d+)\s*-\s*(\d+)\)?', cleaned)
            if nm:
                try:
                    a = int(nm.group(1))
                    b = int(nm.group(2))
                    range_max = max(a, b)
                    range_raw_token = nm.group(2) if b >= a else nm.group(1)
                except:
                    range_max = None

        # 2) Lấy số sau dấu '+' dạng phần trăm: +x.x%
        plus_percent = re.search(r'\+\s*(\d+(?:\.\d+)?)\s*%\b', cleaned)
        plus_number  = re.search(r'\+\s*(\d+(?:\.\d+)?)\b(?!%)', cleaned)
        current_raw_token = None
        if plus_percent:
            current_value = float(plus_percent.group(1))
            is_percent = True
            current_raw_token = plus_percent.group(1)
        elif plus_number:
            if is_percent:
                # Nếu đã xác định là phần trăm từ cặp (min%-max%) mà dấu % sau dấu + bị mất,
                # vẫn đọc giá trị dạng số thực để so sánh chính xác A == C
                current_value = float(plus_number.group(1))
                current_raw_token = plus_number.group(1)
            else:
                current_value = int(float(plus_number.group(1)))
        else:
            # Fallback an toàn
            nums = re.findall(r'(\d+(?:\.\d+)?)', cleaned)
            if nums:
                if is_percent:
                    current_value = float(nums[0])
                    current_raw_token = nums[0]
                else:
                    current_value = int(float(nums[0]))
            else:
                return (0.0 if is_percent else 0), None, is_percent

        # Sanity cho phần trăm: phục hồi giá trị thực nếu OCR dính thừa chữ số (ví dụ 19604 -> 196.04)
        # Đồng bộ kiểu dữ liệu current/range_max
        if is_percent:
            if isinstance(current_value, int):
                current_value = float(current_value)
            if isinstance(range_max, int):
                range_max = float(range_max)
            current_value = self.normalize_percent_value(
                current_value,
                range_max,
                raw_token=current_raw_token,
                reference_token=range_raw_token,
            )
            if range_max is not None:
                range_max = self.normalize_percent_value(
                    range_max,
                    current_value,
                    raw_token=range_raw_token,
                    reference_token=current_raw_token,
                )
            # Sửa lỗi rơi dấu chấm nếu lệch xa max
            current_value = self.fix_percent_current_with_max(current_value, range_max)
        else:
            if isinstance(current_value, float):
                current_value = int(round(current_value))
            if isinstance(range_max, float):
                range_max = int(round(range_max))

        # BỎ Fallback: KHÔNG tự suy luận MAX từ current đối với %.
        # Yêu cầu phải đọc được cả A và C để so sánh A == C.

        return current_value, range_max, is_percent

    def is_meeting_target(self, current_value, range_max, desired_value, is_percent: bool) -> bool:
        # YÊU CẦU NGHIÊM NGẶT: Chỉ khóa khi đọc được MAX trong ngoặc và A == C
        if range_max is None or (range_max is not None and range_max <= 0):
            return False
        if is_percent:
            # So sánh chính xác theo giá trị OCR: chỉ khóa khi A == C tuyệt đối.
            try:
                return float(current_value) == float(range_max)
            except (TypeError, ValueError):
                return False
        else:
            # Số nguyên: bắt buộc bằng đúng
            return int(current_value) == int(range_max)

    def automation_loop(self):
        self.log("=== BẮT ĐẦU QUÁ TRÌNH TỰ ĐỘNG ===")
        cycle_count = 0
        
        while self.is_running:
            try:
                cycle_count += 1
                self.log(f"--- Chu kỳ {cycle_count} ---")

                if not self.game_window or not self.game_window.isActive:
                    self.log("Cửa sổ game không hoạt động. Tạm dừng.")
                    time.sleep(1.0)
                    continue

                # Theo yêu cầu: KHÔNG bỏ tích bất kỳ ô khóa nào trước khi tẩy luyện.
                # Chỉ thực hiện bỏ tích sau khi thăng cấp thành công.

                # Nhấp nút Tẩy Luyện với delay dài hơn
                pyautogui.click(self.config["refine_button"])
                self.log(">> Đã nhấn Tẩy Luyện")
                time.sleep(1.6) # Rút ngắn thời gian chờ UI load hoàn toàn

                all_done = True
                locked_this_cycle = False
                # Theo dõi dòng đạt MAX trong chu kỳ hiện tại (kể cả đã khóa)
                max_flags = [False] * len(self.config["stats"])
                for i, stat in enumerate(self.config["stats"]):
                    if self.locked_stats[i]:
                        self.log(f"   Chỉ số {i+1}: Đã khóa")
                        # Dòng đã khóa được coi là đang ở trạng thái MAX
                        max_flags[i] = True
                        continue
                    
                    # Bỏ qua nếu chưa thiết lập
                    if sum(stat["area"]) == 0 or sum(stat["lock_button"]) == 0:
                        continue
                    
                    all_done = False
                    
                    # Chụp và đọc chỉ số với delay để UI ổn định
                    x, y, w, h = stat["area"]
                    time.sleep(0.2) # Chờ UI ổn định trước khi chụp
                    screenshot = pyautogui.screenshot(region=(x, y, w, h))
                    processed_img = self.process_image_for_ocr(screenshot)
                    
                    # Đọc OCR với cơ chế dự phòng
                    text = self.ocr_read_text(processed_img, debug_tag=f"stat_{i+1}")
                    current_value, range_max, is_percent = self.parse_ocr_result(text)

                    # Bỏ qua nếu đọc nhiễu/không hợp lệ để tránh khóa sai
                    if not self.is_read_valid(current_value, range_max, is_percent):
                        self.log(f"   Chỉ số {i+1}: dữ liệu OCR bất thường, bỏ qua vòng này")
                        continue
                    
                    # Tính mục tiêu và đánh giá đạt/chưa với tolerance
                    target = (range_max if range_max is not None else stat['desired_value'])
                    achieved = self.is_meeting_target(current_value, range_max, stat['desired_value'], is_percent)

                    if target is not None and target > 0:
                        if is_percent:
                            fmt_current = self.format_percent_value(current_value)
                            fmt_target = self.format_percent_value(target)
                            self.log(
                                f"   Chỉ số {i+1}: '{text.strip()}' -> {fmt_current}% / Mục tiêu {fmt_target}%  => {'ĐẠT' if achieved else 'chưa đạt'}"
                            )
                        else:
                            self.log(f"   Chỉ số {i+1}: '{text.strip()}' -> {current_value} / Mục tiêu {target}  => {'ĐẠT' if achieved else 'chưa đạt'}")
                    else:
                        if is_percent:
                            fmt_current = self.format_percent_value(current_value)
                            self.log(f"   Chỉ số {i+1}: '{text.strip()}' -> {fmt_current}%")
                        else:
                            self.log(f"   Chỉ số {i+1}: '{text.strip()}' -> {current_value}")
                    
                    # Cập nhật GUI
                    if range_max is not None:
                        if is_percent:
                            fmt_current = self.format_percent_value(current_value)
                            fmt_range = self.format_percent_value(range_max)
                            display_text = f"Hiện tại: {fmt_current}% / Max: {fmt_range}%"
                            self.root.after(0, lambda i=i, text=display_text: self.stat_entries[i]["current_label"].config(text=text))
                        else:
                            self.root.after(0, lambda i=i, val=current_value, mx=range_max: self.stat_entries[i]["current_label"].config(text=f"Hiện tại: {val} / Max: {mx}"))
                    else:
                        if is_percent:
                            fmt_current = self.format_percent_value(current_value)
                            display_text = f"Giá trị hiện tại: {fmt_current}%"
                            self.root.after(0, lambda i=i, text=display_text: self.stat_entries[i]["current_label"].config(text=text))
                        else:
                            self.root.after(0, lambda i=i, val=current_value: self.stat_entries[i]["current_label"].config(text=f"Giá trị hiện tại: {val}"))

                    # So sánh và khóa: chỉ khóa nếu đạt mục tiêu và (không yêu cầu chữ đỏ hoặc là chữ đỏ)
                    if achieved:
                        require_red = bool(self.require_red_var.get())
                        if (not require_red) or self.has_red_text(screenshot):
                            self.log(f"   !!! Chỉ số {i+1} đạt MAX. Đang khóa...")
                            time.sleep(0.8) # Chờ trước khi click khóa
                            pyautogui.click(stat["lock_button"])
                            self.locked_stats[i] = True
                            locked_this_cycle = True
                            time.sleep(1.0) # Chờ UI cập nhật sau khi khóa
                        else:
                            self.log(f"   → Đạt mục tiêu nhưng chưa xác nhận chữ đỏ, bỏ qua")
                        # Ghi nhận đạt MAX trong chu kỳ
                        max_flags[i] = True

                # Kiểm tra điều kiện thăng cấp: cần 4 dòng đạt MAX, nhưng chỉ có thể khóa 3 dòng
                num_locked = sum(1 for v in self.locked_stats if v)
                total_max = sum(1 for i in range(len(self.config["stats"])) if self.locked_stats[i] or max_flags[i])
                self.log(f"   Số dòng đã khóa: {num_locked}/4 | Tổng dòng MAX (đã khóa + đạt MAX hiện tại): {total_max}/4")

                # Thăng cấp khi đã khóa >= 3 và tổng cộng 4 dòng đạt MAX
                if num_locked >= 3 and total_max >= 4:
                    if self.is_upgrade_available():
                        self.log("🎯 Đủ điều kiện: 3 dòng đã khóa + 1 dòng đạt MAX, nút Thăng Cấp active - Bắt đầu thăng cấp!")
                    else:
                        self.log("🎯 Đủ điều kiện: 3 dòng đã khóa + 1 dòng đạt MAX - Thử thăng cấp (fallback)...")

                    try:
                        if self.game_window:
                            self.game_window.activate()
                            time.sleep(0.2)
                    except Exception:
                        pass

                    upgrade_result = self.perform_upgrade_sequence()
                    if upgrade_result:
                        self.log("🔄 Tự động tiếp tục tẩy luyện với mục tiêu mới...")
                        time.sleep(0.6)
                    
                    # Click nút Thăng Cấp 1 lần duy nhất
                    upgrade_clicked = False
                    if sum(self.config.get("upgrade_button", [0,0])) > 0:
                        bx, by = self.config["upgrade_button"]
                        pyautogui.moveTo(bx, by)
                        pyautogui.click(bx, by)
                        upgrade_clicked = True
                        self.log(f"▶️ Đã click nút Thăng Cấp tại ({bx}, {by})")
                    elif sum(self.config.get("upgrade_area", [0,0,0,0])) > 0:
                        ux, uy, uw, uh = self.config["upgrade_area"]
                        cx, cy = ux + uw//2, uy + uh//2
                        pyautogui.moveTo(cx, cy)
                        pyautogui.click(cx, cy)
                        upgrade_clicked = True
                        self.log(f"▶️ Đã click vùng Thăng Cấp tại ({cx}, {cy})")

                    if upgrade_clicked:
                        # Chờ animation thăng cấp hoàn thành
                        time.sleep(4.0) # Tăng thời gian chờ animation

                        # Bỏ tích và xác nhận bằng template: yêu cầu cả 4 ô là 'chưa tích'
                        success_unlock = self.unlock_all_locks(max_attempts=6, force_click=True)
                        if not success_unlock:
                            if self.brute_force_unlock_locks(cycles=3):
                                time.sleep(0.4)
                                success_unlock = self.unlock_all_locks(max_attempts=4, force_click=True)
                        self.locked_stats = [False] * 4

                        # Sau khi bỏ tích bằng click, kiểm tra bằng template vài lần để chắc chắn
                        check_rounds = 0
                        all_ok = False
                        used_template = (self._tpl_checked is not None) or (self._tpl_unchecked is not None)

                        while check_rounds < 3 and success_unlock and used_template:
                            tpl_status = self.all_locks_unchecked_by_template()
                            if tpl_status is True:
                                all_ok = True
                                break
                            if tpl_status is False:
                                self.log("   ⏳ Template: phát hiện còn ô đang TÍCH, thử bỏ tích lại...")
                                # Thử bỏ tích mạnh lại 1 vòng ngắn
                                success_unlock = self.unlock_all_locks(max_attempts=3, force_click=True)
                            else:
                                self.log("   ⏳ Template: không đủ chắc chắn, sẽ kiểm tra lại sau 0.4s...")
                            check_rounds += 1
                            time.sleep(0.4)

                        if success_unlock and not all_ok:
                            self.log("   🔍 Bỏ qua kiểm tra template hoặc chưa đủ chắc chắn, chuyển sang kiểm tra fallback bằng màu sắc...")
                            all_ok = self.verify_all_locks_unchecked()

                        if success_unlock and all_ok:
                            self.log("✅ Đã thăng cấp thành công và xác nhận 4 ô đều CHƯA TÍCH!")
                            self.log("🔄 Tự động tiếp tục tẩy luyện với mục tiêu mới...")
                            time.sleep(0.6)
                        else:
                            self.log(
                                "⚠️ Không thể xác nhận 4 ô đều CHƯA TÍCH sau thăng cấp. Tránh tẩy luyện sai nên tool sẽ dừng để bạn kiểm tra lại."
                            )
                            self.is_running = False
                            self.root.after(0, self._update_button_states)
                            time.sleep(1.0)
                        continue
                    else:
                        self.log("⏳ Chưa thể hoàn tất thăng cấp, sẽ thử lại sau 1.0s.")
                        time.sleep(1.0)
                    continue

                elif num_locked < 3 or total_max < 4:
                    self.log(f"📊 Chưa đủ điều kiện thăng cấp (khóa {num_locked}/4, tổng MAX {total_max}/4) - Tiếp tục tẩy luyện...")

                # Nếu không có điều kiện thăng cấp, tiếp tục chu kỳ bình thường
                if all_done:
                    self.log("ℹ️ Tất cả chỉ số đã được xử lý trong chu kỳ này")

                time.sleep(1.0) # Rút ngắn thời gian nghỉ giữa các chu kỳ để tăng tốc
            
            except FailSafeException:
                self.log("⛔ PyAutoGUI fail-safe được kích hoạt. Đang dừng luồng tự động để đảm bảo an toàn.")
                self.is_running = False
                break
            except Exception as e:
                self.log(f"❌ Có lỗi xảy ra: {e}")
                time.sleep(3) # Tăng thời gian nghỉ khi có lỗi

        self.log("=== LUỒNG TỰ ĐỘNG ĐÃ DỪNG ===")
        self.root.after(0, self._update_button_states)

    def start_automation(self):
        if self.is_running:
            self.log("⚠️ Đã chạy rồi!")
            return

        if not self.game_window:
            messagebox.showerror("Lỗi", "Vui lòng chọn cửa sổ game!")
            return
        
        # Kiểm tra cấu hình
        if sum(self.config["refine_button"]) == 0:
            messagebox.showerror("Lỗi", "Vui lòng thiết lập nút Tẩy Luyện!")
            return
        
        if not self._sync_stat_entries_to_config(strict=True):
            return

        # Kiểm tra có ít nhất một chỉ số có vùng đọc và nút khóa (không bắt buộc desired_value)
        has_configured_stats = any(sum(stat["area"]) > 0 and sum(stat["lock_button"]) > 0
                                 for stat in self.config["stats"])
        
        if not has_configured_stats:
            messagebox.showerror("Lỗi", "Vui lòng thiết lập ít nhất một chỉ số với vùng đọc và nút khóa!")
            return
        
        self.is_running = True
        self.locked_stats = [False] * 4 # Reset trạng thái khóa
        self.pending_upgrade = False
        self._update_button_states()
        self.automation_thread = threading.Thread(target=self.automation_loop, daemon=True)
        self.automation_thread.start()
        
    def stop_automation(self):
        if not self.is_running:
            return
        self.is_running = False
        self.log("⏹️ Đang yêu cầu dừng...")
        self._update_button_states()

    def _update_button_states(self):
        if self.is_running:
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
        else:
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
    
    def on_closing(self):
        self.save_config()
        self.is_running = False
        try:
            keyboard.unhook_all()
        except:
            pass
        self.root.destroy()

    def save_config(self):
        try:
            self._sync_stat_entries_to_config(strict=False)
            with open(self.config_file, "w", encoding='utf-8') as f:
                # đồng bộ require_red từ checkbox
                self.config["require_red"] = bool(self.require_red_var.get())
                json.dump(self.config, f, indent=4, ensure_ascii=False)
            self.log("✅ Đã lưu cấu hình.")
        except Exception as e:
            self.log(f"❌ Lỗi khi lưu cấu hình: {e}")

    def load_config(self):
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, "r", encoding='utf-8') as f:
                    loaded_config = json.load(f)
                
                # Cập nhật config với dữ liệu đã lưu
                if "refine_button" in loaded_config:
                    self.config["refine_button"] = loaded_config["refine_button"]
                
                if "stats" in loaded_config:
                    for i, stat_conf in enumerate(loaded_config["stats"]):
                        if i < len(self.config["stats"]):
                            self.config["stats"][i].update(stat_conf)
                
                # Cập nhật GUI từ config
                btn_pos = self.config["refine_button"]
                if sum(btn_pos) > 0:
                    self.refine_btn_label.config(text=f"X={btn_pos[0]}, Y={btn_pos[1]}")

                # set checkbox trạng thái yêu cầu chữ đỏ
                self.require_red_var.set(bool(self.config.get("require_red", False)))

                for i, stat in enumerate(self.config.get("stats", [])):
                    if i >= len(self.stat_entries):
                        continue
                    area = stat.get("area", [0, 0, 0, 0])
                    if sum(area) > 0:
                        self.stat_entries[i]["area_label"].config(text=f"Đã đặt ({area[2]}x{area[3]})")
                    else:
                        self.stat_entries[i]["area_label"].config(text="Chưa đặt")

                    lock_btn = stat.get("lock_button", [0, 0])
                    if sum(lock_btn) > 0:
                        self.stat_entries[i]["lock_label"].config(text=f"X={lock_btn[0]}, Y={lock_btn[1]}")
                    else:
                        self.stat_entries[i]["lock_label"].config(text="Chưa đặt")

                    lock_area = stat.get("lock_ocr_area", [0, 0, 0, 0])
                    if sum(lock_area) > 0:
                        self.stat_entries[i]["lock_ocr_label"].config(text=f"Đã đặt ({lock_area[2]}x{lock_area[3]})")
                    else:
                        self.stat_entries[i]["lock_ocr_label"].config(text="Chưa đặt")

                    desired_entry = self.stat_entries[i]["desired_value"]
                    desired_entry.delete(0, tk.END)
                    desired_val = stat.get("desired_value", 0)
                    try:
                        desired_int = int(desired_val)
                    except (TypeError, ValueError):
                        desired_int = 0
                    if desired_int != 0:
                        desired_entry.insert(0, str(desired_int))

                    unchecked_entry = self.stat_entries[i]["lock_unchecked_entry"]
                    unchecked_entry.delete(0, tk.END)
                    unchecked_entry.insert(0, stat.get("lock_unchecked_keyword", ""))

                    checked_entry = self.stat_entries[i]["lock_checked_entry"]
                    checked_entry.delete(0, tk.END)
                    checked_entry.insert(0, stat.get("lock_checked_keyword", ""))

                    self.stat_entries[i]["lock_status_label"].config(text="Trạng thái khóa: --")

                # upgrade button/area labels
                up_btn = self.config.get("upgrade_button", [0,0])
                if sum(up_btn) > 0:
                    self.upgrade_btn_label.config(text=f"X={up_btn[0]}, Y={up_btn[1]}")
                up_area = self.config.get("upgrade_area", [0,0,0,0])
                if sum(up_area) > 0:
                    self.upgrade_area_label.config(text=f"Đã đặt ({up_area[2]}x{up_area[3]})")

                self.log("✅ Đã tải cấu hình đã lưu.")
            else:
                self.log("ℹ️ Không tìm thấy file cấu hình, sử dụng mặc định.")
        except Exception as e:
            self.log(f"❌ Lỗi khi tải cấu hình: {e}")

    # === Nhận diện mẫu ô khóa ===
    def _load_lock_templates(self) -> None:
        paths = self.config.get("lock_templates", {})
        checked_path = paths.get("checked")
        unchecked_path = paths.get("unchecked")
        def _load_one(p):
            if not p:
                return None
            if not os.path.isabs(p):
                p = os.path.join(os.getcwd(), p)
            if not os.path.exists(p):
                return None
            img = Image.open(p).convert('L').resize((24, 24), Image.LANCZOS)
            if np is None:
                return img
            arr = np.asarray(img, dtype=np.float32)
            m = arr.mean()
            s = arr.std() if arr.std() > 1e-5 else 1.0
            arr = (arr - m) / s
            return arr
        self._tpl_checked = _load_one(checked_path)
        self._tpl_unchecked = _load_one(unchecked_path)
        if self._tpl_checked is not None or self._tpl_unchecked is not None:
            self.log("🔎 Đã tải template nhận diện ô khóa.")

    def capture_lock_template(self, is_checked: bool) -> None:
        if not self.game_window:
            messagebox.showerror("Lỗi", "Vui lòng chọn cửa sổ game trước!")
            return
        try:
            self.game_window.activate()
        except Exception:
            pass
        time.sleep(0.3)

        info_window = tk.Toplevel(self.root)
        info_window.title("Lấy mẫu ô khóa")
        info_window.geometry("420x140")
        info_window.transient(self.root)
        info_window.grab_set()
        msg = "Đưa chuột vào giữa ô khóa ĐÃ TÍCH và nhấn F8" if is_checked else "Đưa chuột vào giữa ô khóa CHƯA TÍCH và nhấn F8"
        tk.Label(info_window, text=msg, padx=16, pady=16, justify=tk.LEFT).pack()

        pos: list[tuple[int,int]] = []
        def on_f8(event):
            if event.name == 'f8':
                pos.append(pyautogui.position())
                keyboard.unhook_all()
                info_window.destroy()

        keyboard.on_press(on_f8)
        self.root.wait_window(info_window)

        if not pos:
            return
        cx, cy = pos[0]
        box = 28
        half = box // 2
        left = max(0, cx - half)
        top = max(0, cy - half)
        snap = pyautogui.screenshot(region=(left, top, box, box))
        out_name = self.config.get("lock_templates", {}).get("checked" if is_checked else "unchecked")
        if not out_name:
            out_name = "lock_checked.png" if is_checked else "lock_unchecked.png"
            self.config.setdefault("lock_templates", {})["checked" if is_checked else "unchecked"] = out_name
        try:
            snap.save(out_name)
            self.save_config()
            self._load_lock_templates()
            state_txt = "ĐÃ TÍCH" if is_checked else "CHƯA TÍCH"
            self.log(f"✅ Đã lưu mẫu ô khóa {state_txt}: {out_name}")
        except Exception as e:
            self.log(f"❌ Lỗi lưu mẫu ô khóa: {e}")

    def _template_similarity(self, img: Image.Image, tpl_norm) -> float:
        if tpl_norm is None:
            return -1.0
        try:
            # So khớp trung tâm để tránh viền
            gray_full = img.convert('L').resize((24, 24), Image.LANCZOS)
            gray = gray_full.crop((3, 3, 21, 21))  # 18x18
            if np is None:
                # Fallback: negative MSE (để so sánh tương đối)
                garr = list(gray.getdata())
                if isinstance(tpl_norm, np.ndarray):
                    timg = Image.fromarray((tpl_norm*32+128).clip(0,255).astype('uint8'))
                else:
                    timg = tpl_norm
                timg = timg if timg.size == (18, 18) else timg.resize((18, 18), Image.LANCZOS)
                tarr = list(timg.getdata())
                n = min(len(garr), len(tarr))
                if n == 0:
                    return -1.0
                mse = sum((garr[i]-tarr[i])**2 for i in range(n))/n
                return -mse
            g = np.asarray(gray, dtype=np.float32)
            gm = g.mean(); gs = g.std() if g.std() > 1e-5 else 1.0
            g = (g - gm) / gs
            tpl = tpl_norm
            if isinstance(tpl_norm, np.ndarray):
                tpl = tpl_norm
            else:
                tpl_img = tpl_norm if tpl_norm.size == (18, 18) else tpl_norm.resize((18, 18), Image.LANCZOS)
                tpl = np.asarray(tpl_img, dtype=np.float32)
                tm = tpl.mean(); ts = tpl.std() if tpl.std() > 1e-5 else 1.0
                tpl = (tpl - tm) / ts
            num = float((g * tpl).sum())
            den = float(math.sqrt((g*g).sum()) * math.sqrt((tpl*tpl).sum()))
            if den < 1e-6:
                return -1.0
            return num / den
        except Exception:
            return -1.0

    def _is_unchecked_by_template(self, snap: Image.Image) -> bool | None:
        """Trả về True nếu ảnh ô khóa khớp mẫu 'chưa tích', False nếu khớp 'đã tích'.
        Trả về None nếu không đủ mẫu để kết luận.
        """
        try:
            if self._tpl_checked is None and self._tpl_unchecked is None:
                return None
            sim_checked = self._template_similarity(snap, self._tpl_checked)
            sim_unchecked = self._template_similarity(snap, self._tpl_unchecked)
            # Ngưỡng nới lỏng, dựa trên chênh lệch
            if sim_unchecked >= 0.60 and (sim_unchecked - max(-1.0, sim_checked)) >= 0.08:
                return True
            if sim_checked >= 0.60 and (sim_checked - max(-1.0, sim_unchecked)) >= 0.08:
                return False
            return None
        except Exception:
            return None

    def all_locks_unchecked_by_template(self) -> bool | None:
        """Kiểm tra tất cả ô khóa theo template. True nếu tất cả 'chưa tích'.
        False nếu có ít nhất một ô khớp 'đã tích'. None nếu không đủ mẫu để kết luận.
        """
        results: list[bool | None] = []
        diffs: list[float] = []
        for stat_cfg in self.config.get("stats", []):
            lock_pos = stat_cfg.get("lock_button", [0, 0])
            if sum(lock_pos) == 0:
                continue
            lx, ly = int(lock_pos[0]), int(lock_pos[1])
            box = 28
            half = box // 2
            left = max(0, lx - half)
            top = max(0, ly - half)
            snap = pyautogui.screenshot(region=(left, top, box, box))
            # Tính luôn chênh lệch để dùng khi không chắc chắn
            sim_c = self._template_similarity(snap, self._tpl_checked)
            sim_u = self._template_similarity(snap, self._tpl_unchecked)
            diffs.append(sim_u - sim_c)
            res = None
            try:
                res = self._is_unchecked_by_template(snap)
            except Exception:
                res = None
            results.append(res)

        if not results:
            return None
        if any(r is False for r in results):
            return False
        if all(r is True for r in results):
            return True
        # Chấp nhận nếu >= 3 ô là True và các ô còn lại có chênh lệch nghiêng về unchecked
        true_count = sum(1 for r in results if r is True)
        if true_count >= 3 and sum(1 for d in diffs if d >= 0.06) >= 3:
            return True
        # Nếu trung bình chênh lệch rõ rệt về unchecked, cũng coi là True
        if len(diffs) >= 3 and (sum(diffs)/max(1, len(diffs))) >= 0.09:
            return True
        return None

    def verify_all_locks_unchecked(self, retries: int = 2, delay: float = 0.4, allow_bruteforce: bool = True) -> bool:
        """Kiểm tra lại trạng thái bỏ tích của các ô khóa bằng phân tích màu sắc.

        Hàm này dùng ``is_lock_checked`` để xác nhận thủ công trong trường hợp
        thiếu mẫu template hoặc kết quả so khớp chưa rõ ràng.
        """

        indices = [idx for idx, stat_cfg in enumerate(self.config.get("stats", []))
                   if sum(stat_cfg.get("lock_button", [0, 0])) > 0]
        if not indices:
            return True

        still_checked: list[int] = []
        for attempt in range(retries):
            still_checked = []
            for idx in indices:
                lock_pos = self.config["stats"][idx].get("lock_button", [0, 0])
                try:
                    if self.is_lock_checked(lock_pos, stat_index=idx):
                        still_checked.append(idx)
                except Exception as exc:
                    self.log(f"   ⚠️ Fallback: lỗi khi kiểm tra ô khóa {idx+1}: {exc}")
                    still_checked.append(idx)

            if not still_checked:
                if attempt > 0:
                    self.log("   ✅ Fallback màu sắc: xác nhận tất cả ô đã bỏ tích.")
                else:
                    self.log("   ✅ Fallback màu sắc: tất cả ô đang ở trạng thái bỏ tích.")
                return True

            if attempt < retries - 1:
                self.log(
                    "   ⏳ Fallback màu sắc: còn {} ô nghi ngờ đang TÍCH, chờ {:.1f}s rồi kiểm tra lại...".format(
                        len(still_checked), delay
                    )
                )
                time.sleep(delay)

        if still_checked and allow_bruteforce:
            self.log(
                "   🔁 Fallback màu sắc: thử nhấp mạnh các ô khóa rồi kiểm tra lại..."
            )
            if self.brute_force_unlock_locks(cycles=3):
                time.sleep(delay)
                return self.verify_all_locks_unchecked(
                    retries=retries,
                    delay=delay,
                    allow_bruteforce=False,
                )

        self.log(
            "   ⚠️ Fallback màu sắc: phát hiện {} ô vẫn đang TÍCH sau {} lần kiểm tra.".format(
                len(still_checked), retries
            )
        )
        return False


if __name__ == "__main__":
    root = tk.Tk()
    app = AutoRefineApp(root)
    root.mainloop()
