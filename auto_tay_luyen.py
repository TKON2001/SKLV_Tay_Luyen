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
                {"name": "Chỉ số 1", "area": [0, 0, 0, 0], "lock_button": [0, 0], "desired_value": 0},
                {"name": "Chỉ số 2", "area": [0, 0, 0, 0], "lock_button": [0, 0], "desired_value": 0},
                {"name": "Chỉ số 3", "area": [0, 0, 0, 0], "lock_button": [0, 0], "desired_value": 0},
                {"name": "Chỉ số 4", "area": [0, 0, 0, 0], "lock_button": [0, 0], "desired_value": 0},
            ],
            "upgrade_area": [0, 0, 0, 0],
            "upgrade_button": [0, 0],
            "require_red": False
        }
        self.locked_stats = [False] * 4
        self.config_file = "config_tay_luyen.json"
        self.require_red_var = tk.BooleanVar(value=False)

        # --- Tạo giao diện ---
        self.create_widgets()
        self.load_config()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

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

        # Nút Tẩy Luyện
        ttk.Label(coords_frame, text="Nút Tẩy Luyện:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.refine_btn_label = ttk.Label(coords_frame, text="Chưa thiết lập")
        self.refine_btn_label.grid(row=0, column=1, sticky=tk.W)
        ttk.Button(coords_frame, text="Thiết lập", command=lambda: self.setup_coord("refine_button")).grid(row=0, column=2, padx=5)

        # Các chỉ số
        self.stat_entries = []
        for i in range(4):
            ttk.Separator(coords_frame, orient=tk.HORIZONTAL).grid(row=1 + i * 3, columnspan=6, sticky="ew", pady=5)
            
            # Tên và giá trị mong muốn
            ttk.Label(coords_frame, text=f"Chỉ số {i+1}:").grid(row=2 + i * 3, column=0, sticky=tk.W, pady=2)
            desired_val_entry = ttk.Entry(coords_frame, width=10)
            desired_val_entry.grid(row=2 + i * 3, column=1, sticky=tk.W)
            
            # Vùng đọc chỉ số
            area_label = ttk.Label(coords_frame, text="Vùng đọc: Chưa đặt")
            area_label.grid(row=2 + i * 3, column=2, sticky=tk.W, padx=10)
            ttk.Button(coords_frame, text="Đặt vùng", command=lambda i=i: self.setup_coord("stat_area", i)).grid(row=2 + i * 3, column=3, padx=5)
            
            # Nút khóa
            lock_label = ttk.Label(coords_frame, text="Nút khóa: Chưa đặt")
            lock_label.grid(row=2 + i * 3, column=4, sticky=tk.W, padx=10)
            ttk.Button(coords_frame, text="Đặt nút", command=lambda i=i: self.setup_coord("stat_lock", i)).grid(row=2 + i * 3, column=5, padx=5)
            
            # Hiển thị giá trị hiện tại
            current_label = ttk.Label(coords_frame, text="Giá trị hiện tại: --")
            current_label.grid(row=3 + i * 3, column=0, columnspan=6, sticky=tk.W)
            
            self.stat_entries.append({
                "desired_value": desired_val_entry,
                "area_label": area_label,
                "lock_label": lock_label,
                "current_label": current_label
            })

        # Khu vực nhận diện nút Thăng Cấp
        ttk.Separator(coords_frame, orient=tk.HORIZONTAL).grid(row=14, columnspan=6, sticky="ew", pady=6)
        ttk.Label(coords_frame, text="Vùng nút Thăng Cấp:").grid(row=15, column=0, sticky=tk.W, pady=2)
        self.upgrade_area_label = ttk.Label(coords_frame, text="Chưa đặt")
        self.upgrade_area_label.grid(row=15, column=1, columnspan=3, sticky=tk.W)
        ttk.Button(coords_frame, text="Đặt vùng thăng cấp", command=lambda: self.setup_coord("upgrade_area")).grid(row=15, column=4, padx=5)
        
        ttk.Label(coords_frame, text="Nút Thăng Cấp:").grid(row=16, column=0, sticky=tk.W, pady=2)
        self.upgrade_btn_label = ttk.Label(coords_frame, text="Chưa đặt")
        self.upgrade_btn_label.grid(row=16, column=1, sticky=tk.W)
        ttk.Button(coords_frame, text="Thiết lập", command=lambda: self.setup_coord("upgrade_button")).grid(row=16, column=2, padx=5)

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
                if (coord_type == "stat_area" and len(positions) == 2) or \
                   (coord_type != "stat_area" and len(positions) == 1):
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
            self.stat_entries[index]["area_label"].config(text=f"Vùng đọc: Đã đặt ({area[2]}x{area[3]})")
            self.save_config()
        elif coord_type == "stat_lock" and positions:
            self.config["stats"][index]["lock_button"] = list(positions[0])
            self.stat_entries[index]["lock_label"].config(text=f"Nút khóa: Đã đặt")
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
                        self.log(f"Chỉ số {i+1}: '{text.strip()}' -> {current_value:.2f}% / MAX {range_max:.2f}%")
                        self.stat_entries[i]["current_label"].config(text=f"Hiện tại: {current_value:.2f}% / Max: {range_max:.2f}%")
                    else:
                        self.log(f"Chỉ số {i+1}: '{text.strip()}' -> {current_value} / MAX {range_max}")
                        self.stat_entries[i]["current_label"].config(text=f"Hiện tại: {current_value} / Max: {range_max}")
                else:
                    if is_percent:
                        self.log(f"Chỉ số {i+1}: '{text.strip()}' -> {current_value:.2f}%")
                        self.stat_entries[i]["current_label"].config(text=f"Giá trị hiện tại: {current_value:.2f}%")
                    else:
                        self.log(f"Chỉ số {i+1}: '{text.strip()}' -> {current_value}")
                        self.stat_entries[i]["current_label"].config(text=f"Giá trị hiện tại: {current_value}")
                
                # Lưu ảnh để debug
                processed_img.save(f"debug_stat_{i+1}.png")
                self.log(f"Đã lưu ảnh debug: debug_stat_{i+1}.png")
                
            except Exception as e:
                self.log(f"Lỗi khi đọc chỉ số {i+1}: {e}")

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
            if success_unlock:
                self.locked_stats = [False] * 4
                self.log("✅ Đã thăng cấp thành công và bỏ tích các dòng!")
                return True

            self.log("⚠️ Đã thăng cấp nhưng không bỏ tích hết các dòng, sẽ thử lại.")
            time.sleep(0.8)

        self.log("❌ Thử thăng cấp nhiều lần nhưng chưa thành công hoàn toàn.")
        return False

    def is_lock_checked(self, lock_pos: list[int] | tuple[int, int]) -> bool:
        # Phân tích hình ảnh của ô khóa để xác định trạng thái: tìm dấu tích vàng
        try:
            lx, ly = int(lock_pos[0]), int(lock_pos[1])
        except Exception:
            return False
        
        # Tăng kích thước vùng chụp để bắt được dấu tích rõ hơn
        box_size = 34
        half = box_size // 2
        left = max(0, lx - half)
        top = max(0, ly - half)
        snap = pyautogui.screenshot(region=(left, top, box_size, box_size))

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
                if r > 180 and g > 180 and b < 120:
                    yellow_pixels += 1
                    # Vàng sáng (dấu tích)
                    if r > 220 and g > 220 and b < 80:
                        bright_yellow_pixels += 1

                # Kiểm tra theo HSV để bao phủ trường hợp màu vàng đậm/nhạt
                h, s, v = colorsys.rgb_to_hsv(r / 255.0, g / 255.0, b / 255.0)
                if 0.11 <= h <= 0.20 and s >= 0.35 and v >= 0.50:
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

        # Có dấu tích vàng nếu có đủ pixel vàng sáng
        has_checkmark = (
            bright_yellow_ratio > 0.025
            or yellow_ratio > 0.10
            or hsv_yellow_ratio > 0.045
        )
        
        status = "TÍCH" if has_checkmark else "TRỐNG"
        self.log(f"   Kết quả Lock {lock_pos}: {status}")
        
        return has_checkmark

    def ensure_unchecked(self, lock_pos: list[int] | tuple[int, int], *, force: bool = False) -> bool:
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
            if not force and not self.is_lock_checked(lock_pos):
                self.log(f"   ✅ Lock {lock_pos} đã ở trạng thái bỏ tích")
                return True
            elif force:
                self.log(f"   🔁 Force bỏ tích Lock {lock_pos} bất kể trạng thái nhận diện")

            # Thử click với nhiều vị trí khác nhau để tăng độ chính xác
            click_positions = [
                (x, y),           # Vị trí chính xác
                (x+1, y),         # Lệch phải 1px
                (x-1, y),         # Lệch trái 1px
                (x, y+1),         # Lệch xuống 1px
                (x, y-1),         # Lệch lên 1px
                (x+2, y+2),       # Lệch chéo
            ]
            
            for attempt in range(5):  # Tăng số lần thử
                self.log(f"   Thử bỏ tích lần {attempt + 1}/5...")
                
                for offset_x, offset_y in click_positions:
                    try:
                        # Click với vị trí offset
                        pyautogui.moveTo(offset_x, offset_y)
                        time.sleep(0.2) # Chờ trước khi click
                        pyautogui.click(offset_x, offset_y)
                        time.sleep(0.8)  # Tăng thời gian chờ UI cập nhật
                        
                        # Kiểm tra kết quả
                        if not self.is_lock_checked(lock_pos):
                            self.log(f"   ✅ Đã bỏ tích thành công Lock {lock_pos}")
                            return True
                            
                    except Exception as e:
                        self.log(f"   ⚠️ Lỗi khi click Lock {lock_pos}: {e}")
                        continue
                
                # Nếu vẫn chưa bỏ tích được, thử click mạnh hơn
                if attempt < 4:
                    time.sleep(0.5)
                    try:
                        # Double click để chắc chắn
                        pyautogui.doubleClick(x, y)
                        time.sleep(0.4)
                        if not self.is_lock_checked(lock_pos):
                            self.log(f"   ✅ Đã bỏ tích bằng double click Lock {lock_pos}")
                            return True
                    except Exception:
                        pass
            
            # Kiểm tra lần cuối
            final_check = not self.is_lock_checked(lock_pos)
            if final_check:
                self.log(f"   ✅ Cuối cùng đã bỏ tích Lock {lock_pos}")
                return True
            else:
                self.log(f"   ❌ Không thể bỏ tích Lock {lock_pos} sau 5 lần thử")
                return False

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
                if self.is_lock_checked(lock_pos):
                    pending.append((idx, lock_pos))

        if not pending:
            # Không có ô nào cần bỏ tích
            return True

        self.log("🔄 Đang bỏ tích các ô khóa...")

        for attempt in range(max_attempts):
            self.log(f"   Lần thử bỏ tích: {attempt + 1}/{max_attempts}")
            next_pending: list[tuple[int, list[int] | tuple[int, int]]] = []

            for idx, lock_pos in pending:
                if self.ensure_unchecked(lock_pos, force=force_click):
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

    def normalize_vi(self, s: str) -> str:
        # Bỏ dấu tiếng Việt để so khớp văn bản đơn giản
        nfkd = unicodedata.normalize('NFKD', s)
        ascii_str = ''.join([c for c in nfkd if not unicodedata.combining(c)])
        return ascii_str.lower()

    def clean_ocr_text(self, text: str) -> str:
        # Làm sạch một số lỗi OCR phổ biến
        s = text
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

    def fix_percent_current_with_max(self, current_value: float, range_max: float | None) -> float:
        # Sửa lỗi rơi dấu chấm: 1485 -> 148.5 hoặc 14.85 nếu gần range_max
        if range_max is None:
            return current_value
        candidates = [current_value, current_value / 10.0, current_value / 100.0]
        best = current_value
        best_delta = abs(current_value - range_max)
        for c in candidates:
            delta = abs(c - range_max)
            if delta < best_delta:
                best_delta = delta
                best = c
        return best

    def is_read_valid(self, current_value, range_max, is_percent: bool) -> bool:
        if is_percent:
            # Giá trị % hợp lệ trong [0, 200]
            if current_value < 0 or current_value > 200:
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
        is_percent = '%' in cleaned

        # 1) Tìm cặp min-max dạng phần trăm trong ngoặc: (a%-b%)
        pm = re.search(r'\((\d+(?:\.\d+)?)%\s*-\s*(\d+(?:\.\d+)?)%\)?', cleaned)
        if pm:
            try:
                a = float(pm.group(1))
                b = float(pm.group(2))
                range_max = max(a, b)
                is_percent = True
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
                except:
                    range_max = None

        # 2) Lấy số sau dấu '+' dạng phần trăm: +x.x%
        plus_percent = re.search(r'\+\s*(\d+(?:\.\d+)?)\s*%\b', cleaned)
        plus_number  = re.search(r'\+\s*(\d+(?:\.\d+)?)\b(?!%)', cleaned)
        if plus_percent:
            current_value = float(plus_percent.group(1))
            is_percent = True
        elif plus_number:
            if is_percent:
                # Nếu đã xác định là phần trăm từ cặp (min%-max%) mà dấu % sau dấu + bị mất,
                # vẫn đọc giá trị dạng số thực để so sánh chính xác A == C
                current_value = float(plus_number.group(1))
            else:
                current_value = int(float(plus_number.group(1)))
        else:
            # Fallback an toàn
            nums = re.findall(r'(\d+(?:\.\d+)?)', cleaned)
            if nums:
                if is_percent:
                    current_value = float(nums[0])
                else:
                    current_value = int(float(nums[0]))
            else:
                return (0.0 if is_percent else 0), None, is_percent

        # Sanity cho phần trăm: đưa về khoảng 0..200 nếu OCR dính thừa chữ số (ví dụ 19604 -> 196.04)
        def normalize_percent(x: float) -> float:
            val = x
            # Nếu quá lớn, chia 10 cho đến khi <= 200 hoặc 2 lần
            for _ in range(3):
                if val <= 200:
                    break
                val = val / 10.0
            return val

        # Đồng bộ kiểu dữ liệu current/range_max
        if is_percent:
            if isinstance(current_value, int):
                current_value = float(current_value)
            if isinstance(range_max, int):
                range_max = float(range_max)
            current_value = normalize_percent(current_value)
            if range_max is not None:
                range_max = normalize_percent(range_max)
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
            # So sánh theo định dạng hiển thị (2 chữ số thập phân). A phải bằng C sau khi làm tròn 2 số.
            return round(float(current_value), 2) == round(float(range_max), 2)
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

                # Kiểm tra xem có ô nào đang tích không (sau khi thăng cấp)
                leftover_indices: list[int] = []
                for idx, stat_cfg in enumerate(self.config["stats"]):
                    if self.locked_stats[idx]:
                        continue
                    if sum(stat_cfg.get("lock_button", [0, 0])) > 0 and self.is_lock_checked(stat_cfg["lock_button"]):
                        leftover_indices.append(idx)

                if leftover_indices:
                    self.log(
                        f"⚠️ Phát hiện {len(leftover_indices)} ô khóa vẫn đang tích - đang bỏ tích lại trước khi tẩy luyện..."
                    )
                    if self.unlock_all_locks(max_attempts=3, force_click=True, target_indices=leftover_indices):
                        self.log("🔄 Đã bỏ tích các ô còn lại, tiếp tục tẩy luyện sau 0.6s...")
                        time.sleep(0.6)
                        self.log("🔄 Đã bỏ tích các ô còn lại, tiếp tục tẩy luyện sau 1s...")
                        time.sleep(1.0)
                    else:
                        self.log("❌ Không thể bỏ tích toàn bộ ô khóa, tạm dừng 2s rồi thử lại...")
                        time.sleep(2.0)
                    continue

                # Nhấp nút Tẩy Luyện với delay dài hơn
                pyautogui.click(self.config["refine_button"])
                self.log(">> Đã nhấn Tẩy Luyện")
                time.sleep(1.6) # Rút ngắn thời gian chờ UI load hoàn toàn

                all_done = True
                for i, stat in enumerate(self.config["stats"]):
                    if self.locked_stats[i]:
                        self.log(f"   Chỉ số {i+1}: Đã khóa")
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
                            self.log(f"   Chỉ số {i+1}: '{text.strip()}' -> {current_value:.2f}% / Mục tiêu {target:.2f}%  => {'ĐẠT' if achieved else 'chưa đạt'}")
                        else:
                            self.log(f"   Chỉ số {i+1}: '{text.strip()}' -> {current_value} / Mục tiêu {target}  => {'ĐẠT' if achieved else 'chưa đạt'}")
                    else:
                        if is_percent:
                            self.log(f"   Chỉ số {i+1}: '{text.strip()}' -> {current_value:.2f}%")
                        else:
                            self.log(f"   Chỉ số {i+1}: '{text.strip()}' -> {current_value}")
                    
                    # Cập nhật GUI
                    if range_max is not None:
                        if is_percent:
                            self.root.after(0, lambda i=i, val=current_value, mx=range_max: self.stat_entries[i]["current_label"].config(text=f"Hiện tại: {val:.2f}% / Max: {mx:.2f}%"))
                        else:
                            self.root.after(0, lambda i=i, val=current_value, mx=range_max: self.stat_entries[i]["current_label"].config(text=f"Hiện tại: {val} / Max: {mx}"))
                    else:
                        if is_percent:
                            self.root.after(0, lambda i=i, val=current_value: self.stat_entries[i]["current_label"].config(text=f"Giá trị hiện tại: {val:.2f}%"))
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
                            time.sleep(1.0) # Chờ UI cập nhật sau khi khóa
                        else:
                            self.log(f"   → Đạt mục tiêu nhưng chưa xác nhận chữ đỏ, bỏ qua")

                # Kiểm tra điều kiện thăng cấp: CHỈ khi đủ 4 dòng MAX trở lên
                num_locked = sum(1 for v in self.locked_stats if v)
                self.log(f"   Số dòng đã khóa: {num_locked}/4")
                
                # Thăng cấp khi đủ 3 dòng MAX trở lên
                if num_locked >= 3:
                    if self.is_upgrade_available():
                        self.log("🎯 Đủ 3 dòng MAX và nút Thăng Cấp active - Bắt đầu thăng cấp!")
                    else:
                        self.log("🎯 Đủ 3 dòng MAX - Thử thăng cấp (fallback)...")

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

                        success_unlock = self.unlock_all_locks(max_attempts=6, force_click=True)
                        self.locked_stats = [False] * 4

                        if success_unlock:
                            self.log("✅ Đã thăng cấp thành công và bỏ tích các dòng!")
                            self.log("🔄 Tự động tiếp tục tẩy luyện với mục tiêu mới...")
                            self.log("💡 Tool sẽ tự động tẩy luyện liên tục cho đến khi bạn dừng thủ công.")
                            time.sleep(1.0)
                            continue
                        else:
                            self.log(
                                "⚠️ Không thể xác nhận bỏ tích hết các dòng sau thăng cấp. Tránh tẩy luyện sai nên tool sẽ dừng để bạn kiểm tra lại."
                            )
                            self.is_running = False
                            self.root.after(0, self._update_button_states)
                            time.sleep(1.0)
                            continue
                    else:
                        self.log("⏳ Chưa thể hoàn tất thăng cấp, sẽ thử lại sau 1.0s.")
                        time.sleep(1.0)

                    continue

                elif num_locked < 3:
                    self.log(f"📊 Chưa đủ 3 dòng MAX ({num_locked}/4) - Tiếp tục tẩy luyện...")

                # Nếu không có điều kiện thăng cấp, tiếp tục chu kỳ bình thường
                if all_done:
                    self.log("ℹ️ Tất cả chỉ số đã được xử lý trong chu kỳ này")

                time.sleep(1.0) # Rút ngắn thời gian nghỉ giữa các chu kỳ để tăng tốc
            
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
        
        # Cập nhật giá trị mong muốn từ GUI
        for i in range(4):
            try:
                val_text = self.stat_entries[i]["desired_value"].get()
                if val_text.strip():
                    val = int(val_text)
                    self.config["stats"][i]["desired_value"] = val
                else:
                    self.config["stats"][i]["desired_value"] = 0
            except ValueError:
                messagebox.showerror("Lỗi", f"Giá trị mong muốn của Chỉ số {i+1} không hợp lệ!")
                return
        
        # Kiểm tra có ít nhất một chỉ số có vùng đọc và nút khóa (không bắt buộc desired_value)
        has_configured_stats = any(sum(stat["area"]) > 0 and sum(stat["lock_button"]) > 0 
                                 for stat in self.config["stats"])
        
        if not has_configured_stats:
            messagebox.showerror("Lỗi", "Vui lòng thiết lập ít nhất một chỉ số với vùng đọc và nút khóa!")
            return
        
        self.is_running = True
        self.locked_stats = [False] * 4 # Reset trạng thái khóa
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


if __name__ == "__main__":
    root = tk.Tk()
    app = AutoRefineApp(root)
    root.mainloop()
