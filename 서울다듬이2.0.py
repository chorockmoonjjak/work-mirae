import customtkinter as ctk
from tkinter import filedialog, messagebox, Menu, simpledialog
from PIL import Image
import os
import win32com.client as win32
import pandas as pd
import re
import threading
import shutil
import sys
import json
import ctypes
from ctypes import wintypes

import google.generativeai as genai

class AdminLingoPurifierApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- 디자인 시스템 (Design System) ---
        self.MORNING_YELLOW = "#FE9C00"
        self.DARK_YELLOW = "#E08900" 
        self.LIGHT_YELLOW = "#FFF0D4"
        self.MIDNIGHT_BLUE = "#2C3E50"
        self.BG_COLOR = "#F5F6F8"
        self.SURFACE_COLOR = "#FFFFFF"

        # --- 윈도우 기본 설정 ---
        self.title("서울 다듬이 v2.0")
        self.geometry("1400x880") 
        self.configure(fg_color=self.BG_COLOR)
        ctk.set_appearance_mode("Light")
        ctk.set_default_color_theme("blue")

        # --- 내부 상태 변수 ---
        self.admin_term_db = {}
        self.pair_metadata = {} 
        self.ai_metadata = {}
        self.loaded_hwp_path = None
        self.replacements_to_apply = [] 
        self.cached_original_content = ""
        self.text_font_size = 15 # 폰트 조절용 상태 변수
        self.load_database()

        # --- 메인 프레임 설정 ---
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)

        # === 1. 타이틀 & 로고 이미지 프레임 ===
        self.header_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, pady=(20, 0), sticky="ew")
        
        try:
            base_dir = sys._MEIPASS
        except Exception:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            
        ui_dir = os.path.join(base_dir, 'ui')

        # 가운데 타이틀 & 서브타이틀 (정중앙 배치)
        self.title_center_frame = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        self.title_center_frame.pack(pady=5)

        self.title_label = ctk.CTkLabel(
            self.title_center_frame, 
            text="✨ 서울 다듬이 2.0", 
            font=ctk.CTkFont(family="Inter", size=36, weight="bold"),
            text_color=self.MIDNIGHT_BLUE
        )
        self.title_label.pack(side="top")
        
        self.subtitle_label = ctk.CTkLabel(
            self.title_center_frame,
            text="공무원을 위한 스마트 문서 파트너!\n복잡한 행정용어를 클릭 한 번으로 쉽고 바르게, AI 문장 다듬기로 자연스럽게 순화해보세요.",
            font=ctk.CTkFont(family="Inter", size=14, weight="normal"),
            text_color=self.MORNING_YELLOW
        )
        self.subtitle_label.pack(side="top", pady=(0, 5))

        # 좌측 로고 이미지 (EN_Basic_5.png) - place()로 절대 배치
        left_logo_path = os.path.join(ui_dir, 'EN_Basic_5.png')
        if os.path.exists(left_logo_path):
            try:
                left_img = Image.open(left_logo_path)
                left_logo = ctk.CTkImage(light_image=left_img, dark_image=left_img, size=(200, 46))
                self.left_logo_label = ctk.CTkLabel(self.header_frame, image=left_logo, text="")
                self.left_logo_label.place(x=40, rely=0.5, anchor="w")
            except Exception as e:
                print(f"좌측 로고 로딩 실패: {e}")

        # 우측 해치 이미지 (img-haechi01-2d.png) - place()로 절대 배치
        right_logo_path = os.path.join(ui_dir, 'img-haechi01-2d.png')
        if os.path.exists(right_logo_path):
            try:
                right_img = Image.open(right_logo_path)
                right_logo = ctk.CTkImage(light_image=right_img, dark_image=right_img, size=(55, 60))
                self.right_logo_label = ctk.CTkLabel(self.header_frame, image=right_logo, text="")
                self.right_logo_label.place(relx=1.0, rely=0.5, anchor="e", x=-40)
            except Exception as e:
                print(f"우측 로고 로딩 실패: {e}")

        # === 2. 버튼 및 컨트롤 프레임 ===
        self.control_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.control_frame.grid(row=1, column=0, padx=40, pady=10, sticky="ew")
        self.control_frame.grid_columnconfigure(0, weight=1)

        self.filepath_entry = ctk.CTkEntry(
            self.control_frame, 
            placeholder_text="🔎 변환할 파일을 찾아 선택하거나 내용을 복사해서 붙여넣어주세요 (.txt, .hwp, .hwpx 지원)",
            height=55,
            font=("Inter", 16),
            border_color=self.MORNING_YELLOW,
            border_width=3,
            fg_color=self.SURFACE_COLOR,
            text_color=self.MIDNIGHT_BLUE,
            corner_radius=12
        )
        self.filepath_entry.grid(row=0, column=0, padx=(0, 15), sticky="ew")

        self.browse_button = ctk.CTkButton(
            self.control_frame, 
            text="📁 파일 찾기", 
            width=120, 
            height=55,
            command=self.browse_file,
            fg_color=self.MORNING_YELLOW,
            text_color="white",
            hover_color=self.DARK_YELLOW,
            corner_radius=12,
            font=ctk.CTkFont(family="Inter", size=15, weight="bold")
        )
        self.browse_button.grid(row=0, column=1, padx=(0, 5))

        self.run_button = ctk.CTkButton(
            self.control_frame, 
            text="🚀 즉시 변환", 
            width=120, 
            height=55,
            command=self.process_text,
            fg_color="#3498DB",
            hover_color="#2980B9",
            text_color="white",
            corner_radius=10,
            font=ctk.CTkFont(family="Inter", size=15, weight="bold")
        )
        self.run_button.grid(row=0, column=2, padx=5)
        
        self.undo_button = ctk.CTkButton(
            self.control_frame, 
            text="↩ 되돌리기", 
            width=120, 
            height=55,
            command=self.undo_text,
            fg_color="#34495E",
            hover_color="#2C3E50",
            text_color="white",
            corner_radius=10,
            font=ctk.CTkFont(family="Inter", size=15, weight="bold")
        )
        self.undo_button.grid(row=0, column=3, padx=5)

        self.save_button = ctk.CTkButton(
            self.control_frame, 
            text="💾 새로 저장", 
            width=120, 
            height=55,
            command=self.save_file,
            fg_color="#27ae60",
            hover_color="#219150",
            text_color="white",
            corner_radius=10,
            font=ctk.CTkFont(family="Inter", size=15, weight="bold")
        )
        self.save_button.grid(row=0, column=4, padx=(5, 10))

        self.memo_var = ctk.BooleanVar(value=True)
        self.memo_checkbox = ctk.CTkCheckBox(
            self.control_frame,
            text="변경메모 삽입",
            variable=self.memo_var,
            font=ctk.CTkFont(family="Inter", size=15, weight="bold"),
            text_color=self.MIDNIGHT_BLUE,
            border_color="#27ae60",
            hover_color="#a8e6cf",
            fg_color="#27ae60"
        )
        self.memo_checkbox.grid(row=0, column=5, padx=(0, 0), sticky="w")

        # === 3. 글자 크기 조정 프레임 (Row 2) ===
        self.settings_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.settings_frame.grid(row=2, column=0, padx=50, pady=(0, 5), sticky="e")
        
        self.font_size_label = ctk.CTkLabel(
            self.settings_frame,
            text="가", font=ctk.CTkFont(size=12, weight="bold"), text_color=self.MIDNIGHT_BLUE
        )
        self.font_size_label.pack(side="left", padx=5)
        
        self.font_slider = ctk.CTkSlider(
            self.settings_frame,
            from_=10, to=30, number_of_steps=20, width=150,
            command=self.on_font_scale_change,
            button_color=self.MORNING_YELLOW,
            button_hover_color=self.DARK_YELLOW
        )
        self.font_slider.set(self.text_font_size)
        self.font_slider.pack(side="left", padx=5)
        
        self.font_size_large_label = ctk.CTkLabel(
            self.settings_frame,
            text="가", font=ctk.CTkFont(size=20, weight="bold"), text_color=self.MIDNIGHT_BLUE
        )
        self.font_size_large_label.pack(side="left", padx=5)

        # === 4. 결과 프레임 (Card Style) ===
        self.result_frame = ctk.CTkFrame(self, fg_color=self.SURFACE_COLOR, corner_radius=18, border_width=1, border_color="#EAECEE")
        self.result_frame.grid(row=4, column=0, padx=40, pady=(0, 20), sticky="nsew")
        self.result_frame.grid_columnconfigure((0, 1), weight=1, uniform="equal_cols")
        self.result_frame.grid_rowconfigure(1, weight=1)

        self.original_label = ctk.CTkLabel(
            self.result_frame, 
            text="📄 원본 텍스트(클릭하여 치환, 드래그 후 우클릭하여 AI 다듬기)", 
            font=ctk.CTkFont(family="Inter", size=18, weight="bold"), 
            text_color=self.MIDNIGHT_BLUE
        )
        self.original_label.grid(row=0, column=0, pady=(20, 10))

        self.modified_label = ctk.CTkLabel(
            self.result_frame, 
            text="✨ 데스크탑 봇 수정 제안", 
            font=ctk.CTkFont(family="Inter", size=18, weight="bold"), 
            text_color=self.DARK_YELLOW
        )
        self.modified_label.grid(row=0, column=1, pady=(20, 10))

        self.original_text = ctk.CTkTextbox(
            self.result_frame, 
            font=("Malgun Gothic", 15), 
            border_width=1, 
            border_color="#D5DBDB", 
            fg_color="#FBFCFC",
            text_color="#2C3E50",
            wrap="word",
            corner_radius=12
        )
        self.original_text.grid(row=1, column=0, padx=(25, 12), pady=(0, 25), sticky="nsew")
        self.original_text._textbox.configure(undo=True, autoseparators=True, maxundo=-1)

        self.modified_text = ctk.CTkTextbox(
            self.result_frame, 
            font=("Malgun Gothic", 15), 
            border_width=1, 
            border_color=self.MORNING_YELLOW, 
            fg_color="#FFF9EF",
            text_color="#2C3E50",
            wrap="word",
            corner_radius=12
        )
        self.modified_text.grid(row=1, column=1, padx=(12, 25), pady=(0, 25), sticky="nsew")

        # --- 우클릭 컨텍스트 메뉴 설정 ---
        self.context_menu = Menu(self, tearoff=0)
        self.context_menu.add_command(label="✨ 선택한 문장 AI로 더 좋게 다듬기", command=self.ai_refine_text)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="↩ 변경사항 되돌리기", command=self.undo_text)

        self.original_text._textbox.bind("<Button-3>", self.show_context_menu)
        self.original_text._textbox.bind("<Control-z>", self.undo_text)

        # === 5. 하단 풋터 프레임 ===
        self.footer_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.footer_frame.grid(row=5, column=0, padx=40, pady=(0, 15), sticky="ew")
        self.footer_frame.grid_columnconfigure(2, weight=1)
        
        self.link_db = ctk.CTkLabel(self.footer_frame, text="🔗 순화어 라이브러리", font=ctk.CTkFont(size=13, underline=True), text_color="#2980b9", cursor="hand2")
        self.link_db.grid(row=0, column=0, padx=(0, 20), sticky="w")
        self.link_db.bind("<Button-1>", lambda e: self.open_file('행정순화용어_5,138건.xlsx'))
        
        self.dict_button = ctk.CTkLabel(self.footer_frame, text="📝 사용자 사전", font=ctk.CTkFont(size=13, underline=True), text_color="#2980b9", cursor="hand2")
        self.dict_button.grid(row=0, column=1, padx=(0, 20), sticky="w")
        self.dict_button.bind("<Button-1>", lambda e: self.manage_user_terms())
        
        self.link_readme = ctk.CTkLabel(self.footer_frame, text="🔗 설명서", font=ctk.CTkFont(size=13, underline=True), text_color="#2980b9", cursor="hand2")
        self.link_readme.grid(row=0, column=2, sticky="w")
        self.link_readme.bind("<Button-1>", lambda e: self.open_file('readme.txt'))
        
        sig_text = "Build by dhkim, jspark, bwim, wjjang, dhjeong | '26. 1st, MIRAEINJAEYANGSEONG MIRAE Team 4"
        self.signature_label = ctk.CTkLabel(
            self.footer_frame, 
            text=sig_text, 
            font=ctk.CTkFont(family="Inter", size=11), 
            text_color="#95A5A6", 
            justify="right"
        )
        self.signature_label.grid(row=0, column=3, sticky="e")



        self._last_y1 = None
        self._last_y2 = None
        self.after(50, self.sync_scrolling_poll)

    def on_font_scale_change(self, value):
        self.text_font_size = int(value)
        self.original_text.configure(font=("Malgun Gothic", self.text_font_size))
        self.modified_text.configure(font=("Malgun Gothic", self.text_font_size))
        
        # update colored tag fonts too
        for tag_name in self.pair_metadata.keys():
            self.original_text._textbox.tag_config(tag_name, font=ctk.CTkFont(size=self.text_font_size, weight="bold"))
            self.modified_text._textbox.tag_config(tag_name, font=ctk.CTkFont(size=self.text_font_size, weight="bold"))
            
        for tag_name in self.ai_metadata.keys():
            self.original_text._textbox.tag_config(tag_name, font=ctk.CTkFont(size=self.text_font_size, weight="bold"))

    def open_file(self, filename):
        try:
            try:
                base_dir = sys._MEIPASS
            except Exception:
                base_dir = os.path.dirname(os.path.abspath(__file__))
            filepath = os.path.join(base_dir, filename)
            
            if os.path.exists(filepath):
                os.startfile(filepath)
            else:
                messagebox.showwarning("오류", f"'{filename}' 파일을 찾을 수 없습니다.")
        except Exception as e:
            print(f"파일 열기 실패: {e}")

    def sync_scrolling_poll(self):
        try:
            y1 = self.original_text._textbox.yview()
            y2 = self.modified_text._textbox.yview()
            x1 = self.original_text._textbox.xview()
            x2 = self.modified_text._textbox.xview()

            if self._last_y1 != y1:
                self.modified_text._textbox.yview_moveto(y1[0])
                self._last_y1 = y1
                self._last_y2 = self.modified_text._textbox.yview()
            elif self._last_y2 != y2:
                self.original_text._textbox.yview_moveto(y2[0])
                self._last_y2 = y2
                self._last_y1 = self.original_text._textbox.yview()

            if not hasattr(self, '_last_x1'):
                self._last_x1 = x1
                self._last_x2 = x2

            if self._last_x1 != x1:
                self.modified_text._textbox.xview_moveto(x1[0])
                self._last_x1 = x1
                self._last_x2 = self.modified_text._textbox.xview()
            elif self._last_x2 != x2:
                self.original_text._textbox.xview_moveto(x2[0])
                self._last_x2 = x2
                self._last_x1 = self.original_text._textbox.xview()
        except:
            pass
        self.after(30, self.sync_scrolling_poll)

    def show_context_menu(self, event):
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def restore_tags(self):
        self.replacements_to_apply.clear()
        
        for tag_name, meta in self.pair_metadata.items():
            mark_name = meta["mark"]
            before = meta["before"]
            after = meta["after"]
            m_mark = meta["m_mark"]
            after_options = [opt.strip() for opt in after.split(",") if opt.strip()]
            
            try:
                start_idx = self.original_text._textbox.index(mark_name)
                m_start_idx = self.modified_text._textbox.index(m_mark)
            except Exception: continue
            
            self.original_text._textbox.tag_remove(tag_name, "1.0", "end")
            self.modified_text._textbox.tag_remove(tag_name, "1.0", "end")
            
            if self.original_text._textbox.get(start_idx, f"{start_idx}+{len(before)}c") == before:
                self.original_text._textbox.tag_add(tag_name, start_idx, f"{start_idx}+{len(before)}c")
                self.original_text._textbox.tag_config(tag_name, background="", foreground="#e74c3c", font=ctk.CTkFont(size=self.text_font_size, weight="bold"))
                
                self.modified_text._textbox.tag_add(tag_name, m_start_idx, f"{m_start_idx}+{len(after)}c")
                self.modified_text._textbox.tag_config(tag_name, background="", foreground="#2980b9", font=ctk.CTkFont(size=self.text_font_size, weight="bold"))
            else:
                matched_opt = None
                for opt in after_options:
                    if self.original_text._textbox.get(start_idx, f"{start_idx}+{len(opt)}c") == opt:
                        matched_opt = opt
                        break

                if matched_opt:
                    self.original_text._textbox.tag_add(tag_name, start_idx, f"{start_idx}+{len(matched_opt)}c")
                    self.original_text._textbox.tag_config(tag_name, background="", foreground="#27ae60", font=ctk.CTkFont(size=self.text_font_size, weight="bold"))
                    
                    offset = after.find(matched_opt)
                    if offset != -1:
                        m_opt_start = f"{m_start_idx}+{offset}c"
                        m_opt_end = f"{m_start_idx}+{offset+len(matched_opt)}c"
                        self.modified_text._textbox.tag_add(tag_name, m_opt_start, m_opt_end)
                        self.modified_text._textbox.tag_config(tag_name, background="", foreground="#27ae60", font=ctk.CTkFont(size=self.text_font_size, weight="bold"))
                    
                    if matched_opt != before:
                        self.replacements_to_apply.append({
                            "before": before, "after": matched_opt, "memo": f"행정순화어 반영: {before} → {matched_opt}"
                        })
                
        for tag_name, meta in self.ai_metadata.items():
            mark_name = meta["mark"]
            before = meta["before"]
            after = meta["after"]
            showing_ai = meta.get("showing_ai", True)
            try:
                start_idx = self.original_text._textbox.index(mark_name)
            except Exception: continue
            
            self.original_text._textbox.tag_remove(tag_name, "1.0", "end")
            
            if showing_ai:
                if self.original_text._textbox.get(start_idx, f"{start_idx}+{len(after)}c") == after:
                    self.original_text._textbox.tag_add(tag_name, start_idx, f"{start_idx}+{len(after)}c")
                    self.original_text._textbox.tag_config(tag_name, foreground="#8e44ad", background="#f4edfb", font=ctk.CTkFont(size=self.text_font_size, weight="bold"))
                    if after != before:
                        self.replacements_to_apply.append({
                            "before": before, "after": after, "memo": "💡 AI 문장 다듬기 결과"
                        })
            else:
                if self.original_text._textbox.get(start_idx, f"{start_idx}+{len(before)}c") == before:
                    self.original_text._textbox.tag_add(tag_name, start_idx, f"{start_idx}+{len(before)}c")
                    self.original_text._textbox.tag_config(tag_name, foreground="#999999", background="#f0f0f0", font=ctk.CTkFont(size=self.text_font_size, weight="bold"))

    def _get_current_replacements(self):
        result = []
        for tag_name, meta in self.pair_metadata.items():
            mark_name = meta["mark"]
            before = meta["before"]
            after = meta["after"]
            after_options = [opt.strip() for opt in after.split(",") if opt.strip()]

            try:
                start_idx = self.original_text._textbox.index(mark_name)
            except Exception:
                continue

            current_as_before = self.original_text._textbox.get(start_idx, f"{start_idx}+{len(before)}c")
            if current_as_before == before:
                continue

            matched_opt = None
            for opt in after_options:
                if opt == before: continue
                current = self.original_text._textbox.get(start_idx, f"{start_idx}+{len(opt)}c")
                if current == opt:
                    matched_opt = opt
                    break

            if matched_opt:
                result.append({
                    "before": before,
                    "after": matched_opt,
                    "memo": f"행정순화어 반영: {before} \u2192 {matched_opt}",
                    "type": "admin"
                })

        for tag_name, meta in self.ai_metadata.items():
            mark_name = meta["mark"]
            before = meta["before"]
            after = meta["after"]
            showing_ai = meta.get("showing_ai", True)

            if not showing_ai: continue

            try:
                start_idx = self.original_text._textbox.index(mark_name)
            except Exception: continue

            current_as_before = self.original_text._textbox.get(start_idx, f"{start_idx}+{len(before)}c")
            if current_as_before == before: continue

            current = self.original_text._textbox.get(start_idx, f"{start_idx}+{len(after)}c")
            if current == after and after != before:
                result.append({
                    "before": before,
                    "after": after,
                    "memo": "\U0001f4a1 AI 문장 다듬기 결과",
                    "type": "ai"
                })

        return result

    def undo_text(self, event=None):
        try:
            self.original_text._textbox.edit_undo()
        except Exception:
            if not event:
                messagebox.showinfo("안내", "더 이상 되돌릴 작업이 없습니다.")
            return "break"
            
        self.restore_tags()
        return "break"

    def ai_refine_text(self):
        try:
            sel_start = self.original_text._textbox.index("sel.first")
            sel_end = self.original_text._textbox.index("sel.last")
            selected_text = self.original_text._textbox.get(sel_start, sel_end)
        except Exception:
            messagebox.showwarning("주의", "먼저 다듬을 문장을 드래그하여 선택해 주세요.")
            return
            
        if not selected_text.strip(): return
            
        api_key = os.environ.get("GOOGLE_API_KEY", "")
        if not api_key:
            api_key = simpledialog.askstring("API 키 필요", "Gemini API 키를 입력해주세요 (최초 1회):\n발급처: https://aistudio.google.com/app/apikey")
            if not api_key: return
            os.environ["GOOGLE_API_KEY"] = api_key

        def run_ai():
            try:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel('gemini-2.5-flash')
                self.after(0, lambda: self.title("✨ 서울 다듬이 v2.0 - 🤖 AI가 처리 중..."))
                prompt = (
                    "다음 주어진 문장을 공공기관 및 일반 사용자가 읽기 쉽고, 자연스럽고, 명확한 표현으로 수정해줘. 단, 원문의 형식을 유지해주세요 "
                    "설명 없이 오직 수정된 최종 결과 문장만을 출력해:\n\n"
                    f"{selected_text}"
                )
                response = model.generate_content(prompt)
                ai_text = response.text.strip()
                self.after(0, self.apply_ai_text, sel_start, sel_end, ai_text, selected_text)
            except Exception as e:
                err_msg = str(e)
                korean_msg = self._translate_api_error(err_msg)
                self.after(0, lambda msg=korean_msg: messagebox.showerror("AI 오류", msg))
            finally:
                self.after(0, lambda: self.title("✨ 서울 다듬이 v2.0"))
                
        threading.Thread(target=run_ai, daemon=True).start()

    def _translate_api_error(self, err_msg: str) -> str:
        msg_lower = err_msg.lower()
        if "429" in err_msg or "resource_exhausted" in msg_lower or "quota" in msg_lower or "rate limit" in msg_lower:
            return "🚫 AI API 사용량(할당량)이 초과되었습니다."
        if "401" in err_msg or "403" in err_msg or "api_key" in msg_lower or "invalid" in msg_lower or "permission" in msg_lower:
            return "🔑 API 키 인증에 실패하였습니다."
        if "connection" in msg_lower or "timeout" in msg_lower or "network" in msg_lower or "unreachable" in msg_lower:
            return "🌐 네트워크 연결 오류가 발생하였습니다."
        if "500" in err_msg or "503" in err_msg or "internal" in msg_lower or "unavailable" in msg_lower:
            return "⚠️ AI 서버 내부 오류가 발생하였습니다."
        if "safety" in msg_lower or "blocked" in msg_lower or "harm" in msg_lower:
            return "🚨 AI 안전 정책에 의해 요청이 차단되었습니다."
        return f"❌ AI 처리 중 오류가 발생하였습니다.\n\n• 오류 내용: {err_msg}"

    def apply_ai_text(self, sel_start, sel_end, ai_text, original_text):
        textbox = self.original_text._textbox
        textbox.edit_separator()
        textbox.delete(sel_start, sel_end)
        textbox.insert(sel_start, ai_text)
        
        ai_tag = f"ai_tag_{len(self.ai_metadata)}"
        ai_mark = f"mark_{ai_tag}"
        textbox.mark_set(ai_mark, sel_start)
        textbox.mark_gravity(ai_mark, ctk.LEFT)
        
        self.ai_metadata[ai_tag] = {
            "mark": ai_mark,
            "before": original_text,
            "after": ai_text,
            "showing_ai": True
        }
        
        textbox.edit_separator()
        self.restore_tags()
        self._bind_ai_tag(ai_tag)

    def _bind_ai_tag(self, tag_name):
        textbox = self.original_text._textbox

        def on_ai_enter(e):
            meta = self.ai_metadata.get(tag_name)
            if not meta: return
            textbox.tag_config(tag_name, background="#e8d5f7", foreground="#6c1fa6")

        def on_ai_leave(e):
            self.restore_tags()

        def on_ai_click(e):
            meta = self.ai_metadata.get(tag_name)
            if not meta: return
            ranges = textbox.tag_ranges(tag_name)
            if not ranges: return
            start_idx, end_idx = ranges[0], ranges[1]

            showing_ai = meta.get("showing_ai", True)
            new_text = meta["before"] if showing_ai else meta["after"]

            textbox.edit_separator()
            textbox.delete(start_idx, end_idx)
            textbox.insert(start_idx, new_text)
            meta["showing_ai"] = not showing_ai
            textbox.edit_separator()
            self.restore_tags()
            self._bind_ai_tag(tag_name)

        textbox.tag_bind(tag_name, "<Enter>", on_ai_enter)
        textbox.tag_bind(tag_name, "<Leave>", on_ai_leave)
        textbox.tag_bind(tag_name, "<Button-1>", on_ai_click)

    def load_database(self):
        self.admin_term_db.clear()
        try:
            base_dir = sys._MEIPASS
        except Exception:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            
        db_path = os.path.join(base_dir, '행정순화용어_5,138건.xlsx')
        try:
            if os.path.exists(db_path):
                # 새로운 DB 구조에 맞게 수정 (0: 번호, 1: 순화대상어, 2: 순화어)
                df = pd.read_excel(db_path)
                for _, row in df.iterrows():
                    if len(row) >= 3:
                        before = str(row.iloc[1]).split('(')[0].strip()
                        after = str(row.iloc[2]).strip()
                        if before and after and before != 'nan' and after != 'nan':
                            self.admin_term_db[before] = after
        except Exception as e:
            print(f"DB Load Error: {e}")
            
        user_db_path = "user_terms.json"
        if os.path.exists(user_db_path):
            try:
                with open(user_db_path, "r", encoding="utf-8") as f:
                    user_terms = json.load(f)
                    for before, after in user_terms.items():
                        if before and after:
                            self.admin_term_db[before] = after
            except Exception as e:
                print(f"User DB Load Error: {e}")

    def manage_user_terms(self):
        user_db_path = "user_terms.json"
        
        user_terms = {}
        if os.path.exists(user_db_path):
            try:
                with open(user_db_path, "r", encoding="utf-8") as f:
                    user_terms = json.load(f)
            except: pass

        toplevel = ctk.CTkToplevel(self)
        toplevel.title("📝 사용자 사전 관리")
        toplevel.geometry("500x600")
        toplevel.configure(fg_color=self.BG_COLOR)
        
        # 창이 메인 윈도우 뒤로 숨지 않도록 포커스 및 상단 이동 강제
        toplevel.after(200, lambda: toplevel.focus_force())
        toplevel.after(200, lambda: toplevel.lift())

        header_lbl = ctk.CTkLabel(
            toplevel, text="사용자 맞춤 순화어 관리",
            font=ctk.CTkFont(family="Inter", size=20, weight="bold"),
            text_color=self.MIDNIGHT_BLUE
        )
        header_lbl.pack(pady=(20, 10))

        add_frame = ctk.CTkFrame(toplevel, fg_color=self.SURFACE_COLOR, corner_radius=10, border_width=1, border_color="#EAECEE")
        add_frame.pack(padx=20, fill="x", pady=10)

        before_entry = ctk.CTkEntry(add_frame, placeholder_text="순화 대상어 (바꿀 말)", width=150)
        before_entry.grid(row=0, column=0, padx=10, pady=10)

        arrow_lbl = ctk.CTkLabel(add_frame, text="→")
        arrow_lbl.grid(row=0, column=1)

        after_entry = ctk.CTkEntry(add_frame, placeholder_text="행정순화어 (권하는 말)", width=150)
        after_entry.grid(row=0, column=2, padx=10, pady=10)

        def add_term():
            b = before_entry.get().strip()
            a = after_entry.get().strip()
            if not b or not a: return
            user_terms[b] = a
            refresh_list()
            before_entry.delete(0, "end")
            after_entry.delete(0, "end")

        add_btn = ctk.CTkButton(
            add_frame, text="추가", width=60, 
            fg_color=self.MORNING_YELLOW, hover_color=self.DARK_YELLOW, 
            text_color="white", command=add_term, font=ctk.CTkFont(weight="bold")
        )
        add_btn.grid(row=0, column=3, padx=10, pady=10)

        list_frame = ctk.CTkScrollableFrame(toplevel, fg_color=self.SURFACE_COLOR, corner_radius=10, border_width=1, border_color="#EAECEE")
        list_frame.pack(padx=20, fill="both", expand=True, pady=(0, 10))

        def delete_term(b):
            if b in user_terms:
                del user_terms[b]
                refresh_list()

        def refresh_list():
            for widget in list_frame.winfo_children():
                widget.destroy()
            if not user_terms:
                ctk.CTkLabel(list_frame, text="등록된 사용자 단어가 없습니다.", text_color="gray").pack(pady=20)
                return
            for b, a in user_terms.items():
                row_f = ctk.CTkFrame(list_frame, fg_color="transparent")
                row_f.pack(fill="x", pady=2)
                lbl = ctk.CTkLabel(row_f, text=f"{b} → {a}", anchor="w", font=ctk.CTkFont(size=14))
                lbl.pack(side="left", padx=10)
                del_btn = ctk.CTkButton(row_f, text="삭제", width=40, fg_color="#e74c3c", hover_color="#c0392b", command=lambda x=b: delete_term(x))
                del_btn.pack(side="right", padx=10)

        refresh_list()

        def save_and_close():
            try:
                with open(user_db_path, "w", encoding="utf-8") as f:
                    json.dump(user_terms, f, ensure_ascii=False, indent=2)
                self.load_database()
                toplevel.destroy()
                if self.cached_original_content:
                    self.process_text()
            except Exception as e:
                messagebox.showerror("저장 오류", f"사용자 사전을 저장하는데 실패했습니다:\n{e}")

        action_frame = ctk.CTkFrame(toplevel, fg_color="transparent")
        action_frame.pack(padx=20, fill="x", pady=(0, 10))

        def upload_excel():
            filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
            if filepath:
                try:
                    df = pd.read_excel(filepath)
                    for _, row in df.iterrows():
                        if len(row) >= 2:
                            b = str(row.iloc[0]).split('(')[0].strip()
                            a = str(row.iloc[1]).strip()
                            if b and a and b != 'nan' and a != 'nan':
                                user_terms[b] = a
                    refresh_list()
                    messagebox.showinfo("완료", "엑셀 파일에서 순화어를 성공적으로 불러왔습니다.")
                    # 강제로 앞으로 다시 가져오기
                    toplevel.lift()
                    toplevel.focus_force()
                except Exception as e:
                    messagebox.showerror("오류", f"엑셀 파일을 반영하는 중 오류가 발생했습니다:\n{e}")

        def reset_to_default():
            if messagebox.askyesno("초기화", "모든 사용자 추가 순화어를 삭제하고 기본 라이브러리만 사용하시겠습니까?"):
                user_terms.clear()
                refresh_list()
                toplevel.lift()
                toplevel.focus_force()

        def download_template():
            filepath = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile="사용자_순화어_업로드_서식.xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="엑셀 서식 저장"
            )
            if filepath:
                try:
                    df = pd.DataFrame(columns=["바꿔야 할 말(순화대상어)", "권하는 말(행정순화어)"])
                    # 예시 데이터 하나 추가
                    df.loc[0] = ["리스트", "목록"]
                    df.to_excel(filepath, index=False)
                    messagebox.showinfo("완료", "엑셀 업로드용 서식이 저장되었습니다.")
                    toplevel.lift()
                    toplevel.focus_force()
                except Exception as e:
                    messagebox.showerror("오류", f"서식을 저장하는 중 오류가 발생했습니다:\n{e}")

        template_btn = ctk.CTkButton(
            toplevel, text="📥 엑셀 서식 다운로드", height=30,
            fg_color="transparent", text_color="#2980b9", hover_color="#ecf0f1",
            command=download_template, font=ctk.CTkFont(size=12, underline=True)
        )
        template_btn.pack(padx=20, anchor="e", pady=(0, 5))

        upload_btn = ctk.CTkButton(
            action_frame, text="📁 엑셀 일괄 업로드", height=35,
            fg_color="#27ae60", hover_color="#219150", 
            command=upload_excel, font=ctk.CTkFont(weight="bold")
        )
        upload_btn.pack(side="left", padx=(0, 5), expand=True, fill="x")

        reset_btn = ctk.CTkButton(
            action_frame, text="🔄 기본으로 초기화", height=35,
            fg_color="#e74c3c", hover_color="#c0392b", 
            command=reset_to_default, font=ctk.CTkFont(weight="bold")
        )
        reset_btn.pack(side="right", padx=(5, 0), expand=True, fill="x")

        save_btn = ctk.CTkButton(
            toplevel, text="💾 저장 및 닫기", height=40,
            fg_color=self.MIDNIGHT_BLUE, hover_color="#1A252F",
            command=save_and_close, font=ctk.CTkFont(weight="bold")
        )
        save_btn.pack(padx=20, fill="x", pady=(0, 20))

    def browse_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("지원 파일", "*.txt *.hwp *.hwpx"), ("Text files", "*.txt"), ("HWP files", "*.hwp *.hwpx")])
        if filepath:
            self.filepath_entry.delete(0, "end")
            self.filepath_entry.insert(0, filepath)

    def process_text(self):
        filepath = self.filepath_entry.get().strip()
        original_content = ""

        if not filepath:
            pasted_text = self.original_text.get("1.0", "end-1c").strip()
            if not pasted_text:
                messagebox.showwarning("주의", "파일을 선택하거나 원본 텍스트 창에 내용을 붙여넣으세요.")
                return
            original_content = pasted_text
            self.loaded_hwp_path = None
        else:
            ext = os.path.splitext(filepath)[1].lower()
            self.loaded_hwp_path = filepath if ext in [".hwp", ".hwpx"] else None

            try:
                if ext == ".txt":
                    with open(filepath, "r", encoding="utf-8") as f:
                        original_content = f.read()
                elif ext in [".hwp", ".hwpx"]:
                    original_content = self.extract_text_from_hwp(filepath)
                else:
                    messagebox.showerror("오류", "지원하지 않는 파일 형식입니다. (.txt, .hwp, .hwpx만 지원)")
                    return
            except Exception as e:
                messagebox.showerror("오류", f"파일을 읽는 중 오류 발생: {e}")
                return

        self.replacements_to_apply.clear()
        self.pair_metadata.clear()
        self.ai_metadata.clear()

        if not original_content:
            messagebox.showwarning("결과", "파일 또는 붙여넣은 텍스트의 내용을 읽을 수 없거나 비어 있습니다.")
            return

        self.cached_original_content = original_content

        self.original_text.delete("1.0", "end")
        self.modified_text.delete("1.0", "end")

        if not self.admin_term_db:
             self.original_text.insert("end", original_content)
             self.modified_text.insert("end", original_content)
             self.original_text._textbox.edit_reset()
             return

        keys = sorted(self.admin_term_db.keys(), key=len, reverse=True)
        keys_joined = "|".join(map(re.escape, keys))
        pattern = re.compile(r"(?<![가-힣a-zA-Z0-9])(?:" + keys_joined + r")")
        
        last_idx = 0
        pair_id = 0
        
        for match in pattern.finditer(original_content):
            start, end = match.span()
            before = match.group()
            after = self.admin_term_db[before]
            
            pre_text = original_content[last_idx:start]
            self.original_text.insert("end", pre_text)
            self.modified_text.insert("end", pre_text)
            
            tag_name = f"pair_{pair_id}"
            mark_name = f"mark_{tag_name}"
            
            o_start = self.original_text.index("end-1c")
            self.original_text.insert("end", before)
            
            self.original_text._textbox.mark_set(mark_name, o_start)
            self.original_text._textbox.mark_gravity(mark_name, ctk.LEFT)
            
            m_start = self.modified_text.index("end-1c")
            self.modified_text.insert("end", after)
            m_mark = f"m_mark_{tag_name}"
            self.modified_text._textbox.mark_set(m_mark, m_start)
            self.modified_text._textbox.mark_gravity(m_mark, ctk.LEFT)
            
            self.pair_metadata[tag_name] = {"mark": mark_name, "before": before, "after": after, "m_mark": m_mark}
            
            self.bind_hover_and_click(tag_name, before, after)
            pair_id += 1
            last_idx = end
            
        rem_text = original_content[last_idx:]
        self.original_text.insert("end", rem_text)
        self.modified_text.insert("end", rem_text)
        
        self.restore_tags()
        self.original_text._textbox.edit_reset()

    def bind_hover_and_click(self, tag_name, before_text, after_text):
        def on_enter(e):
            self.original_text._textbox.tag_config(tag_name, background="#FFF1AA", foreground="black")
            self.modified_text._textbox.tag_config(tag_name, background="#FFF1AA", foreground="black")
            
        def on_leave(e=None):
            self.restore_tags()

        def on_click(e):
            try:
                ranges = self.original_text._textbox.tag_ranges(tag_name)
                if not ranges: return
                start_idx, end_idx = ranges[0], ranges[1]
                
                current_text = self.original_text._textbox.get(start_idx, end_idx)
                after_options = [opt.strip() for opt in after_text.split(",") if opt.strip()]
                
                def apply_choice(choice):
                    self.original_text._textbox.edit_separator()
                    
                    real_ranges = self.original_text._textbox.tag_ranges(tag_name)
                    if not real_ranges: return
                    r_start, r_end = real_ranges[0], real_ranges[1]
                    
                    self.original_text.delete(r_start, r_end)
                    self.original_text.insert(r_start, choice)
                    self.original_text._textbox.edit_separator()
                    self.restore_tags()

                if current_text in after_options:
                    apply_choice(before_text)
                    return 
                
                if len(after_options) > 1:
                    choice_menu = Menu(self, tearoff=0)
                    for opt in after_options:
                        choice_menu.add_command(label=f"✔ '{opt}'(으)로 순화하기", command=lambda o=opt: apply_choice(o))
                    choice_menu.tk_popup(e.x_root, e.y_root)
                else:
                    apply_choice(after_options[0])

            except Exception as ex:
                print(f"클릭 반영 에러: {ex}")

        self.original_text._textbox.tag_bind(tag_name, "<Enter>", on_enter)
        self.original_text._textbox.tag_bind(tag_name, "<Leave>", on_leave)
        self.original_text._textbox.tag_bind(tag_name, "<Button-1>", on_click)
        self.modified_text._textbox.tag_bind(tag_name, "<Enter>", on_enter)
        self.modified_text._textbox.tag_bind(tag_name, "<Leave>", on_leave)
        self.modified_text._textbox.tag_bind(tag_name, "<Button-1>", on_click)

    def _trim_replacement_pair(self, before, after):
        if before == after:
            return None

        prefix = 0
        min_len = min(len(before), len(after))
        while prefix < min_len and before[prefix] == after[prefix]:
            prefix += 1

        suffix = 0
        while suffix < (min_len - prefix) and before[len(before) - 1 - suffix] == after[len(after) - 1 - suffix]:
            suffix += 1

        before_end = len(before) - suffix if suffix else len(before)
        after_end = len(after) - suffix if suffix else len(after)
        core_before = before[prefix:before_end]
        core_after = after[prefix:after_end]

        if not core_before:
            if prefix > 0:
                prefix -= 1
            elif suffix > 0:
                suffix -= 1
            else:
                return None

            before_end = len(before) - suffix if suffix else len(before)
            after_end = len(after) - suffix if suffix else len(after)
            core_before = before[prefix:before_end]
            core_after = after[prefix:after_end]

        if not core_before or core_before == core_after:
            return None

        return prefix, core_before, core_after

    def _build_hwp_replacements(self, old_doc, new_doc):
        import difflib

        unique_reps = []
        old_lines = old_doc.split("\n")
        new_lines = new_doc.split("\n")
        line_starts = []
        cursor = 0
        for line in old_lines:
            line_starts.append(cursor)
            cursor += len(line) + 1

        sm = difflib.SequenceMatcher(None, old_lines, new_lines)
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == 'equal':
                continue

            for k in range(max(i2 - i1, j2 - j1)):
                idx_old = i1 + k
                idx_new = j1 + k

                old_l = old_lines[idx_old] if idx_old < i2 else ""
                new_l = new_lines[idx_new] if idx_new < j2 else ""
                if old_l == new_l:
                    continue

                sm_char = difflib.SequenceMatcher(None, old_l, new_l)
                char_opcodes = sm_char.get_opcodes()
                changes = [(idx, op) for idx, op in enumerate(char_opcodes) if op[0] != 'equal']

                for k_idx, (_, (_, ci1, ci2, cj1, cj2)) in enumerate(changes):
                    prev_mod_ci2 = changes[k_idx - 1][1][2] if k_idx > 0 else 0
                    start_c = max(prev_mod_ci2, ci1 - 10)

                    next_mod_ci1 = changes[k_idx + 1][1][1] if k_idx < len(changes) - 1 else len(old_l)
                    end_c = min(next_mod_ci1, ci2 + 10)

                    raw_before = old_l[start_c:end_c]
                    if not raw_before:
                        if old_l:
                            raw_before = old_l
                            raw_after = new_l
                            start_c = 0
                        else:
                            continue
                    else:
                        raw_after = old_l[start_c:ci1] + new_l[cj1:cj2] + old_l[ci2:end_c]

                    trimmed = self._trim_replacement_pair(raw_before, raw_after)
                    if not trimmed:
                        continue

                    prefix_trimmed, core_before, core_after = trimmed
                    effective_start = start_c + prefix_trimmed
                    char_offset = line_starts[idx_old] + effective_start if idx_old < len(line_starts) else len(old_doc)
                    skip = old_doc.count(core_before, 0, char_offset)

                    unique_reps.append({
                        "before": core_before,
                        "after": core_after,
                        "skip": skip
                    })

        unique_reps.reverse()
        return unique_reps

    def save_file(self):
        content = self.original_text.get("1.0", "end-1c")
        if not content.strip():
            messagebox.showwarning("주의", "저장할 내용이 없습니다.")
            return
            
        initial_file = ""
        in_path = self.filepath_entry.get()
        if in_path and os.path.exists(in_path):
            base_name, _ = os.path.splitext(os.path.basename(in_path))
            initial_file = f"{base_name}_수정1"
            
        filepath = filedialog.asksaveasfilename(
            initialfile=initial_file,
            defaultextension=".hwp", 
            filetypes=[("HWP file", "*.hwp"), ("HWPX file", "*.hwpx"), ("Text file", "*.txt")],
            title="반영된 문서를 파일로 저장"
        )
        if not filepath: return
        
        ext = os.path.splitext(filepath)[1].lower()
        
        progress_win = ctk.CTkToplevel(self)
        progress_win.title("저장 중")
        progress_win.geometry("400x150")
        progress_win.attributes("-topmost", True)
        progress_win.protocol("WM_DELETE_WINDOW", lambda: None)
        
        lbl = ctk.CTkLabel(progress_win, text="백그라운드 봇이 문서의 단어를 교체하고 있습니다.\n잠시만 기다려주세요...", font=("", 14), text_color=self.DARK_YELLOW)
        lbl.pack(pady=(20, 10))
        p_bar = ctk.CTkProgressBar(progress_win, width=300, mode="indeterminate", progress_color=self.MORNING_YELLOW)
        p_bar.pack(pady=10)
        p_bar.start()

        saved_reps_for_txt = self._get_current_replacements()

        def save_threadTask():
            try:
                import pythoncom
                pythoncom.CoInitialize() 
                
                if ext in [".hwp", ".hwpx"]:
                    if self.loaded_hwp_path and os.path.exists(self.loaded_hwp_path):
                        target_save_path = os.path.abspath(filepath)
                        original_hwp_path = os.path.abspath(self.loaded_hwp_path)
                        
                        if target_save_path != original_hwp_path:
                            shutil.copy2(original_hwp_path, target_save_path)
                        
                        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
                        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
                        
                        if not hwp.Open(target_save_path, "", "forceopen:true"):
                            raise Exception(f"복사된 HWP/HWPX 원본 파일을 열 수 없습니다.")
                            
                        hwp.XHwpWindows.Item(0).Visible = False
                            
                        old_doc = self.cached_original_content.replace("\r", "")
                        new_doc = content.replace("\r", "")

                        total_memo_count = 0
                        reps_count = 0

                        if old_doc.strip() != new_doc.strip():
                            reps = self._build_hwp_replacements(old_doc, new_doc)
                            reps_count = len(reps)

                            for rep in reps:
                                hwp.HAction.Run("Cancel")
                                hwp.MovePos(2) 

                                clean_before = rep["before"]
                                after_text = rep["after"]
                                if not clean_before: continue

                                find_str = clean_before[:200] if len(clean_before) > 200 else clean_before
                                target_skip = rep["skip"]
                                current_skip = 0
                                
                                # ★ [핵심 수정] 무한 루프 판정 로직: Set을 이용하여 방문한 절대 좌표 기록
                                visited_pos = set()

                                while True:
                                    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
                                    hwp.HParameterSet.HFindReplace.FindString = find_str
                                    hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
                                    hwp.HParameterSet.HFindReplace.Direction = 0 

                                    if not hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet):
                                        break

                                    pos = hwp.GetPos()
                                    
                                    # 이전에 찾은 좌표를 또 방문했다면 문서 전체를 한 바퀴 돈 것
                                    if pos in visited_pos:
                                        break
                                    visited_pos.add(pos)

                                    if current_skip < target_skip:
                                        current_skip += 1
                                        continue

                                    if len(clean_before) > 200:
                                        for _ in range(len(clean_before) - 200):
                                            hwp.HAction.Run("MoveSelRight")

                                    # [핵심 수정] 남아있는 글자 삭제 보장(Delete 액션 도입)
                                    hwp.HAction.Run("Delete")
                                    
                                    if after_text:
                                        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
                                        hwp.HParameterSet.HInsertText.Text = after_text.replace("\n", "\r\n")
                                        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

                                    if self.memo_var.get():
                                        b_disp = clean_before if len(clean_before) <= 40 else clean_before[:20] + "..." + clean_before[-17:]
                                        if not after_text or not after_text.strip():
                                            a_disp = "(삭제)"
                                        else:
                                            a_disp = after_text if len(after_text) <= 40 else after_text[:20] + "..." + after_text[-17:]
                                        memo_text = f"변경: {b_disp} \u2192 {a_disp}"

                                        hwp.HAction.Run("InsertFieldMemo")
                                        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
                                        hwp.HParameterSet.HInsertText.Text = memo_text
                                        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
                                        hwp.HAction.Run("CloseEx")
                                        total_memo_count += 1

                                    break

                        try:
                            hwp.Save()
                        except Exception as e:
                            save_fmt = "HWPX" if ext == ".hwpx" else "HWP"
                            hwp.SaveAs(target_save_path, save_fmt, "")
                            
                        hwp.Clear(1)
                        hwp.Quit()
                        
                        memo_msg = f"HWP 자동화 봇이 치환 작업을 성공적으로 마쳤습니다!\n\n총 {reps_count}곳의 텍스트 수정 완료\n"
                        if self.memo_var.get():
                            memo_msg += f"총 {total_memo_count}개의 변경 메모 삽입 완료"
                        else:
                            memo_msg += "변경 메모 삽입 옵션이 비활성화되어 메모를 생략했습니다."
                            
                        self.after(0, lambda m=memo_msg: messagebox.showinfo("완료", m))
                    else:
                        target_save_path = os.path.abspath(filepath)
                        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
                        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
                        hwp.Clear(1)
                        act = hwp.CreateAction("InsertText")
                        pset = act.CreateSet()
                        act.GetDefault(pset)
                        pset.SetItem("Text", content.replace("\n", "\r\n"))
                        act.Execute(pset)
                        
                        save_fmt = "HWPX" if ext == ".hwpx" else "HWP"
                        if not hwp.SaveAs(target_save_path, save_fmt, ""):
                            raise Exception("임시 HWP 저장을 실패했습니다.")
                        hwp.Clear(1)
                        hwp.Quit()
                        self.after(0, lambda: messagebox.showinfo("완료", "텍스트 기반의 새 HWP/HWPX 문서로 저장되었습니다."))
                else:
                    change_log = ""
                    if saved_reps_for_txt:
                        log_lines = []
                        for rep in saved_reps_for_txt:
                            before_w = rep.get("before", "")
                            after_w = rep.get("after", "")
                            rep_type = rep.get("type", "")
                            if not before_w or not after_w or before_w == after_w:
                                continue
                            if rep_type == "ai":
                                b_disp = before_w if len(before_w) <= 30 else before_w[:15] + "..." + before_w[-12:]
                                a_disp = after_w if len(after_w) <= 30 else after_w[:15] + "..." + after_w[-12:]
                                log_lines.append(f"  [AI 다듬기] {b_disp} \u2192 {a_disp}")
                            else:
                                log_lines.append(f"  [행정순화어] {before_w} \u2192 {after_w}")
                        if log_lines:
                            unique_lines = list(dict.fromkeys(log_lines))
                            change_log = "\n\n" + "=" * 50 + "\n[변경 이력 요약]\n" + "=" * 50 + "\n"
                            change_log += "\n".join(unique_lines)
                            change_log += f"\n\n총 {len(saved_reps_for_txt)}건 변경"

                    with open(filepath, "w", encoding="utf-8") as f:
                        f.write(content + change_log)
                    log_count = len(saved_reps_for_txt) if saved_reps_for_txt else 0
                    msg = f"성공적으로 TXT 파일로 저장되었습니다."
                    if log_count > 0:
                        msg += f"\n({log_count}건의 변경 이력이 함께 저장됨)"
                    self.after(0, lambda m=msg: messagebox.showinfo("완료", m))
            except Exception as e:
                err_msg = str(e)
                self.after(0, lambda m=err_msg: messagebox.showerror("오류", f"저장 중 오류:\n{m}"))
            finally:
                pythoncom.CoUninitialize()
                self.after(0, progress_win.destroy)

        threading.Thread(target=save_threadTask, daemon=True).start()

    def extract_text_from_hwp(self, filepath):
        hwp = None
        try:
            filepath = os.path.abspath(filepath)
            hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            
            if not hwp.Open(filepath, "", "forceopen:true"):
                raise Exception("한글 파일을 열 수 없습니다. 파일이 이미 열려 있는지 확인하세요.")

            # ★ [핵심 수정 2] TEXT 대신 UNICODE 포맷 추출 지정 (특수문자 및 기호 매칭 오류 방지)
            extracted_text = hwp.GetTextFile("UNICODE", "")
            hwp.Clear(1)
            hwp.Quit()
            return extracted_text.strip()
        except Exception as e:
            if hwp:
                try: hwp.Quit()
                except: pass
            raise Exception(f"한글 파일 추출 처리 중 오류 발생: {e}")

if __name__ == "__main__":
    app = AdminLingoPurifierApp()
    app.mainloop()
