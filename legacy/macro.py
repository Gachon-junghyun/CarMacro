# -*- coding: utf-8 -*-
"""
순차 입력 매크로 (Sequential Form Filler) v0.1
- 탭1: 순차 입력 필드 — 신청서 등에 이름/생년월일 등을 단축키로 한 칸씩 붙여넣기
- 탭2: 메시지 스니펫 — 메시지마다 단축키 지정 → 복사 또는 복사+붙여넣기

한글 입력 안정성을 위해 "키 타이핑"이 아니라
"클립보드 복사 + Ctrl+V 자동 붙여넣기" 방식을 사용한다.
"""

import json
import os
import sys
import time
import threading
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

import keyboard
import pyperclip


def app_dir():
    """실행 파일(또는 스크립트)이 위치한 폴더.
    PyInstaller로 묶인 .exe면 exe 옆, 아니면 .py 옆을 가리킨다.
    → config.json이 USB의 exe 옆에 저장되도록 보장."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_PATH = os.path.join(app_dir(), "config.json")

DEFAULT_CONFIG = {
    "settings": {
        "next_field_hotkey": "f8",      # 현재 칸 입력하고 다음으로 (메인)
        "prev_field_hotkey": "f7",      # 전거: 한 칸 뒤로 가서 다시 입력
        "fwd_field_hotkey": "f9",       # 뒤거: 한 칸 앞으로 가서 입력
        "clear_back_hotkey": "f6",      # 취소: 칸 지우고 뒤로
        "fill_all_hotkey": "shift+f8",  # 전체 자동 입력
        "reset_hotkey": "f10",          # 포인터 리셋
        "auto_tab": True,               # 붙여넣은 뒤 Tab 자동 전송
        "paste_delay": 0.05,            # 동작 사이 대기(초)
    },
    "fields": [
        {"label": "이름", "value": "홍길동"},
        {"label": "생년월일", "value": "1990-01-01"},
        {"label": "전화번호", "value": "010-1234-5678"},
    ],
    "snippets": [
        {"label": "인사말", "value": "안녕하세요, 잘 부탁드립니다.",
         "hotkey": "ctrl+alt+1", "paste": False},
    ],
}


def load_config():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            # 누락된 기본 키 보강
            for k, v in DEFAULT_CONFIG["settings"].items():
                cfg.setdefault("settings", {}).setdefault(k, v)
            cfg.setdefault("fields", [])
            cfg.setdefault("snippets", [])
            return cfg
        except Exception as e:
            print("config 로드 실패, 기본값 사용:", e)
    return json.loads(json.dumps(DEFAULT_CONFIG))


def save_config(cfg):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def do_paste(value, send_tab, delay):
    """클립보드에 값을 넣고 Ctrl+V 전송. 옵션으로 Tab."""
    pyperclip.copy(value)
    time.sleep(delay)
    keyboard.send("ctrl+v")
    if send_tab:
        time.sleep(delay)
        keyboard.send("tab")


def do_paste_replace(value, send_tab, delay):
    """기존 칸 내용을 전체 선택(Ctrl+A) 후 덮어쓰기.
    이미 값이 있는 칸에 다시 넣어도 중복되지 않는다 (전거/취소 후 재입력용)."""
    keyboard.send("ctrl+a")
    time.sleep(delay)
    pyperclip.copy(value)
    time.sleep(delay)
    keyboard.send("ctrl+v")
    if send_tab:
        time.sleep(delay)
        keyboard.send("tab")


def do_clear(delay):
    """현재 포커스된 입력칸의 내용을 모두 지운다 (Ctrl+A → Delete)."""
    keyboard.send("ctrl+a")
    time.sleep(delay)
    keyboard.send("delete")


class MacroApp:
    def __init__(self, root):
        self.root = root
        self.cfg = load_config()
        self.pointer = 0  # 현재 순차 필드 인덱스
        self._hotkey_handles = []

        root.title("순차 입력 매크로 v0.1")
        root.geometry("560x520")

        nb = ttk.Notebook(root)
        nb.pack(fill="both", expand=True, padx=8, pady=8)

        self.tab_fields = ttk.Frame(nb)
        self.tab_snips = ttk.Frame(nb)
        self.tab_settings = ttk.Frame(nb)
        nb.add(self.tab_fields, text="순차 입력 필드")
        nb.add(self.tab_snips, text="메시지 스니펫")
        nb.add(self.tab_settings, text="설정")

        self._build_fields_tab()
        self._build_snippets_tab()
        self._build_settings_tab()

        self.status = tk.StringVar(value="준비됨")
        ttk.Label(root, textvariable=self.status, relief="sunken",
                  anchor="w").pack(fill="x", side="bottom")

        self.refresh_fields()
        self.refresh_snippets()
        self.register_hotkeys()
        self.update_status()

        root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------- 순차 입력 필드 탭 ----------
    def _build_fields_tab(self):
        f = self.tab_fields
        info = ("F8 = 입력하고 다음 칸   |   F7 = 전거(뒤로 다시)   |   F9 = 뒤거(앞으로)\n"
                "F6 = 취소(칸 지우고 뒤로)   |   Shift+F8 = 전체 자동   |   F10 = 리셋\n"
                "신청 폼의 첫 칸을 클릭한 뒤 단축키를 누르세요.")
        ttk.Label(f, text=info, foreground="#555", justify="left").pack(anchor="w", padx=6, pady=(6, 2))

        self.fields_list = tk.Listbox(f, height=12)
        self.fields_list.pack(fill="both", expand=True, padx=6, pady=4)

        btns = ttk.Frame(f)
        btns.pack(fill="x", padx=6, pady=4)
        ttk.Button(btns, text="추가", command=self.add_field).pack(side="left")
        ttk.Button(btns, text="수정", command=self.edit_field).pack(side="left", padx=3)
        ttk.Button(btns, text="삭제", command=self.del_field).pack(side="left")
        ttk.Button(btns, text="▲", width=3, command=lambda: self.move_field(-1)).pack(side="left", padx=(12, 1))
        ttk.Button(btns, text="▼", width=3, command=lambda: self.move_field(1)).pack(side="left", padx=1)
        ttk.Button(btns, text="리셋", command=self.reset_pointer).pack(side="right")
        ttk.Button(btns, text="전체 삭제", command=self.clear_all_fields).pack(side="right", padx=4)

    def refresh_fields(self):
        self.fields_list.delete(0, "end")
        for i, fld in enumerate(self.cfg["fields"]):
            marker = "▶ " if i == self.pointer else "   "
            self.fields_list.insert("end", f"{marker}{i+1}. {fld['label']}: {fld['value']}")

    def add_field(self):
        label = simpledialog.askstring("필드 추가", "라벨 (예: 이름):", parent=self.root)
        if not label:
            return
        value = simpledialog.askstring("필드 추가", f"'{label}' 값:", parent=self.root)
        if value is None:
            return
        self.cfg["fields"].append({"label": label, "value": value})
        save_config(self.cfg)
        self.refresh_fields()

    def edit_field(self):
        idx = self._sel(self.fields_list)
        if idx is None:
            return
        fld = self.cfg["fields"][idx]
        label = simpledialog.askstring("필드 수정", "라벨:", initialvalue=fld["label"], parent=self.root)
        if label is None:
            return
        value = simpledialog.askstring("필드 수정", "값:", initialvalue=fld["value"], parent=self.root)
        if value is None:
            return
        fld["label"], fld["value"] = label, value
        save_config(self.cfg)
        self.refresh_fields()

    def del_field(self):
        idx = self._sel(self.fields_list)
        if idx is None:
            return
        del self.cfg["fields"][idx]
        if self.pointer >= len(self.cfg["fields"]):
            self.pointer = 0
        save_config(self.cfg)
        self.refresh_fields()
        self.update_status()

    def move_field(self, d):
        idx = self._sel(self.fields_list)
        if idx is None:
            return
        j = idx + d
        if 0 <= j < len(self.cfg["fields"]):
            flds = self.cfg["fields"]
            flds[idx], flds[j] = flds[j], flds[idx]
            save_config(self.cfg)
            self.refresh_fields()
            self.fields_list.selection_set(j)

    def reset_pointer(self):
        self.pointer = 0
        self.refresh_fields()
        self.update_status()

    def clear_all_fields(self):
        if not self.cfg["fields"]:
            return
        if messagebox.askyesno("전체 삭제", "등록된 순차 입력 필드를 모두 삭제할까요?"):
            self.cfg["fields"] = []
            self.pointer = 0
            save_config(self.cfg)
            self.refresh_fields()
            self.update_status()

    # ---------- 메시지 스니펫 탭 ----------
    def _build_snippets_tab(self):
        f = self.tab_snips
        ttk.Label(f, text="메시지마다 단축키를 지정하세요. (복사 / 복사+붙여넣기)",
                  foreground="#555").pack(anchor="w", padx=6, pady=(6, 2))
        self.snips_list = tk.Listbox(f, height=12)
        self.snips_list.pack(fill="both", expand=True, padx=6, pady=4)

        btns = ttk.Frame(f)
        btns.pack(fill="x", padx=6, pady=4)
        ttk.Button(btns, text="추가", command=self.add_snippet).pack(side="left")
        ttk.Button(btns, text="수정", command=self.edit_snippet).pack(side="left", padx=3)
        ttk.Button(btns, text="삭제", command=self.del_snippet).pack(side="left")
        ttk.Button(btns, text="전체 삭제", command=self.clear_all_snippets).pack(side="right")

    def clear_all_snippets(self):
        if not self.cfg["snippets"]:
            return
        if messagebox.askyesno("전체 삭제", "등록된 메시지 스니펫을 모두 삭제할까요?"):
            self.cfg["snippets"] = []
            save_config(self.cfg)
            self.refresh_snippets()
            self.register_hotkeys()

    def refresh_snippets(self):
        self.snips_list.delete(0, "end")
        for s in self.cfg["snippets"]:
            mode = "붙여넣기" if s.get("paste") else "복사"
            self.snips_list.insert("end", f"[{s.get('hotkey','-')}] ({mode}) {s['label']}: {s['value']}")

    def add_snippet(self):
        label = simpledialog.askstring("스니펫 추가", "라벨:", parent=self.root)
        if not label:
            return
        value = simpledialog.askstring("스니펫 추가", "메시지 내용:", parent=self.root)
        if value is None:
            return
        hotkey = simpledialog.askstring("스니펫 추가", "단축키 (예: ctrl+alt+1):", parent=self.root)
        if not hotkey:
            return
        paste = messagebox.askyesno("동작 선택", "붙여넣기까지 자동으로 할까요?\n(아니오 = 클립보드 복사만)")
        self.cfg["snippets"].append({"label": label, "value": value,
                                     "hotkey": hotkey.strip().lower(), "paste": paste})
        save_config(self.cfg)
        self.refresh_snippets()
        self.register_hotkeys()

    def edit_snippet(self):
        idx = self._sel(self.snips_list)
        if idx is None:
            return
        s = self.cfg["snippets"][idx]
        label = simpledialog.askstring("수정", "라벨:", initialvalue=s["label"], parent=self.root)
        if label is None:
            return
        value = simpledialog.askstring("수정", "내용:", initialvalue=s["value"], parent=self.root)
        if value is None:
            return
        hotkey = simpledialog.askstring("수정", "단축키:", initialvalue=s.get("hotkey", ""), parent=self.root)
        if not hotkey:
            return
        s.update(label=label, value=value, hotkey=hotkey.strip().lower())
        save_config(self.cfg)
        self.refresh_snippets()
        self.register_hotkeys()

    def del_snippet(self):
        idx = self._sel(self.snips_list)
        if idx is None:
            return
        del self.cfg["snippets"][idx]
        save_config(self.cfg)
        self.refresh_snippets()
        self.register_hotkeys()

    # ---------- 설정 탭 ----------
    def _build_settings_tab(self):
        f = self.tab_settings
        s = self.cfg["settings"]
        self.var_auto_tab = tk.BooleanVar(value=s["auto_tab"])
        ttk.Checkbutton(f, text="붙여넣은 뒤 자동으로 Tab (다음 칸으로 이동)",
                        variable=self.var_auto_tab,
                        command=self._save_settings).pack(anchor="w", padx=10, pady=10)
        hk = (f"다음 칸: {s['next_field_hotkey'].upper()}    전거: {s['prev_field_hotkey'].upper()}    "
              f"뒤거: {s['fwd_field_hotkey'].upper()}\n"
              f"취소(지우고 뒤로): {s['clear_back_hotkey'].upper()}    "
              f"전체 입력: {s['fill_all_hotkey'].upper()}    리셋: {s['reset_hotkey'].upper()}")
        ttk.Label(f, text=hk, foreground="#555", justify="left").pack(anchor="w", padx=10)
        ttk.Label(f, text=f"설정 파일: {CONFIG_PATH}", foreground="#999",
                  wraplength=520).pack(anchor="w", padx=10, pady=(20, 0))

    def _save_settings(self):
        self.cfg["settings"]["auto_tab"] = self.var_auto_tab.get()
        save_config(self.cfg)

    # ---------- 단축키 등록 / 동작 ----------
    def register_hotkeys(self):
        for h in self._hotkey_handles:
            try:
                keyboard.remove_hotkey(h)
            except Exception:
                pass
        self._hotkey_handles = []
        s = self.cfg["settings"]
        try:
            self._hotkey_handles.append(keyboard.add_hotkey(s["next_field_hotkey"], self.fire_next_field))
            self._hotkey_handles.append(keyboard.add_hotkey(s["prev_field_hotkey"], self.fire_prev_field))
            self._hotkey_handles.append(keyboard.add_hotkey(s["fwd_field_hotkey"], self.fire_fwd_field))
            self._hotkey_handles.append(keyboard.add_hotkey(s["clear_back_hotkey"], self.fire_clear_back))
            self._hotkey_handles.append(keyboard.add_hotkey(s["fill_all_hotkey"], self.fire_fill_all))
            self._hotkey_handles.append(keyboard.add_hotkey(s["reset_hotkey"], self.reset_pointer))
            for snip in self.cfg["snippets"]:
                hk = snip.get("hotkey")
                if hk:
                    self._hotkey_handles.append(
                        keyboard.add_hotkey(hk, lambda sn=snip: self.fire_snippet(sn)))
        except Exception as e:
            messagebox.showerror("단축키 오류", f"단축키 등록 실패: {e}")

    def fire_next_field(self):
        """F8 — 현재 칸 입력하고 다음 칸으로."""
        flds = self.cfg["fields"]
        if not flds:
            return
        if self.pointer >= len(flds):
            self.status.set("끝까지 입력 완료 — F10으로 리셋")
            return
        delay = self.cfg["settings"]["paste_delay"]
        do_paste_replace(flds[self.pointer]["value"], self.cfg["settings"]["auto_tab"], delay)
        self.pointer += 1
        self.root.after(0, self._after_fire)

    def fire_prev_field(self):
        """F7 — 전거: 한 칸 뒤로(Shift+Tab) 가서 이전 칸 다시 입력(덮어쓰기)."""
        flds = self.cfg["fields"]
        if not flds:
            return
        delay = self.cfg["settings"]["paste_delay"]
        keyboard.send("shift+tab")
        time.sleep(delay)
        self.pointer = max(self.pointer - 1, 0)
        do_paste_replace(flds[self.pointer]["value"], False, delay)
        self.root.after(0, self._after_fire)

    def fire_fwd_field(self):
        """F9 — 뒤거: 한 칸 앞으로(Tab) 가서 다음 칸 입력(덮어쓰기)."""
        flds = self.cfg["fields"]
        if not flds:
            return
        delay = self.cfg["settings"]["paste_delay"]
        keyboard.send("tab")
        time.sleep(delay)
        self.pointer = min(self.pointer + 1, len(flds) - 1)
        do_paste_replace(flds[self.pointer]["value"], False, delay)
        self.root.after(0, self._after_fire)

    def fire_clear_back(self):
        """F6 — 취소: 한 칸 뒤로 가서 그 칸 내용을 지우고 대기.
        잘못 입력했을 때 그 값을 없애고, 다시 F8로 올바른 값을 넣을 수 있다."""
        delay = self.cfg["settings"]["paste_delay"]
        keyboard.send("shift+tab")
        time.sleep(delay)
        do_clear(delay)
        self.pointer = max(self.pointer - 1, 0)
        self.root.after(0, self._after_fire)

    def fire_fill_all(self):
        flds = self.cfg["fields"]
        delay = self.cfg["settings"]["paste_delay"]
        for fld in flds:
            do_paste_replace(fld["value"], True, delay)
            time.sleep(delay)
        self.pointer = len(flds)
        self.root.after(0, self._after_fire)

    def fire_snippet(self, snip):
        if snip.get("paste"):
            do_paste(snip["value"], False, self.cfg["settings"]["paste_delay"])
        else:
            pyperclip.copy(snip["value"])
        self.root.after(0, lambda: self.status.set(f"스니펫 실행: {snip['label']}"))

    def _after_fire(self):
        self.refresh_fields()
        self.update_status()

    # ---------- 공통 ----------
    def update_status(self):
        flds = self.cfg["fields"]
        if flds:
            cur = flds[self.pointer]["label"] if self.pointer < len(flds) else "-"
            self.status.set(f"다음 입력: {self.pointer+1}/{len(flds)} → {cur}")
        else:
            self.status.set("필드 없음")

    def _sel(self, listbox):
        sel = listbox.curselection()
        if not sel:
            messagebox.showinfo("선택", "목록에서 항목을 먼저 선택하세요.")
            return None
        return sel[0]

    def on_close(self):
        try:
            keyboard.unhook_all_hotkeys()
        except Exception:
            pass
        self.root.destroy()


def main():
    root = tk.Tk()
    MacroApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
