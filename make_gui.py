# -*- coding: utf-8 -*-
"""신청서 데이터셋 만들기 GUI.

폼에 신청자 정보를 입력하면 실전 양식 xlsx(세로형)를 생성한다.
모든 경우 지원: 개인/개인사업자/단체 · 전기/수소 · 택시 · 생애최초 ·
사회계층 · 전환지원금(폐차 정보) . 생성 후 parser 로 즉시 검증 미리보기.

build_real.build_workbook 을 재사용(생성 로직 단일화).
app.py 는 이 xlsx 를 그대로 불러 자동입력한다.
"""
import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# 생성 로직(스킬 스크립트) 재사용
_ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(_ROOT, ".claude", "skills", "real-dataset", "scripts"))
sys.path.insert(0, _ROOT)
import build_real  # noqa: E402
from parser import parse_vertical  # noqa: E402


class MakerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("신청서 데이터셋 만들기")
        self.geometry("720x860")
        self.vars = {}          # key -> tk 변수
        self._build_ui()
        self._toggle_busi()
        self._toggle_exchange()
        self._toggle_social()

    # ── 위젯 헬퍼 ────────────────────────────────────────────
    def _row(self, parent, key, label, default="", width=28):
        fr = ttk.Frame(parent)
        fr.pack(fill="x", padx=6, pady=1)
        ttk.Label(fr, text=label, width=16, anchor="e").pack(side="left")
        var = tk.StringVar(value=default)
        ttk.Entry(fr, textvariable=var, width=width).pack(side="left", fill="x", expand=True, padx=4)
        self.vars[key] = var
        return fr

    def _combo(self, parent, key, label, values, default):
        fr = ttk.Frame(parent)
        fr.pack(fill="x", padx=6, pady=1)
        ttk.Label(fr, text=label, width=16, anchor="e").pack(side="left")
        var = tk.StringVar(value=default)
        cb = ttk.Combobox(fr, textvariable=var, values=values, width=25, state="readonly")
        cb.pack(side="left", padx=4)
        self.vars[key] = var
        return cb

    def _check(self, parent, key, label):
        var = tk.BooleanVar(value=False)
        ttk.Checkbutton(parent, text=label, variable=var,
                        command=self._on_toggle).pack(side="left", padx=8)
        self.vars[key] = var
        return var

    # ── UI 구성 ──────────────────────────────────────────────
    def _build_ui(self):
        # 스크롤 캔버스(필드가 많음)
        canvas = tk.Canvas(self, highlightthickness=0)
        sb = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.body = ttk.Frame(canvas)
        self.body.bind("<Configure>",
                       lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.body, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))
        b = self.body

        # 기본
        f1 = ttk.LabelFrame(b, text="기본")
        f1.pack(fill="x", padx=8, pady=4)
        self._combo(f1, "gubun", "구분(전기/수소)", ["전기", "수소"], "전기")
        self._row(f1, "contract_day", "계약일자", "2026-07-07")
        self.reqkind_cb = self._combo(f1, "req_kind", "신청유형",
                                      ["개인", "개인사업자", "단체"], "개인")
        self.reqkind_cb.bind("<<ComboboxSelected>>", lambda e: self._toggle_busi())
        self._row(f1, "name", "성명/기관명")
        self._row(f1, "birth", "생년월일", "")
        self._combo(f1, "sex", "성별", ["남", "여"], "남")

        # 사업자(개인사업자/단체일 때)
        self.f_busi = ttk.LabelFrame(b, text="사업자 정보 (개인사업자·단체)")
        self.f_busi.pack(fill="x", padx=8, pady=4)
        self._row(self.f_busi, "busi_no", "사업자번호")
        self._row(self.f_busi, "busi_nm", "사업자명")

        # 차량
        f2 = ttk.LabelFrame(b, text="차량")
        f2.pack(fill="x", padx=8, pady=4)
        self._row(f2, "model", "신청차종", "")
        ttk.Label(f2, text="※ 사이트 정식명에 가깝게 (reference/models_reference.txt)",
                  font=("", 8)).pack(anchor="w", padx=22)
        self._row(f2, "cnt", "신청대수", "1")
        self._row(f2, "delivery", "출고예정일자", "2026-07-07")
        self._row(f2, "subsidy", "보조금금액")

        # 주소
        f3 = ttk.LabelFrame(b, text="주소")
        f3.pack(fill="x", padx=8, pady=4)
        self._row(f3, "addr", "주소(도로명)")
        self._row(f3, "addr_detail", "이하주소(상세)")

        # 연락처
        f4 = ttk.LabelFrame(b, text="연락처")
        f4.pack(fill="x", padx=8, pady=4)
        self._row(f4, "phone", "전화번호", ".")
        self._row(f4, "mobile", "휴대폰(신청자)")
        self._row(f4, "email", "이메일", ".")
        ttk.Label(f4, text="※ 종이 전화번호는 보통 '.'로 두고 실제 번호는 휴대폰에",
                  font=("", 8)).pack(anchor="w", padx=22)

        # 신청조건
        f5 = ttk.LabelFrame(b, text="신청조건")
        f5.pack(fill="x", padx=8, pady=4)
        crow = ttk.Frame(f5)
        crow.pack(fill="x", padx=6, pady=2)
        self._check(crow, "taxi_yn", "택시여부")
        self._check(crow, "first_buy_yn", "생애최초 구매자")
        self._check(crow, "social_yn", "사회계층")
        self._check(crow, "exchange_yn", "전환지원금(폐차)")
        self._row(f5, "social_kind", "사회계층 유형")

        # 전환지원금 폐차 정보
        self.f_ex = ttk.LabelFrame(b, text="전환지원금 폐차 정보 (전환지원금 체크 시)")
        self.f_ex.pack(fill="x", padx=8, pady=4)
        self._row(self.f_ex, "prev_owner", "직전소유주")
        self._row(self.f_ex, "first_reg", "차량 최초등록일")
        self._row(self.f_ex, "own_start", "소유기간 시작일")
        self._row(self.f_ex, "own_end", "소유기간 종료일", "2026-07-07")
        self._combo(self.f_ex, "fuel", "유종", ["", "LPG", "경유", "휘발유"], "")
        self._row(self.f_ex, "scrap_car_no", "폐차 차량번호")

        # 담당자 / 암호
        f6 = ttk.LabelFrame(b, text="담당자 · 계약 · 암호")
        f6.pack(fill="x", padx=8, pady=4)
        self._row(f6, "contact_nm", "담당자성명")
        self._row(f6, "contact_mobile", "담당자휴대폰")
        self._row(f6, "contract_no", "계약번호")
        self._row(f6, "pw", "암호")
        self._row(f6, "pw_rev", "암호 역순")

        # 출력 폴더
        f7 = ttk.LabelFrame(b, text="저장 위치")
        f7.pack(fill="x", padx=8, pady=4)
        orow = ttk.Frame(f7)
        orow.pack(fill="x", padx=6, pady=2)
        ttk.Label(orow, text="폴더(비우면 계약일 real_MMDD)", width=24, anchor="e").pack(side="left")
        self.vars["outdir"] = tk.StringVar(value="")
        ttk.Entry(orow, textvariable=self.vars["outdir"], width=26).pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(orow, text="찾기", command=self._pick_dir).pack(side="left")

        # 버튼
        bar = ttk.Frame(b)
        bar.pack(fill="x", padx=8, pady=8)
        ttk.Button(bar, text="엑셀 생성", command=self.generate).pack(side="left")
        ttk.Button(bar, text="생성 + 검증 미리보기", command=self.generate_verify).pack(side="left", padx=6)
        ttk.Button(bar, text="폼 비우기", command=self.clear_form).pack(side="left")

        # 로그
        ttk.Label(b, text="결과").pack(anchor="w", padx=10)
        self.log = tk.Text(b, height=12, state="disabled")
        self.log.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.log.tag_config("red", foreground="#cc0000")
        self.log.tag_config("green", foreground="#117711")

    # ── 조건부 표시 ──────────────────────────────────────────
    def _on_toggle(self):
        self._toggle_exchange()
        self._toggle_social()

    def _set_frame_state(self, frame, enabled):
        for child in frame.winfo_children():
            for w in (child.winfo_children() or [child]):
                try:
                    w.configure(state=("normal" if enabled else "disabled"))
                except tk.TclError:
                    pass

    def _toggle_busi(self):
        on = self.vars["req_kind"].get() in ("개인사업자", "단체")
        self._set_frame_state(self.f_busi, on)

    def _toggle_exchange(self):
        self._set_frame_state(self.f_ex, self.vars["exchange_yn"].get())

    def _toggle_social(self):
        # 사회계층 체크 시 유형 입력 활성
        pass

    def _pick_dir(self):
        p = filedialog.askdirectory()
        if p:
            self.vars["outdir"].set(p)

    # ── 생성 ────────────────────────────────────────────────
    def _collect(self):
        """폼 → build_real 스펙 dict."""
        def g(k):
            v = self.vars.get(k)
            if isinstance(v, tk.BooleanVar):
                return "Y" if v.get() else ""
            return (v.get() if v else "").strip()

        spec = {}
        for k in ["gubun", "contract_day", "req_kind", "name", "birth", "sex",
                  "busi_no", "busi_nm", "model", "cnt", "delivery", "subsidy",
                  "addr", "addr_detail", "phone", "mobile", "email",
                  "taxi_yn", "first_buy_yn", "social_yn", "social_kind",
                  "exchange_yn", "prev_owner", "first_reg", "own_start", "own_end",
                  "fuel", "scrap_car_no", "contact_nm", "contact_mobile",
                  "contract_no", "pw", "pw_rev"]:
            val = g(k)
            if val != "":
                spec[k] = val
        return spec

    def _log(self, msg, color=None):
        self.log.config(state="normal")
        self.log.insert("end", msg + "\n", color or ())
        self.log.see("end")
        self.log.config(state="disabled")

    def _do_generate(self):
        spec = self._collect()
        if not spec.get("name"):
            messagebox.showwarning("확인", "성명/기관명은 필수입니다.")
            return None
        outdir = self.vars["outdir"].get().strip() or \
            os.path.join(_ROOT, "real_%s" % build_real._mmdd(spec))
        os.makedirs(outdir, exist_ok=True)
        wb = build_real.build_workbook(spec)
        path = os.path.join(outdir, "신청서_%s.xlsx" % spec["name"])
        wb.save(path)
        gubun = build_real.derive_gubun(spec)
        self._log(f"✅ 생성: {path}  [{gubun}]", "green")
        return path

    def generate(self):
        try:
            self._do_generate()
        except Exception as e:
            self._log(f"❌ 생성 오류: {e}", "red")

    def generate_verify(self):
        path = None
        try:
            path = self._do_generate()
        except Exception as e:
            self._log(f"❌ 생성 오류: {e}", "red")
            return
        if not path:
            return
        try:
            res = parse_vertical(path)
        except Exception as e:
            self._log(f"❌ 파서 검증 오류: {e}", "red")
            return
        self._log("── 파서 매핑(자동입력 대상) ──")
        for k, v in res["mapped"].items():
            self._log(f"   {k} = {v}")
        if res["special"]:
            self._log("── 텍스트매칭(차종·사회계층) ──")
            for k, v in res["special"].items():
                self._log(f"   {k} = {v}")
        if res["handoff"]:
            self._log("── 사람이 직접 ──")
            for lbl, val, reason in res["handoff"]:
                self._log(f"   · {lbl} = '{val}' ({reason})")

    def clear_form(self):
        for k, v in self.vars.items():
            if isinstance(v, tk.BooleanVar):
                v.set(False)
            elif k in ("phone", "email"):
                v.set(".")
            elif k not in ("gubun", "req_kind", "sex", "contract_day", "delivery",
                           "cnt", "own_end", "outdir"):
                v.set("")
        self._toggle_busi()
        self._toggle_exchange()
        self._log("폼을 비웠습니다.")


if __name__ == "__main__":
    MakerApp().mainloop()
