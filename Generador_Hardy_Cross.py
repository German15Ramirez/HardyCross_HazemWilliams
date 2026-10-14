#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generador de Excel Hardy Cross - versión final con diámetros automáticos.
- Hazen-Williams con constante 1743.811 (D en pulgadas reales, Q en L/s, L en m).
- Hardy Cross SIMULTÁNEO con cruce de deltas.
- Selección de diámetro comercial a partir de v = 1.974·Q/D²  →  D = √(1.974·Q/v).
- Validación de Hf por tubería (0.5 – 3.0 mca) en la primera iteración.
- Interfaz moderna con tema oscuro.
Requiere Python 3.9+, tkinter y openpyxl.
"""
import math
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# ---------------------------------------------------------------------------
# Hidráulica: Hazen-Williams con constante 1743.811 (D en pulgadas, Q en L/s)
# ---------------------------------------------------------------------------
N_HW = 1.852
K_HW = 1743.811

DIAMETROS_COMERCIALES = [0.5, 0.75, 1, 1.25, 1.5, 2, 2.5, 3, 4, 6, 8, 10, 12, 14, 16, 18]


def headloss(L, C, D_in, Q_lps):
    """Hf en mca. L en m, D en pulgadas (real), Q en L/s."""
    if L <= 0 or C <= 0 or D_in <= 0:
        raise ValueError("L, C y D deben ser mayores que cero.")
    if abs(Q_lps) < 1e-14:
        return 0.0
    return K_HW * L * abs(Q_lps) ** N_HW / (C ** N_HW * D_in ** 4.87)


def diametro_teorico(Q_lps, v_diseno):
    """D teórico en pulgadas: D = sqrt(1.974·Q/v)."""
    if Q_lps <= 0 or v_diseno <= 0:
        return 0.0
    return math.sqrt(1.974 * Q_lps / v_diseno)


def diametro_comercial(D_teorico, criterio="superior"):
    """Selecciona el diámetro comercial según el criterio."""
    if D_teorico <= 0:
        return DIAMETROS_COMERCIALES[0]
    if criterio == "superior":
        for d in DIAMETROS_COMERCIALES:
            if d >= D_teorico:
                return d
        return DIAMETROS_COMERCIALES[-1]
    elif criterio == "inferior":
        for d in reversed(DIAMETROS_COMERCIALES):
            if d <= D_teorico:
                return d
        return DIAMETROS_COMERCIALES[0]
    else:  # "cercano"
        return min(DIAMETROS_COMERCIALES, key=lambda d: abs(d - D_teorico))


def revisar_hf_primera_iteracion(result):
    """
    Revisa los Hf de la primera iteración. Devuelve (fuera, dentro) donde:
    - fuera: lista de (pid, Hf, sugerencia) con Hf fuera de 0.5–3.0
    - dentro: lista de (pid, Hf) con Hf dentro del rango
    """
    if not result["history"]:
        return [], []
    it1 = result["history"][0]
    fuera = []
    dentro = []
    vistos = set()
    for lp in it1["circuits"]:
        for d in lp["rows"]:
            pid = d["pipe"]
            if pid in vistos:
                continue
            vistos.add(pid)
            hf = d["Hf"]  # valor absoluto
            if hf < 0.5:
                fuera.append((pid, hf, "reducir diámetro"))
            elif hf > 3.0:
                fuera.append((pid, hf, "aumentar diámetro"))
            else:
                dentro.append((pid, hf))
    return fuera, dentro


def solve_hardy_cross(q_in, pipes, loops, max_iter=100, tol_factor=0.005):
    """
    Hardy Cross SIMULTÁNEO con cruce de deltas.
    - Δ propio = Δ_circuito = -ΣHf/(1.85·ΣHf/Q)
    - Δ del otro = -Σ Δ_vecinos
    - Δ total aplicado = Δ propio + Δ del otro (igual en todas las filas de una
      misma tubería compartida).
    - La corrección se aplica UNA SOLA VEZ por tubería física, con el dir del
      primer circuito donde aparece; el mismo Q_after se replica en todas las filas.
    """
    q = {pid: float(p["Q"]) for pid, p in pipes.items()}
    threshold = tol_factor * abs(q_in)
    history = []

    for iteration in range(1, max_iter + 1):
        q_start = q.copy()

        # PASO 1: cada circuito calcula su Δ con los caudales del inicio
        circuit_results = []
        pipe_loops = {}
        for li, loop in enumerate(loops):
            rows = []
            sum_hf = 0.0
            sum_hf_q = 0.0
            for m in loop["members"]:
                pid = m["pipe"]
                s = m["dir"]
                p = pipes[pid]
                q_physical = q_start[pid]
                q_loop = s * q_physical
                hf_abs = headloss(p["L"], p["C"], p["D"], q_loop)
                hf_signed = math.copysign(hf_abs, q_loop) if abs(q_loop) > 1e-14 else 0.0
                hf_q = hf_signed / q_loop if abs(q_loop) > 1e-14 else 0.0
                sum_hf += hf_signed
                sum_hf_q += hf_q
                rows.append({
                    "pipe": pid, "dir": s, "L": p["L"], "C": p["C"], "D": p["D"],
                    "Q_before": q_physical, "Q_loop": q_loop,
                    "Hf": hf_abs, "Hf_signed": hf_signed, "Hf_Q": hf_q
                })
            delta = -sum_hf / (1.85 * sum_hf_q) if sum_hf_q else 0.0
            circuit_results.append({
                "name": loop["name"],
                "rows": rows,
                "sum_hf": sum_hf,
                "sum_hf_q": sum_hf_q,
                "delta": delta,
                "abs_delta": abs(delta),
            })
            for m in loop["members"]:
                pipe_loops.setdefault(m["pipe"], []).append({
                    "loop_idx": li, "loop_name": loop["name"],
                    "dir": m["dir"], "delta": delta,
                })

        # PASO 2: Δ propio (tal cual), Δ del otro y Δ total por fila
        for li, lp in enumerate(circuit_results):
            for r in lp["rows"]:
                pid = r["pipe"]
                delta_propio = lp["delta"]      # sin *-1, tal cual el Excel

                others = [x for x in pipe_loops[pid] if x["loop_idx"] != li]
                delta_otro = 0.0
                for o in others:
                    delta_otro += -o["delta"]

                delta_total = delta_propio + (delta_otro if others else 0.0)

                r["delta_propio"] = delta_propio
                r["delta_otro"] = delta_otro if others else None
                r["delta_total"] = delta_total
                r["compartida"] = bool(others)

        # PASO 3: aplicar corrección UNA SOLA VEZ por tubería física
        applied = set()
        for li, lp in enumerate(circuit_results):
            for r in lp["rows"]:
                pid = r["pipe"]
                if pid in applied:
                    r["Q_after"] = q[pid]
                    continue
                dir_principal = pipe_loops[pid][0]["dir"]
                q[pid] += dir_principal * r["delta_total"]
                applied.add(pid)
                r["Q_after"] = q[pid]

        for li, lp in enumerate(circuit_results):
            for r in lp["rows"]:
                if "Q_after" not in r:
                    r["Q_after"] = q[r["pipe"]]

        max_delta = max((x["abs_delta"] for x in circuit_results), default=0.0)
        history.append({
            "iteration": iteration,
            "circuits": circuit_results,
            "flows": q.copy(),
            "max_delta": max_delta,
        })
        if max_delta <= threshold:
            break

    pipe_results = {}
    for pid, p in pipes.items():
        hf = headloss(p["L"], p["C"], p["D"], q[pid])
        pipe_results[pid] = {
            **p, "Q_final": q[pid], "Hf": hf,
            "h_range_ok": 0.5 <= hf <= 3.0
        }

    return {
        "q_in": q_in, "threshold": threshold, "tol_factor": tol_factor,
        "iterations": len(history),
        "converged": bool(history and history[-1]["max_delta"] <= threshold),
        "history": history, "pipes": pipe_results,
    }


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------
def make_excel(filename, q_in, pipes, loops, result, v_diseno=None, modo_diametro="manual"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Entrada"
    hc = wb.create_sheet("Hardy Cross")
    sm = wb.create_sheet("Resumen")

    dark = "1F4E78"; blue = "D9EAF7"; green = "C6EFCE"
    yellow = "FFEB9C"; red = "FFC7CE"; white = "FFFFFF"
    orange = "FCE4D6"

    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def title(sheet, text, last_col):
        sheet["A1"] = text
        sheet["A1"].font = Font(size=16, bold=True, color=white)
        sheet["A1"].fill = PatternFill("solid", fgColor=dark)
        sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
        sheet["A1"].alignment = Alignment(horizontal="center", vertical="center")
        sheet.row_dimensions[1].height = 26

    def header_row(sheet, row, headers):
        for c, h in enumerate(headers, 1):
            cell = sheet.cell(row, c, h)
            cell.font = Font(bold=True, color=white)
            cell.fill = PatternFill("solid", fgColor=dark)
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        sheet.row_dimensions[row].height = 30

    # ---------------- ENTRADA ----------------
    title(ws, "DISEÑO HIDRÁULICO – HARDY CROSS", 10)
    meta = [
        ("Caudal de entrada (L/s)", q_in),
        ("Criterio de convergencia", f"|Δ| ≤ {result['tol_factor']} × Q entrada"),
        ("Tolerancia calculada (L/s)", result["threshold"]),
        ("Rango de Hf por tubería (mca)", "0.5 – 3.0"),
        ("Circuitos", len(loops)),
    ]
    for r, (a, b) in enumerate(meta, 3):
        ws.cell(r, 1, a).font = Font(bold=True)
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=blue)
        ws.cell(r, 2, b)
        ws.cell(r, 1).border = ws.cell(r, 2).border = border
        ws.cell(r, 1).alignment = Alignment(vertical="center")
        ws.cell(r, 2).alignment = Alignment(vertical="center", wrap_text=True)

    start = 12
    headers = ["ID tubería", "Longitud L (m)", "C", "D real (pulg)",
           "Q inicial (L/s)", "Q final (L/s)", "Hf final (mca)",
           "Hf 0.5–3 mca", "Compartida", "Circuitos / sentido"]
    header_row(ws, start, headers)
    circuit_map = {}
    for lp in loops:
        for m in lp["members"]:
            circuit_map.setdefault(m["pipe"], []).append(
                f"{lp['name']} ({'+' if m['dir'] > 0 else '-'})"
            )
    for r, (pid, p) in enumerate(pipes.items(), start + 1):
        pr = result["pipes"][pid]
        vals = [pid, p["L"], p["C"], p["D"], p["Q"], pr["Q_final"], pr["Hf"],
                "OK" if pr["h_range_ok"] else "FUERA",
                "Sí" if len(circuit_map[pid]) > 1 else "No",
                ", ".join(circuit_map[pid])]
        for c, v in enumerate(vals, 1):
            ws.cell(r, c, v).border = border
            ws.cell(r, c).alignment = Alignment(horizontal="center",
                                                vertical="center", wrap_text=True)
        for c in (7, 8):
            ws.cell(r, c).fill = PatternFill("solid",
                                             fgColor=green if pr["h_range_ok"] else red)
    ws.auto_filter.ref = f"A{start}:J{start + len(pipes)}"

    # ---------------- HARDY CROSS ----------------
    title(hc, "PROCEDIMIENTO DE HARDY CROSS – ITERACIONES", 14)
    hc["A3"] = "Caudal de entrada (L/s)"; hc["B3"] = q_in
    hc["A4"] = "Tolerancia Δ (L/s)"; hc["B4"] = result["threshold"]
    hc["A5"] = "Criterio"; hc["B5"] = f"|Δ| ≤ {result['tol_factor']} × Q entrada"
    hc["A6"] = "Iteraciones realizadas"; hc["B6"] = result["iterations"]
    hc["A7"] = "Estado"; hc["B7"] = "CONVERGIÓ" if result["converged"] else "NO CONVERGIÓ"
    hc["B7"].fill = PatternFill("solid", fgColor=green if result["converged"] else red)
    for r in range(3, 8):
        hc.cell(r, 1).font = Font(bold=True)
        hc.cell(r, 1).fill = PatternFill("solid", fgColor=blue)
        hc.cell(r, 1).border = hc.cell(r, 2).border = border

    row = 9
    table_headers = ["Circuito", "Tubería", "Sentido", "L (m)", "C", "D real (pulg)",
                     "Caudal antes", "Hf", "Hf firmado", "Hf/Q",
                     "Δ propio", "Δ del otro circuito",
                     "Δ total aplicado", "Caudal después"]
    for it in result["history"]:
        hc.merge_cells(start_row=row, start_column=1, end_row=row, end_column=14)
        cell = hc.cell(row, 1, f"ITERACIÓN {it['iteration']}")
        cell.font = Font(size=13, bold=True, color=white)
        cell.fill = PatternFill("solid", fgColor=dark)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        hc.row_dimensions[row].height = 22
        row += 1

        for lp in it["circuits"]:
            hc.merge_cells(start_row=row, start_column=1, end_row=row, end_column=14)
            cell = hc.cell(row, 1,
                f"{lp['name']}    |    ΣHf = {lp['sum_hf']:.6f} mca    |    "
                f"Σ(Hf/Q) = {lp['sum_hf_q']:.6f}    |    Δ = {lp['delta']:.6f} L/s")
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor=blue)
            cell.alignment = Alignment(horizontal="left", vertical="center")
            row += 1

            header_row(hc, row, table_headers)
            row += 1

            for d in lp["rows"]:
                delta_otro = d.get("delta_otro")
                delta_otro_str = delta_otro if delta_otro is not None else "—"
                vals = [
                    lp["name"], d["pipe"], "+" if d["dir"] > 0 else "-",
                    d["L"], d["C"], d["D"], d["Q_before"], d["Hf"], d["Hf_signed"],
                    d["Hf_Q"], d["delta_propio"], delta_otro_str,
                    d["delta_total"], d["Q_after"]
                ]
                for c, v in enumerate(vals, 1):
                    hc.cell(row, c, v).border = border
                    hc.cell(row, c).alignment = Alignment(horizontal="center",
                                                          vertical="center")
                if d.get("compartida"):
                    hc.cell(row, 12).fill = PatternFill("solid", fgColor=yellow)
                    hc.cell(row, 13).fill = PatternFill("solid", fgColor=orange)
                row += 1

            hc.cell(row, 1, "SUMA / CORRECCIÓN").font = Font(bold=True)
            hc.cell(row, 8, lp["sum_hf"])
            hc.cell(row, 9, lp["sum_hf"])
            hc.cell(row, 10, lp["sum_hf_q"])
            hc.cell(row, 11, lp["delta"])
            ok = lp["abs_delta"] <= result["threshold"]
            hc.cell(row, 14, "DETENER" if ok else "CONTINUAR")
            for c in range(1, 15):
                hc.cell(row, c).border = border
                hc.cell(row, c).alignment = Alignment(horizontal="center")
            hc.cell(row, 11).fill = PatternFill("solid", fgColor=green if ok else yellow)
            hc.cell(row, 14).fill = PatternFill("solid", fgColor=green if ok else yellow)
            row += 2

        maxd = it["max_delta"]
        ok = maxd <= result["threshold"]
        hc.merge_cells(start_row=row, start_column=1, end_row=row, end_column=14)
        hc.cell(row, 1,
                f"Resultado de iteración {it['iteration']}: |Δ| máximo = {maxd:.6f} L/s  →  "
                f"{'CONVERGE' if ok else 'SE CONTINÚA'}")
        hc.cell(row, 1).font = Font(bold=True)
        hc.cell(row, 1).fill = PatternFill("solid", fgColor=green if ok else yellow)
        row += 2

    # ---------------- RESUMEN ----------------
    title(sm, "RESUMEN DEL DISEÑO", 8)
    summary = [
        ("Caudal de entrada (L/s)", q_in),
        ("Tolerancia Δ (L/s)", result["threshold"]),
        ("Circuitos", len(loops)),
        ("Tuberías físicas", len(pipes)),
        ("Iteraciones", result["iterations"]),
        ("Convergencia", "Sí" if result["converged"] else "No"),
    ]
    for r, (a, b) in enumerate(summary, 3):
        sm.cell(r, 1, a).font = Font(bold=True)
        sm.cell(r, 1).fill = PatternFill("solid", fgColor=blue)
        sm.cell(r, 2, b)
        sm.cell(r, 1).border = sm.cell(r, 2).border = border

    header_row(sm, 11, ["Circuito", "ΣHf final (mca)", "Σ(Hf/Q) final",
                        "|Δ| final (L/s)", "Cumple Δ", "Observación"])
    final = result["history"][-1]["circuits"] if result["history"] else []
    for r, lp in enumerate(final, 12):
        vals = [lp["name"], lp["sum_hf"], lp["sum_hf_q"], lp["abs_delta"],
                "OK" if lp["abs_delta"] <= result["threshold"] else "NO",
                "Detener" if lp["abs_delta"] <= result["threshold"] else "Continuar iterando"]
        for c, v in enumerate(vals, 1):
            sm.cell(r, c, v).border = border
            sm.cell(r, c).alignment = Alignment(horizontal="center")
        sm.cell(r, 5).fill = PatternFill("solid", fgColor=green if vals[4] == "OK" else red)

    # Anchos y formato
    widths = {"A": 18, "B": 18, "C": 12, "D": 13, "E": 10, "F": 14, "G": 16,
              "H": 15, "I": 15, "J": 15, "K": 20, "L": 22, "M": 20, "N": 16}
    for sheet in wb.worksheets:
        sheet.sheet_view.showGridLines = False
        for col, w in widths.items():
            sheet.column_dimensions[col].width = w
        for rr in sheet.iter_rows():
            for cell in rr:
                if isinstance(cell.value, float):
                    cell.number_format = "0.00000"
                cell.alignment = Alignment(
                    vertical="center",
                    horizontal=cell.alignment.horizontal or "center",
                    wrap_text=True
                )
    hc.sheet_properties.pageSetUpPr.fitToPage = True
    hc.page_setup.fitToWidth = 1
    hc.page_setup.fitToHeight = 0
    hc.print_title_rows = "1:9"
    wb.save(filename)


# ---------------------------------------------------------------------------
# Interfaz gráfica
# ---------------------------------------------------------------------------
COLORS = {
    "bg": "#1e1e2e",
    "panel": "#2a2a3e",
    "panel2": "#33334d",
    "accent": "#7aa2f7",
    "accent2": "#bb9af7",
    "text": "#e0e0f0",
    "text_dim": "#a0a0b8",
    "entry_bg": "#3b3b55",
}


class HardyCrossApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Cálculos Hardy Cross → Método Iterativo")
        self.root.geometry("1250x820")
        self.root.minsize(1080, 700)
        self.root.configure(bg=COLORS["bg"])
        self.circuit_frames = []
        self.modo_diametro = "manual"
        self.v_diseno = 1.5
        self.crit_diam = "superior"
        self._setup_style()
        self.build_start()

    def _setup_style(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(".", background=COLORS["bg"], foreground=COLORS["text"],
                        fieldbackground=COLORS["entry_bg"], font=("Segoe UI", 10))
        style.configure("TFrame", background=COLORS["bg"])
        style.configure("Panel.TFrame", background=COLORS["panel"])
        style.configure("TLabel", background=COLORS["bg"], foreground=COLORS["text"])
        style.configure("Panel.TLabel", background=COLORS["panel"], foreground=COLORS["text"])
        style.configure("Dim.TLabel", background=COLORS["bg"], foreground=COLORS["text_dim"])
        style.configure("Title.TLabel", background=COLORS["bg"], foreground=COLORS["accent"],
                        font=("Segoe UI", 22, "bold"))
        style.configure("Sub.TLabel", background=COLORS["bg"], foreground=COLORS["text_dim"],
                        font=("Segoe UI", 10))
        style.configure("Card.TLabelframe", background=COLORS["panel"],
                        foreground=COLORS["accent2"], borderwidth=1, relief="solid")
        style.configure("Card.TLabelframe.Label", background=COLORS["panel"],
                        foreground=COLORS["accent2"], font=("Segoe UI", 11, "bold"))
        style.configure("TEntry", fieldbackground=COLORS["entry_bg"], foreground=COLORS["text"],
                        bordercolor=COLORS["panel2"], insertcolor=COLORS["text"])
        style.configure("TSpinbox", fieldbackground=COLORS["entry_bg"], foreground=COLORS["text"],
                        arrowcolor=COLORS["accent"], bordercolor=COLORS["panel2"])
        style.configure("Accent.TButton", background=COLORS["accent"], foreground="#1e1e2e",
                        font=("Segoe UI", 10, "bold"), borderwidth=0, focusthickness=0)
        style.map("Accent.TButton",
                  background=[("active", COLORS["accent2"]), ("pressed", COLORS["accent2"])])
        style.configure("Ghost.TButton", background=COLORS["panel2"],
                        foreground=COLORS["text"], borderwidth=0)
        style.map("Ghost.TButton", background=[("active", COLORS["panel"])])

    def clear(self):
        for w in self.root.winfo_children():
            w.destroy()

    # ---------------- Pantalla inicial ----------------
    def build_start(self):
        self.clear()
        outer = ttk.Frame(self.root, padding=30)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="Generador de Cálculos para Redes Cerradas por Hardy Cross y Hazem-Williams",
                  style="Title.TLabel").pack(anchor="w")
        ttk.Label(outer,
                  text="Diseño de redes de tuberías · Hazen-Williams · Iteraciones tipo Excel",
                  style="Sub.TLabel").pack(anchor="w", pady=(0, 20))

        card = ttk.Frame(outer, style="Panel.TFrame", padding=25)
        card.pack(fill="x", pady=10)

        ttk.Label(card, text="Parámetros iniciales", style="Panel.TLabel",
                  font=("Segoe UI", 13, "bold"))\
            .grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 15))
        ttk.Label(card, text="Caudal de entrada Q (L/s):", style="Panel.TLabel")\
            .grid(row=1, column=0, padx=(0, 15), pady=8, sticky="e")
        self.q_var = tk.StringVar()
        ttk.Entry(card, textvariable=self.q_var, width=20)\
            .grid(row=1, column=1, pady=8, sticky="w")

        ttk.Label(card, text="Cantidad de circuitos:", style="Panel.TLabel")\
            .grid(row=2, column=0, padx=(0, 15), pady=8, sticky="e")
        self.n_var = tk.IntVar(value=2)
        ttk.Spinbox(card, from_=1, to=20, textvariable=self.n_var, width=18)\
            .grid(row=2, column=1, pady=8, sticky="w")

        ttk.Label(card, text="Criterio de convergencia:", style="Panel.TLabel")\
            .grid(row=3, column=0, padx=(0, 15), pady=8, sticky="e")
        self.tol_var = tk.StringVar(value="0.0005")
        tol_frame = ttk.Frame(card, style="Panel.TFrame")
        tol_frame.grid(row=3, column=1, pady=8, sticky="w")
        ttk.Radiobutton(tol_frame, text="0.0005 × Q (clase)", value="0.005",
                        variable=self.tol_var).pack(side="left")
        ttk.Radiobutton(tol_frame, text="0.0001 × Q (estricto)", value="0.0001",
                        variable=self.tol_var).pack(side="left", padx=(15, 0))
        ttk.Button(outer, text="Continuar  →  Definir circuitos",
                   style="Accent.TButton", command=self.build_circuits)\
            .pack(pady=25)

        ttk.Label(outer,
                  text="Hazen-Williams: Hf = 1743.811 · L · Q^1.852 / (C^1.852 · D^4.87)   |   "
                       "Q en L/s, D en pulg, L en m",
                  style="Dim.TLabel").pack(side="bottom", pady=10)

    # ---------------- Pantalla de circuitos ----------------
    def build_circuits(self):
        try:
            q = float(self.q_var.get().replace(",", "."))
            n = int(self.n_var.get())
            if q <= 0 or n < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("Datos inválidos",
                                 "Ingrese un caudal positivo y una cantidad válida de circuitos.")
            return
        self.q_in = q
        self.tol_factor = float(self.tol_var.get())
        self.n_circuits = n
        self.clear()

        outer = ttk.Frame(self.root, padding=15)
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer)
        header.pack(fill="x", pady=(0, 10))
        ttk.Label(header, text="Definición de los circuitos",
                  style="Title.TLabel", font=("Segoe UI", 16, "bold")).pack(anchor="w")
        ttk.Label(header,
                  text="Use el mismo ID físico para una tubería compartida. "
                       "El sentido + / − indica su dirección respecto al recorrido del circuito.",
                  style="Sub.TLabel").pack(anchor="w", pady=(4, 0))

        # -------- Panel de modo de diámetro --------
        dm_card = ttk.Frame(outer, style="Panel.TFrame", padding=12)
        dm_card.pack(fill="x", pady=(6, 8))

        ttk.Label(dm_card, text="Modo de diámetro:",
                  style="Panel.TLabel", font=("Segoe UI", 10, "bold"))\
            .grid(row=0, column=0, padx=(0, 10), pady=4, sticky="e")

        self.modo_var = tk.StringVar(value=self.modo_diametro)
        mod_frame = ttk.Frame(dm_card, style="Panel.TFrame")
        mod_frame.grid(row=0, column=1, columnspan=4, pady=4, sticky="w")
        ttk.Radiobutton(mod_frame, text="Manual (ingreso yo el D)",
                        value="manual", variable=self.modo_var,
                        command=self._on_modo_change).pack(side="left")
        ttk.Radiobutton(mod_frame, text="Calcular por velocidad de diseño",
                        value="auto", variable=self.modo_var,
                        command=self._on_modo_change).pack(side="left", padx=(15, 0))

        ttk.Label(dm_card, text="Velocidad de diseño (m/s):",
                  style="Panel.TLabel")\
            .grid(row=1, column=0, padx=(0, 10), pady=4, sticky="e")
        self.v_var = tk.StringVar(value=str(self.v_diseno))
        self.v_entry = ttk.Entry(dm_card, textvariable=self.v_var, width=8)
        self.v_entry.grid(row=1, column=1, pady=4, sticky="w")

        ttk.Label(dm_card, text="Criterio de redondeo:",
                  style="Panel.TLabel")\
            .grid(row=1, column=2, padx=(15, 10), pady=4, sticky="e")
        self.crit_var = tk.StringVar(value=self.crit_diam)
        crit_frame = ttk.Frame(dm_card, style="Panel.TFrame")
        crit_frame.grid(row=1, column=3, pady=4, sticky="w")
        self.crit_sup = ttk.Radiobutton(crit_frame, text="Superior",
                                        value="superior", variable=self.crit_var)
        self.crit_sup.pack(side="left")
        self.crit_cer = ttk.Radiobutton(crit_frame, text="Más cercano",
                                        value="cercano", variable=self.crit_var)
        self.crit_cer.pack(side="left", padx=(10, 0))
        self.crit_inf = ttk.Radiobutton(crit_frame, text="Inferior",
                                        value="inferior", variable=self.crit_var)
        self.crit_inf.pack(side="left", padx=(10, 0))

        self.btn_calc_d = ttk.Button(dm_card, text="Calcular D comerciales",
                                     style="Ghost.TButton",
                                     command=self.calcular_diametros)
        self.btn_calc_d.grid(row=1, column=4, padx=(20, 0), pady=4)

        ttk.Label(dm_card,
                  text="Disp.: " + ", ".join(str(d) for d in DIAMETROS_COMERCIALES),
                  style="Panel.TLabel", font=("Segoe UI", 9, "italic"))\
            .grid(row=2, column=0, columnspan=5, pady=(6, 0), sticky="w")

        self._on_modo_change()

        # -------- Barra de herramientas --------
        tools = ttk.Frame(outer)
        tools.pack(fill="x", pady=(4, 8))
        ttk.Button(tools, text="Rellenar ejemplo (2 circuitos del Excel)",
                   style="Ghost.TButton", command=self.fill_example).pack(side="left")
        ttk.Button(tools, text="Limpiar todo", style="Ghost.TButton",
                   command=lambda: [self.create_rows(i, force=True)
                                    for i in range(len(self.circuit_frames))])\
            .pack(side="left", padx=8)

        canvas = tk.Canvas(outer, bg=COLORS["bg"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        body = ttk.Frame(canvas)
        body.bind("<Configure>",
                  lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=body, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        self.circuit_frames = []
        for i in range(n):
            lf = ttk.LabelFrame(body, text=f"  Circuito {i+1}  ",
                                style="Card.TLabelframe", padding=10)
            lf.pack(fill="x", pady=8, padx=4)

            top = ttk.Frame(lf, style="Panel.TFrame")
            top.pack(fill="x", pady=(0, 6))
            ttk.Label(top, text="N° de tuberías:", style="Panel.TLabel").pack(side="left")
            nv = tk.IntVar(value=4)
            ttk.Spinbox(top, from_=2, to=30, textvariable=nv, width=6)\
                .pack(side="left", padx=6)

            container = ttk.Frame(lf, style="Panel.TFrame")
            container.pack(fill="x", pady=4)

            data = {"frame": lf, "nvar": nv, "container": container, "rows": []}
            self.circuit_frames.append(data)

            nv.trace_add("write", lambda *a, idx=i: self.create_rows(idx))
            self.create_rows(i)

        bottom = ttk.Frame(outer)
        bottom.pack(fill="x", pady=12)
        ttk.Button(bottom, text="←  Atrás", style="Ghost.TButton",
                   command=self.build_start).pack(side="left")
        ttk.Button(bottom, text="Generar Excel y calcular Hardy Cross",
                   style="Accent.TButton", command=self.generate).pack(side="right")

    def _on_modo_change(self):
        auto = (self.modo_var.get() == "auto")
        state = "normal" if auto else "disabled"
        try:
            self.v_entry.configure(state=state)
            for w in (self.crit_sup, self.crit_cer, self.crit_inf):
                w.configure(state=state)
            self.btn_calc_d.configure(state=state)
        except Exception:
            pass

    def create_rows(self, idx, force=False):
        """Reconstruye las filas del circuito idx conservando los valores existentes."""
        data = self.circuit_frames[idx]
        container = data["container"]

        old_values = []
        if not force:
            for vars_ in data["rows"]:
                try:
                    old_values.append([v.get() for v in vars_])
                except tk.TclError:
                    pass

        for child in container.winfo_children():
            child.destroy()
        data["rows"] = []

        headers = ["ID físico", "L (m)", "C", "D real (pulg)",
                   "Q inicial (L/s)", "Sentido"]
        for c, h in enumerate(headers):
            ttk.Label(container, text=h, style="Panel.TLabel",
                      font=("Segoe UI", 9, "bold"))\
                .grid(row=0, column=c, padx=4, pady=(0, 4))

        try:
            count = int(data["nvar"].get())
        except (ValueError, tk.TclError):
            count = 4

        for r in range(count):
            vars_ = [tk.StringVar() for _ in range(6)]
            if r < len(old_values):
                for c in range(6):
                    vars_[c].set(old_values[r][c])
                if not vars_[0].get().strip():
                    vars_[0].set(f"P{idx+1:02d}-{r+1:02d}")
            else:
                vars_[0].set(f"P{idx+1:02d}-{r+1:02d}")
                vars_[5].set("+")

            for c, var in enumerate(vars_):
                e = ttk.Entry(container, textvariable=var,
                              width=8 if c == 5 else 14)
                e.grid(row=r+1, column=c, padx=4, pady=3)
            data["rows"].append(vars_)

    def calcular_diametros(self):
        """Recorre todas las filas, calcula D comercial con la velocidad de diseño."""
        try:
            v = float(self.v_var.get().replace(",", "."))
            if v <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Velocidad inválida",
                                 "Ingrese una velocidad de diseño positiva (m/s).")
            return
        self.v_diseno = v
        self.crit_diam = self.crit_var.get()
        n_calc = 0
        for data in self.circuit_frames:
            for vars_ in data["rows"]:
                try:
                    q_val = float(vars_[4].get().replace(",", "."))
                except ValueError:
                    continue
                if q_val == 0:
                    continue
                D_teor = diametro_teorico(abs(q_val), v)
                D_com = diametro_comercial(D_teor, self.crit_diam)
                vars_[3].set(f"{D_com:.2f}")
                n_calc += 1
        messagebox.showinfo(
            "Diámetros calculados",
            f"Se actualizaron {n_calc} diámetros usando v = {v} m/s "
            f"y criterio '{self.crit_diam}'.")

    def fill_example(self):
        """Rellena los circuitos con datos del Excel original."""
        if len(self.circuit_frames) < 2:
            messagebox.showinfo(
                "Ejemplo",
                "Cambia la cantidad de circuitos a 2 o más para cargar el ejemplo.")
            return
        ejemplo = [
            [("P01-01", "80", "140", "3", "-1.6", "-"),
             ("P01-02", "140", "140", "4", "-2.6", "-"),
             ("P01-03", "80", "140", "4", "6.4", "+"),
             ("P01-04", "140", "140", "2", "0.6", "+")],
            [("P02-01", "80", "140", "2", "-1.2", "-"),
             ("P02-02", "140", "140", "2", "-0.6", "-"),
             ("P02-03", "80", "140", "3", "4.8", "+"),
             ("P02-04", "140", "140", "2", "1.2", "+")],
        ]
        for i, rows_data in enumerate(ejemplo):
            if i >= len(self.circuit_frames):
                break
            data = self.circuit_frames[i]
            data["nvar"].set(len(rows_data))
            self.create_rows(i, force=True)
            for r, vals in enumerate(rows_data):
                for c, v in enumerate(vals):
                    data["rows"][r][c].set(v)

    def generate(self):
        try:
            # Actualizar velocidad y criterio (por si cambiaron)
            if self.modo_var.get() == "auto":
                try:
                    v = float(self.v_var.get().replace(",", "."))
                    if v <= 0:
                        raise ValueError
                except ValueError:
                    raise ValueError("Velocidad de diseño inválida.")
                self.v_diseno = v
                self.crit_diam = self.crit_var.get()
            self.modo_diametro = self.modo_var.get()

            pipes = {}
            loops = []
            for ci, data in enumerate(self.circuit_frames, 1):
                if not data["rows"]:
                    raise ValueError(f"El Circuito {ci} no tiene filas de tuberías.")
                members = []
                for ri, vars_ in enumerate(data["rows"], 1):
                    pid = vars_[0].get().strip()
                    if not pid:
                        raise ValueError(
                            f"Circuito {ci}, fila {ri}: falta el ID de tubería.")
                    try:
                        L = float(vars_[1].get().replace(",", "."))
                        C = float(vars_[2].get().replace(",", "."))
                        D = float(vars_[3].get().replace(",", "."))
                        Q = float(vars_[4].get().replace(",", "."))
                    except ValueError:
                        raise ValueError(
                            f"Circuito {ci}, fila {ri} ({pid}): "
                            f"L, C, D y Q deben ser números válidos.")
                    direction = 1 if vars_[5].get().strip() in ("+", "1", "+1") else -1
                    if L <= 0 or C <= 0 or D <= 0:
                        raise ValueError(
                            f"Circuito {ci}, fila {ri} ({pid}): "
                            f"L, C y D deben ser positivos.")
                    if pid not in pipes:
                        pipes[pid] = {"L": L, "C": C, "D": D, "Q": Q}
                    else:
                        p = pipes[pid]
                        if not (math.isclose(p["L"], L, rel_tol=1e-8, abs_tol=1e-8)
                                and math.isclose(p["C"], C, rel_tol=1e-8, abs_tol=1e-8)
                                and math.isclose(p["D"], D, rel_tol=1e-8, abs_tol=1e-8)):
                            raise ValueError(
                                f"La tubería compartida {pid} tiene L/C/D "
                                f"diferentes entre circuitos.")
                    members.append({"pipe": pid, "dir": direction})
                loops.append({"name": f"Circuito {ci}", "members": members})

            if not pipes:
                raise ValueError(
                    "No se ingresaron tuberías. "
                    "Verifica que cada circuito tenga al menos una fila con ID.")

            result = solve_hardy_cross(self.q_in, pipes, loops,
                                       tol_factor=self.tol_factor)

            # Aviso de Hf fuera de rango en la primera iteración
            fuera, dentro = revisar_hf_primera_iteracion(result)
            if fuera:
                lineas = []
                for pid, hf, sug in fuera:
                    lineas.append(f"  • {pid}: Hf = {hf:.3f} mca  →  {sug}")
                msg = ("En la primera iteración hay tuberías con Hf fuera del "
                       "rango 0.5 – 3.0 mca:\n\n" + "\n".join(lineas) +
                       "\n\nConsejo:\n"
                       "  – Si Hf < 0.5 mca, el diámetro es muy grande "
                       "(conviene reducir diámetro / subir velocidad).\n"
                       "  – Si Hf > 3.0 mca, el diámetro es muy pequeño "
                       "(conviene aumentar diámetro / bajar velocidad).\n\n"
                       "Ajustar los diámetros te ayudará a converger en menos "
                       "iteraciones.")
                messagebox.showwarning("Hf fuera de rango en iteración 1", msg)
            else:
                messagebox.showinfo(
                    "Hf en rango",
                    "Todas las tuberías tienen Hf dentro del rango 0.5 – 3.0 mca "
                    "en la primera iteración.")

            filename = filedialog.asksaveasfilename(
                title="Guardar Excel",
                defaultextension=".xlsx",
                initialfile="Hardy_Cross_generado.xlsx",
                filetypes=[("Excel", "*.xlsx")])
            if not filename:
                return

            make_excel(filename, self.q_in, pipes, loops, result,
                       v_diseno=self.v_diseno,
                       modo_diametro=self.modo_diametro)

            messagebox.showinfo(
                "Listo",
                f"Excel generado correctamente.\n\n"
                f"Circuitos: {len(loops)}\n"
                f"Tuberías físicas: {len(pipes)}\n"
                f"Iteraciones: {result['iterations']}\n"
                f"Convergencia: {'Sí' if result['converged'] else 'No'}")
        except Exception as exc:
            messagebox.showerror("No se pudo generar", str(exc))


def main():
    root = tk.Tk()
    HardyCrossApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
