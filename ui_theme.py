"""Shared Tk theme for Resolve Whisper windows.

Design tokens from apps/DESIGN.md (LAB37 internal baseline, charcoal shell),
OKLCH converted to hex for Tk. Chroma-0 neutrals, color only for status,
0.90 off-white ceiling (anti-glare).

macOS Aqua ignores bg/fg on tk.Button and tk.OptionMenu (they render as
native white pills), so interactive widgets here are Label/Frame-based with
button behavior: hover lift, focus ring, keyboard activation. Verified
pattern shared with LineCut.
"""
import tkinter as tk

BRAND = "LAB37 TOOLS"

# --- Tokens (apps/DESIGN.md Part 2) ---
BG = "#171717"          # canvas (oklch 0.205)
CARD = "#242424"        # L1 surface (0.26)
INPUT = "#2B2B2B"       # inputs, menus, tracks (0.29)
SECONDARY = "#313131"   # secondary button fill (0.315)
HOVER = "#333333"       # neutral hover/active lift (0.32)
BORDER = "#393939"      # (0.345)
EDGE = "#2B2B2B"        # 1px top edge light on canvas
FG = "#DEDEDE"          # foreground, 0.90 ceiling
FG_ACTIVE = "#D1D1D1"   # primary button pressed (0.86)
MUTED = "#989898"       # muted-foreground (0.68, AA)
DISABLED = "#696969"    # (0.52)
RING = "#A4A4A4"        # focus ring (0.72)
RED = "#FF6467"         # destructive
AMBER = "#DFA11A"       # warning
SUCCESS = "#4CC157"     # success

FONT_UI = "Helvetica Neue"
FONT_MONO = "Menlo"


def edge_light(root):
    """1px top edge light: the one depth cue Tk can render (DESIGN.md 1.2)."""
    tk.Frame(root, bg=EDGE, height=1).pack(fill="x", side="top")


def brand_header(parent, title, subtitle=None):
    """Brand eyebrow + window title (+ optional muted subtitle)."""
    tk.Label(parent, text=BRAND, font=(FONT_MONO, 9), fg=MUTED, bg=BG
             ).pack(anchor="w")
    tk.Label(parent, text=title, font=(FONT_UI, 13, "bold"), fg=FG, bg=BG
             ).pack(anchor="w", pady=(6, 0))
    if subtitle:
        tk.Label(parent, text=subtitle, font=(FONT_UI, 10), fg=MUTED, bg=BG
                 ).pack(anchor="w", pady=(2, 0))


def section_label(parent, text, pady=(14, 6)):
    tk.Label(parent, text=text.upper(), font=(FONT_UI, 9, "bold"),
             fg=MUTED, bg=BG).pack(anchor="w", pady=pady)


def make_button(parent, text, kind="secondary", command=None):
    """Label styled as a button (Aqua ignores colors on real tk.Button)."""
    styles = {
        "primary": dict(bg=FG, fg=BG, hover=FG_ACTIVE,
                        font=(FONT_UI, 11, "bold"), padx=16, pady=6),
        "secondary": dict(bg=SECONDARY, fg=FG, hover=BORDER,
                          font=(FONT_UI, 11), padx=16, pady=6),
        "ghost": dict(bg=BG, fg=MUTED, hover=BG,
                      font=(FONT_UI, 10), padx=4, pady=2),
    }
    st = styles[kind]
    btn = tk.Label(
        parent, text=text, font=st["font"], fg=st["fg"], bg=st["bg"],
        padx=st["padx"], pady=st["pady"], cursor="hand2", takefocus=1,
        highlightthickness=1, highlightbackground=parent["bg"],
        highlightcolor=RING,
    )
    state = {"enabled": True}

    def _hover(_e):
        if state["enabled"]:
            btn.config(bg=st["hover"], fg=FG if kind == "ghost" else st["fg"])

    def _rest(_e=None):
        if state["enabled"]:
            btn.config(bg=st["bg"], fg=st["fg"])

    def _click(_e):
        if state["enabled"] and command:
            command()
        return "break"

    btn.bind("<Enter>", _hover)
    btn.bind("<Leave>", _rest)
    btn.bind("<Button-1>", _click)
    btn.bind("<Return>", _click)
    btn.bind("<space>", _click)

    def set_enabled(flag):
        state["enabled"] = flag
        if flag:
            btn.config(bg=st["bg"], fg=st["fg"], cursor="hand2")
        else:
            btn.config(bg=SECONDARY if kind == "primary" else st["bg"],
                       fg=DISABLED, cursor="arrow")

    btn.set_enabled = set_enabled
    btn.set_text = lambda t: btn.config(text=t)
    btn.is_enabled = lambda: state["enabled"]
    return btn


def make_dropdown(parent, label_text, options, initial_value,
                  width=20, label_width=18):
    """Labeled select. options: list of (label, value); returns a get() fn.

    Aqua ignores colors on OptionMenu's menubutton, so the face is a
    token-styled Frame + Labels; options post as a native popup menu,
    which follows the system dark theme.
    """
    row = tk.Frame(parent, bg=BG)
    row.pack(fill="x", pady=(0, 6))
    tk.Label(row, text=label_text, font=(FONT_UI, 10),
             fg=MUTED, bg=BG, width=label_width, anchor="w").pack(side="left")

    by_label = dict(options)
    labels = [lbl for lbl, _ in options]
    start = labels[0]
    for lbl, val in options:
        if val == initial_value:
            start = lbl
            break
    var = tk.StringVar(value=start)

    face = tk.Frame(row, bg=INPUT, highlightthickness=1,
                    highlightbackground=BORDER, cursor="hand2")
    face.pack(side="left")
    value_lbl = tk.Label(face, textvariable=var, font=(FONT_UI, 10),
                         fg=FG, bg=INPUT, width=width, anchor="w",
                         padx=8, pady=3)
    value_lbl.pack(side="left")
    chevron = tk.Label(face, text="▾", font=(FONT_UI, 9),
                       fg=MUTED, bg=INPUT, padx=6)
    chevron.pack(side="left")

    menu = tk.Menu(face, tearoff=0, font=(FONT_UI, 12))
    for lbl in labels:
        menu.add_command(label=lbl, command=lambda l=lbl: var.set(l))

    def _post(_e=None):
        try:
            menu.tk_popup(face.winfo_rootx(),
                          face.winfo_rooty() + face.winfo_height())
        finally:
            menu.grab_release()
        return "break"

    def _hover(_e):
        face.config(highlightbackground=RING)

    def _rest(_e):
        face.config(highlightbackground=BORDER)

    for wdg in (face, value_lbl, chevron):
        wdg.bind("<Button-1>", _post)
        wdg.bind("<Enter>", _hover)
        wdg.bind("<Leave>", _rest)

    return lambda: by_label[var.get()]


def make_toggle(parent, label_text, initial=False, label_width=18):
    """Labeled checkbox (native Aqua checkbuttons ignore dark bg).

    Box + check glyph built from Labels; the whole row is clickable.
    Returns a get() fn.
    """
    state = {"on": bool(initial)}
    row = tk.Frame(parent, bg=BG, cursor="hand2")
    row.pack(fill="x", pady=(0, 6))
    tk.Label(row, text=label_text, font=(FONT_UI, 10),
             fg=MUTED, bg=BG, width=label_width, anchor="w").pack(side="left")

    box = tk.Frame(row, bg=INPUT, highlightthickness=1,
                   highlightbackground=BORDER, cursor="hand2")
    box.pack(side="left")
    mark = tk.Label(box, text="", font=(FONT_UI, 11, "bold"),
                    fg=BG, bg=INPUT, width=2, pady=0)
    mark.pack()

    def _paint():
        if state["on"]:
            mark.config(text="✓", bg=FG, fg=BG)
            box.config(bg=FG)
        else:
            mark.config(text="", bg=INPUT)
            box.config(bg=INPUT)

    def _flip(_e=None):
        state["on"] = not state["on"]
        _paint()
        return "break"

    def _hover(_e):
        box.config(highlightbackground=RING)

    def _rest(_e):
        box.config(highlightbackground=BORDER)

    for wdg in (row, box, mark):
        wdg.bind("<Button-1>", _flip)
        wdg.bind("<Enter>", _hover)
        wdg.bind("<Leave>", _rest)

    _paint()
    return lambda: state["on"]


def place_window(root, y_divisor=3):
    """Size to content and center horizontally, upper third vertically."""
    root.update_idletasks()
    w = root.winfo_reqwidth()
    h = root.winfo_reqheight()
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    x = (sw - w) // 2
    y = (sh - h) // y_divisor
    root.geometry(f"+{x}+{y}")
