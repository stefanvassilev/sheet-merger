#!/usr/bin/env python3
"""
Excel Sheet Merger
==================
Interactively merge sheets from multiple Excel files into one workbook,
preserving all formatting, formulas, charts, and styles via Excel's COM API.

Usage:
    python merger.py [directory]

    directory  Path to scan for Excel files (default: directory of this script)

Requirements:
    Run setup.ps1, or manually:
        pip install pywin32 windows-curses
        python Scripts/pywin32_postinstall.py -install
"""

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from typing import Dict, List, Set

# ── Dependency check ──────────────────────────────────────────────────────────

try:
    import win32com.client
    import pywintypes
except ImportError:
    print("ERROR: pywin32 is not installed.")
    print()
    print("Run setup.ps1, or manually:")
    print("  pip install pywin32 windows-curses")
    print("  python Scripts/pywin32_postinstall.py -install")
    sys.exit(1)

import curses


# ── Data model ────────────────────────────────────────────────────────────────

@dataclass
class SheetEntry:
    file_path: str   # absolute path to the source Excel file
    file_key: str    # unique identifier for this file (used as default prefix)
    sheet_name: str  # original sheet name inside the source file
    custom_name: str = ""  # if non-empty, overrides the prefix-computed final name


# ── Excel / name helpers ──────────────────────────────────────────────────────

def find_excel_files(directory: str) -> List[str]:
    result: List[str] = []
    extensions = {".xlsx", ".xls", ".xlsm", ".xlsb", ".csv"}
    for root, _dirs, files in os.walk(directory):
        for f in files:
            if os.path.splitext(f)[1].lower() in extensions:
                result.append(os.path.join(root, f))
    return sorted(result)


def get_sheet_names(path: str, excel) -> List[str]:
    wb = excel.Workbooks.Open(path, UpdateLinks=False, ReadOnly=True)
    try:
        return [wb.Sheets(i + 1).Name for i in range(wb.Sheets.Count)]
    finally:
        wb.Close(False)


def sanitize(name: str) -> str:
    """Sanitize a string to be a valid Excel sheet name (max 31 chars)."""
    for ch in r'\/:*?[]':
        name = name.replace(ch, '_')
    return name[:31]


def unique_name(base: str, used: Set[str]) -> str:
    """Return `base` if unused, otherwise append _2, _3, … (respecting 31-char limit)."""
    if base not in used:
        return base
    n = 2
    while True:
        suf = f"_{n}"
        candidate = base[: 31 - len(suf)] + suf
        if candidate not in used:
            return candidate
        n += 1


# ── Colour pair IDs ──────────────────────────────────────────────────────────

_CP_HEADER = 1   # title bar: bold white on blue
_CP_COLHDR = 2   # column header text: bold cyan
_CP_SEL    = 3   # selected / cursor row: white on blue
_CP_GRAB   = 4   # grabbed row: bold yellow
_CP_BORDER = 5   # box-drawing characters: cyan
_CP_GREEN  = 6   # status ready: green
_CP_YELLOW = 7   # status grabbed / edit prompt: yellow
# Cycling file-group tints (4 colours)
_CP_FILE   = [8, 9, 10, 11]


# ── TUI ───────────────────────────────────────────────────────────────────────

class TUI:
    """
    Full-screen curses TUI.

    Layout (row numbers relative to terminal height H):
        0          : title bar (bold white-on-blue)
        1          : ├──#──┬──Final Name──┬──Source──┬──Prefix──┤
        2          : column header labels
        3          : ├──────┼─────────────┼──────────┼──────────┤
        4 … L+3    : data rows  (L = list_height)
        L+4        : ├──────┴─────────────┴──────────┴──────────┤
        L+5        : │ Showing X–Y of N …                       │
        L+6        : ├──────────────────────────────────────────┤
        L+7        : │ ↑↓ Navigate  SPACE Grab/Drop  …          │
        L+8        : │ status line                               │
    """

    # Box-drawing character table
    _B = dict(
        tl="┌", tr="┐", bl="└", br="┘",
        h="─",  v="│",
        lm="├", rm="┤",
        td="┬", tu="┴", x="┼",
    )

    # Footer rows below the list (sep + count + sep + ctrl + status)
    _FOOTER = 5
    # Rows above the list (title + sep + colhdr + sep)
    _HEADER = 4

    def __init__(self, stdscr, sheets: List[SheetEntry], prefixes: Dict[str, str]):
        self.stdscr = stdscr
        self.sheets = sheets
        self.prefixes = prefixes
        self.cursor = 0
        self.offset = 0
        self.grabbed       = False   # single-sheet grab (Space)
        self.grabbed_group = False   # whole-file grab (G)
        self.grabbed_key: str = ""   # file_key of the grabbed group
        self.confirmed = False
        self.alive = True
        self.flash = ""        # one-shot status message

        # Assign a cycling tint index to each unique file_key
        self.tint: Dict[str, int] = {}
        idx = 0
        for e in sheets:
            if e.file_key not in self.tint:
                self.tint[e.file_key] = idx % len(_CP_FILE)
                idx += 1

        self.colors = False

    # ── Colour setup ─────────────────────────────────────────────────────────

    def _init_colors(self) -> None:
        if not curses.has_colors():
            return
        try:
            curses.start_color()
            curses.use_default_colors()
            curses.init_pair(_CP_HEADER, curses.COLOR_WHITE,   curses.COLOR_BLUE)
            curses.init_pair(_CP_COLHDR, curses.COLOR_CYAN,    -1)
            curses.init_pair(_CP_SEL,    curses.COLOR_WHITE,   curses.COLOR_BLUE)
            curses.init_pair(_CP_GRAB,   curses.COLOR_YELLOW,  -1)
            curses.init_pair(_CP_BORDER, curses.COLOR_CYAN,    -1)
            curses.init_pair(_CP_GREEN,  curses.COLOR_GREEN,   -1)
            curses.init_pair(_CP_YELLOW, curses.COLOR_YELLOW,  -1)
            for i, fg in enumerate(
                [curses.COLOR_WHITE, curses.COLOR_CYAN,
                 curses.COLOR_MAGENTA, curses.COLOR_GREEN]
            ):
                curses.init_pair(_CP_FILE[i], fg, -1)
            self.colors = True
        except Exception:
            pass

    def _a(self, cp: int, bold: bool = False, dim: bool = False) -> int:
        """Return curses attribute for colour pair + optional bold/dim."""
        attr = curses.color_pair(cp) if self.colors else 0
        if bold:
            attr |= curses.A_BOLD
        if dim:
            attr |= curses.A_DIM
        return attr

    # ── Entry point ──────────────────────────────────────────────────────────

    def run(self) -> bool:
        self._init_colors()
        curses.curs_set(0)
        self.stdscr.keypad(True)
        while self.alive:
            self._draw()
            self._handle_key()
        return self.confirmed

    # ── Drawing helpers ──────────────────────────────────────────────────────

    def _put(self, row: int, col: int, text: str, attr: int = 0) -> None:
        H, W = self.stdscr.getmaxyx()
        if row < 0 or row >= H or col < 0 or col >= W:
            return
        text = text[: W - col]
        try:
            self.stdscr.addstr(row, col, text, attr)
        except curses.error:
            pass

    def _col_widths(self, W: int):
        """Return (num_w, sht_w, src_w, pfx_w) that together fill width W."""
        # Fixed or bounded column widths
        num_w = 5                           # "▶ 12 "
        pfx_w = max(8,  min(16, W // 7))
        src_w = max(14, min(26, W // 4))
        # Remaining space goes to the sheet-name column
        # Total borders: 1+num_w+1+(sht_w+2)+1+(src_w+2)+1+(pfx_w+2)+1 = 11+num+sht+src+pfx
        sht_w = max(10, W - 11 - num_w - src_w - pfx_w)
        return num_w, sht_w, src_w, pfx_w

    def _hline_cols(self, W: int, num_w: int, sht_w: int,
                    src_w: int, pfx_w: int, cross: str) -> str:
        """Build a full-width column-separated horizontal line."""
        B = self._B
        return (
            B["lm"]
            + B["h"] * num_w
            + cross
            + B["h"] * (sht_w + 2)
            + cross
            + B["h"] * (src_w + 2)
            + cross
            + B["h"] * (pfx_w + 2)
            + B["rm"]
        )[:W]

    def _hline_full(self, W: int) -> str:
        """Build a simple full-width horizontal line (no column divisions)."""
        B = self._B
        return (B["lm"] + B["h"] * (W - 2) + B["rm"])[:W]

    # ── Main draw ────────────────────────────────────────────────────────────

    def _draw(self) -> None:
        self.stdscr.erase()
        H, W = self.stdscr.getmaxyx()

        MIN_H, MIN_W = self._HEADER + 1 + self._FOOTER, 44
        if H < MIN_H or W < MIN_W:
            self._put(0, 0,
                      f"Terminal too small ({W}x{H}) — resize to at least {MIN_W}x{MIN_H}.",
                      curses.A_BOLD)
            self.stdscr.refresh()
            return

        num_w, sht_w, src_w, pfx_w = self._col_widths(W)
        list_h = H - self._HEADER - self._FOOTER

        # Row indices
        r_title = 0
        r_sep1  = 1
        r_chdr  = 2
        r_sep2  = 3
        r_list0 = 4
        r_sep3  = r_list0 + list_h
        r_count = r_sep3  + 1
        r_sep4  = r_count + 1
        r_ctrl  = r_sep4  + 1
        r_stat  = r_ctrl  + 1

        B = self._B

        # ── Title ─────────────────────────────────────────────────────────
        title = " \u2726 Excel Sheet Merger \u2726 "
        self._put(r_title, 0, title.center(W), self._a(_CP_HEADER, bold=True))

        # ── Column separators ─────────────────────────────────────────────
        self._put(r_sep1, 0,
                  self._hline_cols(W, num_w, sht_w, src_w, pfx_w, B["td"]),
                  self._a(_CP_BORDER))
        self._put(r_sep2, 0,
                  self._hline_cols(W, num_w, sht_w, src_w, pfx_w, B["x"]),
                  self._a(_CP_BORDER))
        self._put(r_sep3, 0,
                  self._hline_cols(W, num_w, sht_w, src_w, pfx_w, B["tu"]),
                  self._a(_CP_BORDER))

        # ── Column header row ─────────────────────────────────────────────
        bar = self._a(_CP_BORDER)
        hdr = self._a(_CP_COLHDR, bold=True)
        c = 0
        self._put(r_chdr, c, B["v"], bar);                  c += 1
        self._put(r_chdr, c, f"{'#':>{num_w-1}} ",          hdr); c += num_w
        self._put(r_chdr, c, B["v"], bar);                  c += 1
        self._put(r_chdr, c, f" {'Final Sheet Name':<{sht_w}} ", hdr); c += sht_w + 2
        self._put(r_chdr, c, B["v"], bar);                  c += 1
        self._put(r_chdr, c, f" {'Source File':<{src_w}} ", hdr); c += src_w + 2
        self._put(r_chdr, c, B["v"], bar);                  c += 1
        self._put(r_chdr, c, f" {'Prefix':<{pfx_w}} ",      hdr); c += pfx_w + 2
        self._put(r_chdr, c, B["v"], bar)

        # ── List rows ─────────────────────────────────────────────────────
        self._draw_list(r_list0, list_h, W, num_w, sht_w, src_w, pfx_w)

        # ── Count row ─────────────────────────────────────────────────────
        vis_end = min(self.offset + list_h, len(self.sheets))
        count_str = f" Showing {self.offset + 1}\u2013{vis_end} of {len(self.sheets)} "
        self._put(r_count, 0,
                  (B["v"] + count_str.ljust(W - 2) + B["v"])[:W],
                  self._a(_CP_BORDER, dim=True))

        # ── Dividers for footer ───────────────────────────────────────────
        self._put(r_sep4, 0, self._hline_full(W), self._a(_CP_BORDER))

        # ── Controls bar ──────────────────────────────────────────────────
        parts = [
            "\u2191\u2193 Navigate",
            "SPACE Grab sheet",
            "G Grab file group",
            "DEL Remove sheet",
            "SDEL Remove file",
            "E Edit prefix",
            "R Rename sheet",
            "ENTER Confirm",
            "Q Quit",
        ]
        ctrl = "  \u2502  ".join(parts)
        ctrl_line = (B["v"] + " " + ctrl)[: W - 1] + B["v"]
        self._put(r_ctrl, 0, ctrl_line, self._a(_CP_BORDER, dim=True))

        # ── Status line ───────────────────────────────────────────────────
        self._draw_status(r_stat, W)

        self.stdscr.refresh()

    def _draw_list(self, top: int, list_h: int, W: int,
                   num_w: int, sht_w: int, src_w: int, pfx_w: int) -> None:
        # Keep cursor inside the visible window
        if self.cursor < self.offset:
            self.offset = self.cursor
        elif self.cursor >= self.offset + list_h:
            self.offset = max(0, self.cursor - list_h + 1)

        B = self._B
        bar = self._a(_CP_BORDER)

        for i in range(list_h):
            idx = self.offset + i
            row = top + i

            if idx >= len(self.sheets):
                # Empty filler row (just borders)
                self._put(row, 0,
                          B["v"] + " " * (W - 2) + B["v"],
                          bar)
                continue

            e = self.sheets[idx]
            pfx    = self.prefixes.get(e.file_key, e.file_key)
            final  = e.custom_name if e.custom_name else sanitize(f"{pfx}_{e.sheet_name}")
            src    = os.path.basename(e.file_path)

            is_sel        = idx == self.cursor
            is_grab       = is_sel and self.grabbed
            is_grab_group = self.grabbed_group and e.file_key == self.grabbed_key

            if is_grab or is_grab_group:
                row_attr = self._a(_CP_GRAB, bold=True)
            elif is_sel:
                row_attr = self._a(_CP_SEL, bold=True)
            else:
                cp = _CP_FILE[self.tint.get(e.file_key, 0)]
                row_attr = self._a(cp)

            indicator = "\u25b6 " if is_sel else "  "
            num_str   = f"{indicator}{idx + 1}"
            fn_cell   = f"[{final}]" if (is_grab or is_grab_group) else final

            c = 0
            self._put(row, c, B["v"], bar);                          c += 1
            self._put(row, c, f"{num_str:>{num_w-1}} "[:num_w],    row_attr); c += num_w
            self._put(row, c, B["v"], bar);                          c += 1
            self._put(row, c, f" {fn_cell:<{sht_w}} "[:sht_w+2],   row_attr); c += sht_w + 2
            self._put(row, c, B["v"], bar);                          c += 1
            self._put(row, c, f" {src:<{src_w}} "[:src_w+2],        row_attr); c += src_w + 2
            self._put(row, c, B["v"], bar);                          c += 1
            self._put(row, c, f" {pfx:<{pfx_w}} "[:pfx_w+2],        row_attr); c += pfx_w + 2
            self._put(row, c, B["v"], bar)

    def _draw_status(self, row: int, W: int) -> None:
        B = self._B
        if self.flash:
            msg  = f" \u2192 {self.flash} "
            attr = self._a(_CP_YELLOW, bold=True)
            self.flash = ""
        elif self.grabbed and self.sheets:
            e   = self.sheets[self.cursor]
            pfx = self.prefixes.get(e.file_key, e.file_key)
            nm  = e.custom_name if e.custom_name else sanitize(f"{pfx}_{e.sheet_name}")
            msg  = (f" \u25ba Grabbed \"{nm}\" "
                    f"\u2014 \u2191\u2193 to move, SPACE to drop, ESC to cancel ")
            attr = self._a(_CP_YELLOW, bold=True)
        elif self.grabbed_group:
            count = sum(1 for e in self.sheets if e.file_key == self.grabbed_key)
            src   = next((os.path.basename(e.file_path)
                          for e in self.sheets if e.file_key == self.grabbed_key), "")
            msg  = (f" \u25ba Grabbed {count} sheet(s) from \"{src}\" "
                    f"\u2014 \u2191\u2193 to move group, SPACE to drop, ESC to cancel ")
            attr = self._a(_CP_YELLOW, bold=True)
        else:
            n_files = len({e.file_key for e in self.sheets})
            msg  = (f" \u2714 Ready \u2014 {len(self.sheets)} sheet(s)"
                    f" from {n_files} file(s)  |  ENTER to merge ")
            attr = self._a(_CP_GREEN)

        line = (B["v"] + msg.ljust(W - 2) + B["v"])[:W]
        self._put(row, 0, line, attr)

    # ── Key handling ─────────────────────────────────────────────────────────

    def _handle_key(self) -> None:
        key = self.stdscr.getch()
        if key == curses.KEY_RESIZE:
            return
        if self.grabbed:
            self._key_grab(key)
        elif self.grabbed_group:
            self._key_grab_group(key)
        else:
            self._key_normal(key)

    def _key_normal(self, key: int) -> None:
        n = len(self.sheets)
        if key == curses.KEY_UP:
            self.cursor = max(0, self.cursor - 1)
        elif key == curses.KEY_DOWN:
            self.cursor = min(n - 1, self.cursor + 1)
        elif key == curses.KEY_HOME:
            self.cursor = 0
        elif key == curses.KEY_END:
            self.cursor = max(0, n - 1)
        elif key == ord(' ') and n > 0:
            self.grabbed = True
        elif key in (ord('g'), ord('G')) and n > 0:
            self._grab_group()
        elif key == curses.KEY_DC and n > 0:
            self._remove_sheet()
        elif key == curses.KEY_SDC and n > 0:
            self._remove_file_group()
        elif key in (ord('e'), ord('E')) and n > 0:
            self._edit_prefix()
        elif key in (ord('r'), ord('R')) and n > 0:
            self._edit_sheet_name()
        elif key in (10, 13, curses.KEY_ENTER):
            if not self.sheets:
                self.flash = "Nothing to merge \u2014 add files first"
            else:
                self.confirmed = True
                self.alive = False
        elif key in (ord('q'), ord('Q'), 27):   # q / Q / Esc
            self.alive = False

    def _key_grab(self, key: int) -> None:
        n = len(self.sheets)
        if key == curses.KEY_UP and self.cursor > 0:
            i = self.cursor
            self.sheets[i], self.sheets[i - 1] = self.sheets[i - 1], self.sheets[i]
            self.cursor -= 1
        elif key == curses.KEY_DOWN and self.cursor < n - 1:
            i = self.cursor
            self.sheets[i], self.sheets[i + 1] = self.sheets[i + 1], self.sheets[i]
            self.cursor += 1
        elif key in (ord(' '), 27):   # Space = drop, Esc = release
            self.grabbed = False

    def _grab_group(self) -> None:
        """Grab all sheets from the current sheet's file, compact them into a
        contiguous block at the first occurrence, then enter group-grab mode."""
        if not self.sheets:
            return
        fk = self.sheets[self.cursor].file_key

        # Count non-group sheets that appear before the first group sheet
        first_group_idx = next(i for i, e in enumerate(self.sheets) if e.file_key == fk)
        non_group_before = sum(
            1 for e in self.sheets[:first_group_idx] if e.file_key != fk
        )

        # Rebuild list: others (in original order) with group inserted as a block
        group  = [e for e in self.sheets if e.file_key == fk]
        others = [e for e in self.sheets if e.file_key != fk]
        self.sheets[:] = others[:non_group_before] + group + others[non_group_before:]

        self.cursor       = non_group_before   # first sheet of the group
        self.grabbed_group = True
        self.grabbed_key   = fk

    def _remove_sheet(self) -> None:
        """Remove only the sheet at the cursor."""
        if not self.sheets:
            return
        e = self.sheets[self.cursor]
        self.sheets.pop(self.cursor)
        self.cursor = min(self.cursor, max(0, len(self.sheets) - 1))
        self.flash = f'Removed "{e.sheet_name}"'

    def _remove_file_group(self) -> None:
        """Remove all sheets belonging to the current sheet's source file."""
        if not self.sheets:
            return
        fk      = self.sheets[self.cursor].file_key
        src     = os.path.basename(self.sheets[self.cursor].file_path)
        removed = sum(1 for e in self.sheets if e.file_key == fk)
        self.sheets[:] = [e for e in self.sheets if e.file_key != fk]
        self.cursor = min(self.cursor, max(0, len(self.sheets) - 1))
        self.flash = f'Removed {removed} sheet(s) from "{src}"'

    def _key_grab_group(self, key: int) -> None:
        """Move the grabbed file group as a contiguous block."""
        fk    = self.grabbed_key
        count = sum(1 for e in self.sheets if e.file_key == fk)
        first = next((i for i, e in enumerate(self.sheets) if e.file_key == fk), None)
        if first is None:
            self.grabbed_group = False
            return
        last = first + count - 1

        if key == curses.KEY_UP and first > 0:
            # Take the sheet just above the block and move it below the block
            elem = self.sheets.pop(first - 1)
            self.sheets.insert(last, elem)   # last index shifted down by 1 after pop
            self.cursor = first - 1
        elif key == curses.KEY_DOWN and last < len(self.sheets) - 1:
            # Take the sheet just below the block and move it above the block
            elem = self.sheets.pop(last + 1)
            self.sheets.insert(first, elem)
            self.cursor = first + 1
        elif key in (ord(' '), 27):   # Space = drop, Esc = release
            self.grabbed_group = False
            self.grabbed_key   = ""

    def _edit_prefix(self) -> None:
        if not self.sheets:
            return
        e       = self.sheets[self.cursor]
        fk      = e.file_key
        current = self.prefixes.get(fk, fk)
        src     = os.path.basename(e.file_path)
        H, W    = self.stdscr.getmaxyx()
        prompt  = f" Prefix for \"{src}\": "
        value   = list(current)

        curses.curs_set(1)
        while True:
            last = H - 1
            self.stdscr.move(last, 0)
            self.stdscr.clrtoeol()
            display = prompt + ''.join(value) + '\u2588'   # ▌ block cursor
            self._put(last, 0, display[: W - 1], self._a(_CP_YELLOW, bold=True))
            cx = min(len(prompt) + len(value), W - 2)
            try:
                self.stdscr.move(last, cx)
            except curses.error:
                pass
            self.stdscr.refresh()

            k = self.stdscr.getch()
            if k in (10, 13, curses.KEY_ENTER):
                break
            elif k == 27:                              # Esc → cancel
                value = list(current)
                break
            elif k in (curses.KEY_BACKSPACE, 127, 8):
                if value:
                    value.pop()
            elif 32 <= k <= 126:
                value.append(chr(k))

        curses.curs_set(0)
        new_pfx = ''.join(value).strip()
        if new_pfx:
            self.prefixes[fk] = new_pfx
            self.flash = f'Prefix updated to "{new_pfx}"'
        else:
            self.flash = "Prefix cannot be empty \u2014 keeping original"

    def _edit_sheet_name(self) -> None:
        if not self.sheets:
            return
        e       = self.sheets[self.cursor]
        pfx     = self.prefixes.get(e.file_key, e.file_key)
        current = e.custom_name if e.custom_name else sanitize(f"{pfx}_{e.sheet_name}")
        H, W    = self.stdscr.getmaxyx()
        prompt  = f" Name for \"{e.sheet_name}\": "
        value   = list(current)
        cancelled = False

        curses.curs_set(1)
        while True:
            last = H - 1
            self.stdscr.move(last, 0)
            self.stdscr.clrtoeol()
            display = prompt + ''.join(value) + '\u2588'
            self._put(last, 0, display[: W - 1], self._a(_CP_YELLOW, bold=True))
            cx = min(len(prompt) + len(value), W - 2)
            try:
                self.stdscr.move(last, cx)
            except curses.error:
                pass
            self.stdscr.refresh()

            k = self.stdscr.getch()
            if k in (10, 13, curses.KEY_ENTER):
                break
            elif k == 27:                              # Esc → cancel
                cancelled = True
                break
            elif k in (curses.KEY_BACKSPACE, 127, 8):
                if value:
                    value.pop()
            elif 32 <= k <= 126:
                value.append(chr(k))

        curses.curs_set(0)
        if cancelled:
            self.flash = "Rename cancelled"
            return
        new_name = sanitize(''.join(value).strip())
        if new_name:
            e.custom_name = new_name
            self.flash = f'Sheet renamed to "{new_name}"'
        else:
            e.custom_name = ""
            self.flash = "Name cleared \u2014 using prefix-computed name"


# ── Merge ─────────────────────────────────────────────────────────────────────

_MERGE_VBS = (
    "Option Explicit\n"
    "\n"
    "' merger.py helper: reads tab-delimited manifest and merges sheets via Excel COM.\n"
    "' DEST<TAB>output_path\n"
    "' SHEET<TAB>src_path<TAB>sheet_name<TAB>final_name\n"
    "\n"
    "Dim fso, ts, ln, parts\n"
    "Dim excel, destWb, destWbPath, initial\n"
    "Dim srcPath, sheetName, finalName, srcWb\n"
    "Dim openWbs, k, i\n"
    "\n"
    "Set fso     = CreateObject(\"Scripting.FileSystemObject\")\n"
    "Set openWbs = CreateObject(\"Scripting.Dictionary\")\n"
    "\n"
    "Set excel = CreateObject(\"Excel.Application\")\n"
    "excel.Visible = False\n"
    "excel.DisplayAlerts = False\n"
    "On Error Resume Next\n"
    "excel.AskToUpdateLinks = False\n"
    "On Error Resume Next\n"
    "\n"
    "Set destWb = Nothing\n"
    "destWbPath = \"\"\n"
    "initial    = 0\n"
    "\n"
    "Set ts = fso.OpenTextFile(WScript.Arguments(0), 1, False, -1)\n"
    "\n"
    "Do While Not ts.AtEndOfStream\n"
    "    ln = Trim(ts.ReadLine())\n"
    "    If Len(ln) > 0 Then\n"
    "        parts = Split(ln, Chr(9))\n"
    "        If UBound(parts) >= 1 Then\n"
    "            If parts(0) = \"DEST\" Then\n"
    "                destWbPath = parts(1)\n"
    "            ElseIf parts(0) = \"SHEET\" And UBound(parts) >= 3 Then\n"
    "                srcPath   = parts(1)\n"
    "                sheetName = parts(2)\n"
    "                finalName = parts(3)\n"
    "                If destWb Is Nothing Then\n"
    "                    Set destWb = excel.Workbooks.Add()\n"
    "                    initial = destWb.Sheets.Count\n"
    "                End If\n"
    "                If Not openWbs.Exists(srcPath) Then\n"
    "                    openWbs.Add srcPath, excel.Workbooks.Open(srcPath, False, True)\n"
    "                End If\n"
    "                Set srcWb = openWbs(srcPath)\n"
    "                srcWb.Sheets(sheetName).Copy , destWb.Sheets(destWb.Sheets.Count)\n"
    "                destWb.Sheets(destWb.Sheets.Count).Name = finalName\n"
    "                WScript.Echo \"OK:\" & finalName\n"
    "            End If\n"
    "        End If\n"
    "    End If\n"
    "Loop\n"
    "ts.Close\n"
    "\n"
    "For Each k In openWbs.Keys\n"
    "    openWbs(k).Close False\n"
    "Next\n"
    "\n"
    "If Not (destWb Is Nothing) Then\n"
    "    For i = 1 To initial\n"
    "        If destWb.Sheets.Count > 1 Then\n"
    "            destWb.Sheets(1).Delete\n"
    "        End If\n"
    "    Next\n"
    "    destWb.CheckCompatibility = False\n"
    "    destWb.SaveAs destWbPath, 51\n"
    "    destWb.Close False\n"
    "End If\n"
    "\n"
    "excel.Quit\n"
    "WScript.Quit 0\n"
)


def merge_files(
    sheets: List[SheetEntry],
    prefixes: Dict[str, str],
    output_path: str,
    progress_cb=None,
) -> None:
    """
    Merge sheets into a new workbook via a VBScript helper.

    Using a VBScript (run through cscript.exe) sidesteps COM-reference issues
    that prevent Python's win32com from reliably copying sheets across workbooks.
    VBScript's native COM layer handles cross-workbook Copy correctly.
    """
    abs_out = os.path.abspath(output_path)

    # Pre-compute final sheet names
    used: Set[str] = set()
    entries = []
    for entry in sheets:
        pfx  = prefixes.get(entry.file_key, entry.file_key)
        base = entry.custom_name if entry.custom_name else sanitize(f"{pfx}_{entry.sheet_name}")
        name = unique_name(base, used)
        used.add(name)
        entries.append((entry.file_path, entry.sheet_name, name))

    tmp_dir = tempfile.mkdtemp()
    try:
        # Manifest — tab-delimited, written as UTF-16 so VBScript reads it correctly
        manifest = os.path.join(tmp_dir, "manifest.txt")
        lines = [f"DEST\t{abs_out}"]
        for fp, sname, fname in entries:
            lines.append(f"SHEET\t{fp}\t{sname}\t{fname}")
        with open(manifest, "w", encoding="utf-16") as f:
            f.write("\n".join(lines))

        # VBScript
        vbs = os.path.join(tmp_dir, "merge.vbs")
        with open(vbs, "w", encoding="utf-8") as f:
            f.write(_MERGE_VBS)

        # Stream stdout for real-time progress
        proc = subprocess.Popen(
            ["cscript", "//Nologo", vbs, manifest],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        err_lines: List[str] = []
        for raw in proc.stdout:             # type: ignore[union-attr]
            raw = raw.rstrip()
            if raw.startswith("OK:") and progress_cb:
                progress_cb(raw[3:])
            elif raw.startswith("ERR:"):
                err_lines.append(raw[4:])
        proc.wait()
        stderr_out = proc.stderr.read()     # type: ignore[union-attr]

        if proc.returncode != 0 or err_lines:
            msg = "\n".join(err_lines) or stderr_out or "VBScript merge failed"
            raise RuntimeError(msg.strip())

    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ── Main ──────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Merge sheets from multiple Excel files into one workbook, "
            "preserving all formatting."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="Example:\n  python merger.py C:\\Reports",
    )
    parser.add_argument(
        "directory",
        nargs="?",
        default=os.path.dirname(os.path.abspath(__file__)),
        help="Directory to scan for Excel files (default: script location)",
    )
    args = parser.parse_args()

    directory = os.path.abspath(args.directory)
    if not os.path.isdir(directory):
        sys.exit(f"Error: '{directory}' is not a valid directory.")

    print(f"Scanning (recursively): {directory}")
    files = find_excel_files(directory)
    if not files:
        sys.exit("No Excel or CSV files found in that directory (searched recursively).")

    print(f"Found {len(files)} file(s) (Excel + CSV). Reading sheet names via Excel...")

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
    except Exception as exc:
        sys.exit(
            f"Could not start Excel COM server: {exc}\n"
            "Make sure Microsoft Excel is installed and try running\n"
            "  python Scripts/pywin32_postinstall.py -install"
        )

    sheets: List[SheetEntry] = []
    prefixes: Dict[str, str] = {}
    used_keys: Set[str] = set()

    try:
        for fp in files:
            base = os.path.splitext(os.path.basename(fp))[0]
            key, suffix = base, 1
            while key in used_keys:
                key = f"{base}_{suffix}"
                suffix += 1
            used_keys.add(key)
            prefixes[key] = key

            try:
                names = get_sheet_names(fp, excel)
                for n in names:
                    sheets.append(SheetEntry(fp, key, n))
                print(f"  {os.path.basename(fp)}: {', '.join(names)}")
            except pywintypes.com_error as exc:
                print(f"  Warning: skipping {os.path.basename(fp)}: {exc}")
    finally:
        excel.Quit()

    if not sheets:
        sys.exit("No readable sheets found.")

    # ── Launch TUI ────────────────────────────────────────────────────────────
    confirmed = False

    def _tui(stdscr):
        nonlocal confirmed
        confirmed = TUI(stdscr, sheets, prefixes).run()

    curses.wrapper(_tui)

    if not confirmed:
        sys.exit("\nCancelled.")

    # ── Output path ───────────────────────────────────────────────────────────
    print()
    raw = input("Output filename [merged_output.xlsx]: ").strip() or "merged_output.xlsx"
    out = raw if os.path.isabs(raw) else os.path.join(directory, raw)
    if not out.lower().endswith((".xlsx", ".xls", ".xlsm")):
        out += ".xlsx"

    if os.path.exists(out):
        ans = input(
            f"'{os.path.basename(out)}' already exists. Overwrite? [y/N] "
        ).strip().lower()
        if ans != "y":
            sys.exit("Aborted.")
        # Delete before merging — Excel cannot SaveAs over a file that was
        # opened (even read-only) in the same session; deleting it first
        # means Excel always creates a fresh file with no locking conflict.
        try:
            os.remove(out)
        except OSError as e:
            sys.exit(f"Cannot delete existing '{os.path.basename(out)}': {e}\n"
                     "Close it in Excel and try again.")

    # ── Merge ─────────────────────────────────────────────────────────────────
    print(f"\nMerging {len(sheets)} sheet(s) \u2192 {out}")
    print()

    def _progress(name: str) -> None:
        print(f"  \u2192 {name}")

    try:
        merge_files(sheets, prefixes, out, progress_cb=_progress)
    except (RuntimeError, pywintypes.com_error) as exc:
        sys.exit(f"\nMerge failed: {exc}")
    except PermissionError:
        sys.exit(
            f"\nCannot write '{out}'.\n"
            "Is the file already open in Excel?"
        )

    # ── Verify output ─────────────────────────────────────────────────────────
    print("\nVerifying output...")
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            wb = excel.Workbooks.Open(out, UpdateLinks=False, ReadOnly=True)
            actual   = wb.Sheets.Count
            expected = len(sheets)
            sheet_names = [wb.Sheets(i + 1).Name for i in range(actual)]
            wb.Close(False)
        finally:
            excel.Quit()

        if actual == expected:
            print(f"\u2713 Verified: {actual} sheet(s) in output")
            for n in sheet_names:
                print(f"  \u2022 {n}")
        else:
            print(
                f"Warning: expected {expected} sheet(s) but found {actual}.\n"
                f"Output may be incomplete."
            )
            for n in sheet_names:
                print(f"  \u2022 {n}")
    except Exception as exc:
        print(f"Warning: could not verify output file: {exc}")

    print(f"\n\u2713 Done!  Saved to: {out}")


if __name__ == "__main__":
    main()
