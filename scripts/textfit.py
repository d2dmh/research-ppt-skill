"""Approximate CJK/Latin line wrapping used by layout stacking and lint.

Widths are in em (1 em = font size). Calibrated against 微软雅黑/Arial renders of the
reference style: about 0.36 in per 16 pt line including the 3 pt paragraph gap.
"""
import re

CLOSING = set('，。；：、！？）》」』】〕”’%．,.;:!?)]}…°')
OPENING = set('（《「『【〔“‘([{')
# Latin words chained by 、 (e.g. ARIMA、XGBoost、LSTM) are kept together: renderers often won't break there.
TOKEN = re.compile(r'[A-Za-z0-9][A-Za-z0-9.\-+/_%°×·]*(?:、[A-Za-z0-9][A-Za-z0-9.\-+/_%°×·]*)*|\s+|.', re.S)


def char_width(ch):
    o = ord(ch)
    if ch == ' ':
        return 0.28
    if o < 128:
        if ch in '.,;:!|\'`il':
            return 0.28
        if ch.isupper() or ch in 'mwMW@%':
            return 0.68
        return 0.55
    if 0x2070 <= o <= 0x209F:      # super/subscripts
        return 0.38
    if 0x0370 <= o <= 0x03FF or 0x2190 <= o <= 0x22FF or o in (0x00B0, 0x00B1, 0x00D7, 0x00B7, 0x2212):
        return 0.6
    if 0x2010 <= o <= 0x2027:      # dashes, quotes, ellipsis (narrow in YaHei)
        return 0.6
    return 1.0                     # CJK, fullwidth punctuation


def text_width(value):
    return sum(char_width(c) for c in value)


def wrap(line, width_em):
    """Split one paragraph into display lines (Latin words unbroken, CJK per char)."""
    if not line:
        return ['']
    lines, current, used = [], '', 0.0
    for token in TOKEN.findall(line):
        w = text_width(token)
        if current and used + w > width_em:
            stripped = token.strip()
            # Closing punctuation hangs on the previous line (hangingPunct=1).
            if stripped and all(c in CLOSING for c in stripped):
                current += token
                used += w
                continue
            if not token.strip():
                lines.append(current)
                current, used = '', 0.0
                continue
            # Opening punctuation should not stay alone at the end of a line.
            if current and current[-1] in OPENING:
                lines.append(current[:-1])
                current, used = current[-1], char_width(current[-1])
            else:
                lines.append(current)
                current, used = '', 0.0
        current += token
        used += w
    if current or not lines:
        lines.append(current)
    return lines


def line_height(size_pt, spacing=1.18, after_pt=3):
    return size_pt * 1.23 * spacing / 72 + after_pt / 72


def measure(value, size_pt, width_in, bold=False, bullet=False, spacing=None):
    """Return (display_lines, height_in) for a text box with the deck's paragraph settings."""
    em = size_pt / 72 * (1.04 if bold else 1.0)
    width_em = max(width_in - (0.28 if bullet else 0), 0.1) / em
    shown = []
    for paragraph in value.split('\n'):
        shown.extend(wrap(paragraph, width_em))
    return shown, len(shown) * line_height(size_pt, spacing or (1.3 if bullet else 1.18)) + 0.04


def visible_count(fragment):
    return sum(1 for c in fragment if not c.isspace() and c not in CLOSING and c not in OPENING)
