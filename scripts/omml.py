"""Small LaTeX subset -> native Office Math (OMML) plus a linear Unicode fallback.

Supported: \\frac{a}{b}, x_{i}, x^{2}, \\sqrt{x}, \\hat/\\bar/\\tilde/\\dot/\\overline{x},
\\text{..}/\\mathrm{..}, Greek letters, common operators (\\sum \\cdot \\times \\le \\ge \\approx
\\langle \\rangle \\Delta \\partial \\infty \\to ...), upright functions (\\max \\min \\log \\exp ...).
Anything else is written literally, so check the render.
"""
import re
from lxml import etree

M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

SYMBOLS = {
    'alpha': 'α', 'beta': 'β', 'gamma': 'γ', 'delta': 'δ', 'epsilon': 'ε', 'varepsilon': 'ε', 'zeta': 'ζ',
    'eta': 'η', 'theta': 'θ', 'lambda': 'λ', 'mu': 'μ', 'nu': 'ν', 'xi': 'ξ', 'pi': 'π', 'rho': 'ρ',
    'sigma': 'σ', 'tau': 'τ', 'phi': 'φ', 'varphi': 'φ', 'chi': 'χ', 'psi': 'ψ', 'omega': 'ω', 'kappa': 'κ',
    'Gamma': 'Γ', 'Delta': 'Δ', 'Theta': 'Θ', 'Lambda': 'Λ', 'Pi': 'Π', 'Sigma': 'Σ', 'Phi': 'Φ',
    'Psi': 'Ψ', 'Omega': 'Ω',
    'sum': '∑', 'prod': '∏', 'int': '∫', 'partial': '∂', 'infty': '∞', 'nabla': '∇',
    'cdot': '·', 'times': '×', 'pm': '±', 'mp': '∓', 'div': '÷', 'le': '≤', 'leq': '≤', 'ge': '≥',
    'geq': '≥', 'neq': '≠', 'ne': '≠', 'approx': '≈', 'sim': '∼', 'propto': '∝', 'equiv': '≡',
    'in': '∈', 'notin': '∉', 'subset': '⊂', 'cup': '∪', 'cap': '∩', 'forall': '∀', 'exists': '∃',
    'to': '→', 'rightarrow': '→', 'leftarrow': '←', 'Rightarrow': '⇒', 'leftrightarrow': '↔',
    'langle': '⟨', 'rangle': '⟩', 'lfloor': '⌊', 'rfloor': '⌋', 'lceil': '⌈', 'rceil': '⌉',
    'ldots': '…', 'cdots': '⋯', 'dots': '…', 'prime': '′', 'degree': '°', '%': '%', '{': '{', '}': '}',
    ',': ' ', ';': ' ', ':': ' ', '!': '', 'quad': ' ', 'qquad': '  ',
    '\\': ' ',
}
FUNCTIONS = {'max', 'min', 'log', 'ln', 'exp', 'sin', 'cos', 'tan', 'arg', 'lim', 'det', 'mod', 'sgn'}
ACCENTS = {'hat': '̂', 'bar': '̅', 'overline': '̅', 'tilde': '̃', 'dot': '̇',
           'vec': '⃗'}
SUB = str.maketrans('0123456789+-=()aehijklmnoprstuvx', '₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓ')
SUP = str.maketrans('0123456789+-=()abcdefghijklmnoprstuvwxyz', '⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ᵃᵇᶜᵈᵉᶠᵍʰⁱʲᵏˡᵐⁿᵒᵖʳˢᵗᵘᵛʷˣʸᶻ')
SUB_KEYS = set('0123456789+-=()aehijklmnoprstuvx')
SUP_KEYS = set('0123456789+-=()abcdefghijklmnoprstuvwxyz')


class Parser:
    def __init__(self, source):
        self.s, self.i = source, 0

    def peek(self):
        return self.s[self.i] if self.i < len(self.s) else None

    def group(self):
        while self.peek() is not None and self.peek().isspace():
            self.i += 1
        if self.peek() == '{':
            self.i += 1
            nodes = self.sequence('}')
            self.i += 1
            return nodes
        return [self.atom()]

    def command(self):
        self.i += 1
        start = self.i
        if self.peek() is not None and not self.peek().isalpha():
            self.i += 1
            return self.s[start:self.i]
        while self.peek() is not None and self.peek().isalpha():
            self.i += 1
        return self.s[start:self.i]

    def atom(self):
        ch = self.peek()
        if ch == '\\':
            name = self.command()
            if name == 'frac':
                return ('frac', self.group(), self.group())
            if name == 'sqrt':
                return ('sqrt', self.group())
            if name in ACCENTS:
                return ('acc', ACCENTS[name], self.group())
            if name in ('text', 'mathrm', 'operatorname', 'mathrm'):
                return ('up', flatten(self.group()))
            if name in ('left', 'right', 'big', 'Big', 'bigl', 'bigr'):
                nxt = self.atom()
                return nxt
            if name in FUNCTIONS:
                return ('up', name)
            return ('t', SYMBOLS.get(name, name))
        if ch == '{':
            return ('grp', self.group())
        self.i += 1
        return ('t', {'-': '−', '*': '·'}.get(ch, ch))

    def sequence(self, stop=None):
        nodes = []
        while self.peek() is not None and self.peek() != stop:
            if self.peek().isspace():   # LaTeX math ignores source spaces
                self.i += 1
                continue
            if self.peek() in '_^':
                base = nodes.pop() if nodes else ('t', '')
                if base[0] == 't' and len(base[1]) > 1:
                    # Scripts bind to the last symbol only (TeX semantics): split merged text.
                    nodes.append(('t', base[1][:-1]))
                    base = ('t', base[1][-1])
                sub = sup = None
                while self.peek() in ('_', '^'):
                    mark = self.peek()
                    self.i += 1
                    if mark == '_':
                        sub = self.group()
                    else:
                        sup = self.group()
                nodes.append(('script', base, sub, sup))
                continue
            node = self.atom()
            if node[0] == 't' and nodes and nodes[-1][0] == 't' and node[1] not in '⟨⟩' and nodes[-1][1][-1:] not in '⟨⟩':
                nodes[-1] = ('t', nodes[-1][1] + node[1])
            else:
                nodes.append(node)
        return nodes


def flatten(nodes):
    return ''.join(linear([n]) for n in nodes)


def linear(nodes):
    out = []
    for node in nodes:
        kind = node[0]
        if kind in ('t', 'up'):
            out.append(node[1])
        elif kind == 'grp':
            out.append(linear(node[1]))
        elif kind == 'frac':
            num, den = linear(node[1]), linear(node[2])
            wrap = lambda v: v if len(v) <= 2 else '(' + v + ')'
            out.append(wrap(num) + '/' + wrap(den))
        elif kind == 'sqrt':
            out.append('√(' + linear(node[1]) + ')')
        elif kind == 'acc':
            out.append(linear(node[2]) + node[1])
        elif kind == 'script':
            base = linear([node[1]])
            for part, table, keys, mark in ((node[2], SUB, SUB_KEYS, '_'), (node[3], SUP, SUP_KEYS, '^')):
                if part is None:
                    continue
                value = linear(part)
                if value and all(c in keys for c in value):
                    base += value.translate(table)
                else:
                    base += mark + ('(' + value + ')' if len(value) > 1 else value)
            out.append(base)
    return ''.join(out)


def _m(tag, parent=None):
    return etree.Element('{%s}%s' % (M, tag)) if parent is None else etree.SubElement(parent, '{%s}%s' % (M, tag))


def _run(parent, value, rpr_factory, upright=False):
    run = _m('r', parent)
    if upright:
        mrpr = _m('rPr', run)
        sty = _m('sty', mrpr)
        sty.set('{%s}val' % M, 'p')
    run.append(rpr_factory())
    t = _m('t', run)
    t.text = value
    if value != value.strip():
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return run


def _emit(nodes, parent, rpr):
    for node in nodes:
        kind = node[0]
        if kind == 't':
            _run(parent, node[1], rpr)
        elif kind == 'up':
            _run(parent, node[1], rpr, upright=True)
        elif kind == 'grp':
            _emit(node[1], parent, rpr)
        elif kind == 'frac':
            f = _m('f', parent)
            _emit(node[1], _m('num', f), rpr)
            _emit(node[2], _m('den', f), rpr)
        elif kind == 'sqrt':
            rad = _m('rad', parent)
            props = _m('radPr', rad)
            hide = _m('degHide', props)
            hide.set('{%s}val' % M, '1')
            _m('deg', rad)
            _emit(node[1], _m('e', rad), rpr)
        elif kind == 'acc':
            acc = _m('acc', parent)
            props = _m('accPr', acc)
            chr_ = _m('chr', props)
            chr_.set('{%s}val' % M, node[1])
            _emit(node[2], _m('e', acc), rpr)
        elif kind == 'script':
            _, base, sub, sup = node
            tag = 'sSubSup' if sub is not None and sup is not None else 'sSub' if sub is not None else 'sSup'
            holder = _m(tag, parent)
            _emit([base], _m('e', holder), rpr)
            if sub is not None:
                _emit(sub, _m('sub', holder), rpr)
            if sup is not None:
                _emit(sup, _m('sup', holder), rpr)


def to_omml(source, rpr_factory, justify='left'):
    """Return (m:oMathPara element, linear unicode fallback)."""
    nodes = Parser(source).sequence()
    para = etree.Element('{%s}oMathPara' % M, nsmap={'m': M})
    props = _m('oMathParaPr', para)
    jc = _m('jc', props)
    jc.set('{%s}val' % M, {'left': 'left', 'center': 'centerGroup', 'right': 'right'}.get(justify, 'left'))
    math = _m('oMath', para)
    _emit(nodes, math, rpr_factory)
    fallback = re.sub(r'\s*([=+−≤≥≈≠×→])\s*', r' \1 ', linear(nodes)).strip()
    return para, re.sub(r'([(\[]) ([−+]) ', r'\1\2', fallback)


def plain_omml(value, rpr_factory, justify='left'):
    para = etree.Element('{%s}oMathPara' % M, nsmap={'m': M})
    props = _m('oMathParaPr', para)
    jc = _m('jc', props)
    jc.set('{%s}val' % M, {'left': 'left', 'center': 'centerGroup', 'right': 'right'}.get(justify, 'left'))
    math = _m('oMath', para)
    _run(math, value, rpr_factory)
    return para, value
