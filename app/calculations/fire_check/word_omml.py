"""
Уравнения Word через Office Math ML (parse_xml): дроби, корни, степени, индексы.
Шрифт: Cambria Math в блоке формулы.
"""

from __future__ import annotations

import xml.sax.saxutils as xml_esc
from typing import Any, Literal

MixedPart = tuple[Literal["text", "math"], str]

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

MNS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Как в Raschet_SN.docx: Cambria Math 12 pt (sz 24) внутри m:r
_MATH_W_RPR = (
    f'<w:rPr xmlns:w="{WNS}">'
    f'<w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math" w:cs="Arial"/>'
    f'<w:sz w:val="24"/><w:szCs w:val="24"/>'
    f"</w:rPr>"
)


def _xe(s: str) -> str:
    return xml_esc.escape(s)


def _math_plain_text(s: str) -> str:
    """Cambria Math плохо рисует ℃ (U+2103) и длинное тире в m:t — даёт «?» и пустые квадраты."""
    return (
        s.replace("\u2103", "\u00B0C")
        .replace("℃", "\u00B0C")
        .replace("\u2014", "-")
        .replace("\u2013", "-")
    )


def _omml_text_needs_nor(s: str) -> bool:
    """Cambria Math: кириллица, греческие Θ/θ, цифры и «,» — через m:nor (иначе пустые квадраты в Word)."""
    for c in s:
        if "\u0370" <= c <= "\u04ff":
            return True
        if c in ",.;·−":
            return True
        if c.isdigit():
            return True
    return False


def _wrap_omath(inner: str, *, math_para_center: bool = False) -> str:
    """Обёртка m:oMathPara — в WordprocessingML этот узел должен быть прямым потомком w:p, не w:r."""
    jc = ""
    if math_para_center:
        jc = '<m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>'
    return (
        f'<m:oMathPara xmlns:m="{MNS}">'
        f"{jc}"
        f"<m:oMath>{inner}</m:oMath>"
        f"</m:oMathPara>"
    )


def append_omath_xml(
    doc: Any,
    inner_omath: str,
    *,
    centered: bool = True,
    font_pt: int = 14,
    body_indent: bool = False,
) -> None:
    """Добавляет абзац с m:oMathPara (числовые формулы Θн, B, Θк в эталоне)."""
    p = doc.add_paragraph()
    _apply_body_paragraph_format(p, body_indent=body_indent, centered=centered)
    _clear_paragraph_runs(p)

    # Обёртка <p xmlns:m> — иначе parse_xml(m:oMathPara) даёт OMML, который Word не открывает.
    wrapped = _wrap_omath(inner_omath, math_para_center=centered and not body_indent)
    omp = parse_xml(f'<p xmlns:m="{MNS}">{wrapped}</p>')[0]
    p._p.append(omp)


def _apply_body_paragraph_format(
    p: Any,
    *,
    body_indent: bool,
    centered: bool,
) -> None:
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = 1.15
    pf.space_before = Pt(3)
    pf.space_after = Pt(3)
    if body_indent:
        pf.first_line_indent = Cm(1.25)
        pf.left_indent = Cm(0)
        pf.right_indent = Cm(0)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    else:
        pf.first_line_indent = Cm(0)
        pf.left_indent = Cm(0)
        pf.right_indent = Cm(0)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if centered else WD_ALIGN_PARAGRAPH.JUSTIFY


def _clear_paragraph_runs(p: Any) -> None:
    for child in list(p._p):
        if child.tag == qn("w:r"):
            p._p.remove(child)


def _append_text_run(p: Any, text: str) -> None:
    r = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), "Times New Roman")
    fonts.set(qn("w:hAnsi"), "Times New Roman")
    r_pr.append(fonts)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "28")
    r_pr.append(sz)
    sz_cs = OxmlElement("w:szCs")
    sz_cs.set(qn("w:val"), "28")
    r_pr.append(sz_cs)
    r.append(r_pr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    p._p.append(r)


def append_inline_omath(
    doc: Any,
    inner_omath: str,
    *,
    body_indent: bool = True,
    centered: bool = True,
) -> None:
    """Одна формула как m:oMath в абзаце (как схемы Iп0 и Θн в эталоне)."""
    p = doc.add_paragraph()
    _apply_body_paragraph_format(p, body_indent=body_indent, centered=centered)
    _clear_paragraph_runs(p)
    omath = parse_xml(f'<m:oMath xmlns:m="{MNS}">{inner_omath}</m:oMath>')
    p._p.append(omath)


def append_mixed_paragraph(
    doc: Any,
    parts: list[MixedPart],
    *,
    body_indent: bool = True,
    centered: bool = False,
) -> None:
    """Абзац: Times New Roman + встроенные m:oMath (легенды «где…»)."""
    p = doc.add_paragraph()
    _apply_body_paragraph_format(p, body_indent=body_indent, centered=centered)
    _clear_paragraph_runs(p)
    for kind, content in parts:
        if kind == "text":
            if content:
                _append_text_run(p, content)
        elif kind == "math":
            if content:
                omath = parse_xml(f'<m:oMath xmlns:m="{MNS}">{content}</m:oMath>')
                p._p.append(omath)


def xml_mrun(text: str) -> str:
    if text is None:
        text = ""
    safe = _math_plain_text(str(text))
    if safe.strip() == "":
        safe = "-"
    t = _xe(safe)
    m_rpr = "<m:rPr><m:nor/></m:rPr>" if _omml_text_needs_nor(safe) else ""
    return f"<m:r>{_MATH_W_RPR}{m_rpr}<m:t xml:space=\"preserve\">{t}</m:t></m:r>"


def xml_msubsup(base: str, sub: str, sup: str) -> str:
    """Iп0⁽³⁾ — m:sSubSup как в Raschet_SN.docx."""
    return (
        "<m:sSubSup>"
        f"<m:e>{xml_mrun(base)}</m:e>"
        f"<m:sub>{xml_mrun(sub)}</m:sub>"
        f"<m:sup>{xml_mrun(sup)}</m:sup>"
        "</m:sSubSup>"
    )


def xml_mdelim(inner: str) -> str:
    """Круглые скобки m:d — как (ΘДД−ΘОКР) в эталоне."""
    return f"<m:d><m:dPr/><m:e>{inner}</m:e></m:d>"


def xml_msub(base: str, sub: str) -> str:
    """Подстрочный индекс: корень должен быть m:sSub (а не m:sub), иначе Word не рисует формулу."""
    return (
        "<m:sSub>"
        f"<m:e>{xml_mrun(base)}</m:e>"
        f"<m:sub>{xml_mrun(sub)}</m:sub>"
        "</m:sSub>"
    )


def xml_msup(base: str, sup: str) -> str:
    # m:sup — как m:sub: без лишнего m:e, иначе Word не открывает docx.
    return (
        "<m:sSup>"
        f"<m:e>{xml_mrun(base)}</m:e>"
        f"<m:sup>{xml_mrun(sup)}</m:sup>"
        "</m:sSup>"
    )


def xml_msup_expr(base_xml: str, sup: str) -> str:
    """Степень, когда основание — уже фрагмент OMML."""
    return (
        "<m:sSup>"
        f"<m:e>{base_xml}</m:e>"
        f"<m:sup>{xml_mrun(sup)}</m:sup>"
        "</m:sSup>"
    )


def xml_mfrac(num_inner: str, den_inner: str) -> str:
    return (
        "<m:f>"
        f"<m:num><m:e>{num_inner}</m:e></m:num>"
        f"<m:den><m:e>{den_inner}</m:e></m:den>"
        "</m:f>"
    )


def xml_msqrt(inner: str) -> str:
    """Корень без видимого индекса степени (degHide) — иначе в Word пустые квадраты."""
    return (
        "<m:rad>"
        "<m:radPr><m:degHide m:val=\"1\"/></m:radPr>"
        "<m:deg/>"
        f"<m:e>{inner}</m:e>"
        "</m:rad>"
    )


def xml_mexp(power_inner: str) -> str:
    return (
        "<m:sSup>"
        f"<m:e>{xml_mrun('e')}</m:e>"
        f"<m:sup>{power_inner}</m:sup>"
        "</m:sSup>"
    )


def omml_B_string(
    i_ka: str,
    t_s: str,
    tae_s: str,
    *,
    b_result: str | None = None,
) -> str:
    """Bтер = I²∙t + Tаэ; при b_result — итог и размерность в конце: … = 27,55 кА²·с."""
    b = xml_msub("B", "тер")
    i2 = xml_msup_expr(xml_mrun(i_ka), "2")
    dot = xml_mrun("\u2219")
    s = (
        f"{b}{xml_mrun(' = ')}{i2}{dot}{xml_mrun(t_s)}"
        f"{xml_mrun(' + ')}{xml_mrun(tae_s)}"
    )
    if b_result is not None:
        ka2 = xml_msup_expr(xml_mrun("кА"), "2")
        s += f"{xml_mrun(' = ')}{xml_mrun(b_result)}{ka2}{xml_mrun('\u00b7')}{xml_mrun('с')}"
    return s


def omml_B_schema_string() -> str:
    """Bтер = Iп0² · tоткл + Tаэ (без чисел)."""
    b = xml_msub("B", "тер")
    ip0 = xml_msub("I", "п0")
    i2 = xml_msup_expr(ip0, "2")
    toff = xml_msub("t", "откл")
    ta = xml_msub("T", "аэ")
    return f"{b}{xml_mrun(' = ')}{i2}{xml_mrun('·')}{toff}{xml_mrun(' + ')}{ta}"


def omml_theta_n_string(to_c: str, tdd_c: str, tok_c: str, iw: str, ia: str, tn: str) -> str:
    th = xml_msub("Θ", "н")
    frac = xml_mfrac(xml_mrun(iw), xml_mrun(ia))
    sq = xml_msup_expr(xml_mdelim(frac), "2")
    diff = f"{xml_mrun(tdd_c)}{xml_mrun('−')}{xml_mrun(tok_c)}"
    return (
        f"{th}{xml_mrun(' = ')}"
        f"{xml_mrun(to_c)}{xml_mrun('+')}{xml_mdelim(diff)}"
        f"{xml_mrun('·')}{sq}{xml_mrun(' = ')}{xml_mrun(tn)}{xml_mrun(' \u00B0C')}"
    )


def omml_Ip0_schema_string() -> str:
    """Iп0⁽³⁾ = U_ном / (√3 · √(R₁эк²+X₁эк²)) — как в тексте до подстановки чисел."""
    num = xml_msub("U", "ном")
    rsq = xml_msup_expr(xml_msub("R", "1эк"), "2")
    xsq = xml_msup_expr(xml_msub("X", "1эк"), "2")
    den = f"{xml_msqrt(xml_mrun('3'))}{xml_mrun('·')}{xml_msqrt(rsq + xml_mrun('+') + xsq)}"
    ip = xml_msubsup("I", "п0", "(3)")
    return f"{ip}{xml_mrun(' = ')}{xml_mfrac(num, den)}"


def omml_Ip0_numeric_mohm(
    u_nom_v: str,
    r1_mohm: str,
    x1_mohm: str,
    *,
    ik_ka: str | None = None,
) -> str:
    """Iп0⁽³⁾ = Uном / (√3·√(R₁²+X₁²)); при ik_ka — итог в конце: … = 5,247 кА."""
    ip = xml_msubsup("I", "п0", "(3)")
    rsq = xml_msup_expr(xml_mrun(r1_mohm), "2")
    xsq = xml_msup_expr(xml_mrun(x1_mohm), "2")
    den = f"{xml_msqrt(xml_mrun('3'))}{xml_mrun('\u2219')}{xml_msqrt(rsq + xml_mrun('+') + xsq)}"
    s = f"{ip}{xml_mrun(' = ')}{xml_mfrac(xml_mrun(u_nom_v), den)}"
    if ik_ka is not None:
        s += f"{xml_mrun(' = ')}{xml_mrun(ik_ka)}{xml_mrun(' кА')}"
    return s


def _omml_theta_k_exp_pair(b_inner: str, s_mm: str) -> tuple[str, str]:
    """e^(19,58·B/S²) и e^(19,58·B/S²−1) — стековая дробь, как в численной Θк."""
    s2 = xml_msup_expr(xml_mrun(s_mm), "2")
    num = f"{xml_mrun('19,58')}{xml_mrun('·')}{b_inner}"
    frac = xml_mfrac(num, s2)
    pow2 = f"{frac}{xml_mrun('−1')}"
    return xml_mexp(frac), xml_mexp(pow2)


def omml_theta_k_schema_string(s_mm: str = "S") -> str:
    """Θк = Θн·e^(19,58·Bтер/S²) + a·e^(19,58·Bтер/S²−1)."""
    thk = xml_msub("Θ", "к")
    thn = xml_msub("Θ", "н")
    e1, e2 = _omml_theta_k_exp_pair(xml_msub("B", "тер"), s_mm)
    return f"{thk}{xml_mrun(' = ')}{thn}{xml_mrun('·')}{e1}{xml_mrun(' + ')}{xml_mrun('a')}{xml_mrun('·')}{e2}"


def omml_theta_k_string(
    theta_k: str,
    s_mm: str,
    *,
    theta_n_num: str | None = None,
    b_ter_num: str | None = None,
) -> str:
    """Θк; при theta_n_num и b_ter_num — числа в показателях и перед e, иначе символы Θн и Bтер."""
    thk = xml_msub("Θ", "к")
    lead = xml_mrun(theta_n_num) if theta_n_num else xml_msub("Θ", "н")
    b_piece = xml_mrun(b_ter_num) if b_ter_num else xml_msub("B", "тер")
    e1, e2 = _omml_theta_k_exp_pair(b_piece, s_mm)
    return (
        f"{thk}{xml_mrun(' = ')}{lead}{xml_mrun('·')}{e1}{xml_mrun(' + 228·')}{e2}"
        f"{xml_mrun(' = ')}{xml_mrun(theta_k)}{xml_mrun(' \u00B0C')}"
    )


def omml_theta_n_schema_string() -> str:
    """Θн = Θо + (ΘДД − ΘОКР)·(Iраб/IДД)²."""
    thn = xml_msub("Θ", "н")
    tho = xml_msub("Θ", "о")
    tdd = xml_msub("Θ", "ДД")
    tok = xml_msub("Θ", "ОКР")
    irab = xml_msub("I", "раб")
    idd = xml_msub("I", "ДД")
    frac = xml_mfrac(irab, idd)
    sq = xml_msup_expr(xml_mdelim(frac), "2")
    diff = f"{tdd}{xml_mrun('−')}{tok}"
    return f"{thn}{xml_mrun(' = ')}{tho}{xml_mrun('+')}{xml_mdelim(diff)}{xml_mrun('·')}{sq}{xml_mrun(';')}"


def legend_where_ik_schema_parts() -> list[MixedPart]:
    """где Uном … R1эк и X1эк … (как в Raschet_SN)."""
    return [
        ("text", "где "),
        ("math", xml_msub("U", "ном")),
        ("text", "- номинальное напряжение сети (фазное значение); "),
        ("math", xml_msub("R", "1эк")),
        ("text", " и "),
        ("math", xml_msub("X", "1эк")),
        (
            "text",
            " - соответственно эквивалентное активное и индуктивное сопротивление "
            "прямой последовательности до точки к.з.",
        ),
    ]


def legend_where_theta_n_o_parts() -> list[MixedPart]:
    return [
        ("text", "где  "),
        ("math", xml_msub("Θ", "о")),
        ("text", " - температура окружающей среды во время КЗ,  "),
        ("math", xml_msub("Θ", "о")),
        ("text", " принимаем равной 25"),
        ("math", xml_mrun("\u00B0C")),
        ("text", ";"),
    ]


def legend_where_theta_n_dd_parts() -> list[MixedPart]:
    return [
        ("math", xml_msub("Θ", "ДД")),
        (
            "text",
            " – значение расчетной длительно допустимой температуры жилы, для кабелей "
            "с поливинилхлоридной изоляцией по циркулярам № Ц-02-98 (Э) принимаем равной 70 ",
        ),
        ("math", xml_mrun("\u00B0C")),
        ("text", ";"),
    ]


def legend_where_theta_n_okr_parts() -> list[MixedPart]:
    return [
        ("math", xml_msub("Θ", "ОКР")),
        ("text", " – значение расчетной температуры окружающей среды 25 "),
        ("math", xml_mrun("\u00B0C")),
        ("text", ";"),
    ]


def legend_where_theta_n_irab_parts() -> list[MixedPart]:
    return [
        ("math", xml_msub("I", "раб")),
        ("text", " – значение тока до КЗ, А(выбран по макс. нагрузке на ТСН);"),
    ]


def legend_where_theta_n_idd_parts() -> list[MixedPart]:
    return [
        ("math", xml_msub("I", "ДД")),
        ("text", " – значение расчетного длительно допустимого тока в соответствии с ПУЭ, А."),
    ]


def legend_where_B_parts() -> list[MixedPart]:
    return [
        ("text", "где "),
        ("math", xml_msub("I", "п0")),
        ("text", "– значение трехфазного тока КЗ в конце участка кабельной линии от ТСН до ЩСН, кА; "),
        ("math", xml_msub("t", "откл")),
        ("text", " – время отключения тока КЗ, с;"),
    ]


def legend_where_B_tae_parts() -> list[MixedPart]:
    return [
        ("text", "- эквивалентная постоянная времени затухания апериодической составляющей тока КЗ от удаленных источников "),
        ("text", "по циркуляру № Ц-02-98 (Э) для сети 0,4 кВ принимаем равной "),
        ("math", xml_mrun("0,02")),
        ("text", " с."),
    ]


def legend_where_theta_k_parts() -> list[MixedPart]:
    return [
        ("text", "где "),
        ("math", xml_msub("Θ", "н")),
        ("text", "– температура жилы до КЗ, "),
        ("math", xml_mrun("\u00B0C")),
        ("text", "; в – постоянная, характеризующая теплофизические характеристики материала жилы, для меди 19,58 "),
        ("math", xml_mrun("мм\u2074/(кА\u00B2\u2219с)")),
        ("text", ";"),
    ]


def legend_where_theta_k_b_parts() -> list[MixedPart]:
    return [
        ("text", " "),
        ("math", xml_msub("B", "тер")),
        ("text", " – тепловой импульс от тока КЗ, "),
        ("math", xml_mrun("кА\u00B2\u2219с")),
        ("text", ";"),
    ]


def legend_where_theta_k_s_parts() -> list[MixedPart]:
    return [
        ("text", "S – сечение жилы кабельной линии, "),
        ("math", xml_mrun("мм\u00B2")),
        ("text", ";"),
    ]


def legend_where_theta_k_a_parts() -> list[MixedPart]:
    return [
        ("text", "a – величина, обратная температурному коэффициенту электрического сопротивления при 0 "),
        ("math", xml_mrun("\u00B0C")),
        ("text", ", равная 228 "),
        ("math", xml_mrun("\u00B0C")),
        ("text", ". "),
    ]
