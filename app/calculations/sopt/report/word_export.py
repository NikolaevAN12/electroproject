"""Формирование документа Word для расчёта СОПТ."""

from __future__ import annotations

import re
import xml.sax.saxutils as xml_esc
from typing import Any

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import parse_xml
from docx.shared import Cm, Pt

from ..models import EQUIPMENT_LABEL_BY_KEY, SoptEquipmentItem, SoptInput, SoptResult
from ..resistance import item_resistance_ohm
from ..schematic_draw import (
    akb_input_schematic_png_bytes,
    k1_substitution_schematic_png_bytes,
    subsection_kz_schematic_png_bytes,
)
from ..selectivity_map import SelectivityMapResult, build_subsection_selectivity_maps
from .numbering import (
    REPORT_SECTION_RESISTANCE,
    REPORT_SECTION_SELECTIVITY,
    REPORT_SECTION_TKZ,
    REPORT_SECTION_TKZ_BUS,
    selectivity_section,
    selectivity_subsection,
    tkz_user_section,
    tkz_user_subsection,
)

_FIG1_WIDTH_CM = 7.02
_FIG1_HEIGHT_CM = 3.87
_FIG2_AND_NEXT_WIDTH_CM = 9.6
_FIG2_AND_NEXT_HEIGHT_CM = 5.29

_JUMPER_LENGTH_BETWEEN_BANKS = 0.25
_JUMPER_LENGTH_TO_NEIGHBOR_CABINET = 1.5
_MNS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_MATH_W_RPR = (
    f'<w:rPr xmlns:w="{_WNS}">'
    f'<w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math" w:cs="Arial"/>'
    f'<w:sz w:val="24"/><w:szCs w:val="24"/>'
    f"</w:rPr>"
)

def _fmt_fixed(value: float, digits: int) -> str:
    return f"{value:.{digits}f}".replace(".", ",")


def _fmt_intish(value: float) -> str:
    rounded = round(value)
    if abs(value - rounded) < 1e-9:
        return str(int(rounded))
    return _fmt_fixed(value, 2).rstrip("0").rstrip(",")


def _xe(s: str) -> str:
    return xml_esc.escape(s)


def _xml_mrun(text: str) -> str:
    safe = _xe(text if text else "-")
    return f'<m:r>{_MATH_W_RPR}<m:rPr><m:nor/></m:rPr><m:t xml:space="preserve">{safe}</m:t></m:r>'


def _xml_mfrac(num_inner: str, den_inner: str) -> str:
    return (
        "<m:f>"
        f"<m:num><m:e>{num_inner}</m:e></m:num>"
        f"<m:den><m:e>{den_inner}</m:e></m:den>"
        "</m:f>"
    )


def _xml_msub(base: str, sub: str) -> str:
    return (
        "<m:sSub>"
        f"<m:e>{_xml_mrun(base)}</m:e>"
        f"<m:sub>{_xml_mrun(sub)}</m:sub>"
        "</m:sSub>"
    )


def _style_paragraph(
    p: Any,
    *,
    alignment: WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH.JUSTIFY,
    first_line_indent_cm: float = 1.25,
) -> None:
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = 1.15
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.first_line_indent = Cm(first_line_indent_cm)
    pf.left_indent = Cm(0)
    pf.right_indent = Cm(0)
    p.alignment = alignment


def _add_body_paragraph(doc: Any, text: str) -> Any:
    p = doc.add_paragraph(text)
    _style_paragraph(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent_cm=1.25)
    return p


def _add_figure_caption(doc: Any, text: str) -> Any:
    p = doc.add_paragraph(text)
    _style_paragraph(p, alignment=WD_ALIGN_PARAGRAPH.CENTER, first_line_indent_cm=0)
    return p


def _add_blank_line(doc: Any) -> None:
    p = doc.add_paragraph("")
    _style_paragraph(p, alignment=WD_ALIGN_PARAGRAPH.LEFT, first_line_indent_cm=0)


def _append_formula(doc: Any, inner_omml: str, *, first_line_indent_cm: float = 0) -> None:
    p = doc.add_paragraph()
    _style_paragraph(p, alignment=WD_ALIGN_PARAGRAPH.LEFT, first_line_indent_cm=first_line_indent_cm)
    omath = parse_xml(f'<m:oMath xmlns:m="{_MNS}">{inner_omml}</m:oMath>')
    p._p.append(omath)


def _kz_points_with_chain_r(items: list[SoptEquipmentItem]) -> list[tuple[str, float]]:
    points: list[tuple[str, float]] = []
    chain_r = 0.0
    for item in items:
        if item.item_type == "kz_point":
            points.append((item.designation.strip() or f"KЗ-{len(points) + 1}", chain_r))
        else:
            chain_r += item_resistance_ohm(item)
    if not points:
        points.append(("K1", chain_r))
    return points


def _collect_device_resistances(
    input_model: SoptInput, item_type: str
) -> tuple[list[tuple[int, float, int]], list[str]]:
    groups: dict[tuple[int, float], int] = {}
    unknown: list[str] = []
    for section in input_model.sections:
        for subsection in section.subsections:
            for item in subsection.items:
                if item.item_type != item_type:
                    continue
                if item.rated_current_a <= 0 or item.resistance_ohm <= 0:
                    unknown.append(item.designation.strip() or "без обозначения")
                    continue
                current = int(round(item.rated_current_a))
                key = (current, item.resistance_ohm)
                groups[key] = groups.get(key, 0) + 1
    entries = sorted(
        ((current, resistance, count) for (current, resistance), count in groups.items()),
        key=lambda row: (row[0], row[1]),
    )
    return entries, unknown


def _add_heading_paragraph(doc: Any, text: str, *, font_size: int = 14) -> Any:
    p = doc.add_paragraph()
    _style_paragraph(p, alignment=WD_ALIGN_PARAGRAPH.LEFT, first_line_indent_cm=0)
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(font_size)
    return p


def _selectivity_section_heading(section_name: str, section_idx: int) -> str:
    title = section_name.strip() or f"Раздел {section_idx}"
    return title.replace(
        "РАСЧЕТ ТОКОВ КОРОТКОГО ЗАМЫКАНИЯ",
        "ПРОВЕРКУ СЕЛЕКТИВНОСТИ",
    )


def _append_selectivity_section(
    doc: Any,
    input_model: SoptInput,
    result_model: SoptResult,
    *,
    n_total_elements: int,
    kc: float,
    fig_num: int,
) -> int:
    _add_heading_paragraph(doc, f"{REPORT_SECTION_SELECTIVITY}\tКАРТЫ СЕЛЕКТИВНОСТИ")

    processed_pairs: set[tuple[str, str]] = set()
    selectivity_error: str | None = None
    any_maps = False

    for section_idx, section in enumerate(input_model.sections, start=1):
        if not section.subsections:
            continue

        section_maps: list[SelectivityMapResult] = []
        for subsection in section.subsections:
            if not subsection.items:
                continue
            try:
                selectivity_maps = build_subsection_selectivity_maps(
                    subsection.items,
                    r_ab=result_model.r_ab,
                    r_per=result_model.r_per,
                    r_gr=result_model.r_gr,
                    n_total_elements=n_total_elements,
                    kc=kc,
                )
            except RuntimeError as exc:
                selectivity_error = str(exc)
                continue

            for map_result in selectivity_maps:
                pair_key = (map_result.upstream_label, map_result.downstream_label)
                if pair_key in processed_pairs:
                    continue
                processed_pairs.add(pair_key)
                section_maps.append(map_result)

        if not section_maps:
            continue

        any_maps = True
        section_heading = _selectivity_section_heading(section.name, section_idx)
        _add_heading_paragraph(doc, f"{selectivity_section(section_idx)}\t{section_heading}")

        for map_idx, map_result in enumerate(section_maps, start=1):
            _add_heading_paragraph(
                doc,
                f"{selectivity_subsection(section_idx, map_idx)}\t{map_result.block_title}",
                font_size=12,
            )
            pic_para = doc.add_paragraph()
            _style_paragraph(pic_para, alignment=WD_ALIGN_PARAGRAPH.CENTER, first_line_indent_cm=0)
            pic_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pic_para.add_run().add_picture(
                map_result.image,
                width=Cm(12.0),
            )
            _add_figure_caption(
                doc,
                f"Рис.{fig_num} — Карта селективности пары {map_result.upstream_label} "
                f"и {map_result.downstream_label} "
                f"(Iкз.макс={_fmt_fixed(map_result.i_max_a, 1)} А, "
                f"Iкз.мин={_fmt_fixed(map_result.i_min_a, 1)} А).",
            )
            fig_num += 1
            status = (
                "СЕЛЕКТИВНОСТЬ НЕ СОБЛЮДАЕТСЯ"
                if map_result.overlap
                else "Селективность обеспечена"
            )
            p_res = doc.add_paragraph(f"Вывод: {status}")
            _style_paragraph(p_res, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent_cm=1.25)
            if map_result.overlap and p_res.runs:
                p_res.runs[0].bold = True

    if not any_maps:
        if selectivity_error:
            _add_body_paragraph(doc, selectivity_error)
        else:
            _add_body_paragraph(doc, "Карты селективности не построены: нет пар автоматических выключателей.")
    else:
        _add_blank_line(doc)

    return fig_num


def populate_sopt_document(doc: Any, input_model: SoptInput, result_model: SoptResult) -> None:
    calc_title = doc.add_paragraph()
    _style_paragraph(calc_title, alignment=WD_ALIGN_PARAGRAPH.LEFT, first_line_indent_cm=0)
    calc_title_run = calc_title.add_run(
        f"{REPORT_SECTION_RESISTANCE}   РАСЧЕТ СОПРОТИВЛЕНИЙ ЭЛЕМЕНТОВ ЦЕПИ ПОСТОЯННОГО ТОКА"
    )
    calc_title_run.bold = True
    calc_title_run.font.size = Pt(14)
    _add_blank_line(doc)

    battery_name = input_model.battery_name.strip() or "АКБ"
    battery_count = input_model.battery_count
    battery_cells_count = input_model.battery_cells_count
    n_total_elements = input_model.battery_cells_count * input_model.battery_count
    r_el = input_model.r_el
    rho = input_model.rho
    jumper_section = input_model.jumper_section
    jumpers_count = max(0, battery_count - 1)
    total_jumper_length = (
        jumpers_count * _JUMPER_LENGTH_BETWEEN_BANKS
        + _JUMPER_LENGTH_TO_NEIGHBOR_CABINET * 2
    )
    r_ab_mohm = r_el * battery_count
    r_ab = r_ab_mohm / 1000.0
    r_per = rho * total_jumper_length / jumper_section if jumper_section > 0 else 0.0

    _add_body_paragraph(
        doc,
        f"1) Аккумуляторная батарея {battery_name} "
        f"(12 В, количество элементов n={battery_cells_count}), {battery_count} штук.",
    )
    _add_body_paragraph(doc, "Технические данные АБ:")
    _add_body_paragraph(
        doc,
        f"Rэл = {_fmt_fixed(r_el, 1)} мОм – внутреннее сопротивление блока 12 В;",
    )
    _add_body_paragraph(
        doc,
        f"Rаб = Rэл · n = {_fmt_fixed(r_el, 1)} · {battery_count} = {_fmt_fixed(r_ab_mohm, 1)} мОм = {_fmt_fixed(r_ab, 4)} Ом – внутреннее сопротивление АБ.",
    )
    _add_body_paragraph(doc, "Перемычки:")
    _add_body_paragraph(doc, f"Количество перемычек: {jumpers_count} штук;")
    _add_body_paragraph(doc, f"Сечение S = {_fmt_intish(jumper_section)} мм;")
    _add_body_paragraph(
        doc,
        "L = 0,25 м – длина между банками, в соседний шкаф – 1,5 м."
    )
    _add_body_paragraph(doc, "Сопротивление перемычек АБ:")
    _append_formula(
        doc,
        f"{_xml_mrun('RПЕР')}{_xml_mrun(' = ')}{_xml_mrun('p')}{_xml_mrun('·')}"
        f"{_xml_mfrac(_xml_mrun('L'), _xml_mrun('S'))}"
        f"{_xml_mrun(' = ')}{_xml_mrun(_fmt_fixed(rho, 4))}{_xml_mrun('·')}"
        f"{_xml_mfrac(_xml_mrun(_fmt_intish(total_jumper_length)), _xml_mrun(_fmt_intish(jumper_section)))}"
        f"{_xml_mrun(' = ')}{_xml_mrun(_fmt_fixed(r_per, 4))}{_xml_mrun(' Ом')}",
    )
    _add_body_paragraph(doc, f"Где: p = {_fmt_fixed(rho, 4)} – удельное сопротивление меди;")
    _add_body_paragraph(doc, f"S = {_fmt_intish(jumper_section)} мм – сечение перемычек;")
    _add_body_paragraph(
        doc,
        f"L = {jumpers_count} · 0,25 + 1,5 · 2 = {_fmt_intish(total_jumper_length)} м – общая длинна перемычек.",
    )
    _add_blank_line(doc)

    fuse_entries, unknown_fuses = _collect_device_resistances(input_model, "fuse")
    breaker_entries, unknown_breakers = _collect_device_resistances(input_model, "breaker")

    if fuse_entries or unknown_fuses:
        _add_body_paragraph(doc, "2) Сопротивления предохранителей (Табл.2.9 [1]):")
        for current, value, _count in fuse_entries:
            _add_body_paragraph(
                doc,
                f"- сопротивление предохранителя с номинальным током {current} А."
            )
            _append_formula(
                doc,
                f"{_xml_mrun('R')}{_xml_mrun(' = ')}{_xml_mrun(_fmt_fixed(value, 3))}{_xml_mrun(' Ом')}",
            )
        if unknown_fuses:
            _add_body_paragraph(
                doc,
                "- Не указаны номинальный ток и/или R для: "
                + "; ".join(unknown_fuses)
                + "."
            )

    if breaker_entries or unknown_breakers:
        _add_body_paragraph(doc, "3) Сопротивления автоматических выключателей (Табл.2.7 [1]):")
        for current, value, _count in breaker_entries:
            _add_body_paragraph(
                doc,
                f"- сопротивление автоматического выключателя с номинальным током {current} А."
            )
            _append_formula(
                doc,
                f"{_xml_mrun('R')}{_xml_mrun(' = ')}{_xml_mrun(_fmt_fixed(value, 3))}{_xml_mrun(' Ом')}",
            )
        if unknown_breakers:
            _add_body_paragraph(
                doc,
                "- Не указаны номинальный ток и/или R для: "
                + "; ".join(unknown_breakers)
                + "."
            )

    _add_blank_line(doc)

    input_fuse_label = input_model.input_fuse_label.strip() or "Вводной предохранитель"
    input_fuse_r = input_model.input_fuse_resistance
    qn1_ah = input_model.qn1_ah
    q_calc_ah = input_model.q_calc_ah
    n_required = result_model.n_required
    n_accepted = result_model.n_accepted
    r_kz = result_model.r_kz
    r_gr = result_model.r_gr
    e_calc = 1.7 if result_model.selective_ok else 1.93
    r_gr_mohm = r_gr * 1000.0

    _add_heading_paragraph(
        doc,
        f"{REPORT_SECTION_TKZ}\tРАСЧЕТ ТОКОВ КОРОТКОГО ЗАМЫКАНИЯ, ПРОВЕРКА ЧУВСТВИТЕЛЬНОСТИ.",
    )
    _add_heading_paragraph(
        doc,
        f"{REPORT_SECTION_TKZ_BUS}\tРАСЧЕТ ТОКОВ КОРОТКОГО ЗАМЫКАНИЯ НА ШИНАХ ЩИТА ПОСТОЯННОГО ТОКА",
    )
    pic_para = doc.add_paragraph()
    _style_paragraph(pic_para, alignment=WD_ALIGN_PARAGRAPH.CENTER, first_line_indent_cm=0)
    pic_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic_para.add_run().add_picture(
        akb_input_schematic_png_bytes(),
        width=Cm(_FIG1_WIDTH_CM),
        height=Cm(_FIG1_HEIGHT_CM),
    )
    _add_figure_caption(doc, "Рис.1 – Схема питания от АКБ.")
    pic2_para = doc.add_paragraph()
    _style_paragraph(pic2_para, alignment=WD_ALIGN_PARAGRAPH.CENTER, first_line_indent_cm=0)
    pic2_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic2_para.add_run().add_picture(
        k1_substitution_schematic_png_bytes(result_model.r_ab, result_model.r_per, input_fuse_r),
        width=Cm(_FIG2_AND_NEXT_WIDTH_CM),
        height=Cm(_FIG2_AND_NEXT_HEIGHT_CM),
    )
    _add_figure_caption(doc, "Рис.2 – Схема замещения до точки K1.")
    _add_blank_line(doc)
    _add_body_paragraph(doc, "Сопротивление цепи тока короткого замыкания в т. K1:")
    _append_formula(
        doc,
        f"{_xml_msub('R', 'кз')}{_xml_mrun(' = (')}{_xml_msub('R', 'аб')}{_xml_mrun(' + ')}{_xml_msub('R', 'пер')}{_xml_mrun(') + 2·')}{_xml_msub('R', 'пр')}{_xml_mrun(' = ')}"
        f"{_xml_mrun('(')}{_xml_mrun(_fmt_fixed(result_model.r_ab, 4))}{_xml_mrun(' + ')}{_xml_mrun(_fmt_fixed(result_model.r_per, 4))}{_xml_mrun(') + 2·')}"
        f"{_xml_mrun(_fmt_fixed(input_fuse_r, 3))}{_xml_mrun(' = ')}{_xml_mrun(_fmt_fixed(r_kz, 3))}"
        f"{_xml_mrun(' Ом')}",
    )
    p_e_calc = doc.add_paragraph()
    _style_paragraph(p_e_calc, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent_cm=1.25)
    p_e_calc.add_run("E")
    run_e_sub = p_e_calc.add_run("расч")
    run_e_sub.font.subscript = True
    p_e_calc.add_run(f" = {_fmt_fixed(e_calc, 2)} В т.к. ")
    p_e_calc.add_run("R")
    run_rkz_sub = p_e_calc.add_run("кз")
    run_rkz_sub.font.subscript = True
    p_e_calc.add_run(f" {'<' if result_model.selective_ok else '≥'} ")
    p_e_calc.add_run("R")
    run_rgr_sub = p_e_calc.add_run("гр")
    run_rgr_sub.font.subscript = True
    p_e_calc.add_run(".")
    _append_formula(
        doc,
        f"{_xml_msub('R', 'гр')}{_xml_mrun(' = 7,5·')}{_xml_mfrac(_xml_mrun('n'), _xml_mrun('N'))}{_xml_mrun(' мОм')}",
    )
    _add_body_paragraph(
        doc,
        f"где n - количество элементов в цепи n= {battery_cells_count}·{battery_count} = {n_total_elements} эл.",
    )
    _add_body_paragraph(
        doc,
        "N - номер аккумуляторной батареи, определяется с учетом понижения энергии батареи за период эксплуатации",
    )
    _append_formula(
        doc,
        f"{_xml_msub('N', 'усл')}{_xml_mrun(' ≥ 1,1·')}{_xml_mfrac(_xml_msub('Q', 'расч'), _xml_msub('Q', 'N=1'))}",
    )
    _add_body_paragraph(doc, "где 1,1 - коэффициент учитывающий понижение энергии батареи")
    p_qn1_desc = doc.add_paragraph()
    _style_paragraph(p_qn1_desc, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent_cm=1.25)
    p_qn1_desc.add_run("Q")
    run_qn1_desc_sub = p_qn1_desc.add_run("N=1")
    run_qn1_desc_sub.font.subscript = True
    p_qn1_desc.add_run(f" - энергия аккумулятора {battery_name} (при одночасовом разряде)")
    p_qn1_val = doc.add_paragraph()
    _style_paragraph(p_qn1_val, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent_cm=1.25)
    p_qn1_val.add_run("Q")
    run_qn1_val_sub = p_qn1_val.add_run("N=1")
    run_qn1_val_sub.font.subscript = True
    p_qn1_val.add_run(f" = {_fmt_fixed(qn1_ah, 3)} Ач")
    _add_body_paragraph(doc, f"Qрасч - энергия аккумулятора {battery_name} (при одночасовом разряде)")
    _add_body_paragraph(doc, f"Qрасч = {_fmt_fixed(q_calc_ah, 0)} Ач")
    _append_formula(
        doc,
        f"{_xml_msub('N', 'усл')}{_xml_mrun(' ≥ 1,1·')}"
        f"{_xml_mfrac(_xml_mrun(_fmt_fixed(q_calc_ah, 0)), _xml_mrun(_fmt_fixed(qn1_ah, 3)))}"
        f"{_xml_mrun(' = ')}{_xml_mrun(_fmt_fixed(n_required, 2))}",
    )
    _add_body_paragraph(doc, f"Принимаем N={n_accepted}")
    _append_formula(
        doc,
        f"{_xml_msub('R', 'гр')}{_xml_mrun(' = 7,5·')}{_xml_mfrac(_xml_mrun(str(n_total_elements)), _xml_mrun(str(n_accepted)))}"
        f"{_xml_mrun(' = ')}{_xml_mrun(_fmt_fixed(r_gr_mohm, 1))}{_xml_mrun(' мОм = ')}{_xml_mrun(_fmt_fixed(r_gr, 3))}{_xml_mrun(' Ом')}",
    )
    _add_body_paragraph(doc, f"{_fmt_fixed(r_kz, 3)} Ом {'<' if result_model.selective_ok else '≥'} {_fmt_fixed(r_gr, 3)} Ом")
    _add_blank_line(doc)

    kc = 0.6
    fuse_trip_time_s = 0.001
    i_kz = (e_calc * n_total_elements / r_kz) if r_kz > 0 else 0.0
    i_kzd = kc * i_kz
    nominal_match = re.search(r"(\d+)", input_fuse_label)
    fuse_nominal_a = nominal_match.group(1) if nominal_match else ""
    fuse_nominal_text = f"{fuse_nominal_a}А" if fuse_nominal_a else (input_fuse_label.split("(")[0].strip() or input_fuse_label)

    _add_body_paragraph(doc, "Ток короткого замыкания в т. K1:")
    _append_formula(
        doc,
        f"{_xml_msub('I', 'кз')}{_xml_mrun(' = ')}"
        f"{_xml_mfrac(_xml_mrun('Eрасч·n'), _xml_msub('R', 'кз'))}"
        f"{_xml_mrun(' = ')}{_xml_mfrac(_xml_mrun(_fmt_fixed(e_calc, 2) + '·' + str(n_total_elements)), _xml_mrun(_fmt_fixed(r_kz, 3)))}"
        f"{_xml_mrun(' = ')}{_xml_mrun(_fmt_fixed(i_kz, 1))}{_xml_mrun(' А')}",
    )
    _add_body_paragraph(doc, "Ток короткого замыкания в т. K1 с учётом сопротивления дуги:")
    _append_formula(
        doc,
        f"{_xml_msub('I', 'кзд')}{_xml_mrun(' = ')}{_xml_msub('K', 'c')}{_xml_mrun('·')}{_xml_msub('I', 'кз')}{_xml_mrun(' = ')}{_xml_mrun(_fmt_fixed(kc, 1))}"
        f"{_xml_mrun('·')}{_xml_mrun(_fmt_fixed(i_kz, 1))}{_xml_mrun(' = ')}{_xml_mrun(_fmt_fixed(i_kzd, 1))}{_xml_mrun(' А')}",
    )
    p_kc_desc = doc.add_paragraph()
    _style_paragraph(p_kc_desc, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent_cm=1.25)
    p_kc_desc.add_run("K")
    run_kc_sub = p_kc_desc.add_run("c")
    run_kc_sub.font.subscript = True
    p_kc_desc.add_run(" - коэффициент снижения тока КЗ, находится по кривой зависимости")
    _append_formula(
        doc,
        f"{_xml_msub('K', 'c')}{_xml_mrun(' = f(')}{_xml_msub('R', 'кз')}{_xml_mrun(')')}",
    )
    _append_formula(
        doc,
        f"{_xml_msub('K', 'c')}{_xml_mrun(' = ')}{_xml_mrun(_fmt_fixed(kc, 1))}{_xml_mrun(' для ')}{_xml_msub('R', 'кз')}{_xml_mrun(' = ')}{_xml_mrun(_fmt_fixed(r_kz, 3))}{_xml_mrun(' Ом')}",
    )
    _add_body_paragraph(
        doc,
        f"Время перегорания {fuse_nominal_text} предохранителя не более {_fmt_fixed(fuse_trip_time_s, 3)} с согласно время-токовой характеристике gG -{fuse_nominal_text}",
    )
    _add_blank_line(doc)

    if not input_model.sections:
        _add_body_paragraph(doc, "Разделы не заданы.")
        return

    fig_num = 3
    for section_idx, section in enumerate(input_model.sections, start=1):
        if not section.subsections:
            continue
        section_name = section.name.strip() or f"Раздел {section_idx}"
        _add_heading_paragraph(doc, f"{tkz_user_section(section_idx)}\t{section_name}")

        for subsection_idx, subsection in enumerate(section.subsections, start=1):
            subsection_name = subsection.name.strip() or f"Подраздел {subsection_idx}"
            _add_heading_paragraph(
                doc,
                f"{tkz_user_subsection(section_idx, subsection_idx)}\t{subsection_name}",
                font_size=12,
            )

            if not subsection.items:
                _add_body_paragraph(doc, "(оборудование не добавлено)")
                _add_body_paragraph(doc, "")
                continue

            stream = subsection_kz_schematic_png_bytes(subsection.items, result_model.r_ab, result_model.r_per)
            chain_items = [it for it in subsection.items if it.item_type != "kz_point"]
            kz_count = max(1, sum(1 for it in subsection.items if it.item_type == "kz_point"))
            pic_width_cm = min(18.0, max(9.6, 1.35 * max(1, len(chain_items)) + 1.1 * kz_count + 4.4))
            pic_para = doc.add_paragraph()
            _style_paragraph(pic_para, alignment=WD_ALIGN_PARAGRAPH.CENTER, first_line_indent_cm=0)
            pic_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pic_para.add_run().add_picture(
                stream,
                width=Cm(pic_width_cm),
            )
            point_title = _kz_points_with_chain_r(subsection.items)[-1][0]
            _add_figure_caption(doc, f"Рис.{fig_num} – Схема замещения до точки {point_title}.")
            fig_num += 1

            cable_items = [it for it in subsection.items if it.item_type == "cable"]
            for cable in cable_items:
                if cable.cable_length_m <= 0 or cable.cable_gamma <= 0 or cable.cable_section_mm2 <= 0:
                    continue
                cable_name = cable.designation.strip() or "ВВГЭ"
                _add_body_paragraph(
                    doc,
                    f"Сопротивление кабеля {cable_name} 2х{_fmt_intish(cable.cable_section_mm2)} "
                    f"L={_fmt_intish(cable.cable_length_m)} м",
                )
                rk = item_resistance_ohm(cable)
                _append_formula(
                    doc,
                    f"{_xml_msub('R', 'к')}{_xml_mrun(' = ')}"
                    f"{_xml_mfrac(_xml_mrun('L'), _xml_mrun('Y·S'))}"
                    f"{_xml_mrun(' = ')}"
                    f"{_xml_mfrac(_xml_mrun(_fmt_intish(cable.cable_length_m)), _xml_mrun(_fmt_intish(cable.cable_gamma) + '·' + _fmt_intish(cable.cable_section_mm2)))}"
                    f"{_xml_mrun(' = ')}{_xml_mrun(_fmt_fixed(rk, 3))}{_xml_mrun(' Ом')}",
                )

            point_chains: list[tuple[str, list[float]]] = []
            running_components: list[float] = []
            for item in subsection.items:
                if item.item_type == "kz_point":
                    point_name = item.designation.strip() or f"KЗ-{len(point_chains) + 1}"
                    point_chains.append((point_name, list(running_components)))
                    continue
                comp_r = item_resistance_ohm(item)
                if comp_r > 0:
                    running_components.append(comp_r)
            if not point_chains:
                point_chains.append(("K1", list(running_components)))

            point_rkz_values: list[tuple[str, float]] = []
            for point_name, chain_components in point_chains:
                chain_r = sum(chain_components)
                rkz_point = result_model.r_ab + result_model.r_per + 2 * chain_r
                e_calc_point = 1.7 if rkz_point < result_model.r_gr else 1.93
                ikz_point = (e_calc_point * n_total_elements / rkz_point) if rkz_point > 0 else 0.0
                ikzd_point = kc * ikz_point
                point_rkz_values.append((point_name, rkz_point))
                _add_body_paragraph(doc, f"Сопротивление цепи короткого замыкания в т. {point_name}:")
                inner_sum = " + ".join(_fmt_fixed(value, 3) for value in chain_components) if chain_components else "0"
                _append_formula(
                    doc,
                    f"{_xml_msub('R', 'кз')}{_xml_mrun(' = ')}{_xml_mrun(_fmt_fixed(result_model.r_ab, 4))}"
                    f"{_xml_mrun(' + ')}{_xml_mrun(_fmt_fixed(result_model.r_per, 4))}{_xml_mrun(' + 2·(')}"
                    f"{_xml_mrun(inner_sum)}{_xml_mrun(') = ')}{_xml_mrun(_fmt_fixed(rkz_point, 3))}{_xml_mrun(' Ом')}",
                )
                p_e_point = doc.add_paragraph()
                _style_paragraph(p_e_point, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent_cm=1.25)
                p_e_point.add_run("E")
                run_e_sub = p_e_point.add_run("расч")
                run_e_sub.font.subscript = True
                p_e_point.add_run(f" = {_fmt_fixed(e_calc_point, 2)} В т.к. ")
                p_e_point.add_run("R")
                run_rkz_sub = p_e_point.add_run("кз")
                run_rkz_sub.font.subscript = True
                p_e_point.add_run(f" {'<' if rkz_point < result_model.r_gr else '≥'} ")
                p_e_point.add_run("R")
                run_rgr_sub = p_e_point.add_run("гр")
                run_rgr_sub.font.subscript = True
                p_e_point.add_run(".")
                _add_body_paragraph(doc, f"Ток короткого замыкания в т. {point_name}:")
                _append_formula(
                    doc,
                    f"{_xml_msub('I', 'кз')}{_xml_mrun(' = ')}{_xml_mfrac(_xml_mrun(_fmt_fixed(e_calc_point, 2) + '·' + str(n_total_elements)), _xml_mrun(_fmt_fixed(rkz_point, 3)))}"
                    f"{_xml_mrun(' = ')}{_xml_mrun(_fmt_fixed(ikz_point, 1))}{_xml_mrun(' А')}",
                )
                _add_body_paragraph(doc, f"Ток короткого замыкания в т. {point_name} с учётом сопротивления дуги:")
                _append_formula(
                    doc,
                    f"{_xml_msub('I', 'кзд')}{_xml_mrun(' = ')}{_xml_msub('K', 'c')}{_xml_mrun('·')}{_xml_msub('I', 'кз')}{_xml_mrun(' = ')}"
                    f"{_xml_mrun(_fmt_fixed(kc, 1))}{_xml_mrun('·')}{_xml_mrun(_fmt_fixed(ikz_point, 1))}{_xml_mrun(' = ')}{_xml_mrun(_fmt_fixed(ikzd_point, 1))}{_xml_mrun(' А')}",
                )

            p_kc_desc = doc.add_paragraph()
            _style_paragraph(p_kc_desc, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent_cm=1.25)
            p_kc_desc.add_run("K")
            run_kc_sub = p_kc_desc.add_run("c")
            run_kc_sub.font.subscript = True
            p_kc_desc.add_run(" - коэффициент снижения тока КЗ, находится по кривой зависимости")
            _append_formula(
                doc,
                f"{_xml_msub('K', 'c')}{_xml_mrun(' = f(')}{_xml_msub('R', 'кз')}{_xml_mrun(')')}",
            )
            for _, rkz_point in point_rkz_values:
                _append_formula(
                    doc,
                    f"{_xml_msub('K', 'c')}{_xml_mrun(' = 0,6 для ')}{_xml_msub('R', 'кз')}{_xml_mrun(' = ')}{_xml_mrun(_fmt_fixed(rkz_point, 3))}{_xml_mrun(' Ом')}",
                )

            _add_body_paragraph(doc, "")

    fig_num = _append_selectivity_section(
        doc,
        input_model,
        result_model,
        n_total_elements=n_total_elements,
        kc=kc,
        fig_num=fig_num,
    )
