"""Формирование Word-отчёта «Выбор и проверка КЛ 6 кВ» (свыше 1 кВ)."""

from __future__ import annotations

from collections.abc import Sequence
from io import BytesIO

from docx import Document
from docx.enum.text import WD_BREAK

from app.calculations.fire_check.word_omml import normalize_mul_sign

from .calculator import MvCableFireCalculator
from .conductor_phrase import (
    conductor_genitive,
    conductor_instrumental_plural,
    conductor_kind_from_nominative,
    validate_cores_count,
)
from .doc_format import (
    add_body_formula,
    add_body_mixed_paragraph,
    add_body_paragraph,
    add_centered_formula,
    add_section_heading,
    add_source_data_item,
    add_source_data_mixed_item,
    add_subsection_heading,
    add_body_verdict_paragraph,
    open_template_document,
)
from .models import MvCableFireInput, MvCableFireResult


def _fmt(x: float, nd: int = 2) -> str:
    return f"{x:.{nd}f}".replace(".", ",")


_DEG_C = "°C"
_TAE_SOURCE_SUFFIX = " (согласно ГОСТ Р 52735-2007 таблица Е.1)"


def _deg(value: str, *, spaced: bool = True) -> str:
    return f"{value} {_DEG_C}" if spaced else f"{value}{_DEG_C}"


def _mul_text(text: str) -> str:
    """Знак умножения в формулах — как в отчёте СОПТ (·)."""
    return normalize_mul_sign(text)


def _mat(text: str) -> str:
    return _mul_text(text)


def _center_f(doc: Document, text: str) -> None:
    add_centered_formula(doc, _mul_text(text))


def _body_f(doc: Document, text: str) -> None:
    add_body_formula(doc, _mul_text(text))


def _verdict(
    doc: Document,
    *,
    passed: bool,
    u: float,
    designation: str,
    suffix: str,
) -> None:
    add_body_verdict_paragraph(
        doc,
        passed=passed,
        prefix=f"Вывод: силовой кабель {int(u)} кВ {designation} ",
        suffix=suffix,
    )


def _add_i_nom_source_item(
    doc: Document,
    model: MvCableFireInput,
    i_nom: float,
) -> None:
    if model.i_nom_override_a is not None:
        add_source_data_mixed_item(
            doc,
            [
                ("text", "номинальный ток "),
                ("math", _mat(f"Iн = {_fmt(i_nom, 1)} А")),
                ("text", "."),
            ],
        )
        return

    u = model.voltage_kv
    p = model.tsn_power_kva
    add_source_data_mixed_item(
        doc,
        [
            ("text", "номинальный ток "),
            (
                "math",
                _mat(f"Iн = P/(Uн·√3) = {int(p)}/{int(u)}·√3 = {_fmt(i_nom, 1)} А"),
            ),
            ("text", " (где "),
            ("math", _mat(f"Р={int(p)} кВА")),
            ("text", f" – мощность {model.load_name})."),
        ],
    )


def _section_title(model: MvCableFireInput) -> str:
    if not model.is_standard_load:
        return f"РАСЧЕТНАЯ ПРОВЕРКА {model.load_name.upper()} НА НЕВОЗГОРАНИЕ"
    u = model.voltage_kv
    return (
        f"РАСЧЕТНАЯ ПРОВЕРКА КАБЕЛЯ {int(u)} КВ ОТ {model.source_zru.upper()} "
        f"ДО {model.load_name.upper()} НА НЕВОЗГОРАНИЕ"
    )


def _add_page_break(doc: Document) -> None:
    paragraph = doc.add_paragraph()
    paragraph.add_run().add_break(WD_BREAK.PAGE)


def _append_mv_cable_section(
    doc: Document,
    model: MvCableFireInput,
    result: MvCableFireResult,
) -> None:
    validate_cores_count(model.cores_count)
    conductor_kind = conductor_kind_from_nominative(model.conductor_material)
    mat_instr = conductor_instrumental_plural(conductor_kind)
    mat_gen = conductor_genitive(conductor_kind)

    u = model.voltage_kv
    ikz = model.ikz3_ka
    i_nom = result.i_nom_a
    s = model.section_mm2
    i_dd = model.i_dd_a
    tae = model.tae_s

    add_section_heading(doc, _section_title(model))

    add_body_paragraph(doc, "Исходные данные, исходя из параметров устанавливаемого оборудования:")
    add_source_data_mixed_item(
        doc,
        [
            ("text", "номинальное напряжение "),
            ("math", _mat(f"Uн = {int(u)} кВ")),
            ("text", ";"),
        ],
    )
    _add_i_nom_source_item(doc, model, i_nom)
    add_source_data_mixed_item(
        doc,
        [
            (
                "text",
                f"максимальный ток трехфазного короткого замыкания на шинах {int(u)} кВ "
                f"на {model.substation_name}: ",
            ),
            ("math", _mat(f"Iкз(3) = {_fmt(ikz, 3)} кА")),
            ("text", "."),
        ],
    )
    add_source_data_item(doc, model.cable_route_text)

    if not model.is_existing_cable:
        add_subsection_heading(
            doc,
            f"ПРОВЕРКА КАБЕЛЯ {int(u)} КВ ПО УСЛОВИЮ МИНИМАЛЬНО ДОПУСТИМОГО СЕЧЕНИЯ",
        )
        add_body_mixed_paragraph(
            doc,
            [
                (
                    "text",
                    "Согласно пункту 2.1 циркуляра Ц-02-98(Э) [17] на вновь проектируемых и реконструируемых "
                    "объектах рекомендуется применять кабели сечением не менее ",
                ),
                ("math", _mat(f"Smin= {_fmt(model.s_min_mm2, 0)} мм2")),
                ("text", " ."),
            ],
        )
        _center_f(doc, result.section_check.left_text)
        _verdict(
            doc,
            passed=result.section_ok,
            u=u,
            designation=result.cable_designation,
            suffix="условию минимально допустимого сечения.",
        )

    add_subsection_heading(
        doc,
        f"ПРОВЕРКА КАБЕЛЯ {int(u)} КВ ПО УСЛОВИЮ ДЛИТЕЛЬНО ДОПУСТИМЫХ ТОКОВ",
    )
    add_body_mixed_paragraph(
        doc,
        [
            (
                "text",
                f"Согласно таблицы П1.2 циркуляра Ц-02-98(Э) длительный допустимый ток кабеля с "
                f"{mat_instr} жилами с изоляцией из {model.sheath_material_gen} сечением "
                f"{int(s)} мм2, при прокладке в воздухе составляет ",
            ),
            ("math", _mat(f"Iдд = {_fmt(i_dd, 0)} А")),
            ("text", ", что больше номинального тока "),
            ("math", _mat(f"Iн = {_fmt(i_nom, 1)} А")),
            ("text", "."),
        ],
    )
    cmp = ">" if result.i_dd_ok else "<"
    _center_f(doc, _mat(f"Iдд = {_fmt(i_dd, 0)} А {cmp} Iн = {_fmt(i_nom, 1)} А"))
    _verdict(
        doc,
        passed=result.i_dd_ok,
        u=u,
        designation=result.cable_designation,
        suffix="условию длительно допустимых токов.",
    )

    add_subsection_heading(
        doc,
        "ПРОВЕРКА КАБЕЛЯ ПО УСЛОВИЮ НАГРЕВА ПРИ ВОЗДЕЙСТВИИ ТОКА КОРОТКОГО ЗАМЫКАНИЯ",
    )
    add_body_paragraph(
        doc,
        f"1) Значение начальной температуры жилы кабеля {int(u)} кВ до короткого замыкания (КЗ), {_DEG_C}:",
    )
    _center_f(
        doc,
        f"Qн = Q0 + (Qдд − Qос)·(Iраб./Iдд)^2 = {_fmt(model.theta_0_c, 0)}+"
        f"({_fmt(model.theta_dd_c, 0)}−{_fmt(model.theta_okr_c, 0)})·"
        f"({_fmt(i_nom, 1)}/{_fmt(i_dd, 0)})^2 = {_deg(_fmt(result.theta_n_c, 2), spaced=False)}",
    )
    _body_f(
        doc,
        f"Q0 = {_deg(_fmt(model.theta_0_c, 0))} – фактическая температура окружающей среды во время КЗ,",
    )
    add_body_paragraph(doc, "принята средняя максимальная температура наиболее теплого месяца;")
    _body_f(
        doc,
        f"Qдд = {_deg(_fmt(model.theta_dd_c, 0))} – значение расчетной длительной допустимой температуры "
        f"жилы кабеля, согласно приложения 1 циркуляра Ц-02-98(Э) [17];",
    )
    _body_f(
        doc,
        f"Qос = {_deg(_fmt(model.theta_okr_c, 0))} – значение расчетной температуры окружающей среды (воздуха);",
    )
    _body_f(doc, f"Iраб = Iн = {_fmt(i_nom, 1)} А – значение тока перед КЗ.")

    add_body_paragraph(
        doc,
        "2) Интеграл Джоуля, тепловой импульс от тока КЗ, кА2с, (согласно ф.42 ГОСТ Р 52736-2007 [15]):",
    )
    t_h = result.t_off_heating_s
    _center_f(
        doc,
        "Втер = Вк = Вк.п. + Вк.а. = Iп.с.^2 · (tоткл. + Та.эк.(1 − e-((2·tотк.)/(Та.эк.)))) =",
    )
    _center_f(
        doc,
        f"= {_fmt(ikz, 3)}^2 · ({_fmt(t_h, 2)} + {_fmt(tae, 2)}·(1 − e-((2·{_fmt(t_h, 2)})/({_fmt(tae, 2)})))) "
        f"= {_fmt(result.b_heating_kas2, 2)} кА2с",
    )
    _body_f(
        doc,
        f"Iп.с. = Iкз(3) = {_fmt(ikz, 3)}кА – действующее значение периодической составляющей тока КЗ;",
    )
    _body_f(
        doc,
        f"tоткл.  = {_fmt(t_h, 2)} с – сумма времени действия основной защиты "
        f"({_fmt(model.t_main_protection_s, 1)} с) и времени срабатывания выключателя "
        f"({_fmt(model.t_breaker_s, 2)} с);",
    )
    _body_f(
        doc,
        f"Та.эк. = {_fmt(tae, 2)} с – эквивалентная постоянная времени затухания "
        f"апериодической составляющей тока КЗ{_TAE_SOURCE_SUFFIX};",
    )

    add_body_paragraph(doc, "3) Коэффициент, учитывающий связь теплового импульса и сечения жил кабеля:")
    _center_f(
        doc,
        f"К = (b·Втер.) / S^2 = ({_fmt(model.b_const, 2)}·{_fmt(result.b_heating_kas2, 2)}) / "
        f"{int(s)}^2 = {_fmt(result.k_heating, 2)}",
    )
    _body_f(
        doc,
        f"b={_fmt(model.b_const, 2)} мм2/кА2·с – постоянная, характеризующая теплофизические "
        f"характеристики материала жилы кабеля (согласно циркуляру) (для {mat_gen}).",
    )

    add_body_paragraph(doc, f"4) Температура жил кабеля {int(u)} кВ в конце КЗ, {_DEG_C}:")
    _center_f(
        doc,
        f"Qк = Qн·е^k + а·(е^k − 1) = {_fmt(result.theta_n_c, 2)}·е^{_fmt(result.k_heating, 2)} + "
        f"{_fmt(model.a_const, 0)}·(е^{_fmt(result.k_heating, 2)} − 1) = {_deg(_fmt(result.theta_k_heating_c, 2), spaced=False)}",
    )
    _body_f(
        doc,
        f"а={_deg(_fmt(model.a_const, 0))} – величина, обратная температурному коэффициенту "
        f"электрического сопротивления при {_deg('0')} (согласно циркуляру Ц-02-98(Э) [17]).",
    )
    add_body_paragraph(
        doc,
        f"Согласно таблице 6 ГОСТ Р 52736-2007 [15], температура нагрева проводников кабеля "
        f"с {mat_instr} жилами и изоляцией из {model.sheath_material_gen} при КЗ "
        f"должна быть не выше {_deg(_fmt(model.theta_limit_heating_c, 0))}.",
    )
    h_cmp = "<" if result.heating_ok else ">"
    _center_f(
        doc,
        f"Qк = {_deg(_fmt(result.theta_k_heating_c, 2))} {h_cmp} {_deg(_fmt(model.theta_limit_heating_c, 0))}",
    )
    _verdict(
        doc,
        passed=result.heating_ok,
        u=u,
        designation=result.cable_designation,
        suffix="условию нагрева при воздействии тока короткого замыкания.",
    )

    add_subsection_heading(
        doc,
        "ПРОВЕРКА КАБЕЛЯ НА НЕВОЗГОРАНИЕ ПРИ ДЕЙСТВИИ ТОКА КОРОТКОГО ЗАМЫКАНИЯ",
    )
    add_body_paragraph(
        doc,
        "1) Интеграл Джоуля, тепловой импульс от тока КЗ, кА2с, (согласно ф.42 ГОСТ Р 52736-2007 [15]):",
    )
    t_f = result.t_off_fire_s
    _center_f(
        doc,
        f"Втер = Вк = Вк.п. + Вк.а. = Iп.с.^2 · (tоткл. + Та.эк.(1 − e-((2·tотк.)/(Та.эк.))) = "
        f"{_fmt(ikz, 3)}^2 · ({_fmt(t_f, 2)} +",
    )
    _center_f(
        doc,
        f"+ {_fmt(tae, 2)}·(1 − e-((2·{_fmt(t_f, 2)})/({_fmt(tae, 2)})))) = {_fmt(result.b_fire_kas2, 0)} кА2с",
    )
    _body_f(
        doc,
        f"tоткл. = {_fmt(t_f, 2)} с – сумма времени действия резервной защиты "
        f"({_fmt(model.t_backup_protection_s, 1)} с) и времени срабатывания выключателя "
        f"({_fmt(model.t_breaker_s, 2)} с).",
    )
    _body_f(
        doc,
        f"Та.эк. = {_fmt(tae, 2)} с – эквивалентная постоянная времени затухания "
        f"апериодической составляющей тока КЗ{_TAE_SOURCE_SUFFIX}",
    )

    add_body_paragraph(doc, "2) Коэффициент, учитывающий связь теплового импульса и сечения жил кабеля:")
    _center_f(
        doc,
        f"К = (b·Втер.) / S^2 = ({_fmt(model.b_const, 2)}·{_fmt(result.b_fire_kas2, 0)}) / "
        f"{int(s)}^2 = {_fmt(result.k_fire, 2)}",
    )

    add_body_paragraph(doc, f"3) Температура жил кабеля {int(u)} кВ в конце КЗ, {_DEG_C}:")
    _center_f(
        doc,
        f"Qк = Qн·е^k + а·(е^k − 1) = {_fmt(result.theta_n_c, 2)}·е^{_fmt(result.k_fire, 2)} + "
        f"{_fmt(model.a_const, 0)}·(е^{_fmt(result.k_fire, 2)} − 1) = {_deg(_fmt(result.theta_k_fire_c, 2))}",
    )
    add_body_paragraph(
        doc,
        "Согласно пункту 1.1 циркуляра Ц-02-98(Э) [17] температура нагрева токопроводящих жил "
        f"кабелей с изоляцией из {model.sheath_material_gen} при проверке на невозгорание не должна превышать "
        f"{_deg(_fmt(model.theta_limit_fire_c, 0))}.",
    )
    f_cmp = "<" if result.fire_ok else ">"
    _center_f(
        doc,
        f"Qк = {_deg(_fmt(result.theta_k_fire_c, 2), spaced=False)} {f_cmp} {_deg(_fmt(model.theta_limit_fire_c, 0))}",
    )

    _verdict(
        doc,
        passed=result.fire_ok,
        u=u,
        designation=result.cable_designation,
        suffix="условиям выбора кабеля при воздействии тока короткого замыкания.",
    )


def build_mv_cable_word_report(
    model: MvCableFireInput,
    result: MvCableFireResult | None = None,
) -> Document:
    if result is None:
        result = MvCableFireCalculator().calculate(model)
    doc = open_template_document()
    _append_mv_cable_section(doc, model, result)
    return doc


def build_mv_cable_word_report_multi(
    cables: Sequence[tuple[MvCableFireInput, MvCableFireResult | None]],
) -> Document:
    if not cables:
        raise ValueError("Нет кабелей для формирования отчёта")
    calc = MvCableFireCalculator()
    doc = open_template_document()
    for index, (model, result) in enumerate(cables):
        if index > 0:
            _add_page_break(doc)
        if result is None:
            result = calc.calculate(model)
        _append_mv_cable_section(doc, model, result)
    return doc


def render_mv_cable_report_bytes(model: MvCableFireInput) -> bytes:
    doc = build_mv_cable_word_report(model)
    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()
