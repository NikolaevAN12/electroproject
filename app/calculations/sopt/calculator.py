from __future__ import annotations

import math

from .models import SoptInput, SoptResult

_JUMPER_LENGTH_BETWEEN_BANKS = 0.25
_JUMPER_LENGTH_TO_NEIGHBOR_CABINET = 1.5


class SoptCalculator:
    def calculate(self, input_model: SoptInput) -> SoptResult:
        n_sections = len(input_model.sections)
        n_items = sum(len(subsection.items) for section in input_model.sections for subsection in section.subsections)
        name = input_model.project_name.strip() or "без названия"
        battery = input_model.battery_name.strip() or "АКБ"
        jumpers_count = max(0, input_model.battery_count - 1)
        total_jumper_length = (
            jumpers_count * _JUMPER_LENGTH_BETWEEN_BANKS
            + _JUMPER_LENGTH_TO_NEIGHBOR_CABINET * 2
        )
        r_ab = (input_model.r_el * input_model.battery_count) / 1000.0
        r_per = (
            input_model.rho * total_jumper_length / input_model.jumper_section
            if input_model.jumper_section > 0
            else 0.0
        )
        r_kz = r_ab + r_per + 2 * input_model.input_fuse_resistance
        n_required = (
            1.1 * input_model.q_calc_ah / input_model.qn1_ah
            if input_model.qn1_ah > 0
            else 0.0
        )
        n_accepted = max(1, math.ceil(n_required))
        n_total_elements = input_model.battery_cells_count * input_model.battery_count
        r_gr_mohm = (
            7.5 * n_total_elements / n_accepted
            if n_accepted > 0
            else 0.0
        )
        r_gr = r_gr_mohm / 1000.0
        selective_ok = r_kz < r_gr
        return SoptResult(
            message=(
                f"Проект «{name}», {battery}: элементов — {input_model.battery_cells_count}, "
                f"АКБ — {input_model.battery_count}, Rкз={r_kz:.3f} Ом, "
                f"Rгр={r_gr:.3f} Ом, N={n_accepted}, "
                f"{'условие выполняется' if selective_ok else 'условие не выполняется'}, "
                f"разделов — {n_sections}, позиций оборудования — {n_items}."
            ),
            r_ab=r_ab,
            r_per=r_per,
            r_kz=r_kz,
            r_gr=r_gr,
            n_required=n_required,
            n_accepted=n_accepted,
            selective_ok=selective_ok,
        )
