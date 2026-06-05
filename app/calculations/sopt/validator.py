from __future__ import annotations

from app.core.contracts import ValidationError

from .models import SoptInput


class SoptValidator:
    def validate(self, input_model: SoptInput) -> list[ValidationError]:
        errors: list[ValidationError] = []
        if not input_model.battery_name.strip():
            errors.append(ValidationError(field="battery_name", message="Укажите название АКБ."))
        if input_model.battery_cells_count <= 0:
            errors.append(
                ValidationError(
                    field="battery_cells_count",
                    message="АКБ: количество элементов должно быть больше 0.",
                )
            )
        if input_model.battery_count <= 0:
            errors.append(
                ValidationError(
                    field="battery_count",
                    message="АКБ: количество АКБ должно быть больше 0.",
                )
            )
        if input_model.r_el <= 0:
            errors.append(ValidationError(field="r_el", message="АКБ: Rэл должно быть больше 0."))
        if input_model.rho <= 0:
            errors.append(
                ValidationError(
                    field="rho",
                    message="АКБ: ρ (удельное сопротивление) должно быть больше 0.",
                )
            )
        if input_model.jumper_section <= 0:
            errors.append(
                ValidationError(
                    field="jumper_section",
                    message="АКБ: S (сечение перемычки) должно быть больше 0.",
                )
            )
        if input_model.input_fuse_resistance <= 0:
            errors.append(
                ValidationError(
                    field="input_fuse_resistance",
                    message="Введите сопротивление вводного предохранителя (Rпр) больше 0.",
                )
            )
        if input_model.qn1_ah <= 0:
            errors.append(
                ValidationError(
                    field="qn1_ah",
                    message="Введите QN=1 (А·ч) больше 0.",
                )
            )
        if input_model.q_calc_ah <= 0:
            errors.append(
                ValidationError(
                    field="q_calc_ah",
                    message="Введите Qрасч (А·ч) больше 0.",
                )
            )
        if not input_model.sections:
            errors.append(ValidationError(field="sections", message="Добавьте хотя бы один раздел."))
        for idx, section in enumerate(input_model.sections, start=1):
            if not section.name.strip():
                errors.append(
                    ValidationError(
                        field=f"section:{idx}:name",
                        message=f"Раздел {idx}: укажите название.",
                    )
                )
            if not section.subsections:
                errors.append(
                    ValidationError(
                        field=f"section:{idx}:subsections",
                        message=f"Раздел {idx}: добавьте хотя бы один подраздел.",
                    )
                )
                continue
            for sub_idx, subsection in enumerate(section.subsections, start=1):
                if not subsection.name.strip():
                    errors.append(
                        ValidationError(
                            field=f"section:{idx}:subsection:{sub_idx}:name",
                            message=f"Раздел {idx}, подраздел {sub_idx}: укажите название.",
                        )
                    )
                for item_idx, item in enumerate(subsection.items, start=1):
                    item_label = f"Раздел {idx}, подраздел {sub_idx}, позиция {item_idx}"
                    if item.item_type in ("fuse", "breaker"):
                        if item.rated_current_a <= 0:
                            errors.append(
                                ValidationError(
                                    field=f"section:{idx}:subsection:{sub_idx}:item:{item_idx}:rated_current_a",
                                    message=f"{item_label}: укажите Iном, А > 0.",
                                )
                            )
                        if item.resistance_ohm <= 0:
                            errors.append(
                                ValidationError(
                                    field=f"section:{idx}:subsection:{sub_idx}:item:{item_idx}:resistance_ohm",
                                    message=f"{item_label}: укажите R, Ом > 0.",
                                )
                            )
                    if item.item_type == "cable":
                        if item.cable_length_m <= 0:
                            errors.append(
                                ValidationError(
                                    field=f"section:{idx}:subsection:{sub_idx}:item:{item_idx}:cable_length_m",
                                    message=f"{item_label}: для кабеля укажите L, м > 0.",
                                )
                            )
                        if item.cable_gamma <= 0:
                            errors.append(
                                ValidationError(
                                    field=f"section:{idx}:subsection:{sub_idx}:item:{item_idx}:cable_gamma",
                                    message=f"{item_label}: для кабеля укажите Y > 0.",
                                )
                            )
                        if item.cable_section_mm2 <= 0:
                            errors.append(
                                ValidationError(
                                    field=f"section:{idx}:subsection:{sub_idx}:item:{item_idx}:cable_section_mm2",
                                    message=f"{item_label}: для кабеля укажите S, мм² > 0.",
                                )
                            )
        return errors
