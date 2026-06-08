from __future__ import annotations

from app.core.contracts import ValidationError

from .models import ShsnInput


class ShsnValidator:
    def validate(self, input_model: ShsnInput) -> list[ValidationError]:
        errors: list[ValidationError] = []
        if not input_model.elements:
            errors.append(
                ValidationError("elements", "Загрузите модель сети из Excel или сохраните данные в файл.")
            )
            return errors

        names = [el.name.strip() for el in input_model.elements]
        if any(not n for n in names):
            errors.append(ValidationError("Name", "У каждого элемента должно быть поле Name."))

        duplicates = {n for n in names if names.count(n) > 1}
        if duplicates:
            errors.append(
                ValidationError("Name", f"Повторяющиеся наименования: {', '.join(sorted(duplicates))}")
            )

        sc_points = [el.name for el in input_model.elements if el.is_sc_point]
        if not sc_points:
            errors.append(
                ValidationError(
                    "Is_SC_Point",
                    "Отметьте хотя бы одну точку КЗ (Is_SC_Point = TRUE) в Excel.",
                )
            )

        if input_model.k_temp <= 0:
            errors.append(ValidationError("k_temp", "Коэффициент k_temp должен быть больше 0."))

        known = set(names)
        for el in input_model.elements:
            parent = el.parent_name.strip()
            if parent and parent.lower() not in ("нет", "none") and parent not in known:
                errors.append(
                    ValidationError(
                        "Parent_Name",
                        f"Элемент «{el.name}»: не найден родитель «{parent}».",
                    )
                )
        return errors
