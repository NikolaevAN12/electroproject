from __future__ import annotations



from app.core.contracts import ValidationError



from .models import FireCheckInput





class FireCheckValidator:

    def validate(self, input_model: FireCheckInput) -> list[ValidationError]:

        errors: list[ValidationError] = []

        for row in input_model.rows:

            prefix = f"Строка {row.number}"

            if row.section_mm2 is not None and row.section_mm2 <= 0:

                errors.append(ValidationError(field=f"{prefix}:section", message=f"{prefix}: сечение должно быть > 0"))

            if row.allowed_current_a is not None and row.allowed_current_a <= 0:

                errors.append(

                    ValidationError(field=f"{prefix}:allow", message=f"{prefix}: допустимый ток должен быть > 0")

                )

            if row.shutdown_time_s is not None and row.shutdown_time_s < 0:

                errors.append(

                    ValidationError(field=f"{prefix}:shutdown", message=f"{prefix}: время отключения не может быть < 0")

                )

            if row.short_circuit_ka is not None and row.short_circuit_ka < 0:

                errors.append(ValidationError(field=f"{prefix}:ikz", message=f"{prefix}: ток КЗ не может быть < 0"))



        if input_model.tsn is not None:

            t = input_model.tsn

            if t.r1_mohm <= 0 or t.x1_mohm <= 0:

                errors.append(

                    ValidationError(

                        field="line1:impedance",

                        message="Для 1-й линии: R1 и X1 должны быть положительными (мОм).",

                    )

                )

            if t.u_line_v <= 0:

                errors.append(ValidationError(field="tsn:u", message="Uном должно быть > 0."))

            if t.tae_s < 0:

                errors.append(ValidationError(field="tsn:tae", message="Tаэ не может быть отрицательным."))



        return errors


