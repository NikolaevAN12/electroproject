from __future__ import annotations



import math



from .models import FireCheckInput, FireCheckResult, FireCheckRowResult, FireCheckTsnParams





def _fmt_ru(x: float, digits: int = 3) -> str:

    return f"{x:.{digits}f}".replace(".", ",")





class FireCheckCalculator:

    def calculate(self, input_model: FireCheckInput) -> FireCheckResult:

        tsn = input_model.tsn

        tae_s = tsn.tae_s if tsn else 0.0

        theta_o = tsn.theta_o_c if tsn else 25.0

        theta_dd = tsn.theta_dd_c if tsn else 70.0

        theta_okr = tsn.theta_okr_c if tsn else 25.0



        ik3_from_tsn: float | None = None

        if tsn is not None and tsn.r1_mohm > 0 and tsn.x1_mohm > 0:

            ik3_from_tsn = self.three_phase_sc_current_ka(tsn.r1_mohm, tsn.x1_mohm, tsn.u_line_v)



        rows: list[FireCheckRowResult] = []

        for row in input_model.rows:

            wire_temp = self._wire_temperature(

                row.allowed_current_a,

                row.working_current_a,

                theta_o,

                theta_dd,

                theta_okr,

            )

            bk = self._thermal_impulse_k2s(row.short_circuit_ka, row.shutdown_time_s, tae_s)

            bk_display = "" if bk is None else _fmt_ru(bk, 2)

            final_temp = self._fire_formula_result(wire_temp, bk, row.section_mm2)



            rows.append(

                FireCheckRowResult(

                    number=row.number,

                    cable=row.cable,

                    section_mm2_text="" if row.section_mm2 is None else str(row.section_mm2),

                    allowed_current_a_text="" if row.allowed_current_a is None else str(row.allowed_current_a),

                    working_current_a_text="" if row.working_current_a is None else str(row.working_current_a),

                    short_circuit_ka_text="" if row.short_circuit_ka is None else str(row.short_circuit_ka),

                    shutdown_time_s_text="" if row.shutdown_time_s is None else str(row.shutdown_time_s),

                    wire_temperature_c=wire_temp,

                    bk_value=bk,

                    bk_display=bk_display,

                    final_temperature_display="" if final_temp is None else final_temp,

                    sc_three_phase_ka=ik3_from_tsn,

                )

            )



        return FireCheckResult(rows=rows)



    @staticmethod

    def three_phase_sc_current_ka(r1_mohm: float, x1_mohm: float, u_line_v: float = 400.0) -> float | None:

        """I(3) = Uф / Z, Uф = Uном_лин / √3, R,X в Ом (из мОм)."""

        r_ohm = r1_mohm / 1000.0

        x_ohm = x1_mohm / 1000.0

        z = math.hypot(r_ohm, x_ohm)

        if z <= 0:

            return None

        u_ph = u_line_v / math.sqrt(3.0)

        return (u_ph / z) / 1000.0



    @staticmethod

    def _wire_temperature(

        i_allow: float | None,

        i_work: float | None,

        theta_o: float,

        theta_dd: float,

        theta_okr: float,

    ) -> int | None:

        if i_allow is None or i_work is None or i_allow == 0:

            return None

        ratio = i_work / i_allow

        value = theta_o + (theta_dd - theta_okr) * (ratio * ratio)

        return int(round(value))



    @staticmethod

    def _thermal_impulse_k2s(i_kz_ka: float | None, t_off: float | None, tae_s: float) -> float | None:

        if i_kz_ka is None or t_off is None:

            return None

        return (i_kz_ka * i_kz_ka) * t_off + tae_s



    @staticmethod

    def _fire_formula_result(

        t_wire: int | None,

        bk: float | None,

        section_mm2: float | None,

    ) -> str | None:

        if t_wire is None or bk is None or section_mm2 is None or section_mm2 == 0:

            return None

        kappa = 19.58 * bk / (section_mm2 * section_mm2)

        try:

            exp_k = math.exp(kappa)
            result = float(t_wire) * exp_k + 228.0 * (exp_k - 1.0)

        except OverflowError:

            return None

        return f"{result:.2f}".replace(".", ",")


