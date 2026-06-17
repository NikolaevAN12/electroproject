from __future__ import annotations

import math

from .models import MvCableCheckVerdict, MvCableFireInput, MvCableFireResult


def _fmt_ru(x: float, digits: int = 2) -> str:
    return f"{x:.{digits}f}".replace(".", ",")


class MvCableFireCalculator:
    """Расчёт проверки КЛ 6–10 кВ по циркуляру Ц-02-98(Э) и ГОСТ Р 52736-2007 (ф.42)."""

    def calculate(self, model: MvCableFireInput) -> MvCableFireResult:
        u_kv = model.voltage_kv
        if model.i_nom_override_a is not None:
            i_nom = model.i_nom_override_a
        elif model.tsn_power_kva is not None:
            i_nom = model.tsn_power_kva / (u_kv * math.sqrt(3.0))
        else:
            raise ValueError("Не заданы P, кВА или Iн, А")

        cable_designation = (
            f"{model.cable_mark} {model.cores_count}х{_fmt_ru(model.section_mm2, 0).split(',')[0]}"
        )

        theta_n = self._initial_temperature(
            model.theta_0_c,
            model.theta_dd_c,
            model.theta_okr_c,
            i_nom,
            model.i_dd_a,
        )

        t_off_heating = model.t_main_protection_s + model.t_breaker_s
        b_heating = self._thermal_impulse_kas2(model.ikz3_ka, t_off_heating, model.tae_s)
        k_heating = self._kappa(model.b_const, b_heating, model.section_mm2)
        theta_k_heating = self._final_temperature(theta_n, k_heating, model.a_const)

        t_off_fire = model.t_backup_protection_s + model.t_breaker_s
        b_fire = self._thermal_impulse_kas2(model.ikz3_ka, t_off_fire, model.tae_s)
        k_fire = self._kappa(model.b_const, b_fire, model.section_mm2)
        theta_k_fire = self._final_temperature(theta_n, k_fire, model.a_const)

        section_ok = model.section_mm2 >= model.s_min_mm2
        i_dd_ok = model.i_dd_a > i_nom
        heating_ok = theta_k_heating < model.theta_limit_heating_c
        fire_ok = theta_k_fire < model.theta_limit_fire_c

        return MvCableFireResult(
            i_nom_a=i_nom,
            cable_designation=cable_designation,
            section_ok=section_ok,
            i_dd_ok=i_dd_ok,
            theta_n_c=theta_n,
            b_heating_kas2=b_heating,
            k_heating=k_heating,
            theta_k_heating_c=theta_k_heating,
            heating_ok=heating_ok,
            t_off_fire_s=t_off_fire,
            b_fire_kas2=b_fire,
            k_fire=k_fire,
            theta_k_fire_c=theta_k_fire,
            fire_ok=fire_ok,
            t_off_heating_s=t_off_heating,
            section_check=MvCableCheckVerdict(
                passed=section_ok,
                left_text=f"S={_fmt_ru(model.section_mm2, 0)} мм2",
                right_text="",
                limit_text=f"Smin= {_fmt_ru(model.s_min_mm2, 0)} мм2",
            ),
            current_check=MvCableCheckVerdict(
                passed=i_dd_ok,
                left_text=f"Iдд = {_fmt_ru(model.i_dd_a, 0)} А",
                right_text=f"Iн = {_fmt_ru(i_nom, 1)} А",
                limit_text="",
            ),
            heating_check=MvCableCheckVerdict(
                passed=heating_ok,
                left_text=f"Qк = {_fmt_ru(theta_k_heating, 2)} С",
                right_text=f"{_fmt_ru(model.theta_limit_heating_c, 0)} С",
                limit_text="",
            ),
            fire_check=MvCableCheckVerdict(
                passed=fire_ok,
                left_text=f"Qк = {_fmt_ru(theta_k_fire, 2)}С",
                right_text=f"{_fmt_ru(model.theta_limit_fire_c, 0)} С",
                limit_text="",
            ),
        )

    @staticmethod
    def _initial_temperature(
        theta_0: float,
        theta_dd: float,
        theta_okr: float,
        i_work_a: float,
        i_dd_a: float,
    ) -> float:
        if i_dd_a <= 0:
            return theta_0
        ratio = i_work_a / i_dd_a
        return theta_0 + (theta_dd - theta_okr) * (ratio * ratio)

    @staticmethod
    def _thermal_impulse_kas2(i_ps_ka: float, t_off_s: float, tae_s: float) -> float:
        """Bк = Iп.с.² · (tоткл + Та.эк·(1 − e^(−2·tоткл/Та.эк))) — ГОСТ Р 52736-2007, ф.42."""
        if tae_s <= 0:
            return i_ps_ka * i_ps_ka * t_off_s
        decay = math.exp(-2.0 * t_off_s / tae_s)
        return (i_ps_ka * i_ps_ka) * (t_off_s + tae_s * (1.0 - decay))

    @staticmethod
    def _kappa(b: float, b_thermal: float, section_mm2: float) -> float:
        if section_mm2 <= 0:
            return 0.0
        return (b * b_thermal) / (section_mm2 * section_mm2)

    @staticmethod
    def _final_temperature(theta_n: float, kappa: float, a: float) -> float:
        exp_k = math.exp(kappa)
        return theta_n * exp_k + a * (exp_k - 1.0)
