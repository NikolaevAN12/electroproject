"""Коэффициент снижения тока КЗ (Kc) по сопротивлению цепи Rкз."""

from __future__ import annotations


def kc_from_r_kz(r_kz: float) -> float:
    """Kc по Rкз, Ом: <0,1 → 0,5; 0,1…0,5 → 0,55; >0,5 → 0,6."""
    if r_kz < 0.1:
        return 0.5
    if r_kz <= 0.5:
        return 0.55
    return 0.6
