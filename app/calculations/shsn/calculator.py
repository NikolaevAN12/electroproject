from __future__ import annotations

import math

from .models import DEFAULT_K_TEMP, ElementOnPath, ScPointResult, ShsnInput, ShsnResult
from .network import NetworkGraph


class ShsnCalculator:
    def calculate(self, input_model: ShsnInput) -> ShsnResult:
        graph = NetworkGraph(input_model.elements)
        calc = _Calculator(graph, k_temp=input_model.k_temp)
        sc_names = [el.name for el in input_model.elements if el.is_sc_point]
        results = {name: calc.calculate_sc_point(name) for name in sc_names}
        msg = f"Рассчитано точек КЗ: {len(results)}"
        return ShsnResult(sc_results=results, message=msg)


class _Calculator:
    def __init__(self, graph: NetworkGraph, k_temp: float = DEFAULT_K_TEMP) -> None:
        self.graph = graph
        self.k_temp = k_temp

    def get_sc_params(self, path: list[str]) -> dict[str, float]:
        elements = [self.graph.nodes[name] for name in path if name in self.graph.nodes]
        r1, x1, r0, x0 = (
            sum(el[k] for el in elements) / 1000.0 for k in ("R1", "X1", "R0", "X0")
        )
        u_nom = elements[-1]["U_nom"] if elements else 0.4
        if math.isnan(u_nom):
            u_nom = 0.4
        z1 = math.sqrt(r1**2 + x1**2)
        i_sc_3 = u_nom / (math.sqrt(3) * z1) if z1 > 0 else 0.0
        r_total_1ph_cold = 2 * r1 + r0
        r_total_1ph_hot = r_total_1ph_cold * self.k_temp
        x_total_1ph = 2 * x1 + x0
        z_loop_cold = math.sqrt(r_total_1ph_cold**2 + x_total_1ph**2)
        i_sc_1_cold = (u_nom * math.sqrt(3)) / z_loop_cold if z_loop_cold > 0 else 0.0
        z_loop_hot = math.sqrt(r_total_1ph_hot**2 + x_total_1ph**2)
        i_sc_1_hot = (u_nom * math.sqrt(3)) / z_loop_hot if z_loop_hot > 0 else 0.0
        return {
            "r1": r1,
            "x1": x1,
            "r0": r0,
            "x0": x0,
            "i_sc_3": i_sc_3,
            "i_sc_1_cold": i_sc_1_cold,
            "i_sc_1_hot": i_sc_1_hot,
            "u_nom": u_nom,
        }

    def calculate_sc_point(self, node_name: str) -> ScPointResult:
        path = self.graph.get_path_to_source(node_name)
        node_data = self.graph.nodes.get(node_name, {"Phase_Type": "3PH"})
        params = self.get_sc_params(path)
        elements_data = [
            ElementOnPath(
                name=n,
                type=str(self.graph.nodes[n]["Type"]),
                r1=self.graph.nodes[n]["R1"] / 1000.0,
                x1=self.graph.nodes[n]["X1"] / 1000.0,
                r0=self.graph.nodes[n]["R0"] / 1000.0,
                x0=self.graph.nodes[n]["X0"] / 1000.0,
            )
            for n in path
            if n in self.graph.nodes
        ]
        r_total_cold = 2 * params["r1"] + params["r0"]
        return ScPointResult(
            path=path,
            phase_type=str(node_data.get("Phase_Type", "3PH")),
            elements_data=elements_data,
            r_sum_1=params["r1"],
            x_sum_1=params["x1"],
            r_sum_0=params["r0"],
            x_sum_0=params["x0"],
            i_sc_3=params["i_sc_3"],
            i_sc_1_cold=params["i_sc_1_cold"],
            i_sc_1_hot=params["i_sc_1_hot"],
            u_nom=params["u_nom"],
            r_total_1ph_cold=r_total_cold,
            r_total_1ph_hot=r_total_cold * self.k_temp,
            x_total_1ph=2 * params["x1"] + params["x0"],
        )
