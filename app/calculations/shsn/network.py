from __future__ import annotations

from .models import NetworkElement


class NetworkGraph:
    def __init__(self, elements: list[NetworkElement]) -> None:
        self.nodes: dict[str, dict] = {
            el.name: el.as_graph_node() for el in elements
        }

    def get_path_to_source(self, start_node_name: str, *, include_start_node: bool = True) -> list[str]:
        path: list[str] = []
        current: str | None = start_node_name
        visited: set[str] = set()
        while current is not None:
            if current in visited or current not in self.nodes:
                break
            visited.add(current)
            if current != start_node_name or include_start_node:
                path.append(current)
            parent = self.nodes[current]["Parent_Name"]
            current = None if parent in ("", "нет", None) else str(parent)
        return path[::-1]

    def get_end_of_protected_zone(self, start_node: str) -> tuple[str, bool]:
        current = start_node
        while True:
            children = [n for n, d in self.nodes.items() if d["Parent_Name"] == current]
            if not children:
                return current, False
            next_breaker = next(
                (c for c in children if "автомат" in str(self.nodes[c]["Type"]).lower()),
                None,
            )
            if next_breaker:
                return next_breaker, True
            current = children[0]
