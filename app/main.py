from __future__ import annotations

from app.shared.paths import embedded_build_tag
from app.shell.main_window import MainWindow

# Keep version in a stable location for pack.py.
APP_VERSION = "2.9.2"


def main() -> None:
    tag = embedded_build_tag()
    title = f"Электропроект  v{APP_VERSION}  · {tag}" if tag else f"Электропроект  v{APP_VERSION}"
    window = MainWindow(title=title)
    window.build()
    window.run()


if __name__ == "__main__":
    main()

