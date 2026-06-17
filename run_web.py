"""Веб-сервер Electroproject (локально или в локальной сети)."""

from __future__ import annotations

import argparse
import socket
import threading
import webbrowser

import uvicorn

LOCAL_HOST = "127.0.0.1"
LAN_HOST = "0.0.0.0"
PORT = 8765


def _local_ip() -> str | None:
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.connect(("8.8.8.8", 80))
            return sock.getsockname()[0]
    except OSError:
        return None


def main() -> None:
    parser = argparse.ArgumentParser(description="Веб-сервер Electroproject")
    parser.add_argument("--port", type=int, default=PORT, help=f"Порт (по умолчанию {PORT})")
    parser.add_argument(
        "--host",
        default=LAN_HOST,
        help=f"Адрес привязки (по умолчанию {LAN_HOST} — доступ из локальной сети)",
    )
    parser.add_argument(
        "--local",
        action="store_true",
        help="Только этот компьютер (привязка к 127.0.0.1)",
    )
    parser.add_argument(
        "--lan",
        action="store_true",
        help="Доступ из локальной сети (включён по умолчанию, флаг для совместимости)",
    )
    parser.add_argument("--no-browser", action="store_true", help="Не открывать браузер")
    args = parser.parse_args()

    host = LOCAL_HOST if args.local else (LAN_HOST if args.lan or args.host == LAN_HOST else args.host)
    local_url = f"http://{LOCAL_HOST}:{args.port}/"

    if not args.no_browser:
        threading.Timer(1.0, lambda: webbrowser.open(local_url)).start()

    print(local_url)
    if host == LAN_HOST:
        lan_ip = _local_ip()
        if lan_ip:
            print(f"Для других ПК в сети: http://{lan_ip}:{args.port}/")
        else:
            print("Для других ПК в сети: http://<IP-этого-компьютера>:{args.port}/")
        print("При необходимости разрешите порт в брандмауэре Windows.")
    print("Остановка: Ctrl+C")

    uvicorn.run("app.web.app:app", host=host, port=args.port, log_level="info")


if __name__ == "__main__":
    main()
