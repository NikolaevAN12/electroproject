"""
Сборка EXE: только этот скрипт (или build_windows.bat). Не вызывайте вручную
pyinstaller Electroproject.spec — старый spec подмешивал чужие пути и «вечную» сборку.

PyInstaller пишет сразу в dist. Метка сборки: sha256 и mtime main.py в build_tag.txt.
Лог: %%TEMP%%\\electroproject_pack.log
"""
from __future__ import annotations

import hashlib
import os
import re
import shutil
import subprocess
import sys
import traceback
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent
MAIN = ROOT / "app" / "main.py"
DIST_DIR = ROOT / "dist"
# Свой каталог кэша PyInstaller (иначе берётся %%LOCALAPPDATA%%\pyinstaller — там «залипает» старый анализ).
PI_CACHE = ROOT / ".electroproject_pyi_cache"
LAST_BUILD_MARKER = DIST_DIR / "_last_build.txt"
ERROR_HINT = DIST_DIR / "POCHEMU_NE_SOZDALOS.txt"
LOG_FILE = Path(os.environ.get("TEMP", os.environ.get("TMP", r"C:\Windows\Temp"))) / "electroproject_pack.log"


def _log(msg: str) -> None:
    line = f"{datetime.now().isoformat()} {msg}\n"
    try:
        with LOG_FILE.open("a", encoding="utf-8") as f:
            f.write(line)
    except OSError:
        pass
    print(msg)


def _write_error_hint(title: str, detail: str = "") -> None:
    try:
        DIST_DIR.mkdir(parents=True, exist_ok=True)
        body = (
            f"{title}\n\n{detail}\n\n"
            "Частые причины:\n"
            "1. Запуск build_windows.bat из папки с pack.py и main.py.\n"
            "2. pip install -r requirements.txt\n"
            "3. Антивирус удалил exe — см. журнал Защитника.\n"
            "4. Запускайте Electroproject_LATEST.exe из dist (см. dist\\CHTO_ZAPUSKAT.txt).\n"
            "5. Закройте программу перед сборкой — иначе LATEST.exe не перезапишется.\n"
            "6. Полный лог: %TEMP%\\electroproject_pack.log\n"
            "7. Запасная копия: %LOCALAPPDATA%\\Electroproject_build\\\n"
        )
        ERROR_HINT.write_text(body, encoding="utf-8")
        print("Подсказка:", ERROR_HINT.resolve(), file=sys.stderr)
    except OSError:
        pass


def _safe_version_for_filename(ver: str) -> str:
    v = ver.strip()
    if not v or v == "?":
        v = "0"
    return re.sub(r"[^\w.\-]", "_", v)


def _subprocess_nowindow() -> dict:
    if sys.platform != "win32":
        return {}
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    return {"creationflags": flags} if flags else {}


def _try_kill_electroproject_processes() -> None:
    """Чтобы Windows отпустил LATEST.exe и подтянулась новая сборка."""
    if sys.platform != "win32":
        return
    kw = _subprocess_nowindow()
    for im in ("Electroproject_LATEST.exe", "Electroproject.exe"):
        r = subprocess.run(
            ["taskkill", "/F", "/IM", im, "/T"],
            capture_output=True,
            **kw,
        )
        if r.returncode == 0:
            _log(f"Остановлен процесс: {im}")


def _atomic_replace_exe(src: Path, dest: Path) -> None:
    """Копия через .new + os.replace — надёжнее, чем copy2 поверх запущенного файла."""
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp = dest.with_suffix(dest.suffix + ".new")
    try:
        shutil.copy2(src, tmp)
        os.replace(tmp, dest)
    finally:
        if tmp.is_file():
            try:
                tmp.unlink()
            except OSError:
                pass


def _copy_to_localappdata_fallback(dist_file: Path) -> None:
    la = os.environ.get("LOCALAPPDATA", "").strip()
    if not la:
        return
    dest_dir = Path(la) / "Electroproject_build"
    try:
        dest_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(dist_file, dest_dir / dist_file.name)
        latest_fb = dest_dir / "Electroproject_LATEST.exe"
        _atomic_replace_exe(dist_file, latest_fb)
        _log(f"Запасная копия: {latest_fb.resolve()}")
        try:
            (dest_dir / "POSLEDNIY_EXE.txt").write_text(
                str(latest_fb.resolve()) + "\n" + str((dest_dir / dist_file.name).resolve()),
                encoding="utf-8",
            )
        except OSError:
            pass
    except OSError as e:
        _log(f"AppData копия не удалась: {e}")


def main() -> int:
    _log("=== старт pack.py ===")
    _log(f"ROOT={ROOT.resolve()}")

    if ERROR_HINT.is_file():
        try:
            ERROR_HINT.unlink()
        except OSError:
            pass

    if not MAIN.is_file():
        _write_error_hint(f"Нет main.py: {MAIN}")
        return 1

    DIST_DIR.mkdir(parents=True, exist_ok=True)

    probe = DIST_DIR / "_zapis_rabotaet.txt"
    try:
        probe.write_text(datetime.now().isoformat(), encoding="utf-8")
        probe.unlink()
    except OSError as e:
        _write_error_hint(f"Не могу писать в папку dist:\n{DIST_DIR}", str(e))
        _log(f"FAIL запись dist: {e}")
        return 1

    build_dir = ROOT / "build"
    for d in (build_dir, PI_CACHE):
        if d.is_dir():
            shutil.rmtree(d, ignore_errors=True)
    PI_CACHE.mkdir(parents=True, exist_ok=True)
    _log("Очищены папки build и .electroproject_pyi_cache (кэш PyInstaller только в проекте)")

    pyc = ROOT / "__pycache__"
    if pyc.is_dir():
        shutil.rmtree(pyc, ignore_errors=True)

    main_raw = MAIN.read_bytes()
    try:
        main_src = main_raw.decode("utf-8")
    except UnicodeDecodeError:
        main_src = ""
        _write_error_hint("main.py не UTF-8", str(MAIN))
        return 1

    sha16 = hashlib.sha256(main_raw).hexdigest()[:16]
    mtime_human = datetime.fromtimestamp(
        MAIN.stat().st_mtime, tz=timezone.utc
    ).astimezone().strftime("%Y-%m-%d %H:%M:%S %z")
    _log(f"main.py отпечаток: mtime={mtime_human} sha256[0:16]={sha16}")

    ver_m = re.search(
        r'^APP_VERSION\s*=\s*["\']([^"\']+)["\']',
        main_src,
        re.MULTILINE,
    )
    ver_s = ver_m.group(1) if ver_m else "?"

    safe = _safe_version_for_filename(ver_s)
    ts = datetime.now(timezone.utc).astimezone().strftime("%Y%m%d_%H%M%S")
    # Имя без точек — надёжнее для PyInstaller
    build_name = f"Electroproject_v{safe.replace('.', '_')}_{ts}"
    dist_out = DIST_DIR / f"{build_name}.exe"
    root_out = ROOT / f"{build_name}.exe"

    for legacy in (DIST_DIR / "Electroproject.exe", ROOT / "Electroproject.exe"):
        if legacy.is_file():
            try:
                legacy.unlink()
                _log(f"Удалён устаревший (не из pack.py): {legacy.name}")
            except OSError as e:
                _log(f"Не удалось удалить {legacy}: {e}")

    _try_kill_electroproject_processes()

    tag_path = ROOT / "_ep_build_tag.txt"
    build_stamp = datetime.now(timezone.utc).astimezone().strftime("%Y-%m-%d %H:%M")
    try:
        tag_path.write_text(
            f"сборка {ver_s} {build_stamp} · sha256:{sha16} · main mtime {mtime_human}\n",
            encoding="utf-8",
        )
        add_data_arg = f"{tag_path};."

        cmd = [
            sys.executable,
            "-m",
            "PyInstaller",
            "--onefile",
            "--windowed",
            "--noupx",
            "--clean",
            "--name",
            build_name,
            "--add-data",
            add_data_arg,
            "--distpath",
            str(DIST_DIR.resolve()),
            "--workpath",
            str((ROOT / "build").resolve()),
            "--specpath",
            str(ROOT.resolve()),
            "-y",
            # docx подтягивается не из статического import — явно. Без --collect-all lxml (ломает/раздувает сборку).
            "--hidden-import",
            "docx",
            "--hidden-import",
            "docx.oxml",
            "--hidden-import",
            "openpyxl",
            "--hidden-import",
            "PIL",
            "--hidden-import",
            "PIL.Image",
            "--hidden-import",
            "numpy",
            "--hidden-import",
            "matplotlib",
            "--collect-data",
            "docx",
            str(MAIN.resolve()),
        ]
        env = {
            **os.environ,
            "PYINSTALLER_CONFIG_DIR": str(PI_CACHE.resolve()),
            "PYTHONDONTWRITEBYTECODE": "1",
        }
        _log("Запуск PyInstaller (вывод ниже)")
        _log(f"PYINSTALLER_CONFIG_DIR={env['PYINSTALLER_CONFIG_DIR']}")
        print("EXE будет здесь:", dist_out.resolve())
        print("Команда:", " ".join(cmd))
        r = subprocess.run(
            cmd,
            cwd=str(ROOT.resolve()),
            stderr=subprocess.STDOUT,
            env=env,
        )
        _log(f"PyInstaller код выхода: {r.returncode}")

        if tag_path.is_file():
            try:
                tag_path.unlink()
            except OSError:
                pass

        if r.returncode != 0:
            _write_error_hint(
                f"PyInstaller код {r.returncode}",
                "Выполните: python -m pip install -r requirements.txt\nСмотрите лог: %TEMP%\\electroproject_pack.log",
            )
            return r.returncode

        if not dist_out.is_file():
            found = list(DIST_DIR.glob("Electroproject*.exe"))
            found = [p for p in found if p.name != "Electroproject_LATEST.exe"]
            if found:
                found.sort(key=lambda p: p.stat().st_mtime, reverse=True)
                dist_out = found[0]
                root_out = ROOT / dist_out.name
                _log(f"Взят найденный exe: {dist_out}")
            else:
                _write_error_hint(
                    f"После PyInstaller нет файла:\n{dist_out}",
                    "\n".join(str(p) for p in DIST_DIR.iterdir())[:2000],
                )
                return 1

        sz = dist_out.stat().st_size
        if sz < 1_000_000:
            _write_error_hint(f"Файл слишком маленький: {sz} байт")
            return 1

        blob = dist_out.read_bytes()
        # В onefile многое сжато — проверяем несколько якорей из main.py (UTF-8 и ASCII).
        checks = (
            "Проверка на невозгорание".encode("utf-8"),
            b"class FireCheckCalculator",
            f'APP_VERSION = "{ver_s}"'.encode("utf-8"),
        )
        if any(c in blob for c in checks):
            _log("Проверка: в exe найдены фрагменты текущего main.py — сборка не «пустой» кэш.")
        else:
            _log(
                "Проверка: якоря main.py в сыром exe не видны (нормально при сжатии). "
                "Запуск: dist\\RUN_THIS_EXE.bat — там точный путь к этой сборке."
            )

        try:
            shutil.copy2(dist_out, root_out)
        except OSError as e:
            _log(f"Корень проекта не записан: {e}")

    except Exception as e:
        _write_error_hint("Исключение", traceback.format_exc())
        _log(traceback.format_exc())
        print(e, file=sys.stderr)
        return 1

    latest_dist = DIST_DIR / "Electroproject_LATEST.exe"
    latest_root = ROOT / "Electroproject_LATEST.exe"
    compat_dist = DIST_DIR / "Electroproject.exe"
    compat_root = ROOT / "Electroproject.exe"
    try:
        _try_kill_electroproject_processes()
        _atomic_replace_exe(dist_out, latest_dist)
        _atomic_replace_exe(dist_out, latest_root)
        _atomic_replace_exe(dist_out, compat_dist)
        _atomic_replace_exe(dist_out, compat_root)
    except OSError as e:
        _log(f"LATEST/compat exe не обновлён: {e}")

    stamp = datetime.now(timezone.utc).astimezone().strftime("%Y-%m-%d %H:%M:%S %z")
    ver_line = f"Версия в main.py: {ver_s}\n" if ver_m else ""
    fp_line = f"Отпечаток main.py: sha256[0:16]={sha16}  mtime={mtime_human}\n"
    try:
        (DIST_DIR / "последняя_сборка.txt").write_text(
            f"Время сборки: {stamp}\n"
            f"{ver_line}"
            f"{fp_line}"
            f"Запуск (надёжно): dist\\RUN_THIS_EXE.bat\n"
            f"Свежий exe: {dist_out.resolve()}\n"
            f"Копия: dist\\Electroproject_LATEST.exe\n"
            f"Совместимость: dist\\Electroproject.exe\n"
            f"Корень (архив): {root_out.resolve()}\n",
            encoding="utf-8",
        )
    except OSError:
        pass

    try:
        (DIST_DIR / "SOBRALO_OK.txt").write_text(
            f"OK\nЗАПУСК (свежая сборка): dist\\RUN_THIS_EXE.bat\n"
            f"Путь к exe: {dist_out.resolve()}\n"
            f"LATEST: {latest_dist.resolve()}\n"
            f"Electroproject.exe: {compat_dist.resolve()}\n\n"
            f"Архив: {dist_out.name}\nразмер: {dist_out.stat().st_size}\n"
            f"{fp_line}"
            "Сборка только через pack.py или build_windows.bat (не pyinstaller *.spec).\n",
            encoding="utf-8",
        )
        LAST_BUILD_MARKER.write_text(str(latest_dist.resolve()), encoding="utf-8")
    except OSError:
        pass

    try:
        (DIST_DIR / "CHTO_ZAPUSKAT.txt").write_text(
            "1) Главный способ — двойной щелчок по RUN_THIS_EXE.bat (в этой папке).\n"
            "   Он всегда запускает exe, собранный последним запуском pack.py.\n\n"
            "2) Либо откройте в проводнике файл из LAST_EXE_PATH.txt (одна строка — полный путь к .exe).\n\n"
            "3) Electroproject_LATEST.exe — удобная копия; если программа была открыта, она могла не обновиться.\n\n"
            "В заголовке окна: версия + «сборка … sha256:…» — сравните с последняя_сборка.txt.\n",
            encoding="utf-8",
        )
    except OSError:
        pass

    exe_abs = str(dist_out.resolve())
    try:
        (DIST_DIR / "LAST_EXE_PATH.txt").write_text(exe_abs + "\n", encoding="utf-8")
        run_bat = DIST_DIR / "RUN_THIS_EXE.bat"
        run_bat.write_text(
            "@echo off\r\n"
            "chcp 65001 >nul\r\n"
            f'start "" "{exe_abs}"\r\n',
            encoding="utf-8",
        )
        _log(f"Создан {run_bat.name} — запуск свежего: {exe_abs}")
    except OSError as e:
        _log(f"Не записан RUN_THIS_EXE.bat / LAST_EXE_PATH: {e}")

    _copy_to_localappdata_fallback(dist_out)

    if ver_m:
        print("Версия из main.py:", ver_m.group(1))
    print("Отпечаток main.py:", sha16, "|", mtime_human)
    print("--- УСПЕХ ---")
    print("ЗАПУСК СВЕЖЕЙ СБОРКИ: двойной щелчок по файлу")
    print(" ", (DIST_DIR / "RUN_THIS_EXE.bat").resolve())
    print("или exe:", exe_abs)
    print("LATEST (если программа не была запущена при сборке):", latest_dist.resolve())
    print("Лог при проблемах:", LOG_FILE.resolve())

    if not dist_out.is_file():
        _write_error_hint("Файл в dist пропал после сборки", str(dist_out))
        return 1
    _log("=== успех ===")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
