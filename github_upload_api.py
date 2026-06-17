#!/usr/bin/env python3
"""Выгрузка проекта на GitHub без установки Git (REST API).

Требуется Personal Access Token с правом repo.
Токен: переменная GITHUB_TOKEN или файл .github_token (не коммитится).

Пример:
  set GITHUB_TOKEN=ghp_...
  python github_upload_api.py https://github.com/user/electroproject
"""

from __future__ import annotations

import argparse
import base64
import fnmatch
import json
import os
import re
import ssl
import sys
import urllib.error
import urllib.request
from pathlib import Path

API_VERSION = "2022-11-28"
DEFAULT_BRANCH = "web-app"
TOKEN_FILE = ".github_token"


def _ssl_context() -> ssl.SSLContext:
    ctx = ssl.create_default_context()
    # На части корпоративных ПК нет корневых сертификатов — можно отключить проверку:
    if os.environ.get("GITHUB_SSL_NO_VERIFY", "").strip() in ("1", "true", "yes"):
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
    return ctx


def parse_repo(url: str) -> tuple[str, str]:
    url = url.strip().rstrip("/")
    if url.endswith(".git"):
        url = url[:-4]
    m = re.match(r"(?:https?://)?(?:www\.)?github\.com/([^/]+)/([^/]+)$", url, re.I)
    if m:
        return m.group(1), m.group(2)
    if "/" in url and "://" not in url:
        owner, repo = url.split("/", 1)
        return owner.strip(), repo.strip()
    raise ValueError(f"Неверный URL репозитория: {url!r}")


def load_token(explicit: str | None) -> str:
    if explicit and explicit.strip():
        return explicit.strip()
    env = os.environ.get("GITHUB_TOKEN", "").strip()
    if env:
        return env
    token_path = Path(TOKEN_FILE)
    if token_path.is_file():
        return token_path.read_text(encoding="utf-8").strip()
    raise SystemExit(
        "Нужен токен GitHub:\n"
        "  set GITHUB_TOKEN=ghp_...\n"
        f"  или положите токен в файл {TOKEN_FILE}\n"
        "  или: python github_upload_api.py URL --token ghp_..."
    )


class GitHubClient:
    def __init__(self, owner: str, repo: str, token: str) -> None:
        self.base = f"https://api.github.com/repos/{owner}/{repo}"
        self.token = token
        self._ctx = _ssl_context()

    def request(self, method: str, path: str, payload: dict | None = None) -> dict | list | None:
        url = path if path.startswith("http") else f"{self.base}{path}"
        data = None
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": API_VERSION,
            "User-Agent": "electroproject-uploader",
        }
        if payload is not None:
            data = json.dumps(payload).encode("utf-8")
            headers["Content-Type"] = "application/json"
        req = urllib.request.Request(url, data=data, headers=headers, method=method)
        try:
            with urllib.request.urlopen(req, context=self._ctx, timeout=120) as resp:
                body = resp.read()
                if not body:
                    return None
                return json.loads(body.decode("utf-8"))
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")
            raise SystemExit(f"GitHub API {method} {path}: HTTP {exc.code}\n{detail}") from exc

    def get_ref_sha(self, branch: str) -> str | None:
        try:
            ref = self.request("GET", f"/git/ref/heads/{branch}")
            assert isinstance(ref, dict)
            return ref["object"]["sha"]
        except SystemExit as exc:
            if "HTTP 404" in str(exc):
                return None
            raise

    def create_blob(self, content: bytes) -> str:
        encoding = "base64" if _needs_base64(content) else "utf-8"
        if encoding == "base64":
            content_str = base64.b64encode(content).decode("ascii")
        else:
            content_str = content.decode("utf-8")
        result = self.request(
            "POST",
            "/git/blobs",
            {"content": content_str, "encoding": encoding},
        )
        assert isinstance(result, dict)
        return result["sha"]

    def create_tree(self, entries: list[dict], base_tree: str | None = None) -> str:
        payload: dict = {"tree": entries}
        if base_tree:
            payload["base_tree"] = base_tree
        result = self.request("POST", "/git/trees", payload)
        assert isinstance(result, dict)
        return result["sha"]

    def create_commit(self, message: str, tree_sha: str, parents: list[str]) -> str:
        payload: dict = {"message": message, "tree": tree_sha}
        if parents:
            payload["parents"] = parents
        result = self.request("POST", "/git/commits", payload)
        assert isinstance(result, dict)
        return result["sha"]

    def update_ref(self, branch: str, sha: str, force: bool = False) -> None:
        ref_name = f"refs/heads/{branch}"
        existing = self.get_ref_sha(branch)
        if existing is None:
            self.request("POST", "/git/refs", {"ref": ref_name, "sha": sha})
        else:
            self.request(
                "PATCH",
                f"/git/refs/heads/{branch}",
                {"sha": sha, "force": force},
            )


def _needs_base64(data: bytes) -> bool:
    try:
        data.decode("utf-8")
        return b"\x00" in data
    except UnicodeDecodeError:
        return True


def load_gitignore_rules(root: Path) -> list[tuple[bool, str]]:
    """(is_negation, pattern) — упрощённый парсер .gitignore."""
    rules: list[tuple[bool, str]] = []
    path = root / ".gitignore"
    if not path.is_file():
        return rules
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        neg = line.startswith("!")
        if neg:
            line = line[1:].strip()
        rules.append((neg, line))
    return rules


def _match_rule(rel_posix: str, name: str, is_dir: bool, pattern: str) -> bool:
    p = pattern.replace("\\", "/")
    if p.endswith("/"):
        if not is_dir:
            return False
        p = p[:-1]
    # Правило без / матчит имя файла в любой папке
    if "/" not in p:
        return fnmatch.fnmatch(name, p) or fnmatch.fnmatch(rel_posix, p)
    return fnmatch.fnmatch(rel_posix, p) or fnmatch.fnmatch(rel_posix, p.lstrip("/"))


def is_ignored(rel_path: Path, is_dir: bool, rules: list[tuple[bool, str]]) -> bool:
    rel_posix = rel_path.as_posix()
    name = rel_path.name
    ignored = False
    for neg, pattern in rules:
        if _match_rule(rel_posix, name, is_dir, pattern):
            ignored = not neg
    return ignored


def collect_files(root: Path) -> list[Path]:
    rules = load_gitignore_rules(root)
    always_skip = {
        TOKEN_FILE,
        ".git",
        Path(TOKEN_FILE).name,
    }
    files: list[Path] = []
    for dirpath, dirnames, filenames in os.walk(root):
        rel_dir = Path(dirpath).relative_to(root)
        # Фильтр каталогов in-place для os.walk
        kept_dirs: list[str] = []
        for d in dirnames:
            rel = rel_dir / d
            if d in always_skip or is_ignored(rel, True, rules):
                continue
            kept_dirs.append(d)
        dirnames[:] = kept_dirs

        for fname in filenames:
            rel = rel_dir / fname if rel_dir.parts else Path(fname)
            if rel.as_posix() in always_skip or is_ignored(rel, False, rules):
                continue
            files.append(root / rel)
    return sorted(files)


def build_tree_entries(root: Path, files: list[Path], client: GitHubClient) -> list[dict]:
    entries: list[dict] = []
    total = len(files)
    for idx, path in enumerate(files, 1):
        rel = path.relative_to(root).as_posix()
        data = path.read_bytes()
        blob = client.create_blob(data)
        entries.append({"path": rel, "mode": "100644", "type": "blob", "sha": blob})
        print(f"  [{idx}/{total}] {rel} ({len(data)} байт)")
    return entries


def main() -> None:
    parser = argparse.ArgumentParser(description="Выгрузка на GitHub без Git")
    parser.add_argument("repo", help="https://github.com/owner/repo или owner/repo")
    parser.add_argument("--branch", default=DEFAULT_BRANCH, help=f"ветка (по умолчанию {DEFAULT_BRANCH})")
    parser.add_argument("--token", help="Personal Access Token (или GITHUB_TOKEN)")
    parser.add_argument(
        "--message",
        default="Web version: FastAPI UI, LAN access, fire check MV cable",
        help="сообщение коммита",
    )
    parser.add_argument(
        "--root",
        type=Path,
        default=Path(__file__).resolve().parent,
        help="корень проекта",
    )
    args = parser.parse_args()

    owner, repo = parse_repo(args.repo)
    token = load_token(args.token)
    root = args.root.resolve()

    print(f"Репозиторий: {owner}/{repo}")
    print(f"Ветка: {args.branch}")
    print(f"Папка: {root}")
    print()

    client = GitHubClient(owner, repo, token)
    client.request("GET", "")  # проверка доступа
    print("Доступ к репозиторию OK")

    files = collect_files(root)
    print(f"Файлов к выгрузке: {len(files)}")
    if not files:
        raise SystemExit("Нет файлов для выгрузки")

    print("Загрузка blob-ов...")
    entries = build_tree_entries(root, files, client)

    print("Создание дерева и коммита...")
    tree_sha = client.create_tree(entries)
    parent = client.get_ref_sha(args.branch)
    parents = [parent] if parent else []
    commit_sha = client.create_commit(args.message, tree_sha, parents)
    client.update_ref(args.branch, commit_sha)

    print()
    print("Готово.")
    print(f"https://github.com/{owner}/{repo}/tree/{args.branch}")
    if parent is None:
        print(f"Создана новая ветка «{args.branch}». Ветка main не изменялась.")
    else:
        print(f"Ветка «{args.branch}» обновлена. Ветка main не изменялась.")


if __name__ == "__main__":
    main()
