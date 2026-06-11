"""
github_store.py
================
A tiny "JSON file on GitHub = database" layer for the HKEFLPL app.

It uses the GitHub Contents REST API (no PyGithub dependency, just
`requests`). The JSON file lives in a repo at JSON_PATH and is read on
load and written back on every save as a normal git commit.

Why this works for 2-3 users with rare concurrent writes
---------------------------------------------------------
GitHub's Contents API requires the current file `sha` to update a file.
If two people save at the same time, the second PUT fails with HTTP 409
(sha mismatch). We handle that with OPTIMISTIC LOCKING + RETRY:

    1. read latest JSON (+ sha)
    2. apply your change to that latest copy
    3. PUT with the sha
    4. on 409, go back to step 1 (up to N attempts)

This is NOT a real transactional database:
  * heavy concurrent writes can still starve / conflict,
  * each write is a git commit (1-2s latency),
  * keep the JSON well under ~1 MB for snappy reads.

For this workload (a few warehouse users generating pick lists, rarely
at the same instant) it is perfectly adequate.

Secrets expected (in .streamlit/secrets.toml or env vars):
    GITHUB_TOKEN   fine-grained PAT with Contents: read & write on the repo
    GITHUB_REPO    "owner/repo"
    GITHUB_BRANCH  e.g. "main"
    JSON_PATH      e.g. "data/hkeflpl_store.json"
"""

from __future__ import annotations

import base64
import json
import time
from typing import Any, Callable

import requests

API_ROOT = "https://api.github.com"
DEFAULT_STORE = {
    "schema_version": 1,
    "pl_current": [],          # latest generated PL (list of row dicts)
    "history": [],             # list of run dicts (newest first)
    "comparison_summaries": [],  # list of comparison summary dicts (newest first)
}
MAX_HISTORY = 100              # keep the JSON small; trim oldest beyond this


class GitHubStoreError(Exception):
    pass


class GitHubStore:
    def __init__(self, token: str, repo: str, branch: str = "main",
                 path: str = "data/hkeflpl_store.json"):
        if not token or not repo:
            raise GitHubStoreError("GITHUB_TOKEN සහ GITHUB_REPO දෙකම ඕන.")
        self.token = token
        self.repo = repo.strip().strip("/")
        self.branch = branch or "main"
        self.path = path.strip().strip("/")
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": f"Bearer {self.token}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
        })

    # ---- low level ----
    def _contents_url(self) -> str:
        return f"{API_ROOT}/repos/{self.repo}/contents/{self.path}"

    def _get_raw(self) -> tuple[dict, str | None]:
        """Return (data, sha). If the file doesn't exist yet, returns
        (DEFAULT_STORE copy, None)."""
        r = self.session.get(self._contents_url(), params={"ref": self.branch},
                             timeout=30)
        if r.status_code == 404:
            return json.loads(json.dumps(DEFAULT_STORE)), None
        if r.status_code == 401:
            raise GitHubStoreError("GitHub token වැරදියි / expired (401).")
        if r.status_code == 403:
            raise GitHubStoreError("GitHub access denied / rate limit (403).")
        if not r.ok:
            raise GitHubStoreError(f"GitHub read fail: {r.status_code} {r.text[:200]}")
        payload = r.json()
        content = base64.b64decode(payload["content"]).decode("utf-8")
        try:
            data = json.loads(content) if content.strip() else json.loads(json.dumps(DEFAULT_STORE))
        except json.JSONDecodeError:
            data = json.loads(json.dumps(DEFAULT_STORE))
        return data, payload["sha"]

    def _put_raw(self, data: dict, sha: str | None, message: str) -> dict:
        body = {
            "message": message,
            "content": base64.b64encode(
                json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
            ).decode("ascii"),
            "branch": self.branch,
        }
        if sha:
            body["sha"] = sha
        r = self.session.put(self._contents_url(), json=body, timeout=30)
        return r

    # ---- public ----
    def load(self) -> dict:
        data, _ = self._get_raw()
        for k, v in DEFAULT_STORE.items():
            data.setdefault(k, v if not isinstance(v, list) else [])
        return data

    def test_connection(self) -> str:
        """Returns a short human status string, raises on hard failure."""
        r = self.session.get(f"{API_ROOT}/repos/{self.repo}", timeout=20)
        if r.status_code == 404:
            raise GitHubStoreError(f"Repo හමු නොවුණා: {self.repo}")
        if r.status_code in (401, 403):
            raise GitHubStoreError(f"Auth fail ({r.status_code}). Token/permission බලන්න.")
        if not r.ok:
            raise GitHubStoreError(f"GitHub error {r.status_code}")
        data, sha = self._get_raw()
        state = "file එක තියෙනවා" if sha else "file එක තවම නෑ (පළවෙනි save එකේදී හැදෙයි)"
        runs = len(data.get("history", []))
        return f"✅ Connected → {self.repo}@{self.branch} | {state} | runs: {runs}"

    def update(self, mutate_fn: Callable[[dict], dict], message: str,
               retries: int = 5) -> dict:
        """Optimistic-locked update.

        mutate_fn receives the latest store dict and must return the new
        store dict (mutating in place is fine too). Retries on 409.
        """
        last_err = None
        for attempt in range(retries):
            data, sha = self._get_raw()
            for k, v in DEFAULT_STORE.items():
                data.setdefault(k, v if not isinstance(v, list) else [])
            new_data = mutate_fn(data)
            # keep JSON small
            new_data["history"] = new_data.get("history", [])[:MAX_HISTORY]
            new_data["comparison_summaries"] = \
                new_data.get("comparison_summaries", [])[:MAX_HISTORY]

            r = self._put_raw(new_data, sha, message)
            if r.ok:
                return new_data
            if r.status_code == 409:  # sha conflict -> someone else wrote
                last_err = "409 conflict"
                time.sleep(0.6 * (attempt + 1))
                continue
            if r.status_code in (401, 403):
                raise GitHubStoreError(f"GitHub write denied ({r.status_code}). "
                                       "Token permission (Contents: write) බලන්න.")
            raise GitHubStoreError(f"GitHub write fail: {r.status_code} {r.text[:200]}")
        raise GitHubStoreError(
            f"Write conflict — වෙන කෙනෙක් එකවර save කරනවා වෙන්න පුළුවන් ({last_err}). "
            "ආයෙ try කරන්න.")


# ---- store mutations (used by the app) ----
def add_run(store: dict, *, run_id: str, user: str, load_ids: list[str],
            sl_flag: str, pl_records: list[dict],
            comparison: dict | None) -> dict:
    """Set current PL, push a history entry, optionally push a comparison."""
    store["pl_current"] = pl_records
    entry = {
        "run_id": run_id,
        "timestamp": _now(),
        "user": user,
        "load_ids": load_ids,
        "sl_flag": sl_flag,
        "record_count": len(pl_records),
        "pl_records": pl_records,
    }
    store["history"] = [entry] + store.get("history", [])
    if comparison is not None:
        comp = dict(comparison)
        comp.setdefault("run_id", run_id)
        comp.setdefault("timestamp", _now())
        comp.setdefault("user", user)
        store["comparison_summaries"] = \
            [comp] + store.get("comparison_summaries", [])
    return store


def reset_current(store: dict) -> dict:
    """Module7 equivalent: clear the current PL but KEEP history/summaries."""
    store["pl_current"] = []
    return store


def _now() -> str:
    from datetime import datetime
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
