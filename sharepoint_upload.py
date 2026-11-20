#!/usr/bin/env python3
"""Upload a finished deck to the SMT Weekly Deck folder on SharePoint.

Uses the same Edge profile as the Looker captures, so the Microsoft sign-in is kept
between runs. The first time, a sign-in page opens: sign in there and the upload
carries on by itself.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from urllib.parse import quote

from playwright.sync_api import sync_playwright

SHAREPOINT_HOST = "https://eshopworldirl.sharepoint.com"
SITE_PATH = "/sites/OperationsSharepoint"
SIGN_IN_TIMEOUT_MS = 300_000  # five minutes to finish signing in


def odata_string(value: str) -> str:
    """Quote a value for a SharePoint REST URL ('' escapes a quote, then URL-encode)."""
    return quote(value.replace("'", "''"), safe="")


def wait_for_sign_in(page) -> None:
    """Open the site and, if Microsoft asks for a sign-in, wait until the user finishes it."""
    site_url = SHAREPOINT_HOST + SITE_PATH
    page.goto(site_url, wait_until="domcontentloaded", timeout=120_000)
    if page.url.startswith(SHAREPOINT_HOST):
        return
    print("  Sign in to SharePoint (Microsoft) in the Edge window - the upload continues once you're in.", flush=True)
    page.wait_for_url(f"{SHAREPOINT_HOST}/**", timeout=SIGN_IN_TIMEOUT_MS)
    page.wait_for_load_state("domcontentloaded")


def upload(deck: Path, folder: str, profile: str, channel: str) -> str:
    """Upload ``deck`` into the SharePoint ``folder`` (server-relative path) and return its web URL."""
    data = deck.read_bytes()
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            profile, channel=channel, headless=False,
            ignore_default_args=["--enable-automation"],
            args=["--disable-blink-features=AutomationControlled"])
        try:
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            wait_for_sign_in(page)

            # SharePoint needs a request digest for any write made with the browser's cookies.
            api = SHAREPOINT_HOST + SITE_PATH + "/_api"
            info = ctx.request.post(f"{api}/contextinfo", headers={"Accept": "application/json;odata=nometadata"})
            if not info.ok:
                raise RuntimeError(f"SharePoint refused the sign-in check ({info.status}). Are you signed in with your work account?")
            digest = info.json()["FormDigestValue"]

            url = (
                f"{api}/web/GetFolderByServerRelativeUrl('{odata_string(folder)}')"
                f"/Files/add(url='{odata_string(deck.name)}',overwrite=true)"
            )
            resp = ctx.request.post(
                url,
                headers={
                    "Accept": "application/json;odata=nometadata",
                    "X-RequestDigest": digest,
                    "Content-Type": "application/octet-stream",
                },
                data=data,
                timeout=300_000,
            )
            if not resp.ok:
                raise RuntimeError(f"Upload failed ({resp.status}): {resp.text()[:500]}")
            return SHAREPOINT_HOST + resp.json()["ServerRelativeUrl"]
        finally:
            ctx.close()


def main() -> int:
    ap = argparse.ArgumentParser(description="Upload a deck to the SMT Weekly Deck folder on SharePoint.")
    ap.add_argument("--file", required=True, help="The .pptx to upload.")
    ap.add_argument("--folder", required=True, help="Server-relative SharePoint folder path.")
    ap.add_argument("--profile", required=True)
    ap.add_argument("--channel", default="msedge")
    a = ap.parse_args()

    deck = Path(a.file)
    if not deck.is_file():
        print(f"SharePoint upload: file not found: {deck}", file=sys.stderr)
        return 1
    try:
        web_url = upload(deck, a.folder, a.profile, a.channel)
    except Exception as exc:  # report and fail; the local deck is untouched
        print(f"SharePoint upload FAILED: {exc}", file=sys.stderr)
        return 1
    print(f"Uploaded to SharePoint: {web_url}", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
