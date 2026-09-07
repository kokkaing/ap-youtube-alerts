#!/usr/bin/env python3
"""
AP YouTube watcher.

Checks the Associated Press YouTube channel's public Atom feed for newly
published videos and posts each new one to a Microsoft Teams channel via a
Workflows ("incoming webhook") Adaptive Card.

Adapted from cooperinveen/unifeed-alerts (the UN Unifeed watcher). The core
pattern is identical (state.json watermark + seen-ids dedupe, Adaptive Card
per new item, post via a Teams Workflows webhook) but the fetch step is much
simpler here: YouTube publishes a clean, robots-allowed Atom feed per
channel, so there's no HTML scraping/regex involved.

Reads one secret from the environment:
  TEAMS_WEBHOOK_URL - Teams Workflows webhook URL (the AP YouTube channel)

Zero third-party dependencies (stdlib only).
"""

import html
import json
import os
import sys
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime, timezone

# --- config -----------------------------------------------------------------

# Associated Press's main YouTube channel (@AssociatedPress).
CHANNEL_ID = "UC52X5wxOL_s5yw0dQk7NtgA"
FEED_URL = f"https://www.youtube.com/feeds/videos.xml?channel_id={CHANNEL_ID}"
STATE_FILE = os.path.join(os.path.dirname(__file__), "state.json")

# How many recently-posted video IDs to remember (dedupe guard).
SEEN_IDS_CAP = 500

# On the very first run (no state file), seed silently instead of flooding the
# channel with the ~15 most recent videos the feed returns. Override by
# setting SEED_AND_POST=1 on a manual run.
SEED_AND_POST = os.environ.get("SEED_AND_POST") == "1"

# A browser-ish UA — the default urllib agent can get 403'd / challenged.
USER_AGENT = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
             "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36 "
             "ap-youtube-alerts/1.0")

MAX_DESC_CHARS = 400

NS = {
    "atom": "http://www.w3.org/2005/Atom",
    "yt": "http://www.youtube.com/xml/schemas/2015",
    "media": "http://search.yahoo.com/mrss/",
}


# --- helpers ----------------------------------------------------------------

def log(msg):
    print(msg, flush=True)


def load_state():
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return None
    except (json.JSONDecodeError, OSError) as e:
        log(f"WARNING: could not read state file ({e}); treating as first run.")
        return None


def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)
        f.write("\n")


def clean_text(value):
    """Neutralise markdown so untrusted feed text can't render as links/markup
    in a Teams card. Teams TextBlocks render a subset of markdown, so a
    hostile string like '[click](http://evil)' would otherwise become a live
    link. Video titles/descriptions are public but still untrusted input."""
    if not value:
        return ""
    out = str(value)
    for ch in "[]()`*_#>|":
        out = out.replace(ch, "\\" + ch)
    return out


def truncate(value, limit=MAX_DESC_CHARS):
    value = (value or "").strip()
    if len(value) <= limit:
        return value
    return value[:limit].rstrip() + "…"


def safe_https(value):
    """Return the URL only if it's a plain https:// link, else ''.
    Blocks javascript:, data:, http:, etc. from reaching a card button/image."""
    if not value:
        return ""
    v = str(value).strip()
    return v if v.lower().startswith("https://") else ""


def parse_dt(value):
    """Parse an ISO 8601 timestamp (YouTube's <published>, or our saved
    watermark) to an aware datetime, or None."""
    if not value:
        return None
    try:
        dt = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    except (ValueError, AttributeError):
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt


def fetch_items():
    """Fetch and parse the channel's Atom feed into a list of dicts
    (feed order: newest first, per YouTube's convention)."""
    req = urllib.request.Request(
        FEED_URL, headers={"Accept": "application/atom+xml", "User-Agent": USER_AGENT},
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        body = resp.read()

    root = ET.fromstring(body)
    items = []
    for entry in root.findall("atom:entry", NS):
        video_id = entry.findtext("yt:videoId", default="", namespaces=NS).strip()
        if not video_id:
            continue
        title = entry.findtext("atom:title", default="", namespaces=NS)
        published = entry.findtext("atom:published", default="", namespaces=NS)
        link_el = entry.find("atom:link[@rel='alternate']", NS)
        link = link_el.get("href") if link_el is not None else f"https://www.youtube.com/watch?v={video_id}"

        media_group = entry.find("media:group", NS)
        thumb, desc = "", ""
        if media_group is not None:
            thumb_el = media_group.find("media:thumbnail", NS)
            if thumb_el is not None:
                thumb = thumb_el.get("url", "")
            desc = media_group.findtext("media:description", default="", namespaces=NS)

        items.append({
            "id": video_id,
            "title": html.unescape(title or ""),
            "text": html.unescape(desc or ""),
            "link": link,
            "thumb": thumb,
            "pubDate": published,
        })
    return items


def build_card(item):
    """Build the Teams Adaptive Card envelope for one AP YouTube video."""
    title = clean_text(item.get("title"))
    text = clean_text(truncate(item.get("text")))
    link = safe_https(item.get("link"))
    thumb = safe_https(item.get("thumb"))

    when = parse_dt(item.get("pubDate"))
    subtitle = when.strftime("%d %b %Y, %H:%M UTC") if when else ""

    body = [
        {"type": "TextBlock", "text": "▶️ New AP YouTube video",
         "weight": "Bolder", "size": "Medium", "color": "Accent", "wrap": True},
    ]
    if title:
        body.append({"type": "TextBlock", "text": title, "weight": "Bolder",
                     "wrap": True, "spacing": "Small"})
    if subtitle:
        body.append({"type": "TextBlock", "text": subtitle, "isSubtle": True,
                     "spacing": "None", "size": "Small"})
    if thumb:
        body.append({"type": "Image", "url": thumb, "size": "Stretch",
                     "spacing": "Medium"})
    if text:
        body.append({"type": "TextBlock", "text": text, "wrap": True,
                     "spacing": "Medium"})

    actions = []
    if link:
        actions.append({"type": "Action.OpenUrl",
                        "title": "▶️ Watch on YouTube", "url": link})

    card = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.4",
        "body": body,
        "actions": actions,
    }
    return {
        "type": "message",
        "attachments": [{
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": card,
        }],
    }


def post_to_teams(webhook_url, card):
    payload = json.dumps(card).encode("utf-8")
    req = urllib.request.Request(
        webhook_url, data=payload,
        headers={"Content-Type": "application/json"}, method="POST",
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.status


# --- main -------------------------------------------------------------------

def main():
    webhook_url = os.environ.get("TEAMS_WEBHOOK_URL")
    if not webhook_url:
        log("ERROR: TEAMS_WEBHOOK_URL must be set.")
        return 1

    try:
        items = fetch_items()
    except Exception as e:
        log(f"ERROR fetching/parsing feed: {e}")
        return 1

    log(f"Fetched {len(items)} item(s) from feed.")
    if not items:
        log("WARNING: parsed 0 items — treating as a bad fetch, not advancing state.")
        return 1

    state = load_state()
    first_run = state is None
    if first_run:
        state = {"last_published": None, "seen_ids": []}

    seen_list = list(state.get("seen_ids") or [])
    seen_ids = set(seen_list)
    last_published = parse_dt(state.get("last_published"))

    # Feed is newest-first; reverse so the channel reads chronologically.
    items = list(reversed(items))

    new_items = []
    for item in items:
        item_id = item.get("id")
        if not item_id or item_id in seen_ids:
            continue
        pub = parse_dt(item.get("pubDate"))
        if last_published and pub and pub <= last_published:
            continue
        new_items.append(item)

    posting = not (first_run and not SEED_AND_POST)
    if first_run and not SEED_AND_POST:
        log(f"First run: seeding state with {len(new_items)} item(s), posting none. "
            f"(Set SEED_AND_POST=1 to post on a manual run.)")

    posted = 0
    for item in new_items:
        if posting:
            try:
                status = post_to_teams(webhook_url, build_card(item))
                log(f"Posted: {item.get('id')} @ {item.get('pubDate')} (HTTP {status})")
                posted += 1
            except Exception as e:
                log(f"ERROR posting {item.get('id')}: {e} — will retry next run.")
                continue
        iid = item.get("id")
        if iid not in seen_ids:
            seen_ids.add(iid)
            seen_list.append(iid)
        pub = parse_dt(item.get("pubDate"))
        if pub and (last_published is None or pub > last_published):
            last_published = pub

    state["last_published"] = last_published.strftime("%Y-%m-%dT%H:%M:%SZ") if last_published else None
    state["seen_ids"] = seen_list[-SEEN_IDS_CAP:]
    save_state(state)

    log(f"Done. New: {len(new_items)}, posted: {posted}, "
        f"watermark: {state['last_published']}.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
