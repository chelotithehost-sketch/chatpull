"""Core engine for the tawk.to bulk chat export (no Streamlit code in here).

Flow
  1. LIST   : chat.list is called one day at a time (a window that comes back "full" is split in two),
              every chat id is saved to SQLite.
  2. DETAIL : chat.get is called once per chat (messages are not included in chat.list), through a
              shared adaptive rate limiter, by a small pool of worker threads.
  3. EXPORT : rows are streamed from SQLite into an .xlsx (constant memory).

Everything is checkpointed in SQLite, so a crash / redeploy / Stop can be resumed without
re-downloading what is already done.
"""
import datetime as dt
import os
import queue
import re
import sqlite3
import threading
import time
from zoneinfo import ZoneInfo

import requests
import xlsxwriter

BASE_URL = "https://api.tawk.to/v1"
TZ = ZoneInfo("Africa/Nairobi")  # chat dates/times in the spreadsheet are shown in this zone
UTC = dt.timezone.utc
CELL_LIMIT = 32000  # Excel hard limit is 32,767 characters per cell


class Stopped(Exception):
    pass


class FatalError(Exception):
    """Wrong key / no permission: retrying will not help."""


class ApiError(Exception):
    pass


# --------------------------------------------------------------------------- parsing helpers
def pick(o, keys):
    if isinstance(o, dict):
        for k in keys:
            v = o.get(k)
            if v is not None and v != "":
                return v
    return None


def to_dt(v):
    if v is None or v == "":
        return None
    try:
        if isinstance(v, (int, float)) or re.fullmatch(r"\d+(\.\d+)?", str(v)):
            n = float(v)
            return dt.datetime.fromtimestamp(n / 1000 if n >= 1e12 else n, tz=UTC)
        d = dt.datetime.fromisoformat(str(v).replace("Z", "+00:00"))
        return d if d.tzinfo else d.replace(tzinfo=UTC)
    except (ValueError, OverflowError, OSError):
        return None


def find_array(j, depth=0):
    if isinstance(j, list):
        if not j:
            return []
        return j if isinstance(j[0], dict) else None
    if isinstance(j, dict) and depth < 4:
        for k in ("chats", "conversations", "data", "items", "results"):
            if k in j:
                r = find_array(j[k], depth + 1)
                if r is not None:
                    return r
    return None


def chat_obj(j):
    """chat.get can wrap the chat in several ways; unwrap to the chat dict."""
    if isinstance(j, list):
        return {"messages": j}
    if isinstance(j, dict):
        if isinstance(j.get("chat"), dict):
            return j["chat"]
        d = j.get("data")
        if isinstance(d, dict):
            return d
        if isinstance(d, list):
            return {"messages": d}
    return j if isinstance(j, dict) else {}


def norm_msg(m):
    if not isinstance(m, dict):
        return {"type": "unknown", "mtype": "msg", "name": "", "text": str(m or ""), "time": None, "files": []}
    s = m.get("sender") if isinstance(m.get("sender"), dict) else {}
    t = str(
        pick(s, ["t", "type"])
        or pick(m, ["t", "senderType", "sender_type"])
        or (m["sender"] if isinstance(m.get("sender"), str) else "")
    ).lower()
    typ = "agent" if t.startswith("a") else "visitor" if t.startswith("v") else "system" if t.startswith("s") else (t or "unknown")
    name = str(pick(s, ["n", "name"]) or pick(m, ["n", "senderName", "name"]) or "")
    text = str(pick(m, ["msg", "text", "message", "body", "content"]) or "")
    files = []
    att = pick(m, ["attchs", "attachments", "files"])
    if isinstance(att, list):
        for a in att:
            f = a["content"]["file"] if isinstance(a, dict) and isinstance(a.get("content"), dict) and a["content"].get("file") else a
            if isinstance(f, str):
                files.append(f)
            elif isinstance(f, dict):
                files.append(str(pick(f, ["name", "filename"]) or pick(f, ["url", "link", "src"]) or ""))
    if str(m.get("type", "")) == "file" and isinstance(m.get("data"), dict):
        fd = m["data"].get("file") if isinstance(m["data"].get("file"), dict) else m["data"]
        files.append(str(pick(fd, ["name", "filename"]) or pick(fd, ["url", "link", "src"]) or ""))
    return {
        "type": typ,
        "mtype": str(m.get("type", "msg")),
        "name": name,
        "text": text,
        "time": to_dt(pick(m, ["time", "createdAt", "timestamp", "sentAt", "date", "ts"])),
        "files": [f for f in files if f],
    }


def visitor_of(raw):
    v = raw.get("visitor") if isinstance(raw.get("visitor"), dict) else {}
    return str(pick(v, ["name"]) or pick(raw, ["visitorName"]) or ""), str(pick(v, ["email"]) or "")


def parse_list_item(raw):
    name, email = visitor_of(raw)
    started = to_dt(pick(raw, ["createdOn", "createdAt", "startedAt", "startTime", "time", "created"]))
    return {
        "id": str(pick(raw, ["id", "_id", "chatId", "cid"]) or ""),
        "started": started.timestamp() if started else None,
        "name": name,
        "email": email,
    }


def parse_chat(j):
    """chat.get reply -> dict(name, email, started, duration, content)."""
    raw = chat_obj(j)
    mr = pick(raw, ["messages", "msgs", "conversation", "transcript"])
    msgs = [norm_msg(m) for m in mr] if isinstance(mr, list) else []
    msgs.sort(key=lambda m: m["time"] or dt.datetime.min.replace(tzinfo=UTC))
    name, email = visitor_of(raw)
    started = to_dt(pick(raw, ["createdOn", "createdAt", "startedAt", "startTime", "time", "created"]))
    times = [m["time"] for m in msgs if m["time"]]
    if not started and times:
        started = times[0]
    duration = None
    if len(times) >= 2:
        duration = int((times[-1] - times[0]).total_seconds())
    else:
        end = to_dt(pick(raw, ["endedOn", "closedOn", "endedAt", "closedAt", "endTime"]))
        if end and started:
            duration = max(0, int((end - started).total_seconds()))
    lines = []
    for m in msgs:
        parts = [m["text"]] if m["text"] else []
        parts += [f"[file: {f}]" for f in m["files"]]
        if not parts:
            continue
        ts = m["time"].astimezone(TZ).strftime("%H:%M:%S") if m["time"] else "--:--:--"
        who = "system" if m["type"] == "system" else f"{m['name'] or m['type']} ({m['type']})"
        lines.append(f"[{ts}] {who}: " + " ".join(parts))
    return {
        "name": name,
        "email": email,
        "started": started.timestamp() if started else None,
        "duration": duration,
        "content": "\n".join(lines),
    }


# --------------------------------------------------------------------------- rate limiting + client
class Limiter:
    """Shared by all threads. Spaces requests at `rate`/s; halves the rate on a 429 and creeps back up."""

    def __init__(self, target):
        self.target = float(target)
        self.rate = float(target)
        self.lock = threading.Lock()
        self.next = 0.0
        self.streak = 0
        self.last_cut = 0.0
        self.throttles = 0

    def wait(self):
        with self.lock:
            now = time.monotonic()
            t = max(now, self.next)
            self.next = t + 1.0 / self.rate
        if t > now:
            time.sleep(t - now)

    def success(self):
        with self.lock:
            self.streak += 1
            # creep back towards the target after 50 clean calls, and only if no recent 429
            if self.streak >= 50 and self.rate < self.target and time.monotonic() - self.last_cut > 10:
                self.rate = min(self.target, self.rate * 1.2)
                self.streak = 0

    def throttled(self, retry_after=None):
        with self.lock:
            now = time.monotonic()
            self.throttles += 1
            self.streak = 0
            # several threads usually hit the same 429 burst: cut the rate once per burst, not once per thread
            if now - self.last_cut > 5:
                self.rate = max(0.5, self.rate * 0.7)
                self.last_cut = now
            self.next = max(self.next, now + (retry_after or 3.0))


class Tawk:
    def __init__(self, key, limiter, stop, base=BASE_URL):
        self.key, self.limiter, self.stop, self.base = key, limiter, stop, base.rstrip("/")
        self.session = requests.Session()
        adapter = requests.adapters.HTTPAdapter(pool_connections=4, pool_maxsize=32)
        self.session.mount("https://", adapter)
        self.session.mount("http://", adapter)

    def call(self, path, body):
        url = f"{self.base}/{path}"
        last = "unknown error"
        for attempt in range(8):
            if self.stop.is_set():
                raise Stopped()
            self.limiter.wait()
            try:
                r = self.session.post(url, json=body, auth=(self.key, ""), timeout=60,
                                      headers={"Accept": "application/json"})
            except requests.RequestException as e:
                last = str(e)
                time.sleep(min(60, 2 ** attempt))
                continue
            try:
                data = r.json()
            except ValueError:
                data = None
            if r.status_code == 429 or (isinstance(data, dict) and data.get("error") == "rate_limited"):
                try:
                    ra = float(r.headers.get("Retry-After", ""))
                except ValueError:
                    ra = None
                self.limiter.throttled(ra)
                last = "rate limited"
                continue
            if r.status_code >= 500:
                last = f"HTTP {r.status_code}"
                time.sleep(min(60, 2 ** attempt))
                continue
            if r.status_code in (401, 403):
                raise FatalError(f"HTTP {r.status_code} from {path}: check the API key / that it can access this property. {r.text[:200]}")
            if isinstance(data, dict) and data.get("ok") is False:
                raise ApiError(f"{path}: {data.get('error', 'unknown')} {data.get('message', '')}".strip())
            if not r.ok:
                raise ApiError(f"HTTP {r.status_code} from {path}: {r.text[:200]}")
            if data is None:
                raise ApiError(f"{path}: reply was not JSON: {r.text[:200]}")
            self.limiter.success()
            return data
        raise ApiError(f"{path}: gave up after 8 attempts ({last})")


def iso(d):
    return d.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%S.000Z")


# --------------------------------------------------------------------------- the job
class Job:
    def __init__(self, data_dir="data", base_url=BASE_URL):
        os.makedirs(data_dir, exist_ok=True)
        self.data_dir = data_dir
        self.db_path = os.path.join(data_dir, "export.db")
        self.base_url = base_url
        self.lock = threading.Lock()
        self.conn = sqlite3.connect(self.db_path, check_same_thread=False)
        self.conn.execute("PRAGMA journal_mode=WAL")
        self.conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS chats(
              id TEXT PRIMARY KEY, started REAL, name TEXT DEFAULT '', email TEXT DEFAULT '',
              status TEXT DEFAULT 'pending', attempts INTEGER DEFAULT 0, error TEXT,
              duration INTEGER, content TEXT);
            CREATE INDEX IF NOT EXISTS ix_status ON chats(status);
            CREATE TABLE IF NOT EXISTS windows(w_start TEXT, w_end TEXT, n INTEGER, PRIMARY KEY(w_start, w_end));
            """
        )
        self.conn.commit()
        self.stop_ev = threading.Event()
        self.thread = None
        self.state = "idle"  # idle | listing | downloading | finished | stopped | error
        self.message = ""
        self.limiter = None
        self.run_started = None
        self.run_done = 0
        self.win_done = 0
        self.win_total = 0

    # ---- public
    def running(self):
        return self.thread is not None and self.thread.is_alive()

    def start(self, key, property_id, start_date, end_date, rate=5.0, workers=8, size=1000):
        if self.running():
            return
        self.stop_ev.clear()
        self.limiter = Limiter(rate)
        self.run_started, self.run_done, self.message = time.time(), 0, ""
        self.state = "listing"
        self.thread = threading.Thread(
            target=self._run, args=(key, property_id, start_date, end_date, int(workers), int(size)), daemon=True)
        self.thread.start()

    def stop(self):
        self.stop_ev.set()

    def retry_failed(self):
        with self.lock:
            self.conn.execute("UPDATE chats SET status='pending', attempts=0, error=NULL WHERE status='failed'")
            self.conn.commit()

    def reset(self):
        if self.running():
            return
        with self.lock:
            self.conn.executescript("DELETE FROM chats; DELETE FROM windows;")
            self.conn.commit()
        self.state, self.message = "idle", ""

    def stats(self):
        with self.lock:
            rows = dict(self.conn.execute("SELECT status, COUNT(*) FROM chats GROUP BY status").fetchall())
        done, pending, failed = rows.get("done", 0), rows.get("pending", 0), rows.get("failed", 0)
        elapsed = (time.time() - self.run_started) if self.run_started else 0
        speed = self.run_done / elapsed if elapsed > 5 and self.run_done else 0.0
        return {
            "state": self.state, "message": self.message, "total": done + pending + failed,
            "done": done, "pending": pending, "failed": failed,
            "windows_done": self.win_done, "windows_total": self.win_total,
            "speed": speed, "limit_rate": self.limiter.rate if self.limiter else 0.0,
            "throttles": self.limiter.throttles if self.limiter else 0,
            "eta_min": (pending / speed / 60) if speed and self.state == "downloading" else None,
        }

    def preview(self, key, property_id, start_date, end_date, size=1000):
        """Fetch ONE chat end to end so field names / duration / content can be eyeballed before the long run."""
        client = Tawk(key, Limiter(2), threading.Event(), self.base_url)
        lo = max(start_date, end_date - dt.timedelta(days=3))
        data = client.call("chat.list", self._list_body(property_id, lo, end_date + dt.timedelta(days=1), size))
        items = find_array(data) or []
        if not items:
            return {"error": "chat.list returned no chats for the last few days of the range", "list_reply_keys": list(data)[:10] if isinstance(data, dict) else str(type(data))}
        item = items[0]
        cid = parse_list_item(item)["id"]
        got = client.call("chat.get", {"propertyId": property_id, "chatId": cid})
        p = parse_chat(got)
        p["chat_id"] = cid
        p["list_item_keys"] = sorted(item) if isinstance(item, dict) else []
        p["get_reply_keys"] = sorted(chat_obj(got)) if isinstance(chat_obj(got), dict) else []
        p["content"] = p["content"][:800]
        return p

    # ---- internals
    @staticmethod
    def _list_body(prop, a, b, size):
        return {"propertyId": prop, "startDate": iso(a), "endDate": iso(b - dt.timedelta(seconds=1)),
                "dateType": "cso", "sort": "co-new-old", "size": size}

    def _commit(self):
        with self.lock:
            self.conn.commit()

    def _run(self, key, prop, start_date, end_date, workers, size):
        try:
            client = Tawk(key, self.limiter, self.stop_ev, self.base_url)
            self._list_phase(client, prop, start_date, end_date, size)
            self.state = "downloading"
            self._detail_phase(client, prop, workers)
            self._commit()
            with self.lock:
                left = self.conn.execute("SELECT COUNT(*) FROM chats WHERE status!='done'").fetchone()[0]
            self.state = "finished"
            self.message = "" if left == 0 else f"{left} chat(s) could not be downloaded - use 'Retry failed'."
        except Stopped:
            self._commit()
            self.state, self.message = "stopped", "Stopped. Press Start to resume where it left off."
        except FatalError as e:
            self._commit()
            self.state, self.message = "error", str(e)
        except Exception as e:  # noqa: BLE001 - surface anything to the UI
            self._commit()
            self.state, self.message = "error", f"{type(e).__name__}: {e}"

    def _list_phase(self, client, prop, start_date, end_date, size):
        # start_date / end_date are dates (inclusive); windows are whole local (Nairobi) days
        days, d = [], start_date
        while d <= end_date:
            a = dt.datetime.combine(d, dt.time.min, tzinfo=TZ)
            days.append((a, a + dt.timedelta(days=1)))
            d += dt.timedelta(days=1)
        self.win_total, self.win_done = len(days), 0
        stack = list(reversed(days))
        while stack:
            a, b = stack.pop()
            ka, kb = iso(a), iso(b)
            with self.lock:
                seen = self.conn.execute("SELECT 1 FROM windows WHERE w_start=? AND w_end=?", (ka, kb)).fetchone()
            if seen:
                self.win_done += 1
                continue
            data = client.call("chat.list", self._list_body(prop, a, b, size))
            items = find_array(data)
            if items is None:
                raise ApiError(f"chat.list: could not find the chat array in the reply (keys: {list(data)[:8] if isinstance(data, dict) else type(data)})")
            if len(items) >= size and (b - a) > dt.timedelta(minutes=2):
                mid = a + (b - a) / 2
                stack.append((mid, b))
                stack.append((a, mid))
                self.win_total += 1  # one window became two
                continue
            rows = [parse_list_item(i) for i in items if isinstance(i, dict)]
            with self.lock:
                self.conn.executemany(
                    "INSERT OR IGNORE INTO chats(id, started, name, email) VALUES (?,?,?,?)",
                    [(r["id"], r["started"], r["name"], r["email"]) for r in rows if r["id"]])
                self.conn.execute("INSERT OR REPLACE INTO windows VALUES (?,?,?)", (ka, kb, len(items)))
                self.conn.commit()
            self.win_done += 1

    def _detail_phase(self, client, prop, workers):
        with self.lock:
            ids = [r[0] for r in self.conn.execute("SELECT id FROM chats WHERE status='pending' ORDER BY started")]
        q = queue.Queue()
        for i in ids:
            q.put(i)
        fatal = []
        counter = {"n": 0}

        def worker():
            while not self.stop_ev.is_set():
                try:
                    cid = q.get_nowait()
                except queue.Empty:
                    return
                try:
                    p = parse_chat(client.call("chat.get", {"propertyId": prop, "chatId": cid}))
                    with self.lock:
                        self.conn.execute(
                            "UPDATE chats SET status='done', error=NULL, duration=?, content=?, "
                            "name=CASE WHEN ?!='' THEN ? ELSE name END, email=CASE WHEN ?!='' THEN ? ELSE email END, "
                            "started=COALESCE(started, ?) WHERE id=?",
                            (p["duration"], p["content"], p["name"], p["name"], p["email"], p["email"], p["started"], cid))
                        counter["n"] += 1
                        if counter["n"] % 100 == 0:
                            self.conn.commit()
                    self.run_done += 1
                except Stopped:
                    return
                except FatalError as e:
                    fatal.append(e)
                    self.stop_ev.set()
                    return
                except ApiError as e:
                    with self.lock:
                        row = self.conn.execute("SELECT attempts FROM chats WHERE id=?", (cid,)).fetchone()
                        n = (row[0] if row else 0) + 1
                        self.conn.execute("UPDATE chats SET attempts=?, error=?, status=? WHERE id=?",
                                          (n, str(e)[:300], "pending" if n < 3 else "failed", cid))
                    if n < 3:
                        q.put(cid)

        threads = [threading.Thread(target=worker, daemon=True) for _ in range(workers)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()
        if fatal:
            raise fatal[0]
        if self.stop_ev.is_set():
            raise Stopped()


# --------------------------------------------------------------------------- Excel
_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def export_xlsx(db_path, out_path):
    """Stream finished chats from SQLite into an .xlsx without holding them in memory."""
    conn = sqlite3.connect(db_path)
    wb = xlsxwriter.Workbook(out_path, {"constant_memory": True, "strings_to_urls": False, "strings_to_formulas": False})
    ws = wb.add_worksheet("Chats")
    f_date = wb.add_format({"num_format": "yyyy-mm-dd hh:mm:ss", "valign": "top"})
    f_dur = wb.add_format({"num_format": "[h]:mm:ss", "valign": "top"})
    f_txt = wb.add_format({"valign": "top"})
    f_head = wb.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1})
    for c, h in enumerate(["Name of client", "Email of client", "Date of chat", "Content of the chat",
                           "Duration (h:mm:ss)", "Content (continued)", "Chat ID"]):
        ws.write_string(0, c, h, f_head)
    ws.set_column(0, 0, 24); ws.set_column(1, 1, 30); ws.set_column(2, 2, 20)
    ws.set_column(3, 3, 90); ws.set_column(4, 4, 16); ws.set_column(5, 5, 40); ws.set_column(6, 6, 26)
    ws.freeze_panes(1, 0)
    r = 0
    cur = conn.execute("SELECT name, email, started, content, duration, id FROM chats WHERE status='done' ORDER BY started, id")
    for name, email, started, content, duration, cid in cur:
        r += 1
        ws.write_string(r, 0, _ILLEGAL.sub("", name or ""), f_txt)
        ws.write_string(r, 1, _ILLEGAL.sub("", email or ""), f_txt)
        if started:
            local = dt.datetime.fromtimestamp(started, tz=UTC).astimezone(TZ).replace(tzinfo=None)
            ws.write_datetime(r, 2, local, f_date)
        text = _ILLEGAL.sub("", content or "")
        ws.write_string(r, 3, text[:CELL_LIMIT], f_txt)
        if duration is not None:
            ws.write_number(r, 4, duration / 86400.0, f_dur)
        if len(text) > CELL_LIMIT:
            rest = text[CELL_LIMIT:]
            ws.write_string(r, 5, rest[:CELL_LIMIT] + (" ...[truncated]" if len(rest) > CELL_LIMIT else ""), f_txt)
        ws.write_string(r, 6, cid, f_txt)
    ws.autofilter(0, 0, max(r, 1), 6)
    bad = conn.execute("SELECT id, started, error FROM chats WHERE status!='done' ORDER BY started").fetchall()
    if bad:
        wf = wb.add_worksheet("Not downloaded")
        for c, h in enumerate(["Chat ID", "Date", "Error"]):
            wf.write_string(0, c, h, f_head)
        for i, (cid, started, err) in enumerate(bad, 1):
            wf.write_string(i, 0, cid)
            if started:
                wf.write_datetime(i, 1, dt.datetime.fromtimestamp(started, tz=UTC).astimezone(TZ).replace(tzinfo=None), f_date)
            wf.write_string(i, 2, _ILLEGAL.sub("", err or "not downloaded yet"))
    wb.close()
    conn.close()
    return r, len(bad)
