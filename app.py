import datetime as dt
import os

import streamlit as st

from tawk_export import Job, export_xlsx

st.set_page_config(page_title="tawk.to chat export", page_icon="💬")


@st.cache_resource
def get_job():
    # one job for the whole server, so it keeps running when the browser tab is closed
    return Job(os.environ.get("DATA_DIR", "data"))


job = get_job()

# optional password (set APP_PASSWORD as an environment variable on hcode) - the *.hcode.site URL is public
pw = os.environ.get("APP_PASSWORD", "")
if pw and not st.session_state.get("authed"):
    entered = st.text_input("Password", type="password")
    if entered:
        if entered == pw:
            st.session_state["authed"] = True
            st.rerun()
        else:
            st.error("Wrong password.")
    st.stop()

st.title("tawk.to chat export")
st.caption("Downloads every chat in the date range and builds an Excel file: name, email, date, content, duration.")

key = st.text_input("tawk.to API key", value=os.environ.get("TAWK_API_KEY", ""), type="password")
prop = st.text_input("Property ID", value=os.environ.get("TAWK_PROPERTY_ID", ""))
c1, c2 = st.columns(2)
d1 = c1.date_input("From", dt.date(2026, 1, 1))
d2 = c2.date_input("To (inclusive)", dt.date(2026, 11, 12))
with st.expander("Speed settings"):
    rate = st.number_input("Target requests per second", 0.5, 20.0, 5.5, 0.5,
                           help="92,271 chats in 5 hours needs about 5.2/s. The app slows itself down automatically if tawk.to answers 429.")
    workers = st.number_input("Parallel workers", 1, 32, 8)
    size = st.number_input("Chats per list call (size)", 50, 1000, 1000, 50)

running = job.running()
b1, b2, b3, b4 = st.columns(4)
if b1.button("▶ Start / Resume", type="primary", disabled=running):
    if not key or not prop:
        st.error("Enter the API key and the Property ID.")
    elif d1 > d2:
        st.error("'From' is after 'To'.")
    else:
        job.start(key.strip(), prop.strip(), d1, d2, rate, workers, size)
        st.rerun()
if b2.button("⏹ Stop", disabled=not running):
    job.stop()
if b3.button("🧪 Test 1 chat", disabled=running):
    if not key or not prop:
        st.error("Enter the API key and the Property ID.")
    else:
        with st.spinner("Fetching one chat..."):
            try:
                st.json(job.preview(key.strip(), prop.strip(), d1, d2, size))
            except Exception as e:  # noqa: BLE001
                st.error(f"{type(e).__name__}: {e}")
if b4.button("↻ Retry failed", disabled=running):
    job.retry_failed()


@st.fragment(run_every=3)
def progress():
    s = job.stats()
    total = max(s["total"], 1)
    st.progress(min(s["done"] / total, 1.0), text=f"{s['done']:,} of {s['total']:,} chats downloaded")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Status", s["state"])
    m2.metric("Speed", f"{s['speed']:.1f}/s")
    m3.metric("Allowed rate", f"{s['limit_rate']:.1f}/s" if s["limit_rate"] else "-")
    m4.metric("ETA", f"{s['eta_min'] / 60:.1f} h" if s["eta_min"] else "-")
    if s["state"] == "listing":
        st.write(f"Listing chats: day windows {s['windows_done']} of {s['windows_total']}, {s['total']:,} chats found so far.")
    st.caption(f"Pending {s['pending']:,} · failed {s['failed']:,} · rate-limit hits {s['throttles']}")
    if s["message"]:
        (st.error if s["state"] == "error" else st.info)(s["message"])


progress()

st.divider()
st.subheader("Excel file")
st.caption("You can build it at any time; it contains whatever has been downloaded so far.")
if st.button("Build Excel", disabled=running):
    path = os.path.join(job.data_dir, f"chats_{d1}_{d2}.xlsx")
    with st.spinner("Writing Excel..."):
        rows, bad = export_xlsx(job.db_path, path)
    st.session_state["xlsx"] = (path, rows, bad)
if st.session_state.get("xlsx"):
    path, rows, bad = st.session_state["xlsx"]
    st.success(f"{rows:,} chats in the file" + (f" ({bad:,} not downloaded, listed on a second sheet)." if bad else "."))
    with open(path, "rb") as f:
        st.download_button("⬇ Download Excel", f, file_name=os.path.basename(path),
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
