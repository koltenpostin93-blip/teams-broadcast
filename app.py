import streamlit as st
import msal
import requests
from requests.exceptions import Timeout, ConnectionError as ReqConnError
import json
import base64
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from supabase import create_client

CLIENT_ID = "95aa84e5-44b3-4233-94f6-25ca740aff4d"
TENANT_ID = "4a4f2e28-2f12-4cdb-b5eb-9860e3af1045"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = [
    "https://graph.microsoft.com/Chat.ReadWrite",
    "https://graph.microsoft.com/User.Read",
]
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
REQUEST_TIMEOUT = 30  # seconds per Graph API call


# ── Supabase ──────────────────────────────────────────────────────────────────

@st.cache_resource
def get_supabase():
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])


def sb_get(key, default=None):
    try:
        res = get_supabase().table("app_data").select("value").eq("key", key).execute()
        if res.data:
            return json.loads(res.data[0]["value"])
    except Exception as e:
        st.error(f"Supabase read error ({key}): {e}")
    return default


def sb_set(key, value):
    try:
        get_supabase().table("app_data").upsert({"key": key, "value": json.dumps(value)}).execute()
    except Exception as e:
        st.error(f"Supabase write error ({key}): {e}")


def sb_delete(key):
    try:
        get_supabase().table("app_data").delete().eq("key", key).execute()
    except Exception:
        pass


# ── Profile helpers ───────────────────────────────────────────────────────────

def load_profiles():
    return sb_get("profiles", [])


def save_profiles(profiles):
    sb_set("profiles", profiles)


# ── Token cache ───────────────────────────────────────────────────────────────

def load_cache(profile):
    cache = msal.SerializableTokenCache()
    data = sb_get(f"token_{profile}")
    if data:
        cache.deserialize(json.dumps(data))
    return cache


def save_cache(cache, profile):
    if cache.has_state_changed:
        sb_set(f"token_{profile}", json.loads(cache.serialize()))


def delete_cache(profile):
    sb_delete(f"token_{profile}")


def build_app(cache=None):
    return msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)


# ── Auth helpers ──────────────────────────────────────────────────────────────

def get_cached_token(profile):
    cache = load_cache(profile)
    app = build_app(cache)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            save_cache(cache, profile)
            return result["access_token"]
    return None


def refresh_token(profile):
    token = get_cached_token(profile)
    if token:
        st.session_state.token = token
    return token


def start_device_flow(profile):
    cache = load_cache(profile)
    app = build_app(cache)
    return app.initiate_device_flow(scopes=SCOPES)


def finish_device_flow(flow_dict, profile):
    cache = load_cache(profile)
    app = build_app(cache)
    result = app.acquire_token_by_device_flow(flow_dict)
    if "access_token" in result:
        save_cache(cache, profile)
        return result["access_token"], None
    err = result.get("error_description") or result.get("error") or "Unknown error"
    return None, err


# ── Graph API ─────────────────────────────────────────────────────────────────

def headers(token):
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def fetch_chats(token):
    chats, url = [], f"{GRAPH_BASE}/me/chats?$top=50"
    while url:
        resp = requests.get(url, headers=headers(token), timeout=REQUEST_TIMEOUT)
        if resp.status_code != 200:
            break
        data = resp.json()
        filtered = [
            c for c in data.get("value", [])
            if c.get("chatType") != "meeting" and c.get("topic")
        ]
        chats.extend(filtered)
        url = data.get("@odata.nextLink")
    return chats


def send_message(token, chat_id, message, images=None):
    url = f"{GRAPH_BASE}/me/chats/{chat_id}/messages"
    msg_html = message.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br>")

    if images:
        hosted_contents = []
        img_tags = ""
        for i, (img_bytes, img_mime) in enumerate(images, start=1):
            b64 = base64.b64encode(img_bytes).decode()
            hosted_contents.append({
                "@microsoft.graph.temporaryId": str(i),
                "contentBytes": b64,
                "contentType": img_mime or "image/png",
            })
            img_tags += f'<img src="../hostedContents/{i}/$value" style="max-width:800px;"><br>'

        html_body = f"{msg_html}<br>{img_tags}" if msg_html.strip() else img_tags
        body = {
            "body": {"contentType": "html", "content": html_body},
            "hostedContents": hosted_contents,
        }
    else:
        body = {"body": {"contentType": "html", "content": msg_html}}

    try:
        resp = requests.post(url, headers=headers(token), json=body, timeout=REQUEST_TIMEOUT)
    except Timeout:
        return False, "Request timed out"
    except ReqConnError:
        return False, "Connection error"

    if resp.status_code == 201:
        return True, None
    try:
        err = resp.json().get("error", {}).get("message", resp.text)
    except Exception:
        err = resp.text
    return False, f"[{resp.status_code}] {err}"


# ── Groups ────────────────────────────────────────────────────────────────────

def load_groups(profile):
    return sb_get(f"groups_{profile}", {"subgroups": {}, "hidden": []})


def save_groups(groups, profile):
    sb_set(f"groups_{profile}", groups)


# ── Display name for a chat ───────────────────────────────────────────────────

def chat_label(chat):
    if chat.get("topic"):
        return chat["topic"]
    members = chat.get("members", [])
    names = [m.get("displayName", "") for m in members if m.get("displayName")]
    if names:
        return ", ".join(names[:4]) + (" ..." if len(names) > 4 else "")
    return chat["id"][:24] + "..."


# ── App ───────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Teams Broadcast", page_icon="📣", layout="wide")
st.title("📣 Teams Broadcast Tool")

st.markdown("""
<style>
textarea { spellcheck: true; }
</style>
<script>
document.querySelectorAll('textarea').forEach(t => t.spellcheck = true);
</script>
""", unsafe_allow_html=True)

# ── Profile selector ──────────────────────────────────────────────────────────

if "profile" not in st.session_state:
    st.session_state.profile = None

if not st.session_state.profile:
    st.subheader("Who are you?")
    profiles = load_profiles()

    if profiles:
        col1, col2 = st.columns([2, 1])
        with col1:
            selected_profile = st.selectbox("Select your profile", profiles)
        with col1:
            if st.button("Continue", type="primary"):
                st.session_state.profile = selected_profile
                st.rerun()

    st.divider()
    st.write("**Add a new profile**")
    new_profile = st.text_input("Enter your name (e.g. Kolten)")
    if st.button("Create Profile") and new_profile.strip():
        name = new_profile.strip()
        if name not in profiles:
            profiles.append(name)
            save_profiles(profiles)
        st.session_state.profile = name
        st.rerun()

    st.stop()

profile = st.session_state.profile

# ── Session state init (per profile) ─────────────────────────────────────────

if "token" not in st.session_state:
    st.session_state.token = get_cached_token(profile)
if "flow_data" not in st.session_state:
    st.session_state.flow_data = None
if "chats" not in st.session_state:
    st.session_state.chats = []
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0
if "message_key" not in st.session_state:
    st.session_state.message_key = 0

# ── Sign-in screen ────────────────────────────────────────────────────────────

if not st.session_state.token:
    st.subheader(f"Sign in — {profile}")

    if st.session_state.flow_data is None:
        if st.button("Start Sign In", type="primary"):
            flow = start_device_flow(profile)
            st.session_state.flow_data = flow
            st.rerun()
    else:
        flow = st.session_state.flow_data
        st.markdown("**Step 1:** Open [microsoft.com/devicelogin](https://microsoft.com/devicelogin) in your browser")
        st.markdown(f"**Step 2:** Enter this code: &nbsp; `{flow['user_code']}`")
        st.markdown("**Step 3:** Sign in with your work account, then click the button below")

        if st.button("I've signed in — continue", type="primary"):
            with st.spinner("Completing sign in..."):
                token, err = finish_device_flow(flow, profile)
            if token:
                st.session_state.token = token
                st.session_state.flow_data = None
                st.rerun()
            else:
                st.error(f"Sign in failed: {err}")
                st.session_state.flow_data = None

    st.stop()

# ── Main app (authenticated) ──────────────────────────────────────────────────

token = st.session_state.token

if not st.session_state.chats:
    with st.spinner("Loading your Teams chats..."):
        st.session_state.chats = fetch_chats(token)

groups = load_groups(profile)
hidden_ids = set(groups.get("hidden", []))
chats = [c for c in st.session_state.chats if c["id"] not in hidden_ids]
chat_lookup = {c["id"]: chat_label(c) for c in chats}

# Show active profile in sidebar
with st.sidebar:
    st.markdown(f"**Signed in as:** {profile}")
    st.caption(f"Groups key: `groups_{profile}`")
    raw = sb_get(f"groups_{profile}")
    st.caption(f"Subgroups found: {list(raw['subgroups'].keys()) if raw else 'None'}")
    if st.button("Switch User"):
        st.session_state.clear()
        st.rerun()

tab_broadcast, tab_groups, tab_settings = st.tabs(["📣 Broadcast", "👥 Manage Groups", "⚙️ Settings"])

# ── Broadcast tab ─────────────────────────────────────────────────────────────

with tab_broadcast:
    st.subheader("Compose & Send")

    message = st.text_area("Message", height=180, placeholder="Type your message here...",
                           key=f"message_{st.session_state.message_key}")

    uploaded_files = st.file_uploader(
        "Attach images (optional)",
        type=["png", "jpg", "jpeg", "gif", "bmp", "webp"],
        accept_multiple_files=True,
        help="Select or drag & drop one or more images.",
        key=f"uploader_{st.session_state.uploader_key}"
    )

    if uploaded_files:
        cols = st.columns(min(len(uploaded_files), 4))
        for i, f in enumerate(uploaded_files):
            cols[i % 4].image(f, use_container_width=True)

    group_options = ["— All Chats —"] + sorted(groups.get("subgroups", {}).keys())
    selected = st.selectbox("Send to", group_options)

    if selected == "— All Chats —":
        target_ids = [c["id"] for c in chats]
    else:
        target_ids = groups["subgroups"].get(selected, [])

    with st.expander(f"Refine recipients ({len(target_ids)} selected)", expanded=False):
        st.caption("Uncheck any chats to skip them for this send only — does not modify the saved group.")
        valid_ids = [cid for cid in target_ids if cid in chat_lookup]
        target_ids = st.multiselect(
            "Recipients",
            options=valid_ids,
            default=valid_ids,
            format_func=lambda x: chat_lookup.get(x, x),
            label_visibility="collapsed",
            key=f"refine_{st.session_state.message_key}",
        )

    st.caption(f"{len(target_ids)} chat(s) selected")

    has_content = message.strip() or uploaded_files
    send_clicked = st.button("Send Message", type="primary", disabled=not has_content)

    if send_clicked:
        if not target_ids:
            st.warning("No chats in the selected group.")
        else:
            images = [(f.read(), f.type) for f in uploaded_files] if uploaded_files else None

            fresh_token = refresh_token(profile) or token

            bar = st.progress(0, text="Sending...")
            success, failed, done = 0, 0, 0
            total = len(target_ids)
            errors = []

            with ThreadPoolExecutor(max_workers=8) as executor:
                futures = {
                    executor.submit(send_message, fresh_token, cid, message, images): cid
                    for cid in target_ids
                }
                for future in as_completed(futures):
                    done += 1
                    ok, err = future.result()
                    if ok:
                        success += 1
                    else:
                        failed += 1
                        chat_name = chat_lookup.get(futures[future], futures[future])
                        errors.append(f"**{chat_name}**: {err}")
                    bar.progress(done / total, text=f"Sending... {done}/{total}")

            bar.empty()
            if failed == 0:
                st.success(f"Sent to all {success} chats successfully.")
                st.session_state.uploader_key += 1
                st.session_state.message_key += 1
                st.rerun()
            else:
                st.warning(f"Sent: {success}  |  Failed: {failed}")
                with st.expander("Show errors"):
                    for e in errors:
                        st.markdown(e)

# ── Manage Groups tab ─────────────────────────────────────────────────────────

with tab_groups:
    st.subheader("Subgroups")

    col_left, col_right = st.columns([1, 2])

    with col_left:
        st.write("**Create new subgroup**")
        new_name = st.text_input("Group name", key="new_name")
        if st.button("Create") and new_name.strip():
            if new_name not in groups["subgroups"]:
                groups["subgroups"][new_name] = []
                save_groups(groups, profile)
                st.success(f'Created "{new_name}"')
                st.rerun()
            else:
                st.warning("A group with that name already exists.")

        if groups["subgroups"]:
            st.write("**Delete subgroup**")
            del_target = st.selectbox("Select group", list(groups["subgroups"].keys()), key="del_target")
            if st.button("Delete", type="secondary"):
                del groups["subgroups"][del_target]
                save_groups(groups, profile)
                st.rerun()

    with col_right:
        if groups["subgroups"]:
            st.write("**Edit chats in a subgroup**")
            edit_target = st.selectbox("Select group to edit", list(groups["subgroups"].keys()), key="edit_target")
            current = groups["subgroups"].get(edit_target, [])

            chosen = st.multiselect(
                "Chats in this group",
                options=list(chat_lookup.keys()),
                default=[cid for cid in current if cid in chat_lookup],
                format_func=lambda x: chat_lookup.get(x, x),
            )

            if st.button("Save Changes", type="primary"):
                groups["subgroups"][edit_target] = chosen
                save_groups(groups, profile)
                st.success("Saved.")
        else:
            st.info("Create a subgroup on the left to get started.")

    st.divider()
    st.subheader("🚫 Hidden Chats")
    st.caption("Chats hidden here won't appear in broadcasts or groups.")

    all_chat_lookup = {c["id"]: chat_label(c) for c in st.session_state.chats}
    current_hidden = groups.get("hidden", [])

    new_hidden = st.multiselect(
        "Select chats to hide",
        options=list(all_chat_lookup.keys()),
        default=[cid for cid in current_hidden if cid in all_chat_lookup],
        format_func=lambda x: all_chat_lookup.get(x, x),
    )

    if st.button("Save Hidden List", type="primary"):
        groups["hidden"] = new_hidden
        save_groups(groups, profile)
        st.success("Hidden list saved. Refresh the page to see the updated chat list.")

# ── Settings tab ──────────────────────────────────────────────────────────────

with tab_settings:
    st.subheader("Settings")

    if st.button("Refresh chat list"):
        with st.spinner("Reloading..."):
            st.session_state.chats = fetch_chats(token)
        st.success(f"Loaded {len(st.session_state.chats)} chats.")

    st.divider()

    if st.button("Sign out", type="secondary"):
        delete_cache(profile)
        st.session_state.clear()
        st.rerun()
