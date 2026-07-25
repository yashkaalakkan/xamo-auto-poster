import os
import io
import json
import time
import sys
import base64
import smtplib
from email.mime.text import MIMEText
from datetime import datetime, timezone, timedelta

import requests
import openpyxl
from google.oauth2 import service_account
from google.oauth2.credentials import Credentials as UserCredentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

from nacl import encoding, public

# ---------------------------------------------------------------------------
# Configuration (from environment variables / GitHub Secrets)
# ---------------------------------------------------------------------------

IG_USER_ID = os.environ["IG_USER_ID"]
IG_ACCESS_TOKEN = os.environ["IG_ACCESS_TOKEN"]
print(f"DEBUG: IG token length={len(IG_ACCESS_TOKEN)}, starts='{IG_ACCESS_TOKEN[:10]}', ends='{IG_ACCESS_TOKEN[-10:]}', repr_sample={repr(IG_ACCESS_TOKEN[:20])}")
FB_APP_ID = os.environ["FB_APP_ID"]
FB_APP_SECRET = os.environ["FB_APP_SECRET"]

FB_PAGE_ID = os.environ["FB_PAGE_ID"]
FB_PAGE_ACCESS_TOKEN = os.environ["FB_PAGE_ACCESS_TOKEN"]

YT_CLIENT_ID = os.environ["YT_CLIENT_ID"]
YT_CLIENT_SECRET = os.environ["YT_CLIENT_SECRET"]
YT_REFRESH_TOKEN = os.environ["YT_REFRESH_TOKEN"]

GOOGLE_DRIVE_CREDENTIALS_JSON = os.environ["GOOGLE_DRIVE_CREDENTIALS_JSON"]
DRIVE_FOLDER_ID = os.environ["DRIVE_FOLDER_ID"]  # the Xamo_AutoPoster folder

HEALTHCHECK_PING_URL = os.environ["HEALTHCHECK_PING_URL"]

# --- Alerting ---
GMAIL_ADDRESS = os.environ["GMAIL_ADDRESS"]
GMAIL_APP_PASSWORD = os.environ["GMAIL_APP_PASSWORD"]
ALERT_EMAIL_TO = os.environ.get("ALERT_EMAIL_TO", GMAIL_ADDRESS)

CALLMEBOT_PHONE = os.environ["CALLMEBOT_PHONE"]
CALLMEBOT_APIKEY = os.environ["CALLMEBOT_APIKEY"]

# --- GitHub (for auto-updating the IG token secret after refresh) ---
GH_PAT = os.environ["GH_PAT"]
GH_REPO = os.environ["GITHUB_REPOSITORY"]  # auto-provided by Actions, e.g. "user/xamo-auto-poster"

# YouTube quota safety settings: 5 uploads/day, spaced >=4hrs apart
YT_MAX_UPLOADS_PER_DAY = 5
YT_MIN_HOURS_BETWEEN_POSTS = 4

# IG long-lived token: refresh proactively well before the real 60-day expiry
IG_TOKEN_REFRESH_AFTER_DAYS = 50

# A platform gets marked "skipped_error" (and stops being retried) after this many failures
MAX_FAILURES_PER_PLATFORM = 3

QUEUE_FILENAME = "queue.xlsx"
YT_QUOTA_FILENAME = "yt_quota.json"
TOKEN_STATE_FILENAME = "token_state.json"
VIDEOS_SUBFOLDER = "videos"
POSTED_SUBFOLDER = "posted"

LOCAL_WORKDIR = "/tmp/xamo_poster"
os.makedirs(LOCAL_WORKDIR, exist_ok=True)

GRAPH_API_VERSION = "v21.0"


def log(msg):
    print(f"[{datetime.now(timezone.utc).isoformat()}] {msg}", flush=True)


# ---------------------------------------------------------------------------
# Alerting: email primary, WhatsApp (CallMeBot) failsafe
# ---------------------------------------------------------------------------

def send_email_alert(subject, body):
    msg = MIMEText(body)
    msg["Subject"] = f"[Xamo Auto Poster] {subject}"
    msg["From"] = GMAIL_ADDRESS
    msg["To"] = ALERT_EMAIL_TO

    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=20) as server:
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_ADDRESS, [ALERT_EMAIL_TO], msg.as_string())


def send_whatsapp_alert(message):
    resp = requests.get(
        "https://api.callmebot.com/whatsapp.php",
        params={"phone": CALLMEBOT_PHONE, "text": message, "apikey": CALLMEBOT_APIKEY},
        timeout=20,
    )
    resp.raise_for_status()


def send_alert(subject, body):
    """Email first. If email fails for any reason, fall back to WhatsApp so
    you still hear about it. Never lets an alerting failure crash the run."""
    try:
        send_email_alert(subject, body)
        log(f"Alert sent via email: {subject}")
        return
    except Exception as e:
        log(f"Email alert failed ({e}), falling back to WhatsApp...")

    try:
        send_whatsapp_alert(f"{subject}: {body}"[:1000])
        log(f"Alert sent via WhatsApp fallback: {subject}")
    except Exception as e:
        log(f"WhatsApp fallback alert ALSO failed ({e}). No alert could be delivered.")


# ---------------------------------------------------------------------------
# Google Drive helpers
# ---------------------------------------------------------------------------

def get_drive_service():
    creds_dict = json.loads(GOOGLE_DRIVE_CREDENTIALS_JSON)
    creds = service_account.Credentials.from_service_account_info(
        creds_dict, scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds)


def find_child_folder(drive, parent_id, name):
    query = (
        f"'{parent_id}' in parents and name = '{name}' "
        "and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    )
    results = drive.files().list(q=query, fields="files(id, name)").execute()
    files = results.get("files", [])
    if not files:
        raise RuntimeError(f"Could not find subfolder '{name}' inside folder {parent_id}")
    return files[0]["id"]


def find_file_in_folder(drive, folder_id, name):
    query = f"'{folder_id}' in parents and name = '{name}' and trashed = false"
    results = drive.files().list(q=query, fields="files(id, name)").execute()
    files = results.get("files", [])
    return files[0]["id"] if files else None


def download_file(drive, file_id, local_path):
    request = drive.files().get_media(fileId=file_id)
    with io.FileIO(local_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
    return local_path


def upload_or_replace_file(drive, folder_id, filename, local_path, existing_file_id=None):
    media = MediaFileUpload(local_path, resumable=True)
    if existing_file_id:
        drive.files().update(fileId=existing_file_id, media_body=media).execute()
        return existing_file_id
    else:
        file_metadata = {"name": filename, "parents": [folder_id]}
        created = drive.files().create(body=file_metadata, media_body=media, fields="id").execute()
        return created["id"]


def move_file(drive, file_id, from_folder_id, to_folder_id):
    drive.files().update(
        fileId=file_id,
        addParents=to_folder_id,
        removeParents=from_folder_id,
        fields="id, parents",
    ).execute()


def load_json_state(drive, folder_id, filename, default):
    file_id = find_file_in_folder(drive, folder_id, filename)
    if not file_id:
        return dict(default), None
    local_path = os.path.join(LOCAL_WORKDIR, filename)
    download_file(drive, file_id, local_path)
    with open(local_path, "r") as f:
        return json.load(f), file_id


def save_json_state(drive, folder_id, filename, data, file_id):
    local_path = os.path.join(LOCAL_WORKDIR, filename)
    with open(local_path, "w") as f:
        json.dump(data, f)
    return upload_or_replace_file(drive, folder_id, filename, local_path, existing_file_id=file_id)


# ---------------------------------------------------------------------------
# Queue (queue.xlsx) helpers
# ---------------------------------------------------------------------------

COLUMNS = [
    "ID", "VideoFilename", "Caption", "YT_Title", "Hashtags",
    "IG_Status", "FB_Status", "YT_Status",
    "IG_Fails", "FB_Fails", "YT_Fails",
    "IG_PostedAt", "FB_PostedAt", "YT_PostedAt", "YT_FirstPendingAt",
]

DONE_STATES = ("posted", "skipped_error")


def col_index(name):
    return COLUMNS.index(name) + 1  # openpyxl is 1-indexed


def load_queue(local_path):
    wb = openpyxl.load_workbook(local_path)
    ws = wb.active
    return wb, ws


def get_row_dict(ws, row_num):
    return {col: ws.cell(row=row_num, column=col_index(col)).value for col in COLUMNS}


def set_cell(ws, row_num, col_name, value):
    ws.cell(row=row_num, column=col_index(col_name), value=value)


def find_next_pending_row(ws):
    """First row where at least one platform hasn't reached a DONE state yet."""
    for row_num in range(2, ws.max_row + 1):
        row = get_row_dict(ws, row_num)
        if row["VideoFilename"] is None:
            continue
        statuses = [row["IG_Status"], row["FB_Status"], row["YT_Status"]]
        if any(s not in DONE_STATES for s in statuses):
            return row_num
    return None


def row_fully_posted(ws, row_num):
    """True only if ALL three actually posted (not just done-via-skip)."""
    row = get_row_dict(ws, row_num)
    return all(row[s] == "posted" for s in ["IG_Status", "FB_Status", "YT_Status"])


def record_platform_result(ws, row_num, platform_prefix, success, row_num_video, row_num_video_name):
    pass  # placeholder not used; logic handled inline in main() for clarity


# ---------------------------------------------------------------------------
# IG long-lived token auto-refresh + GitHub secret update
# ---------------------------------------------------------------------------

def refresh_ig_token(current_token):
    """Exchange the current token for a fresh 60-day one, using the
    correct endpoint for Instagram-Login-issued tokens (IGAA...)."""
    resp = requests.get(
        "https://graph.instagram.com/refresh_access_token",
        params={
            "grant_type": "ig_refresh_token",
            "access_token": current_token,
        },
        timeout=30,
    ).json()

    if "access_token" not in resp:
        raise RuntimeError(f"IG token refresh failed: {resp}")

    return resp["access_token"]


def encrypt_secret_for_github(public_key_b64, secret_value):
    pk = public.PublicKey(public_key_b64.encode("utf-8"), encoding.Base64Encoder())
    sealed_box = public.SealedBox(pk)
    encrypted = sealed_box.encrypt(secret_value.encode("utf-8"))
    return base64.b64encode(encrypted).decode("utf-8")


def update_github_secret(secret_name, secret_value):
    headers = {
        "Authorization": f"Bearer {GH_PAT}",
        "Accept": "application/vnd.github+json",
    }
    key_resp = requests.get(
        f"https://api.github.com/repos/{GH_REPO}/actions/secrets/public-key",
        headers=headers, timeout=20,
    ).json()
    if "key" not in key_resp:
        raise RuntimeError(f"Could not fetch GitHub public key: {key_resp}")

    encrypted_value = encrypt_secret_for_github(key_resp["key"], secret_value)

    put_resp = requests.put(
        f"https://api.github.com/repos/{GH_REPO}/actions/secrets/{secret_name}",
        headers=headers,
        json={"encrypted_value": encrypted_value, "key_id": key_resp["key_id"]},
        timeout=20,
    )
    if put_resp.status_code not in (201, 204):
        raise RuntimeError(f"Failed to update GitHub secret {secret_name}: {put_resp.status_code} {put_resp.text}")


def maybe_refresh_ig_token(drive):
    """Refreshes the IG token if it's due, updates the GitHub secret, and
    alerts (email -> WhatsApp fallback) if the refresh itself fails.
    Never crashes the run - if refresh fails, we just keep using the old
    token and try again next run."""
    state, state_file_id = load_json_state(
        drive, DRIVE_FOLDER_ID, TOKEN_STATE_FILENAME, {"ig_last_refreshed": None}
    )

    last_refreshed = state.get("ig_last_refreshed")
    due = True
    if last_refreshed:
        last_dt = datetime.fromisoformat(last_refreshed)
        due = (datetime.now(timezone.utc) - last_dt) >= timedelta(days=IG_TOKEN_REFRESH_AFTER_DAYS)

    if not due:
        return IG_ACCESS_TOKEN

    log("IG token refresh is due - attempting refresh...")
    try:
        new_token = refresh_ig_token(IG_ACCESS_TOKEN)
        update_github_secret("IG_ACCESS_TOKEN", new_token)
        state["ig_last_refreshed"] = datetime.now(timezone.utc).isoformat()
        save_json_state(drive, DRIVE_FOLDER_ID, TOKEN_STATE_FILENAME, state, state_file_id)
        log("IG token refreshed successfully and GitHub secret updated.")
        return new_token
    except Exception as e:
        log(f"IG TOKEN REFRESH FAILED: {e}")
        send_alert(
            "IG token refresh FAILED",
            f"The automated Instagram token refresh failed: {e}\n\n"
            f"The bot will keep using the existing token for now, but it may "
            f"expire soon. Please refresh IG_ACCESS_TOKEN manually if this "
            f"keeps happening.",
        )
        return IG_ACCESS_TOKEN  # keep using the old one, don't crash the run


# ---------------------------------------------------------------------------
# Instagram posting
# ---------------------------------------------------------------------------

def upload_temp_github_release_asset(local_video_path, filename):
    """Uploads the video as an asset on a dedicated 'video-hosting' GitHub
    Release, and returns its public download URL. Creates the release if
    it doesn't already exist."""
    headers = {"Authorization": f"Bearer {GH_PAT}", "Accept": "application/vnd.github+json"}

    releases_resp = requests.get(
        f"https://api.github.com/repos/{GH_REPO}/releases/tags/video-hosting",
        headers=headers, timeout=20,
    )
    if releases_resp.status_code == 200:
        release = releases_resp.json()
    else:
        create_resp = requests.post(
            f"https://api.github.com/repos/{GH_REPO}/releases",
            headers=headers,
            json={"tag_name": "video-hosting", "name": "Temporary video hosting", "draft": False, "prerelease": True},
            timeout=20,
        )
        release = create_resp.json()

    upload_url = release["upload_url"].split("{")[0]

    with open(local_video_path, "rb") as f:
        upload_resp = requests.post(
            f"{upload_url}?name={filename}",
            headers={**headers, "Content-Type": "video/mp4"},
            data=f,
            timeout=120,
        )
    upload_resp.raise_for_status()
    asset = upload_resp.json()
    return asset["browser_download_url"], asset["id"]


def delete_github_release_asset(asset_id):
    headers = {"Authorization": f"Bearer {GH_PAT}", "Accept": "application/vnd.github+json"}
    requests.delete(
        f"https://api.github.com/repos/{GH_REPO}/releases/assets/{asset_id}",
        headers=headers, timeout=20,
    )


def post_to_instagram(video_local_path, caption, access_token):
    base_url = "https://graph.instagram.com"

    video_url, asset_id = upload_temp_github_release_asset(video_local_path, os.path.basename(video_local_path))
    log(f"IG: temporarily hosted video at {video_url}")

    try:
        create_resp = requests.post(
            f"{base_url}/{IG_USER_ID}/media",
            data={
                "media_type": "REELS",
                "video_url": video_url,
                "caption": caption,
                "access_token": access_token,
            },
        ).json()

        if "id" not in create_resp:
            log(f"IG: failed to create container: {create_resp}")
            return False

        container_id = create_resp["id"]

        status = None
        for _ in range(30):
            time.sleep(10)
            status_resp = requests.get(
                f"{base_url}/{container_id}",
                params={"fields": "status_code", "access_token": access_token},
            ).json()
            status = status_resp.get("status_code")
            if status == "FINISHED":
                break
            if status == "ERROR":
                log(f"IG: container processing failed: {status_resp}")
                return False

        if status != "FINISHED":
            log(f"IG: container did not finish processing in time (last status: {status})")
            return False

        publish_resp = requests.post(
            f"{base_url}/{IG_USER_ID}/media_publish",
            data={"creation_id": container_id, "access_token": access_token},
        ).json()

        if "id" not in publish_resp:
            log(f"IG: publish failed: {publish_resp}")
            return False

        log(f"IG: published successfully, media id {publish_resp['id']}")
        return True

    finally:
        delete_github_release_asset(asset_id)
        log("IG: cleaned up temporary hosted video.")

# ---------------------------------------------------------------------------
# Facebook Page posting
# ---------------------------------------------------------------------------

def post_to_facebook_page(video_local_path, caption):
    base_url = f"https://graph-video.facebook.com/{GRAPH_API_VERSION}"

    with open(video_local_path, "rb") as f:
        files = {"source": f}
        data = {"description": caption, "access_token": FB_PAGE_ACCESS_TOKEN}
        resp = requests.post(f"{base_url}/{FB_PAGE_ID}/videos", data=data, files=files).json()

    if "id" not in resp:
        log(f"FB: post failed: {resp}")
        return False

    log(f"FB: posted successfully, video id {resp['id']}")
    return True


# ---------------------------------------------------------------------------
# YouTube posting + quota
# ---------------------------------------------------------------------------

def get_youtube_service():
    creds = UserCredentials(
        token=None,
        refresh_token=YT_REFRESH_TOKEN,
        client_id=YT_CLIENT_ID,
        client_secret=YT_CLIENT_SECRET,
        token_uri="https://oauth2.googleapis.com/token",
        scopes=["https://www.googleapis.com/auth/youtube.upload"],
    )
    return build("youtube", "v3", credentials=creds)


def post_to_youtube(video_local_path, yt_title, caption):
    try:
        youtube = get_youtube_service()
        body = {
            "snippet": {
                "title": yt_title[:100],
                "description": caption,
                "categoryId": "22",
            },
            "status": {"privacyStatus": "public", "selfDeclaredMadeForKids": False},
        }
        media = MediaFileUpload(video_local_path, chunksize=-1, resumable=True, mimetype="video/mp4")
        request = youtube.videos().insert(part="snippet,status", body=body, media_body=media)
        response = None
        while response is None:
            status, response = request.next_chunk()
        log(f"YT: uploaded successfully, video id {response['id']}")
        return True
    except Exception as e:
        log(f"YT: upload failed: {e}")
        return False


def yt_batch_timing_allowed(quota_data):
    """Only checks the 4-hour spacing timer - NOT the daily count.
    Quota count is checked separately per-video during the batch."""
    today_str = datetime.now(timezone.utc).date().isoformat()
    if quota_data.get("date") != today_str:
        quota_data["date"] = today_str
        quota_data["count"] = 0

    last_post_time = quota_data.get("last_post_time")
    if last_post_time:
        last_dt = datetime.fromisoformat(last_post_time)
        if datetime.now(timezone.utc) - last_dt < timedelta(hours=YT_MIN_HOURS_BETWEEN_POSTS):
            return False, quota_data
    return True, quota_data


def yt_record_upload(quota_data):
    quota_data["count"] = quota_data.get("count", 0) + 1
    quota_data["last_post_time"] = datetime.now(timezone.utc).isoformat()
    return quota_data


# ---------------------------------------------------------------------------
# Healthcheck ping
# ---------------------------------------------------------------------------

def ping_healthcheck(success=True, message=""):
    try:
        url = HEALTHCHECK_PING_URL if success else f"{HEALTHCHECK_PING_URL}/fail"
        requests.post(url, data=message.encode("utf-8"), timeout=10)
    except Exception as e:
        log(f"Healthcheck ping failed (non-fatal): {e}")


# ---------------------------------------------------------------------------
# Per-platform attempt helper (handles retry counting + skip-after-N)
# ---------------------------------------------------------------------------

def attempt_platform(ws, row_num, row, status_col, fails_col, posted_at_col,
                      platform_label, post_fn):
    """Runs post_fn() if this platform isn't already done. Updates status,
    failure count, and posted-at timestamp. Sends an alert + marks
    skipped_error if MAX_FAILURES_PER_PLATFORM is reached."""
    if row[status_col] in DONE_STATES:
        return False  # nothing to do, not attempted

    log(f"Attempting {platform_label} post...")
    success = post_fn()

    if success:
        set_cell(ws, row_num, status_col, "posted")
        set_cell(ws, row_num, posted_at_col, datetime.now(timezone.utc).isoformat())
        set_cell(ws, row_num, fails_col, 0)
    else:
        fails = (row[fails_col] or 0) + 1
        set_cell(ws, row_num, fails_col, fails)
        if fails >= MAX_FAILURES_PER_PLATFORM:
            set_cell(ws, row_num, status_col, "skipped_error")
            log(f"{platform_label}: hit {fails} failures on row {row_num}, marking skipped_error.")
            send_alert(
                f"{platform_label} post skipped after {fails} failures",
                f"Row {row_num} ({row['VideoFilename']}) failed to post to "
                f"{platform_label} {fails} times in a row and has been marked "
                f"'skipped_error'. It will NOT be retried automatically. "
                f"Check the row in queue.xlsx and fix/reset it manually.",
            )
        else:
            set_cell(ws, row_num, status_col, "failed")

    return True  # was attempted


def find_video_anywhere(drive, videos_folder_id, posted_folder_id, filename):
    """Looks in /videos first, then /posted, since IG+FB may have already
    moved the file by the time YouTube's batch gets to it."""
    file_id = find_file_in_folder(drive, videos_folder_id, filename)
    if file_id:
        return file_id
    return find_file_in_folder(drive, posted_folder_id, filename)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

YT_BACKLOG_MAX_DAYS = 3  # give up on a video for YT after this many days waiting


def run():
    log("Starting poster run.")
    drive = get_drive_service()

    ig_access_token = maybe_refresh_ig_token(drive)

    videos_folder_id = find_child_folder(drive, DRIVE_FOLDER_ID, VIDEOS_SUBFOLDER)
    posted_folder_id = find_child_folder(drive, DRIVE_FOLDER_ID, POSTED_SUBFOLDER)

    queue_file_id = find_file_in_folder(drive, DRIVE_FOLDER_ID, QUEUE_FILENAME)
    if not queue_file_id:
        log(f"ERROR: could not find {QUEUE_FILENAME} in the Drive folder.")
        ping_healthcheck(success=False, message="queue.xlsx not found")
        sys.exit(1)

    local_queue_path = os.path.join(LOCAL_WORKDIR, QUEUE_FILENAME)
    download_file(drive, queue_file_id, local_queue_path)
    wb, ws = load_queue(local_queue_path)

    any_attempted = False

    # ---------------- IG + FB: drives the queue forward, unaffected by YT ----------------
    pending_rows = []
    for row_num in range(2, ws.max_row + 1):
        row = get_row_dict(ws, row_num)
        if row["VideoFilename"] is None:
            continue
        if row["IG_Status"] not in DONE_STATES or row["FB_Status"] not in DONE_STATES:
            pending_rows.append(row_num)

    log(f"Found {len(pending_rows)} row(s) needing IG/FB work: {pending_rows}")

    for row_num in pending_rows:
        row = get_row_dict(ws, row_num)
        video_filename = row["VideoFilename"]
        caption = row["Caption"] or ""

        video_file_id = find_file_in_folder(drive, videos_folder_id, video_filename)
        if not video_file_id:
            log(f"ERROR: video file '{video_filename}' not found in /videos. Skipping row {row_num}.")
            continue

        local_video_path = os.path.join(LOCAL_WORKDIR, video_filename)
        download_file(drive, video_file_id, local_video_path)

        any_attempted |= attempt_platform(
            ws, row_num, row, "IG_Status", "IG_Fails", "IG_PostedAt", "Instagram",
            lambda: post_to_instagram(local_video_path, caption, ig_access_token),
        )
        any_attempted |= attempt_platform(
            ws, row_num, row, "FB_Status", "FB_Fails", "FB_PostedAt", "Facebook",
            lambda: post_to_facebook_page(local_video_path, caption),
        )

        row_after = get_row_dict(ws, row_num)
        if row_after["IG_Status"] == "posted" and row_after["FB_Status"] == "posted":
            log(f"Row {row_num}: IG+FB done - moving to /posted.")
            move_file(drive, video_file_id, videos_folder_id, posted_folder_id)

    # ---------------- YouTube: batch, every ~4 hours, across ALL eligible rows ----------------
    quota_data, quota_file_id = load_json_state(
        drive, DRIVE_FOLDER_ID, YT_QUOTA_FILENAME, {"date": None, "count": 0, "last_post_time": None}
    )
    allowed_by_time, quota_data = yt_batch_timing_allowed(quota_data)

    if allowed_by_time:
        log("YT batch window is open - scanning for all eligible rows.")
        yt_eligible_rows = []
        for row_num in range(2, ws.max_row + 1):
            row = get_row_dict(ws, row_num)
            if row["VideoFilename"] is None:
                continue
            if row["YT_Status"] not in DONE_STATES:
                yt_eligible_rows.append(row_num)

        # Stamp YT_FirstPendingAt for any row seeing this for the first time
        now_iso = datetime.now(timezone.utc).isoformat()
        for row_num in yt_eligible_rows:
            if ws.cell(row=row_num, column=col_index("YT_FirstPendingAt")).value is None:
                set_cell(ws, row_num, "YT_FirstPendingAt", now_iso)

        remaining_quota = YT_MAX_UPLOADS_PER_DAY - quota_data.get("count", 0)
        log(f"YT eligible rows: {yt_eligible_rows}. Remaining quota today: {remaining_quota}")

        for row_num in yt_eligible_rows:
            row = get_row_dict(ws, row_num)
            video_filename = row["VideoFilename"]
            yt_title = row["YT_Title"] or video_filename
            caption = row["Caption"] or ""

            # Give up if this row has been waiting too long
            first_pending = ws.cell(row=row_num, column=col_index("YT_FirstPendingAt")).value
            if first_pending:
                age_days = (datetime.now(timezone.utc) - datetime.fromisoformat(first_pending)).days
                if age_days >= YT_BACKLOG_MAX_DAYS:
                    set_cell(ws, row_num, "YT_Status", "skipped_error")
                    log(f"Row {row_num}: YT backlog exceeded {YT_BACKLOG_MAX_DAYS} days - giving up.")
                    send_alert(
                        "YouTube video skipped (backlog too old)",
                        f"Row {row_num} ({video_filename}) waited {age_days} days for a YouTube "
                        f"upload slot and has been marked 'skipped_error'. It will not be retried.",
                    )
                    continue

            if remaining_quota <= 0:
                set_cell(ws, row_num, "YT_Status", "waiting_quota")
                log(f"Row {row_num}: daily YT quota exhausted, marked 'waiting_quota'.")
                continue

            video_file_id = find_video_anywhere(drive, videos_folder_id, posted_folder_id, video_filename)
            if not video_file_id:
                log(f"ERROR: video file '{video_filename}' not found anywhere for YT. Skipping row {row_num}.")
                continue

            local_video_path = os.path.join(LOCAL_WORKDIR, video_filename)
            download_file(drive, video_file_id, local_video_path)

            def yt_post_and_record():
                ok = post_to_youtube(local_video_path, yt_title, caption)
                if ok:
                    nonlocal quota_data, remaining_quota
                    quota_data = yt_record_upload(quota_data)
                    remaining_quota -= 1
                return ok

            attempted = attempt_platform(
                ws, row_num, row, "YT_Status", "YT_Fails", "YT_PostedAt", "YouTube",
                yt_post_and_record,
            )
            if attempted:
                any_attempted = True

        quota_data["last_post_time"] = datetime.now(timezone.utc).isoformat()
        save_json_state(drive, DRIVE_FOLDER_ID, YT_QUOTA_FILENAME, quota_data, quota_file_id)
    else:
        log("YT batch window not open yet this run (spacing).")

    wb.save(local_queue_path)
    upload_or_replace_file(drive, DRIVE_FOLDER_ID, QUEUE_FILENAME, local_queue_path, existing_file_id=queue_file_id)

    if any_attempted:
        ping_healthcheck(success=True, message="Run completed with activity")
    else:
        ping_healthcheck(success=True, message="Nothing needed posting this run")

    log("Run complete.")


def main():
    """Top-level wrapper: guarantees that ANY unhandled error still pings
    healthchecks.io as failed and sends an alert, instead of silently dying
    with no trace (which is exactly what would break the keep-alive commit
    and leave you finding out weeks late)."""
    try:
        run()
    except Exception as e:
        log(f"FATAL ERROR in run(): {e}")
        ping_healthcheck(success=False, message=f"Fatal error: {e}")
        try:
            send_alert("Xamo Auto Poster run CRASHED", f"The bot run failed with an unhandled error:\n\n{e}")
        except Exception:
            log("Could not send crash alert either.")
        sys.exit(1)


if __name__ == "__main__":
    main()
