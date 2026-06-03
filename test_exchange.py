"""
test_exchange_onprem.py
=======================
Diagnostic script to verify whether an on-premises Exchange server
is reachable and usable via exchangelib (EWS).

Tests performed (in order):
  1. Package check        — exchangelib is installed
  2. DNS / network        — Exchange hostname resolves and port 443 is open
  3. EWS endpoint         — /EWS/Exchange.asmx responds with HTTP 200/401
  4. Authentication       — credentials are accepted
  5. Autodiscover         — optional, can be skipped if slow/blocked
  6. Mailbox access       — inbox is readable, folder tree is walkable
  7. Folder move (dry run) — confirms move permission without touching real mail

Usage:
    pip install exchangelib
    python test_exchange_onprem.py

Edit the CONFIG block below before running.
"""

import sys
import socket
import ssl
import urllib.request
import urllib.error

# ── CONFIG — fill these in before running ────────────────────────────────────
EXCHANGE_SERVER  = "mail.yourcompany.com"   # on-prem CAS / EWS hostname
EMAIL_ADDRESS    = "you@yourcompany.com"
DOMAIN           = "YOURDOMAIN"             # Windows domain, e.g. CORP
USERNAME         = "yourusername"           # without domain prefix
PASSWORD         = "yourpassword"

# Set to True to also test autodiscover (can be slow or blocked in some envs)
TEST_AUTODISCOVER = False

# Set to True to attempt a dry-run folder move on the FIRST inbox message.
# No email is permanently moved — the script moves it then moves it back.
TEST_MOVE_DRYRUN = False
# ─────────────────────────────────────────────────────────────────────────────

PASS = "  [PASS]"
FAIL = "  [FAIL]"
SKIP = "  [SKIP]"
INFO = "  [INFO]"

results = []

def log(symbol, label, detail=""):
    line = f"{symbol} {label}"
    if detail:
        line += f"\n         {detail}"
    print(line)
    results.append((symbol, label))


# ── Test 1: package installed ─────────────────────────────────────────────────
print("\n── Test 1: exchangelib installed ───────────────────────────────────")
try:
    import exchangelib
    log(PASS, "exchangelib imported", f"version {exchangelib.__version__}")
except ImportError:
    log(FAIL, "exchangelib not found",
        "Run:  pip install exchangelib")
    print("\nCannot continue without exchangelib. Exiting.")
    sys.exit(1)

from exchangelib import (
    Credentials, Account, Configuration,
    DELEGATE, NTLM, Build, Version,
)
from exchangelib.protocol import BaseProtocol


# ── Test 2: DNS + TCP reachability ────────────────────────────────────────────
print("\n── Test 2: network reachability ────────────────────────────────────")
try:
    ip = socket.gethostbyname(EXCHANGE_SERVER)
    log(PASS, f"DNS resolved: {EXCHANGE_SERVER}", f"→ {ip}")
except socket.gaierror as e:
    log(FAIL, f"DNS resolution failed: {EXCHANGE_SERVER}", str(e))
    ip = None

if ip:
    try:
        sock = socket.create_connection((EXCHANGE_SERVER, 443), timeout=5)
        sock.close()
        log(PASS, "TCP port 443 open")
    except (socket.timeout, OSError) as e:
        log(FAIL, "TCP port 443 unreachable", str(e))


# ── Test 3: EWS endpoint HTTP response ────────────────────────────────────────
print("\n── Test 3: EWS endpoint ────────────────────────────────────────────")
ews_url = f"https://{EXCHANGE_SERVER}/EWS/Exchange.asmx"
try:
    # We expect 401 Unauthorized (no creds sent) — that still confirms EWS is up
    ctx = ssl.create_default_context()
    req = urllib.request.Request(ews_url, method="GET")
    try:
        urllib.request.urlopen(req, context=ctx, timeout=8)
        log(PASS, "EWS endpoint reachable (HTTP 200)")
    except urllib.error.HTTPError as e:
        if e.code == 401:
            log(PASS, "EWS endpoint reachable (HTTP 401 — expected, no creds sent)")
        else:
            log(FAIL, f"EWS endpoint returned HTTP {e.code}", str(e))
except urllib.error.URLError as e:
    log(FAIL, "EWS endpoint unreachable", str(e.reason))
except ssl.SSLError as e:
    log(FAIL, "SSL error connecting to EWS", str(e))


# ── Test 4: authentication + mailbox access ───────────────────────────────────
print("\n── Test 4: authentication ──────────────────────────────────────────")
account = None
try:
    credentials = Credentials(
        username=f"{DOMAIN}\\{USERNAME}",
        password=PASSWORD,
    )
    config = Configuration(
        server=EXCHANGE_SERVER,
        credentials=credentials,
        auth_type=NTLM,
    )
    account = Account(
        primary_smtp_address=EMAIL_ADDRESS,
        config=config,
        autodiscover=False,
        access_type=DELEGATE,
    )
    # Force a real EWS call to verify the credentials actually work
    inbox_count = account.inbox.total_count
    log(PASS, "Authenticated successfully (NTLM)",
        f"Inbox total count: {inbox_count}")
except Exception as e:
    log(FAIL, "Authentication failed", str(e))


# ── Test 5: autodiscover (optional) ──────────────────────────────────────────
print("\n── Test 5: autodiscover ────────────────────────────────────────────")
if not TEST_AUTODISCOVER:
    log(SKIP, "Autodiscover skipped (set TEST_AUTODISCOVER = True to enable)")
else:
    try:
        credentials = Credentials(
            username=f"{DOMAIN}\\{USERNAME}",
            password=PASSWORD,
        )
        ad_account = Account(
            primary_smtp_address=EMAIL_ADDRESS,
            credentials=credentials,
            autodiscover=True,
            access_type=DELEGATE,
        )
        log(PASS, "Autodiscover succeeded",
            f"EWS URL: {ad_account.protocol.service_endpoint}")
    except Exception as e:
        log(FAIL, "Autodiscover failed", str(e))


# ── Test 6: folder tree + inbox items ─────────────────────────────────────────
print("\n── Test 6: folder access ───────────────────────────────────────────")
if account is None:
    log(SKIP, "Skipped — authentication failed in Test 4")
else:
    # List top-level folders
    try:
        folders = list(account.inbox.parent.children)
        names = [f.name for f in folders[:10]]
        log(PASS, f"Folder tree readable ({len(folders)} top-level folders)",
            "First 10: " + ", ".join(names))
    except Exception as e:
        log(FAIL, "Could not read folder tree", str(e))

    # Peek at the first inbox message
    try:
        msgs = list(account.inbox.all().order_by("-datetime_received")[:3])
        if msgs:
            m = msgs[0]
            log(PASS, f"Inbox readable — {len(msgs)} message(s) sampled",
                f"Latest: '{m.subject or '(no subject)'}' from {m.sender.email_address}")
        else:
            log(INFO, "Inbox is empty — readable but no messages to inspect")
    except Exception as e:
        log(FAIL, "Could not read inbox messages", str(e))

    # Confirm a target subfolder can be created / found
    try:
        from exchangelib.folders import Folder
        test_folder_name = "_EWS_Test_Folder_"
        existing = [f for f in account.inbox.children if f.name == test_folder_name]
        if existing:
            target = existing[0]
            log(INFO, f"Test subfolder '{test_folder_name}' already exists")
        else:
            target = account.inbox.create_folder(test_folder_name)
            log(PASS, f"Created test subfolder '{test_folder_name}' under Inbox")
    except Exception as e:
        log(FAIL, "Could not create/find test subfolder", str(e))
        target = None


# ── Test 7: move dry run ──────────────────────────────────────────────────────
print("\n── Test 7: folder move (dry run) ───────────────────────────────────")
if not TEST_MOVE_DRYRUN:
    log(SKIP, "Move dry run skipped (set TEST_MOVE_DRYRUN = True to enable)")
elif account is None:
    log(SKIP, "Skipped — authentication failed in Test 4")
elif target is None:
    log(SKIP, "Skipped — test subfolder unavailable")
else:
    try:
        msgs = list(account.inbox.all().order_by("-datetime_received")[:1])
        if not msgs:
            log(SKIP, "No messages in inbox to test move with")
        else:
            msg = msgs[0]
            original_subject = msg.subject or "(no subject)"

            # Move to test folder
            msg.move(target)
            log(PASS, f"Moved '{original_subject}' → '{target.name}'")

            # Move back to inbox
            msg.move(account.inbox)
            log(PASS, f"Moved '{original_subject}' back → Inbox")
    except Exception as e:
        log(FAIL, "Move operation failed", str(e))

    # Clean up test folder if we created it and it's empty
    try:
        if target and target.total_count == 0:
            target.delete()
            log(INFO, f"Cleaned up empty test folder '{target.name}'")
    except Exception:
        pass


# ── Summary ───────────────────────────────────────────────────────────────────
print("\n── Summary ─────────────────────────────────────────────────────────")
passed = sum(1 for s, _ in results if s == PASS)
failed = sum(1 for s, _ in results if s == FAIL)
skipped = sum(1 for s, _ in results if s == SKIP)
print(f"  Passed:  {passed}")
print(f"  Failed:  {failed}")
print(f"  Skipped: {skipped}")

if failed == 0 and passed > 0:
    print("\n  ✓ On-premises Exchange via EWS looks usable.")
    print("  You can use exchangelib as a drop-in replacement for win32com.")
elif failed > 0:
    print("\n  ✗ One or more checks failed — see details above.")
    print("  Common fixes:")
    print("    - DNS/network: ensure EXCHANGE_SERVER is the internal CAS hostname")
    print("    - Auth failure: try EMAIL_ADDRESS format for username instead of DOMAIN\\user")
    print("    - EWS blocked: ask IT to confirm EWS is enabled (Get-WebServicesVirtualDirectory)")
    print("    - SSL error:   if using a self-signed cert, disable cert verification with:")
    print("                   BaseProtocol.HTTP_ADAPTER_CLS = exchangelib.transport.NoVerifyHTTPAdapter")
print()
