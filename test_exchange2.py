"""
test_exchange_onprem.py
=======================
Diagnostic script to verify whether an on-premises Exchange server
is reachable and usable via exchangelib (EWS).

Tests performed (in order):
  1. Package check        — exchangelib is installed
  2. DNS / network        — Exchange hostname resolves and port 443 is open
  3. EWS endpoint         — /EWS/Exchange.asmx responds with HTTP 200/401
  4. Authentication       — tries all username formats and auth types automatically
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

# Set to True to disable SSL certificate verification.
# Use this if your Exchange server uses a self-signed or internal CA cert.
DISABLE_SSL_VERIFY = False
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
    log(FAIL, "exchangelib not found", "Run:  pip install exchangelib")
    print("\nCannot continue without exchangelib. Exiting.")
    sys.exit(1)

from exchangelib import (
    Credentials, Account, Configuration,
    DELEGATE, NTLM, BASIC, GSSAPI,
)
from exchangelib.protocol import BaseProtocol

# Apply SSL override early so all subsequent exchangelib calls respect it
if DISABLE_SSL_VERIFY:
    import exchangelib.transport as _t
    BaseProtocol.HTTP_ADAPTER_CLS = _t.NoVerifyHTTPAdapter
    print(f"{INFO} SSL certificate verification disabled (DISABLE_SSL_VERIFY=True)")


# ── Test 2: DNS + TCP reachability ────────────────────────────────────────────
print("\n── Test 2: network reachability ────────────────────────────────────")
ip = None
try:
    ip = socket.gethostbyname(EXCHANGE_SERVER)
    log(PASS, f"DNS resolved: {EXCHANGE_SERVER}", f"→ {ip}")
except socket.gaierror as e:
    log(FAIL, f"DNS resolution failed: {EXCHANGE_SERVER}", str(e))

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
    ctx = ssl.create_default_context()
    if DISABLE_SSL_VERIFY:
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
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
    log(FAIL, "SSL error — try setting DISABLE_SSL_VERIFY = True", str(e))


# ── Test 4: authentication — tries all username formats + auth types ───────────
print("\n── Test 4: authentication ──────────────────────────────────────────")
print(f"{INFO} Trying all username formats and auth types automatically...\n")

# All username formats Exchange servers commonly accept
username_formats = [
    (f"{DOMAIN}\\{USERNAME}",  "DOMAIN\\username  (classic NTLM)"),
    (EMAIL_ADDRESS,             "email address     (UPN / Office365 style)"),
    (USERNAME,                  "bare username     (some internal configs)"),
    (f"{USERNAME}@{DOMAIN}",   "username@DOMAIN   (UPN variant)"),
]

# Auth types to try in order
auth_types = [
    (NTLM,  "NTLM"),
    (BASIC, "Basic"),
]

account = None
winning_combo = None

for auth_type, auth_label in auth_types:
    if account:
        break
    for uname, uname_label in username_formats:
        if account:
            break
        try:
            print(f"  Trying: auth={auth_label:6s}  username={uname_label}")
            creds = Credentials(username=uname, password=PASSWORD)
            config = Configuration(
                server=EXCHANGE_SERVER,
                credentials=creds,
                auth_type=auth_type,
            )
            acct = Account(
                primary_smtp_address=EMAIL_ADDRESS,
                config=config,
                autodiscover=False,
                access_type=DELEGATE,
            )
            # Force a real EWS call — Account creation alone doesn't authenticate
            inbox_count = acct.inbox.total_count
            account = acct
            winning_combo = (uname, auth_label)
            print()
            log(PASS,
                f"Authenticated successfully",
                f"auth_type={auth_label}  username={uname_label}\n"
                f"         Inbox total count: {inbox_count}"
            )
        except Exception as e:
            err = str(e)
            # Trim long tracebacks to just the meaningful part
            short_err = err.splitlines()[-1] if "\n" in err else err
            print(f"           → {short_err}")

if not account:
    print()
    log(FAIL, "All username format + auth type combinations failed",
        "Check DOMAIN / USERNAME / PASSWORD in the CONFIG block.\n"
        "         If tests 2-3 passed but auth fails, EWS may require Kerberos:\n"
        "           pip install exchangelib[kerberos]  then set auth_type=GSSAPI\n"
        "         Or SSL cert is untrusted — set DISABLE_SSL_VERIFY = True to test."
    )


# ── Test 5: autodiscover (optional) ──────────────────────────────────────────
print("\n── Test 5: autodiscover ────────────────────────────────────────────")
if not TEST_AUTODISCOVER:
    log(SKIP, "Autodiscover skipped (set TEST_AUTODISCOVER = True to enable)")
else:
    try:
        uname = winning_combo[0] if winning_combo else f"{DOMAIN}\\{USERNAME}"
        creds = Credentials(username=uname, password=PASSWORD)
        ad_account = Account(
            primary_smtp_address=EMAIL_ADDRESS,
            credentials=creds,
            autodiscover=True,
            access_type=DELEGATE,
        )
        log(PASS, "Autodiscover succeeded",
            f"EWS URL: {ad_account.protocol.service_endpoint}")
    except Exception as e:
        log(FAIL, "Autodiscover failed", str(e))


# ── Test 6: folder tree + inbox items ─────────────────────────────────────────
print("\n── Test 6: folder access ───────────────────────────────────────────")
target = None
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

    # Peek at the most recent inbox messages
    try:
        msgs = list(account.inbox.all().order_by("-datetime_received")[:3])
        if msgs:
            m = msgs[0]
            log(PASS, f"Inbox readable — {len(msgs)} message(s) sampled",
                f"Latest: '{m.subject or '(no subject)'}' "
                f"from {m.sender.email_address if m.sender else 'unknown'}")
        else:
            log(INFO, "Inbox is empty — readable but no messages to inspect")
    except Exception as e:
        log(FAIL, "Could not read inbox messages", str(e))

    # Confirm a subfolder can be created
    try:
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
            msg.move(target)
            log(PASS, f"Moved '{original_subject}' → '{target.name}'")
            msg.move(account.inbox)
            log(PASS, f"Moved '{original_subject}' back → Inbox")
    except Exception as e:
        log(FAIL, "Move operation failed", str(e))

    # Clean up empty test folder
    try:
        if target and target.total_count == 0:
            target.delete()
            log(INFO, f"Cleaned up empty test folder '{target.name}'")
    except Exception:
        pass


# ── Summary ───────────────────────────────────────────────────────────────────
print("\n── Summary ─────────────────────────────────────────────────────────")
passed  = sum(1 for s, _ in results if s == PASS)
failed  = sum(1 for s, _ in results if s == FAIL)
skipped = sum(1 for s, _ in results if s == SKIP)
print(f"  Passed:  {passed}")
print(f"  Failed:  {failed}")
print(f"  Skipped: {skipped}")

if winning_combo:
    print(f"\n  Working combination:")
    print(f"    auth_type : {winning_combo[1]}")
    print(f"    username  : {winning_combo[0]}")
    print(f"\n  Use these values in your InboxSorter config.")

if failed == 0 and passed > 0:
    print("\n  ✓ On-premises Exchange via EWS is usable.")
    print("  exchangelib can replace win32com in InboxSorter.")
elif failed > 0 and passed < 3:
    print("\n  ✗ Connection or auth failed. Things to check with IT:")
    print("    1. Is EWS enabled?   Get-WebServicesVirtualDirectory | fl InternalUrl,ExternalUrl")
    print("    2. Is your account allowed EWS access?   Get-CASMailbox you@domain | fl EwsEnabled")
    print("    3. Is there an EWS throttling policy blocking the connection?")
    print("    4. Does the server use Kerberos only?  pip install exchangelib[kerberos]")
print()
