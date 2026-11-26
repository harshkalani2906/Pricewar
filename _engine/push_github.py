"""
PUSH TO GITHUB - uploads the Price War tracker (scripts + 1_Input Excel files) to your
GitHub repository. Browser logins (_engine/state), results (2_Output) and the built
dashboard (3_Dashboard) are NOT uploaded (see .gitignore).

    python push_github.py          -> push now (asks for the repository link the first time)
    python push_github.py setup    -> change the repository link
"""
import os
import shutil
import subprocess
import sys
from datetime import datetime

ENGINE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(ENGINE)
URL_FILE = os.path.join(ENGINE, "github_repo.txt")


def run(args, check=False, timeout=180):
    try:
        r = subprocess.run(args, cwd=ROOT, capture_output=True, text=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        if check:
            raise RuntimeError("timed out (check internet)")
        return 124, "timed out (check internet)"
    out = (r.stdout + r.stderr).strip()
    if check and r.returncode != 0:
        raise RuntimeError(out.splitlines()[-1] if out else "git error")
    return r.returncode, out


def ask_url():
    print("\n  1. Open https://github.com/new in your browser (logged in as your GitHub account).")
    print("  2. Repository name: price-war   |   choose  Private   |   do NOT tick 'Add a README'.")
    print("  3. Click 'Create repository' and copy the link from the address bar.\n")
    url = input("  Paste the repository link here and press Enter: ").strip().rstrip("/")
    if not url.startswith("https://github.com/"):
        raise SystemExit("  That doesn't look like a GitHub link - run this again and paste the link from the address bar.")
    if not url.endswith(".git"):
        url += ".git"
    with open(URL_FILE, "w", encoding="utf-8") as fh:
        fh.write(url)
    return url


def main():
    print("=" * 66)
    print("  PRICE WAR - push to GitHub")
    print("=" * 66)
    if shutil.which("git") is None:
        raise SystemExit("\n  Git is not installed. Install it from https://git-scm.com/download/win "
                         "(keep every option on default), then run this again.")
    url = None
    if len(sys.argv) < 2 or sys.argv[1] != "setup":
        try:
            url = open(URL_FILE, encoding="utf-8").read().strip() or None
        except OSError:
            pass
    url = url or ask_url()

    if not os.path.isdir(os.path.join(ROOT, ".git")):
        run(["git", "init"], check=True)
        run(["git", "checkout", "-B", "main"], check=True)
    if run(["git", "config", "user.name"])[0] != 0:
        run(["git", "config", "user.name", "Price War Tracker"])
        run(["git", "config", "user.email", "price-war@localhost"])
    if run(["git", "remote", "get-url", "origin"])[0] == 0:
        run(["git", "remote", "set-url", "origin", url], check=True)
    else:
        run(["git", "remote", "add", "origin", url], check=True)

    run(["git", "add", "-A"], check=True)
    code, out = run(["git", "commit", "-m", f"Price War update {datetime.now():%d-%m-%Y %H:%M}"])
    if code != 0 and "nothing to commit" not in out:
        raise SystemExit(f"\n  Could not save the changes: {out.splitlines()[-1] if out else ''}")

    print("\n  Uploading to GitHub ... (a browser window may ask you to log in - do it once)")
    code, out = run(["git", "push", "-u", "origin", "main"])
    if code != 0 and ("rejected" in out or "fetch first" in out):
        run(["git", "pull", "origin", "main", "--allow-unrelated-histories", "--no-rebase", "--no-edit", "-X", "ours"])
        code, out = run(["git", "push", "-u", "origin", "main"])
    if code == 0:
        print(f"\n  DONE - uploaded to {url[:-4]}")
    else:
        print(f"\n  ! Upload failed: {out.splitlines()[-1] if out else 'unknown error'}")
        print("    Send a screenshot of this window to Claude.")


def auto_push(log=print):
    """Called after every dashboard rebuild: uploads whatever changed (dashboard, Excel inputs,
    scripts). Silent no-op until PUSH TO GITHUB.bat has been run once."""
    try:
        if shutil.which("git") is None or not os.path.isdir(os.path.join(ROOT, ".git")):
            return
        if not os.path.exists(URL_FILE):
            return
        run(["git", "add", "-A"], timeout=60)
        code, out = run(["git", "commit", "-m", f"Auto update {datetime.now():%d-%m-%Y %H:%M}"], timeout=60)
        if code != 0 and "nothing to commit" in out:
            return
        code, out = run(["git", "push", "origin", "main"])
        if code != 0 and ("rejected" in out or "fetch first" in out):
            run(["git", "pull", "origin", "main", "--no-rebase", "--no-edit", "-X", "ours"])
            code, out = run(["git", "push", "origin", "main"])
        log("GitHub: changes uploaded" if code == 0 else
            f"GitHub: upload failed ({out.splitlines()[-1] if out else 'unknown'}) - will retry next time")
    except Exception as e:
        log(f"GitHub: upload skipped ({e})")


if __name__ == "__main__":
    main()
