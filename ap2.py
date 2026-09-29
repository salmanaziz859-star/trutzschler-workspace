import os
import sys
import re
import urllib.request

# =====================================================================
# TRUTZSCHLER WORKSPACE — LAUNCHER (Cross-Platform)
# =====================================================================
# Ye file sirf ek "launcher" hai — isay hi PyInstaller se .exe / .app
# banaya jata hai, aur ye kabhi bhi rebuild nahi karni parti jab tak
# koi bahut bara structural change na ho.
#
# Asal software ka poora code "core_app.py" mein hai, jo GitHub par
# rehta hai. Jab bhi ye launcher chalta hai, ye sabse pehle GitHub se
# "core_app.py" ka LATEST version download karta hai, apne data folder
# mein save karta hai, aur usi ko chala deta hai.
#
# Matlab: aap sirf GitHub par "core_app.py" ko edit/commit karo — koi
# naya .exe ya .app banane ki zaroorat nahi. Agli baar koi bhi client
# software kholega, usay khud-ba-khud naya code mil jayega.
# =====================================================================

# --- Data folder wahi hai jo core app bhi use karta hai ---
if sys.platform == "win32":
    APP_DATA_DIR = "C:\\TrutzschlerData"
else:
    APP_DATA_DIR = os.path.join(os.path.expanduser("~"), "TrutzschlerData")

if not os.path.exists(APP_DATA_DIR):
    os.makedirs(APP_DATA_DIR)

CORE_APP_PATH = os.path.join(APP_DATA_DIR, "core_app.py")

# NOTE: Repo PUBLIC honi chahiye taake ye link bina kisi token/login ke
# khul sake. Agar kabhi repo ka naam ya branch badle, ye line update karna.
CORE_APP_URL = "https://raw.githubusercontent.com/salmanaziz859-star/trutzschler-workspace/main/core_app.py"


def extract_version(code_text):
    """core_app.py ke text se CURRENT_VERSION = "x.x" wali line nikaal kar
    sirf version number wapas karta hai. Agar na mile to None."""
    if not code_text:
        return None
    match = re.search(r'CURRENT_VERSION\s*=\s*["\']([^"\']+)["\']', code_text)
    return match.group(1) if match else None


def fetch_latest_code():
    """GitHub se sabse naya core_app.py download karke local cache update karta hai.
    Agar internet na ho ya GitHub down ho, to purana cached wala code chal jayega
    (agar pehle kabhi download ho chuka ho) — is se app kabhi bhi 'internet nahi hai
    isliye bilkul nahi chalegi' wali halat mein nahi jati.

    Returns: (old_version, new_version) — taake caller ye decide kar sake ke
    'Software updated' wala popup dikhana hai ya nahi."""
    old_version = None
    if os.path.exists(CORE_APP_PATH):
        try:
            with open(CORE_APP_PATH, "r", encoding="utf-8") as f:
                old_version = extract_version(f.read())
        except Exception:
            pass

    new_version = old_version  # default: agar download fail ho to "no change" maana jaye

    try:
        req = urllib.request.Request(CORE_APP_URL, headers={"Cache-Control": "no-cache"})
        with urllib.request.urlopen(req, timeout=6) as response:
            new_code = response.read().decode('utf-8')

        # Sirf tab overwrite karo jab download sahi/mukammal laga ho (khaali/corrupt
        # response se purana kaam karta hua code kabhi zaya na ho)
        if new_code and len(new_code.strip()) > 100:
            with open(CORE_APP_PATH, "w", encoding="utf-8") as f:
                f.write(new_code)
            new_version = extract_version(new_code)
            print("[Launcher] Latest code GitHub se mil gaya.")
        else:
            print("[Launcher] GitHub se khaali response mila — purana cached code use ho raha hai.")

    except Exception as e:
        print(f"[Launcher] Update check skipped (offline mode): {e}")
        print("[Launcher] Purana cached code (agar mojood hai) use ho raha hai.")

    return old_version, new_version


def show_update_popup(old_version, new_version):
    """Sirf tab popup dikhata hai jab: pehle se koi version cached tha (first
    install par popup nahi aana chahiye) AUR version waqai badla ho."""
    if old_version and new_version and old_version != new_version:
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()  # Khaali background window na dikhe
            messagebox.showinfo(
                "Software Updated",
                f"Software update ho gaya hai!\n\nPurana Version: v{old_version}\nNaya Version: v{new_version}"
            )
            root.destroy()
        except Exception as e:
            print(f"[Launcher] Update popup dikhane mein masla: {e}")


def run_core_app():
    if not os.path.exists(CORE_APP_PATH):
        print("=" * 70)
        print("ERROR: Koi bhi app code mojood nahi hai — na cached, na online.")
        print("Internet connection check karein aur software dobara chalayein.")
        print("=" * 70)
        input("\nPress Enter to exit...")
        sys.exit(1)

    with open(CORE_APP_PATH, "r", encoding="utf-8") as f:
        code = f.read()

    # Core app ko isi process ke andar, ek nayi/saaf namespace mein chalate hain
    # taake __name__ == "__main__" wala hissa bhi sahi se trigger ho.
    exec(compile(code, CORE_APP_PATH, 'exec'), {"__name__": "__main__", "__file__": CORE_APP_PATH})


if __name__ == "__main__":
    old_ver, new_ver = fetch_latest_code()
    show_update_popup(old_ver, new_ver)
    run_core_app()
