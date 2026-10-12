# Entry point for the main deployment, the app people open. Streamlit Community Cloud won't host two apps from
# the same repo, branch and file, and the original deployment (the approved Salesforce callback) runs
# draw_dashboard.py. Importing draw_dashboard, instead of running it, skips its OAuth relay block, so this app
# always finishes Salesforce logins itself.
#
# Streamlit reloads an imported module only if someone has the app open when a push changes it, so a push that
# lands while nobody is on the app would leave it running the old code until a reboot. Reload draw_dashboard
# whenever the file on disk isn't the one in memory. Its caches (and logins in progress) carry over. The lock keeps
# two sessions that start at the same moment from pulling the module out from under each other.
import os
import sys
import threading

_lock = sys.__dict__.setdefault("_draw_dashboard_reload_lock", threading.Lock())     # one per server process
with _lock:
    _stamp = os.stat(os.path.join(os.path.dirname(os.path.abspath(__file__)), "draw_dashboard.py")).st_mtime_ns
    _loaded = sys.modules.get("draw_dashboard")
    if _loaded is not None and getattr(_loaded, "_loaded_stamp", None) != _stamp:
        sys.modules.pop("draw_dashboard", None)
    try:
        import draw_dashboard  # noqa: E402
    except KeyError:                    # Streamlit's own reloader dropped it mid-import in another session: once more
        import draw_dashboard  # noqa: E402
    draw_dashboard._loaded_stamp = _stamp
draw_dashboard.main()
