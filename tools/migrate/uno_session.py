"""A headless LibreOffice session for the migration and its verification.

Run these scripts with LibreOffice's bundled Python, which carries the `uno` module:

    /opt/libreoffice26.2/program/python tools/migrate/migrate.py ...
"""

import os
import subprocess
import tempfile
import time

import uno
from com.sun.star.beans import PropertyValue

SOFFICE = os.environ.get("SOFFICE", "/opt/libreoffice26.2/program/soffice")


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


class Session:
    """Starts soffice on a private pipe with a throwaway profile, so that it never touches the
    user's own LibreOffice profile and never collides with another running instance."""

    def __enter__(self):
        self.profile = tempfile.mkdtemp(prefix="spatac-lo-profile-")
        self.pipe = "spatac_migrate_%d" % os.getpid()
        self.proc = subprocess.Popen(
            [
                SOFFICE,
                "--headless",
                "--invisible",
                "--norestore",
                "--nologo",
                "-env:UserInstallation=" + uno.systemPathToFileUrl(self.profile),
                "--accept=pipe,name=%s;urp;" % self.pipe,
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local
        )
        last = None
        for _ in range(120):
            try:
                self.ctx = resolver.resolve(
                    "uno:pipe,name=%s;urp;StarOffice.ComponentContext" % self.pipe
                )
                break
            except Exception as e:  # not up yet
                last = e
                time.sleep(0.5)
        else:
            raise RuntimeError("LibreOffice did not start: %s" % last)
        self.smgr = self.ctx.ServiceManager
        self.desktop = self.smgr.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx
        )
        return self

    def open(self, path, hidden=True):
        url = uno.systemPathToFileUrl(os.path.abspath(path))
        doc = self.desktop.loadComponentFromURL(url, "_blank", 0, (prop("Hidden", hidden),))
        if doc is None:
            raise RuntimeError("cannot open " + path)
        return doc

    def new_calc(self, hidden=True):
        """A new spreadsheet. Pass hidden=False when the saved file must keep view settings
        such as frozen panes: LibreOffice drops them from a hidden document. soffice runs
        headless, so nothing is shown either way."""
        return self.desktop.loadComponentFromURL(
            "private:factory/scalc", "_blank", 0, (prop("Hidden", hidden),)
        )

    def save_xlsx(self, doc, path):
        url = uno.systemPathToFileUrl(os.path.abspath(path))
        doc.storeToURL(url, (prop("FilterName", "Calc Office Open XML"),))

    def __exit__(self, *exc):
        try:
            self.desktop.terminate()
        except Exception:
            pass
        try:
            self.proc.wait(timeout=20)
        except Exception:
            self.proc.kill()
        return False
