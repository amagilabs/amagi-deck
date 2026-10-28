"""Slide renderers. Each one turns slide numbers into PNG files.

* ``powerpoint`` drives an installed Microsoft PowerPoint over COM (Windows only).
* ``soffice`` converts the deck to PDF with LibreOffice, then rasterizes it with PDFium.
* ``sketch`` draws an approximation with Pillow. It always works, but it is not a faithful render.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Dict, List, Optional

from .pdf_extract import page_count, render_pdf_pages

TIMEOUT = int(os.environ.get("AMAGI_DECK_TIMEOUT", "300"))
PDF_HIDDEN_FILTER = 'pdf:impress_pdf_Export:{"ExportHiddenSlides":{"type":"boolean","value":"true"}}'


class RenderError(RuntimeError):
    pass


def find_soffice() -> Optional[str]:
    env = os.environ.get("AMAGI_DECK_SOFFICE")
    if env:
        return env if Path(env).exists() or shutil.which(env) else None
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    candidates = []
    if sys.platform == "win32":
        for root in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")):
            if root:
                candidates.append(Path(root) / "LibreOffice" / "program" / "soffice.com")
                candidates.append(Path(root) / "LibreOffice" / "program" / "soffice.exe")
    elif sys.platform == "darwin":
        candidates.append(Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"))
    for c in candidates:
        if c.exists():
            return str(c)
    return None


def has_powerpoint() -> bool:
    if os.environ.get("AMAGI_DECK_NO_POWERPOINT"):
        return False
    if sys.platform != "win32":
        return False
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"PowerPoint.Application\CLSID"):
            return True
    except OSError:
        return False


def available() -> Dict[str, str]:
    """Renderer name -> short status line."""
    soffice = find_soffice()
    return {
        "powerpoint": "found (Microsoft PowerPoint via COM)" if has_powerpoint()
        else "not available (needs Windows with PowerPoint installed)",
        "soffice": "found: %s" % soffice if soffice else "not found (install LibreOffice, or set AMAGI_DECK_SOFFICE)",
        "sketch": "always available (approximate drawing, not a faithful render)",
    }


def _run(cmd: List[str], env: Optional[dict] = None) -> subprocess.CompletedProcess:
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=TIMEOUT, env=env)
    except subprocess.TimeoutExpired:
        raise RenderError("%s timed out after %ds" % (Path(cmd[0]).name, TIMEOUT))
    except OSError as exc:
        raise RenderError("could not start %s: %s" % (cmd[0], exc))
    return proc


# --- LibreOffice -----------------------------------------------------------------------------

def soffice_convert(src: Path, out_dir: Path, target: str) -> Path:
    """Run ``soffice --convert-to <target>`` and return the produced file."""
    soffice = find_soffice()
    if not soffice:
        raise RenderError("LibreOffice (soffice) not found")
    profile = Path(tempfile.mkdtemp(prefix="amagi-deck-lo-"))
    try:
        cmd = [soffice, "-env:UserInstallation=" + profile.as_uri(), "--headless", "--norestore",
               "--nolockcheck", "--convert-to", target, "--outdir", str(out_dir), str(src)]
        proc = _run(cmd)
        ext = target.split(":", 1)[0]
        produced = out_dir / (src.stem + "." + ext)
        if not produced.exists():
            detail = (proc.stderr or proc.stdout or "").strip().splitlines()
            raise RenderError("soffice did not write %s (exit %d)%s" % (
                produced.name, proc.returncode, ": " + detail[-1] if detail else ""))
        return produced
    finally:
        shutil.rmtree(profile, ignore_errors=True)


def render_soffice(src: Path, targets: Dict[int, Path], width: int, hidden: List[bool]) -> Dict[int, Path]:
    """Render via LibreOffice PDF export. ``hidden`` has one flag per slide in the deck."""
    with tempfile.TemporaryDirectory(prefix="amagi-deck-") as tmp:
        pdf = soffice_convert(src, Path(tmp), PDF_HIDDEN_FILTER)
        pages = page_count(pdf)
        total = len(hidden)
        visible = [i + 1 for i, h in enumerate(hidden) if not h]
        if pages == total:
            mapping = {n: n for n in range(1, total + 1)}
        elif pages == len(visible):
            # Older LibreOffice ignores ExportHiddenSlides; hidden slides get no render.
            mapping = {slide: page for page, slide in enumerate(visible, start=1)}
        else:
            raise RenderError("LibreOffice produced %d pages for %d slides" % (pages, total))
        page_targets = {mapping[n]: out for n, out in targets.items() if n in mapping}
        rendered = render_pdf_pages(pdf, page_targets, width)
        back = {page: slide for slide, page in mapping.items()}
        return {back[p]: out for p, out in rendered.items()}


# --- PowerPoint ------------------------------------------------------------------------------

PS_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject PowerPoint.Application
$before = $app.Presentations.Count
$pres = $null
try {
  $pres = $app.Presentations.Open($env:AMAGI_DECK_SRC, -1, 0, 0)
  if ($env:AMAGI_DECK_MODE -eq 'pptx') {
    $pres.SaveAs($env:AMAGI_DECK_DEST, 24)
  } else {
    $w = [int]$env:AMAGI_DECK_W
    $h = [int]$env:AMAGI_DECK_H
    foreach ($n in ($env:AMAGI_DECK_SLIDES -split ',')) {
      $file = Join-Path $env:AMAGI_DECK_OUT ("{0}.png" -f $n)
      $pres.Slides.Item([int]$n).Export($file, 'PNG', $w, $h)
    }
  }
} finally {
  if ($pres) { $pres.Close() }
  if ($before -eq 0 -and $app.Presentations.Count -eq 0) { $app.Quit() }
}
"""


def _powerpoint(env_extra: Dict[str, str]) -> None:
    if not has_powerpoint():
        raise RenderError("Microsoft PowerPoint not available")
    with tempfile.TemporaryDirectory(prefix="amagi-deck-ps-") as tmp:
        script = Path(tmp) / "export.ps1"
        script.write_text(PS_SCRIPT, encoding="utf-8")
        env = dict(os.environ)
        env.update(env_extra)
        # -Mta: under the default STA apartment, powershell.exe waits about a minute for PowerPoint
        # to finish quitting before it exits. In MTA it returns at once and PowerPoint exits by itself.
        proc = _run(["powershell", "-Mta", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                     "-File", str(script)], env=env)
        if proc.returncode != 0:
            lines = [ln for ln in (proc.stderr or proc.stdout or "").splitlines() if ln.strip()]
            raise RenderError("PowerPoint export failed: %s" % (lines[0].strip() if lines else proc.returncode))


def render_powerpoint(src: Path, targets: Dict[int, Path], width: int, height: int) -> Dict[int, Path]:
    with tempfile.TemporaryDirectory(prefix="amagi-deck-pp-") as tmp:
        _powerpoint({
            "AMAGI_DECK_MODE": "png",
            "AMAGI_DECK_SRC": str(src.resolve()),
            "AMAGI_DECK_OUT": tmp,
            "AMAGI_DECK_W": str(width),
            "AMAGI_DECK_H": str(height),
            "AMAGI_DECK_SLIDES": ",".join(str(n) for n in sorted(targets)),
        })
        done = {}
        for n, out in targets.items():
            produced = Path(tmp) / ("%d.png" % n)
            if produced.exists():
                shutil.move(str(produced), str(out))
                done[n] = out
        if not done:
            raise RenderError("PowerPoint exported no slides")
        return done


def powerpoint_to_pptx(src: Path, dest: Path) -> Path:
    _powerpoint({"AMAGI_DECK_MODE": "pptx", "AMAGI_DECK_SRC": str(src.resolve()),
                 "AMAGI_DECK_DEST": str(dest.resolve())})
    if not dest.exists():
        raise RenderError("PowerPoint did not write %s" % dest.name)
    return dest
