import os
import base64
from io import BytesIO
from pathlib import Path
from email.message import EmailMessage
from email.utils import make_msgid
from datetime import date
from typing import List, Optional, Dict, Any, Tuple, Union

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build


# =============================================================================
# CONFIG (easy to tweak)
# =============================================================================

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
GMAIL_TOKEN = "token_gmail.json"

# Buzz theme
BUZZ = {
    "yellow": "#FFF200",
    "green": "#00AE6F",
    "black": "#000000",
    "white": "#FFFFFF",
    "bg": "#F3F4F6",
    "muted": "#374151",
    "muted2": "#6B7280",
    "border": "#E5E7EB",
    "soft": "#F7F7F7",
}

# -----------------------------------------------------------------------------
# Inline preview settings (PDF -> image)
# -----------------------------------------------------------------------------
ENABLE_INLINE_PREVIEWS = True

# Cap how many PDFs we render previews for (PDFs are still attached regardless)
MAX_PREVIEWS = 12

# Render zoom; higher = sharper but larger
PREVIEW_ZOOM = 2.2

# Downscale if rendered image width is too large
PREVIEW_MAX_WIDTH_PX = 980

# ✅ Crop preview to focus on top KPIs
# Show the top X% of page 1 (0.50 = top half)
PREVIEW_CROP_TOP_FRACTION = 0.00          # usually 0.00
PREVIEW_CROP_HEIGHT_FRACTION = 0.55       # ✅ top 55% feels great for KPIs

# Prefer JPEG to keep email size down (requires Pillow for conversion)
PREFER_JPEG = True
JPEG_QUALITY = 72

# -----------------------------------------------------------------------------
# Store images/icons (optional)
# -----------------------------------------------------------------------------
ENABLE_STORE_ICONS = True

# Put store images here (recommended):
#   store_images/MV.jpg
#   store_images/LM.jpg
#   store_images/SV.jpg
#   store_images/LG.jpg
#   store_images/NC.jpg
#   store_images/WP.jpg
STORE_IMAGE_DIR = Path("store_images")

# Optional overrides if your filenames don't match abbr exactly
STORE_IMAGE_OVERRIDES: Dict[str, str] = {
    # "MV": "mission_viejo.jpg",
}

# Icon render size (px). Bigger looks more “premium”
STORE_ICON_SIZE_PX = 44

# -----------------------------------------------------------------------------
# “Color wave” header banner (generated as inline PNG) (requires Pillow)
# -----------------------------------------------------------------------------
ENABLE_WAVE_BANNER = True
WAVE_BANNER_WIDTH_PX = 760
WAVE_BANNER_HEIGHT_PX = 26

# -----------------------------------------------------------------------------
# Layout
# -----------------------------------------------------------------------------
# If you want “one store per row” (big readable KPI preview), set this to 1
# If you want a 2-column grid, set this to 2
CARDS_PER_ROW = 1


# =============================================================================
# Helpers: parsing filenames
# =============================================================================

def _parse_pdf_identity(pdf_path: str) -> Dict[str, Any]:
    """
    Filenames:
      - "ALL STORES - Owner Snapshot - 2026-02-08.pdf"
      - "MV - Owner Snapshot - MISSION VIEJO - 2026-02-08.pdf"
    """
    p = Path(pdf_path)
    stem = p.stem
    parts = stem.split(" - ")

    if parts and parts[0].strip().upper() == "ALL STORES":
        return {
            "is_all": True,
            "abbr": "ALL",
            "store_name": "All Stores",
            "display_title": "All Stores Summary",
            "sort_key": (0, "ALL"),
        }

    abbr = (parts[0].strip() if parts else "STORE").upper()
    store_name = parts[2].strip() if len(parts) >= 3 else abbr
    display_title = f"{abbr} • {store_name.title()}"

    return {
        "is_all": False,
        "abbr": abbr,
        "store_name": store_name,
        "display_title": display_title,
        "sort_key": (1, abbr),
    }


def _chunk(items: List[Any], n: int) -> List[List[Any]]:
    return [items[i:i + n] for i in range(0, len(items), n)]


# =============================================================================
# PDF preview rendering (top-of-page crop)
# =============================================================================

def _try_render_pdf_first_page(pdf_path: str) -> Optional[bytes]:
    """
    Render first page (cropped to top section) to PNG bytes using PyMuPDF (fitz).
    Returns None if rendering isn't possible.

    Crop behavior is controlled by:
      PREVIEW_CROP_TOP_FRACTION
      PREVIEW_CROP_HEIGHT_FRACTION
    """
    if not ENABLE_INLINE_PREVIEWS:
        return None

    try:
        import fitz  # PyMuPDF
    except Exception:
        return None

    doc = None
    try:
        doc = fitz.open(pdf_path)
        if doc.page_count < 1:
            return None

        page = doc.load_page(0)
        rect = page.rect  # (x0, y0, x1, y1)

        # ✅ Clip to top portion for KPI readability
        clip = None
        if PREVIEW_CROP_HEIGHT_FRACTION and PREVIEW_CROP_HEIGHT_FRACTION < 0.999:
            y0 = rect.y0 + rect.height * float(PREVIEW_CROP_TOP_FRACTION)
            y1 = y0 + rect.height * float(PREVIEW_CROP_HEIGHT_FRACTION)
            clip = fitz.Rect(rect.x0, y0, rect.x1, y1)

        # First render
        mat = fitz.Matrix(PREVIEW_ZOOM, PREVIEW_ZOOM)
        pix = page.get_pixmap(matrix=mat, clip=clip, alpha=False)

        # Downscale by re-rendering with a smaller zoom (better than resampling)
        if PREVIEW_MAX_WIDTH_PX and pix.width > PREVIEW_MAX_WIDTH_PX:
            scale = PREVIEW_MAX_WIDTH_PX / float(pix.width)
            mat2 = fitz.Matrix(PREVIEW_ZOOM * scale, PREVIEW_ZOOM * scale)
            pix = page.get_pixmap(matrix=mat2, clip=clip, alpha=False)

        return pix.tobytes("png")
    except Exception:
        return None
    finally:
        try:
            if doc is not None:
                doc.close()
        except Exception:
            pass


def _maybe_convert_png_to_jpeg(png_bytes: bytes) -> Tuple[bytes, str]:
    """
    Convert PNG -> JPEG to reduce size (if Pillow is available).
    Returns (bytes, subtype) where subtype is 'jpeg' or 'png'.
    """
    if not PREFER_JPEG:
        return png_bytes, "png"

    try:
        from PIL import Image
    except Exception:
        return png_bytes, "png"

    try:
        img = Image.open(BytesIO(png_bytes)).convert("RGB")
        out = BytesIO()
        img.save(out, format="JPEG", quality=JPEG_QUALITY, optimize=True)
        return out.getvalue(), "jpeg"
    except Exception:
        return png_bytes, "png"


# =============================================================================
# Store icon loading (optional)
# =============================================================================

def _try_load_store_icon_bytes(abbr: str) -> Optional[Tuple[bytes, str]]:
    """
    Loads a store image from STORE_IMAGE_DIR and returns (bytes, subtype).
    If Pillow is installed, it will:
      - resize to STORE_ICON_SIZE_PX
      - crop to a circle
      - output PNG (best for icons)
    """
    if not ENABLE_STORE_ICONS:
        return None

    abbr = (abbr or "").strip().upper()
    if not abbr or abbr == "ALL":
        return None

    STORE_IMAGE_DIR.mkdir(parents=True, exist_ok=True)

    candidates: List[Path] = []
    override = STORE_IMAGE_OVERRIDES.get(abbr)
    if override:
        candidates.append(STORE_IMAGE_DIR / override)
    else:
        for ext in (".png", ".jpg", ".jpeg", ".webp"):
            candidates.append(STORE_IMAGE_DIR / f"{abbr}{ext}")
            candidates.append(STORE_IMAGE_DIR / f"{abbr.lower()}{ext}")

    img_path = next((p for p in candidates if p.exists()), None)
    if not img_path:
        return None

    raw = img_path.read_bytes()

    # Best: resize + circle mask via Pillow
    try:
        from PIL import Image, ImageDraw

        im = Image.open(BytesIO(raw)).convert("RGBA")
        im = im.resize((STORE_ICON_SIZE_PX, STORE_ICON_SIZE_PX), Image.LANCZOS)

        mask = Image.new("L", (STORE_ICON_SIZE_PX, STORE_ICON_SIZE_PX), 0)
        draw = ImageDraw.Draw(mask)
        draw.ellipse((0, 0, STORE_ICON_SIZE_PX - 1, STORE_ICON_SIZE_PX - 1), fill=255)

        out = Image.new("RGBA", (STORE_ICON_SIZE_PX, STORE_ICON_SIZE_PX), (0, 0, 0, 0))
        out.paste(im, (0, 0), mask=mask)

        buf = BytesIO()
        out.save(buf, format="PNG", optimize=True)
        return buf.getvalue(), "png"
    except Exception:
        # Fallback: embed original bytes (may be larger)
        ext = img_path.suffix.lower().lstrip(".")
        if ext == "jpg":
            ext = "jpeg"
        if ext not in ("png", "jpeg", "webp", "gif"):
            ext = "png"
        return raw, ext


# =============================================================================
# Wave banner (optional)
# =============================================================================

def _try_make_wave_banner_png(width_px: int, height_px: int) -> Optional[bytes]:
    """
    Creates a small “color wave” banner as PNG bytes using Pillow.
    If Pillow isn't installed, returns None.
    """
    if not ENABLE_WAVE_BANNER:
        return None

    try:
        from PIL import Image, ImageDraw
    except Exception:
        return None

    import math

    img = Image.new("RGB", (width_px, height_px), BUZZ["white"])
    draw = ImageDraw.Draw(img)

    # Wave parameters
    amp1 = max(2, int(height_px * 0.18))
    amp2 = max(2, int(height_px * 0.14))
    base1 = int(height_px * 0.45)
    base2 = int(height_px * 0.62)
    cycles = 1.30  # number of sine cycles across width

    # Yellow wave (upper)
    pts_yellow = []
    for x in range(width_px + 1):
        y = base1 + int(amp1 * math.sin((2 * math.pi * cycles * x) / width_px))
        pts_yellow.append((x, y))
    poly_yellow = [(0, height_px)] + pts_yellow + [(width_px, height_px)]
    draw.polygon(poly_yellow, fill=BUZZ["yellow"])

    # Green wave (lower, phase shifted)
    pts_green = []
    for x in range(width_px + 1):
        y = base2 + int(amp2 * math.sin((2 * math.pi * cycles * x) / width_px + 1.4))
        pts_green.append((x, y))
    poly_green = [(0, height_px)] + pts_green + [(width_px, height_px)]
    draw.polygon(poly_green, fill=BUZZ["green"])

    # Thin top border line for crispness
    draw.line([(0, 0), (width_px, 0)], fill=BUZZ["border"], width=1)

    buf = BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


# =============================================================================
# HTML builder
# =============================================================================

def _build_plain_text_email(report_day: date, data_start: date, data_end: date) -> str:
    return (
        "Buzz Cannabis — Owner Snapshot\n\n"
        f"Report Day: {report_day.strftime('%A, %B %d, %Y')} ({report_day.isoformat()})\n"
        f"Data Window: {data_start.isoformat()} → {data_end.isoformat()}\n\n"
        "Attached:\n"
        "• Store Owner Snapshot PDFs\n"
        "• All Stores Summary\n\n"
        "This email was generated automatically.\n"
    )


def _build_html_email(
    report_day: date,
    data_start: date,
    data_end: date,
    cards: List[Dict[str, Any]],
    pdf_count: int,
    wave_src: str = "",
) -> str:
    """
    Build a Gmail-friendly table-based HTML.

    cards entries look like:
      {
        "abbr": "MV",
        "display_title": "MV • Mission Viejo",
        "is_all": False,
        "img_src": "cid:..." or "",
        "icon_src": "cid:..." or "",
        "has_preview": bool
      }
    """
    header_date = report_day.strftime("%A, %B %d, %Y")
    window_str = f"{data_start.isoformat()} → {data_end.isoformat()}"

    all_card = next((c for c in cards if c.get("is_all")), None)
    store_cards = [c for c in cards if not c.get("is_all")]

    def badge(text: str, bg: str, fg: str) -> str:
        return (
            f"<span style=\"display:inline-block;padding:6px 10px;border-radius:999px;"
            f"background:{bg};color:{fg};font-size:12px;font-weight:800;letter-spacing:0.2px;\">{text}</span>"
        )

    def _icon_block(c: Dict[str, Any]) -> str:
        """
        Store icon. If we have icon_src, show image. Else fallback to abbr circle.
        """
        abbr = (c.get("abbr") or "").upper()
        icon_src = c.get("icon_src") or ""

        if icon_src:
            return (
                f"<img src=\"{icon_src}\" width=\"{STORE_ICON_SIZE_PX}\" height=\"{STORE_ICON_SIZE_PX}\" "
                f"style=\"display:block;border-radius:999px;border:2px solid {BUZZ['yellow']};\" "
                f"alt=\"{abbr} icon\">"
            )

        # fallback: colored monogram circle
        return (
            f"<div style=\"width:{STORE_ICON_SIZE_PX}px;height:{STORE_ICON_SIZE_PX}px;border-radius:999px;"
            f"background:{BUZZ['yellow']};color:{BUZZ['black']};text-align:center;"
            f"line-height:{STORE_ICON_SIZE_PX}px;font-weight:900;font-size:12px;letter-spacing:0.6px;\">"
            f"{abbr}"
            f"</div>"
        )

    def _preview_block(c: Dict[str, Any]) -> str:
        title = c.get("display_title", "Snapshot")
        if c.get("has_preview") and c.get("img_src"):
            return (
                f"<img src=\"{c['img_src']}\" alt=\"{title} preview\" "
                f"style=\"width:100%;height:auto;display:block;border-radius:12px;"
                f"border:1px solid {BUZZ['border']};\">"
                f"<div style=\"margin-top:8px;font-size:11px;color:{BUZZ['muted2']};line-height:1.35;\">"
                f"Preview shows the top KPI section of page 1. Open the attached PDF for the full report."
                f"</div>"
            )

        return (
            f"<div style=\"padding:16px;border:1px dashed {BUZZ['border']};"
            f"border-radius:12px;color:{BUZZ['muted2']};font-size:12px;line-height:1.35;\">"
            f"Preview unavailable — see attached PDF."
            f"</div>"
        )

    def card_html(c: Dict[str, Any]) -> str:
        title = c.get("display_title", "Owner Snapshot")

        return f"""
        <table role="presentation" width="100%" cellspacing="0" cellpadding="0"
               style="border:1px solid {BUZZ['border']};border-radius:16px;overflow:hidden;background:{BUZZ['white']};">
          <tr>
            <td style="padding:12px 14px;border-bottom:1px solid {BUZZ['border']};background:{BUZZ['soft']};">
              <table role="presentation" width="100%" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="{STORE_ICON_SIZE_PX}" valign="middle" style="padding-right:10px;">
                    {_icon_block(c)}
                  </td>
                  <td valign="middle">
                    <div style="font-size:13px;font-weight:900;color:#111827;letter-spacing:0.2px;">
                      {title}
                    </div>
                    <div style="font-size:11px;color:{BUZZ['muted2']};margin-top:2px;">
                      Attached PDF • Owner Snapshot
                    </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td style="padding:12px 14px;">
              {_preview_block(c)}
            </td>
          </tr>
        </table>
        """

    # Build store grid rows
    grid_rows_html = ""
    rows = _chunk(store_cards, max(1, int(CARDS_PER_ROW)))

    for row in rows:
        tds = ""
        for c in row:
            tds += f"""
            <td width="{int(100 / max(1, CARDS_PER_ROW))}%" valign="top" style="padding:8px;">
              {card_html(c)}
            </td>
            """

        # pad last row
        if len(row) < CARDS_PER_ROW:
            tds += f"""
            <td width="{int(100 / max(1, CARDS_PER_ROW))}%" valign="top" style="padding:8px;"></td>
            """

        grid_rows_html += f"<tr>{tds}</tr>"

    # Badges
    top_badges = (
        badge(f"REPORT DAY • {report_day.isoformat()}", BUZZ["yellow"], BUZZ["black"])
        + "&nbsp;&nbsp;"
        + badge(f"DATA WINDOW • {window_str}", BUZZ["black"], BUZZ["yellow"])
        + "&nbsp;&nbsp;"
        + badge(f"{pdf_count} PDF ATTACHMENTS", BUZZ["green"], BUZZ["white"])
    )

    # Wave row (inline image) OR fallback stripes
    if wave_src:
        wave_html = (
            f"<img src=\"{wave_src}\" alt=\"\" style=\"width:100%;height:auto;display:block;\">"
        )
    else:
        wave_html = f"""
        <table role="presentation" width="100%" cellspacing="0" cellpadding="0">
          <tr>
            <td style="height:6px;background:{BUZZ['yellow']};"></td>
            <td style="height:6px;background:{BUZZ['green']};"></td>
            <td style="height:6px;background:{BUZZ['yellow']};"></td>
          </tr>
        </table>
        """

    # All Stores block (full width)
    all_card_html = ""
    if all_card:
        all_card_html = f"""
        <tr>
          <td style="padding:14px 20px 4px 20px;">
            <div style="font-size:14px;font-weight:900;color:#111827;margin-bottom:10px;">
              All Stores Summary
            </div>
            {card_html(all_card)}
          </td>
        </tr>
        """

    html = f"""
    <div style="margin:0;padding:0;background:{BUZZ['bg']};">
      <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:{BUZZ['bg']};padding:24px 0;">
        <tr>
          <td align="center">
            <table role="presentation" width="100%" cellspacing="0" cellpadding="0"
                   style="max-width:760px;background:{BUZZ['white']};border:1px solid {BUZZ['border']};border-radius:18px;overflow:hidden;">
              <!-- Header -->
              <tr>
                <td style="background:{BUZZ['black']};padding:18px 20px;">
                  <div style="color:{BUZZ['yellow']};font-size:18px;font-weight:900;letter-spacing:0.6px;">
                    BUZZ CANNABIS
                  </div>
                  <div style="color:{BUZZ['white']};font-size:13px;margin-top:4px;opacity:0.95;">
                    Owner Snapshot • {header_date}
                  </div>
                </td>
              </tr>

              <!-- Wave -->
              <tr>
                <td>
                  {wave_html}
                </td>
              </tr>

              <!-- Badges -->
              <tr>
                <td style="padding:14px 20px;border-bottom:1px solid {BUZZ['border']};">
                  {top_badges}
                </td>
              </tr>

              <!-- Quick note -->
              <tr>
                <td style="padding:16px 20px;">
                  <div style="font-size:13px;color:{BUZZ['muted']};line-height:1.5;">
                    Attached are the full Owner Snapshot PDFs per store, plus the All Stores summary.
                    Below are <b>top-of-page KPI previews</b> for quick scanning.
                  </div>
                </td>
              </tr>

              {all_card_html}

              <!-- Store Grid -->
              <tr>
                <td style="padding:6px 12px 10px 12px;">
                  <div style="padding:8px 8px 4px 8px;font-size:14px;font-weight:900;color:#111827;">
                    Store Snapshots
                  </div>
                  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="padding:0 0 8px 0;">
                    {grid_rows_html}
                  </table>
                </td>
              </tr>

              <!-- Footer -->
              <tr>
                <td style="padding:14px 20px;background:#111827;color:#9CA3AF;font-size:11px;line-height:1.4;">
                  Auto-generated by Buzz Automation • Reply to this email if something looks off.
                  <span style="color:{BUZZ['yellow']};font-weight:800;">#Buzz</span>
                </td>
              </tr>

            </table>
          </td>
        </tr>
      </table>
    </div>
    """
    return html


# =============================================================================
# Main sender
# =============================================================================

def send_owner_snapshot_email(
    pdf_paths: List[str],
    report_day: date,
    data_start: date,
    data_end: date,
    to_email: Union[str, List[str]] = "anthony@buzzcannabis.com",
):
    """
    Sends Owner Snapshot PDFs via Gmail API using JSON token (cron-safe),
    with a branded HTML body + optional inline preview images.

    Features:
      - ✅ Cropped TOP-of-page preview (KPIs readable)
      - ✅ Optional store icon per card
      - ✅ Optional wave banner (inline image)
      - PDFs are still attached normally
    """

    if not os.path.exists(GMAIL_TOKEN):
        raise RuntimeError("❌ token_gmail.json not found — run Gmail auth first")

    creds = Credentials.from_authorized_user_file(GMAIL_TOKEN, SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        with open(GMAIL_TOKEN, "w") as f:
            f.write(creds.to_json())

    service = build("gmail", "v1", credentials=creds)

    # ---------- Sort + identify PDFs ----------
    existing_pdfs = [p for p in pdf_paths if p and os.path.exists(p)]
    pdf_identities = [_parse_pdf_identity(p) for p in existing_pdfs]
    pdf_sorted = sorted(zip(existing_pdfs, pdf_identities), key=lambda x: x[1]["sort_key"])

    # ---------- Inline images registry ----------
    inline_images: List[Dict[str, Any]] = []

    def _register_inline_image(img_bytes: bytes, subtype: str, filename: str) -> str:
        """
        Registers an inline image and returns the HTML src (cid:...).
        """
        cid = make_msgid(domain="buzzcannabis.local")  # includes <...>
        cid_ref = cid[1:-1]  # strip brackets for HTML
        inline_images.append({
            "cid": cid,
            "cid_ref": cid_ref,
            "bytes": img_bytes,
            "subtype": subtype,
            "filename": filename,
        })
        return f"cid:{cid_ref}"

    # ---------- Wave banner ----------
    wave_src = ""
    wave_png = _try_make_wave_banner_png(WAVE_BANNER_WIDTH_PX, WAVE_BANNER_HEIGHT_PX)
    if wave_png:
        wave_src = _register_inline_image(wave_png, "png", "buzz_wave.png")

    # ---------- Build cards (store + all) ----------
    # Dedup icons by store abbr so we don't attach the same image twice
    icon_src_by_abbr: Dict[str, str] = {}

    cards: List[Dict[str, Any]] = []
    preview_renders_used = 0

    for pdf_path, ident in pdf_sorted:
        abbr = ident.get("abbr", "STORE")

        # Store icon (optional)
        icon_src = ""
        if ENABLE_STORE_ICONS and abbr and abbr != "ALL":
            if abbr in icon_src_by_abbr:
                icon_src = icon_src_by_abbr[abbr]
            else:
                icon = _try_load_store_icon_bytes(abbr)
                if icon:
                    icon_bytes, icon_subtype = icon
                    icon_src = _register_inline_image(
                        icon_bytes,
                        icon_subtype,
                        f"{abbr}_icon.{('png' if icon_subtype == 'png' else icon_subtype)}"
                    )
                    icon_src_by_abbr[abbr] = icon_src

        # PDF preview (top-of-page crop)
        img_src = ""
        has_preview = False

        if ENABLE_INLINE_PREVIEWS and preview_renders_used < MAX_PREVIEWS:
            png = _try_render_pdf_first_page(pdf_path)
            if png:
                img_bytes, img_subtype = _maybe_convert_png_to_jpeg(png)
                img_src = _register_inline_image(
                    img_bytes,
                    img_subtype,
                    f"{abbr}_preview.{('jpg' if img_subtype == 'jpeg' else 'png')}"
                )
                has_preview = True

            preview_renders_used += 1

        cards.append({
            **ident,
            "has_preview": has_preview,
            "img_src": img_src,
            "icon_src": icon_src,
        })

    # ---------- Email ----------
    subject = f"Buzz Cannabis Owner Snapshot — {report_day.isoformat()}"
    to_header = ", ".join(to_email) if isinstance(to_email, list) else to_email

    msg = EmailMessage()
    msg["To"] = to_header
    msg["From"] = "me"
    msg["Subject"] = subject

    # Plain text fallback
    msg.set_content(_build_plain_text_email(report_day, data_start, data_end))

    # HTML body
    html = _build_html_email(
        report_day=report_day,
        data_start=data_start,
        data_end=data_end,
        cards=cards,
        pdf_count=len(existing_pdfs),
        wave_src=wave_src,
    )
    msg.add_alternative(html, subtype="html")

    # Attach inline images to the HTML part
    html_part = msg.get_payload()[-1]
    for img in inline_images:
        html_part.add_related(
            img["bytes"],
            maintype="image",
            subtype=img["subtype"],
            cid=img["cid"],  # keep brackets here
            filename=img["filename"],
        )

    # Attach PDFs
    for path in existing_pdfs:
        with open(path, "rb") as f:
            data = f.read()

        msg.add_attachment(
            data,
            maintype="application",
            subtype="pdf",
            filename=os.path.basename(path),
        )

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()

    service.users().messages().send(
        userId="me",
        body={"raw": raw},
    ).execute()

    print(
        f"📧 Owner Snapshot emailed to {to_email} "
        f"(PDFs: {len(existing_pdfs)}, Previews rendered: {preview_renders_used}, Inline images: {len(inline_images)})"
    )
