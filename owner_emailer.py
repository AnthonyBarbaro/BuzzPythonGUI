import os
import re
import base64
from io import BytesIO
from pathlib import Path
from email.message import EmailMessage
from email.utils import make_msgid
from datetime import date
from typing import List, Optional, Dict, Any, Tuple

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

# Inline preview settings (PDF -> image)
ENABLE_INLINE_PREVIEWS = True
MAX_PREVIEWS = 12                 # safety: cap number of inline images
PREVIEW_MAX_WIDTH_PX = 900        # downscale if huge
PREVIEW_ZOOM = 2.0                # base render zoom
PREFER_JPEG = True                # reduce email size if Pillow is available
JPEG_QUALITY = 72                 # 60-80 is usually good

# Layout
CARDS_PER_ROW = 2                 # 2 looks good in Gmail


# =============================================================================
# Helpers: parsing filenames + rendering previews
# =============================================================================

def _parse_pdf_identity(pdf_path: str) -> Dict[str, Any]:
    """
    Your filenames look like:
      - "ALL STORES - Owner Snapshot - 2026-02-08.pdf"
      - "MV - Owner Snapshot - MISSION VIEJO - 2026-02-08.pdf"

    Returns:
      {
        "is_all": bool,
        "abbr": str,
        "store_name": str,
        "display_title": str,
        "sort_key": tuple
      }
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
    display_title = f"{abbr} • {store_name.title()}"  # title-case for email polish

    return {
        "is_all": False,
        "abbr": abbr,
        "store_name": store_name,
        "display_title": display_title,
        "sort_key": (1, abbr),
    }


def _try_render_pdf_first_page(pdf_path: str) -> Optional[bytes]:
    """
    Render first page of a PDF to PNG bytes using PyMuPDF (fitz).
    Returns None if rendering isn't possible.
    """
    if not ENABLE_INLINE_PREVIEWS:
        return None

    try:
        import fitz  # PyMuPDF
    except Exception:
        return None

    try:
        doc = fitz.open(pdf_path)
        if doc.page_count < 1:
            return None

        page = doc.load_page(0)

        # Start with zoom, then downscale if too wide
        mat = fitz.Matrix(PREVIEW_ZOOM, PREVIEW_ZOOM)
        pix = page.get_pixmap(matrix=mat, alpha=False)

        # Downscale if needed to cap width
        if PREVIEW_MAX_WIDTH_PX and pix.width > PREVIEW_MAX_WIDTH_PX:
            scale = PREVIEW_MAX_WIDTH_PX / float(pix.width)
            mat2 = fitz.Matrix(PREVIEW_ZOOM * scale, PREVIEW_ZOOM * scale)
            pix = page.get_pixmap(matrix=mat2, alpha=False)

        png_bytes = pix.tobytes("png")
        return png_bytes
    except Exception:
        return None
    finally:
        try:
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
        img = Image.open(BytesIO(png_bytes))
        img = img.convert("RGB")
        out = BytesIO()
        img.save(out, format="JPEG", quality=JPEG_QUALITY, optimize=True)
        return out.getvalue(), "jpeg"
    except Exception:
        return png_bytes, "png"


def _chunk(items: List[Any], n: int) -> List[List[Any]]:
    return [items[i:i + n] for i in range(0, len(items), n)]


# =============================================================================
# HTML builder
# =============================================================================

def _build_html_email(
    report_day: date,
    data_start: date,
    data_end: date,
    previews: List[Dict[str, Any]],
    pdf_count: int,
) -> str:
    """
    Build a Gmail-friendly table-based HTML.
    previews: list of dicts like:
      {
        "abbr": "MV",
        "display_title": "MV • Mission Viejo",
        "is_all": False,
        "img_src": "cid:...." or "",
        "has_preview": bool
      }
    """
    header_date = report_day.strftime("%A, %B %d, %Y")
    window_str = f"{data_start.isoformat()} → {data_end.isoformat()}"

    # Separate All Stores card if present
    all_card = next((p for p in previews if p.get("is_all")), None)
    store_cards = [p for p in previews if not p.get("is_all")]

    def badge(text: str, bg: str, fg: str) -> str:
        return (
            f"<span style=\"display:inline-block;padding:6px 10px;border-radius:999px;"
            f"background:{bg};color:{fg};font-size:12px;font-weight:700;\">{text}</span>"
        )

    def card_html(p: Dict[str, Any]) -> str:
        title = p["display_title"]
        preview_html = ""
        if p.get("has_preview") and p.get("img_src"):
            preview_html = (
                f"<img src=\"{p['img_src']}\" alt=\"{title} preview\" "
                f"style=\"width:100%;height:auto;display:block;border-radius:10px;\">"
            )
        else:
            preview_html = (
                f"<div style=\"padding:18px;border:1px dashed {BUZZ['border']};"
                f"border-radius:10px;color:{BUZZ['muted2']};font-size:12px;\">"
                f"Preview unavailable — see attached PDF."
                f"</div>"
            )

        return f"""
        <table role="presentation" width="100%" cellspacing="0" cellpadding="0"
               style="border:1px solid {BUZZ['border']};border-radius:14px;overflow:hidden;background:{BUZZ['white']};">
          <tr>
            <td style="padding:12px 14px;border-bottom:1px solid {BUZZ['border']};background:{BUZZ['soft']};">
              <div style="font-size:13px;font-weight:800;color:#111827;letter-spacing:0.2px;">{title}</div>
              <div style="font-size:11px;color:{BUZZ['muted2']};margin-top:2px;">
                Attached PDF • Owner Snapshot
              </div>
            </td>
          </tr>
          <tr>
            <td style="padding:12px 14px;">
              {preview_html}
            </td>
          </tr>
        </table>
        """

    # Build store grid rows
    grid_rows_html = ""
    rows = _chunk(store_cards, CARDS_PER_ROW)

    for row in rows:
        tds = ""
        for p in row:
            tds += f"""
            <td width="{int(100 / CARDS_PER_ROW)}%" valign="top" style="padding:8px;">
              {card_html(p)}
            </td>
            """
        # If last row is short, pad with empty cell for alignment
        if len(row) < CARDS_PER_ROW:
            tds += f"""
            <td width="{int(100 / CARDS_PER_ROW)}%" valign="top" style="padding:8px;"></td>
            """

        grid_rows_html += f"<tr>{tds}</tr>"

    all_card_html = ""
    if all_card:
        all_card_html = f"""
        <tr>
          <td style="padding:16px 20px;">
            <div style="font-size:14px;font-weight:900;color:#111827;margin-bottom:10px;">
              All Stores Summary
            </div>
            {card_html(all_card)}
          </td>
        </tr>
        """

    # Top badges
    top_badges = (
        badge(f"REPORT DAY • {report_day.isoformat()}", BUZZ["yellow"], BUZZ["black"])
        + "&nbsp;&nbsp;"
        + badge(f"DATA WINDOW • {window_str}", BUZZ["black"], BUZZ["yellow"])
        + "&nbsp;&nbsp;"
        + badge(f"{pdf_count} PDF ATTACHMENTS", BUZZ["green"], BUZZ["white"])
    )

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

              <!-- Badges -->
              <tr>
                <td style="padding:14px 20px;border-bottom:1px solid {BUZZ['border']};">
                  {top_badges}
                </td>
              </tr>

              <!-- Quick note -->
              <tr>
                <td style="padding:16px 20px;">
                  <div style="font-size:13px;color:{BUZZ['muted']};line-height:1.45;">
                    Attached are the full Owner Snapshot PDFs per store, plus the All Stores summary.
                    Below are first-page previews for quick scanning.
                  </div>
                </td>
              </tr>

              {all_card_html}

              <!-- Store Grid -->
              <tr>
                <td style="padding:0 12px 6px 12px;">
                  <div style="padding:10px 8px 4px 8px;font-size:14px;font-weight:900;color:#111827;">
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


# =============================================================================
# Main sender
# =============================================================================

from typing import Union, List

def send_owner_snapshot_email(
    pdf_paths: List[str],
    report_day: date,
    data_start: date,
    data_end: date,
    to_email: Union[str, List[str]] = "anthony@buzzcannabis.com",
):
    """
    Sends Owner Snapshot PDFs via Gmail API using JSON token (cron-safe),
    with a branded HTML body + optional inline preview images per store.

    - If PyMuPDF (fitz) is installed, we render first page of each PDF as an inline image.
    - PDFs remain attached as normal.
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

    # ---------- Build previews list ----------
    previews: List[Dict[str, Any]] = []
    inline_images: List[Dict[str, Any]] = []

    for i, (pdf_path, ident) in enumerate(pdf_sorted):
        if i >= MAX_PREVIEWS:
            break

        png = _try_render_pdf_first_page(pdf_path)
        has_preview = bool(png)

        img_src = ""
        cid_ref = ""
        img_bytes = b""
        img_subtype = ""

        if has_preview:
            img_bytes, img_subtype = _maybe_convert_png_to_jpeg(png)

            cid = make_msgid(domain="buzzcannabis.local")  # includes <...>
            cid_ref = cid[1:-1]  # strip brackets for HTML
            img_src = f"cid:{cid_ref}"

            inline_images.append({
                "cid": cid,
                "cid_ref": cid_ref,
                "bytes": img_bytes,
                "subtype": img_subtype,   # 'png' or 'jpeg'
                "filename": f"{ident['abbr']}_preview.{('jpg' if img_subtype=='jpeg' else 'png')}",
            })

        previews.append({
            **ident,
            "has_preview": has_preview,
            "img_src": img_src,
        })

    # If we have more PDFs than previews (MAX_PREVIEWS cap), still list them as cards without previews
    if len(pdf_sorted) > len(previews):
        for (pdf_path, ident) in pdf_sorted[len(previews):]:
            previews.append({
                **ident,
                "has_preview": False,
                "img_src": "",
            })

    # ---------- Email ----------
    subject = f"Buzz Cannabis Owner Snapshot — {report_day.isoformat()}"
    if isinstance(to_email, list):
        to_header = ", ".join(to_email)
    else:
        to_header = to_email

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
        previews=previews,
        pdf_count=len(existing_pdfs),
    )
    msg.add_alternative(html, subtype="html")

    # Attach inline images to the HTML part (the last payload)
    # EmailMessage structure: [plain, html]; html part is msg.get_payload()[-1]
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

    print(f"📧 Owner Snapshot emailed to {to_email} (PDFs: {len(existing_pdfs)}, Previews: {len(inline_images)})")
