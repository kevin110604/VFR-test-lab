
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.graphics.barcode import qr
from reportlab.graphics.shapes import Drawing
from reportlab.graphics import renderPDF
import argparse

def make_qr_label_pdf(output_path: str, prefix: str, start_suffix: int, end_suffix: int,
                      page_w_mm: float = 50.0, page_h_mm: float = 65.0,
                      margins_mm: float = 3.0, text_pt: int = 12,
                      qr_max_mm: float = None):
    """
    Generate a PDF with one label per page.
    - Label size: page_w_mm x page_h_mm (mm).
    - Margins: margins_mm (mm) on all sides.
    - QR size: auto-fit to content box unless qr_max_mm is provided.
    - Text: report code (e.g., 25-5500) centered under the QR.
    """
    page_w = page_w_mm * mm
    page_h = page_h_mm * mm
    c = canvas.Canvas(output_path, pagesize=(page_w, page_h))

    # Content area
    left = margins_mm * mm
    right = page_w - margins_mm * mm
    top = page_h - margins_mm * mm
    bottom = margins_mm * mm
    content_w = right - left
    content_h = top - bottom

    text_gap = 3 * mm
    approx_text_h = text_pt * 1.2

    max_qr_h = content_h - (approx_text_h + text_gap)
    max_qr_w = content_w

    for i in range(start_suffix, end_suffix + 1):
        report_raw = f"{prefix}-{i:04d}"
        url = f"http://103.77.166.187:8246/update?report={report_raw}"

        # determine QR side length
        if qr_max_mm is None:
            qr_side = min(max_qr_w, max_qr_h)
        else:
            qr_side = min(qr_max_mm * mm, max_qr_w, max_qr_h)

        # QR drawing scaled to qr_side
        qrw = qr.QrCodeWidget(url, barLevel='H')
        b = qrw.getBounds()
        w = b[2] - b[0]
        h = b[3] - b[1]
        d = Drawing(qr_side, qr_side, transform=[qr_side / w, 0, 0, qr_side / h, 0, 0])
        d.add(qrw)

        # Center QR, leave room for text below
        qr_x = left + (content_w - qr_side) / 2.0
        qr_y = bottom + (content_h - (qr_side + text_gap + approx_text_h)) / 2.0 + approx_text_h + text_gap
        renderPDF.draw(d, c, qr_x, qr_y)

        # Report text under QR
        c.setFont("Helvetica-Bold", text_pt)
        c.drawCentredString(page_w / 2.0, qr_y - text_gap, report_raw)

        c.showPage()

    c.save()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate QR label PDF for PREFIX-XXXX ranges on custom label sizes.")
    parser.add_argument("--prefix", type=str, default="25", help="Prefix (e.g., 25, 26). Default: 25")
    parser.add_argument("--start", type=int, required=True, help="Start suffix (e.g., 5500)")
    parser.add_argument("--end", type=int, required=True, help="End suffix (e.g., 9999)")
    parser.add_argument("--out", type=str, required=True, help="Output PDF path")
    parser.add_argument("--w", type=float, default=50.0, help="Label width in mm (default 50)")
    parser.add_argument("--h", type=float, default=65.0, help="Label height in mm (default 65)")
    parser.add_argument("--margin", type=float, default=3.0, help="Margins in mm (default 3)")
    parser.add_argument("--text", type=int, default=12, help="Text size in pt (default 12)")
    parser.add_argument("--qrmm", type=float, default=None, help="Max QR size in mm (optional)")
    args = parser.parse_args()

    make_qr_label_pdf(args.out, args.prefix, args.start, args.end,
                      page_w_mm=args.w, page_h_mm=args.h,
                      margins_mm=args.margin, text_pt=args.text,
                      qr_max_mm=args.qrmm)
