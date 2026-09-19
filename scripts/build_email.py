#!/usr/bin/env python3
"""
Builds a standalone "select-all-and-copy" HTML page that mirrors exactly what
the report's own "Copy as HTML for Email" button would put on the clipboard —
so the user can open it, Ctrl+A, Ctrl+C, and paste into Outlook/Gmail without
ever needing to open the interactive report or click a button (both of which
require a real, focused browser tab with clipboard permission that Claude's
in-chat preview does not grant).

Usage:
    python3 build_email.py \
        --title "IOT Testing Progress Report" \
        --gen-date "September 14, 2026 at 03:01 PM" \
        --logo /path/to/logo.png \
        --system "MiddleWare" --cr-number "CR#6893 - Title" \
        --tester "Ahmed Saeed (ahmed.saiid)" --test-date "14 Sep 2026" \
        --environment "Testing" --status "Completed" \
        --pass 2 --fail 0 --blocked 0 --inprogress 0 --invalid 0 --notrun 0 \
        --out /path/to/TMS_Report_XXXX_EMAIL.html

Only stdlib + Pillow (PIL) are required.
"""
import argparse
import base64

FONT = "Segoe UI,Arial,sans-serif"
THEME = dict(cardBg="#ffffff", text="#1c2333", muted="#6b7280", line="#e6e8ee",
             totalBg="#eaf1ff", trackBg="#eef0f4", kpiBg="#f7f8fc")


def build_donut_png(status_data, out_path):
    from PIL import Image, ImageDraw, ImageFont
    total = sum(d["count"] for d in status_data)
    SCALE = 4
    size = 220 * SCALE
    cx = cy = size // 2
    r_outer = 95 * SCALE
    r_inner = 58 * SCALE
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    if total == 0:
        draw.ellipse([cx - r_outer, cy - r_outer, cx + r_outer, cy + r_outer], fill="#eef0f4")
    else:
        start = -90.0
        for d in status_data:
            if d["count"] == 0:
                continue
            sweep = d["count"] / total * 360.0
            draw.pieslice([cx - r_outer, cy - r_outer, cx + r_outer, cy + r_outer],
                          start, start + sweep, fill=d["color"])
            start += sweep
    draw.ellipse([cx - r_inner, cy - r_inner, cx + r_inner, cy + r_inner], fill="#ffffff")

    # Resolve a bold font on both Windows and Linux/macOS.
    # The original Skill used a Linux-only DejaVuSans path, which causes
    # Pillow to raise "cannot open resource" when the app runs on Windows.
    import os

    def load_bold_font(font_size):
        candidates = []

        if os.name == "nt":
            windir = os.environ.get("WINDIR", r"C:\Windows")
            candidates.extend([
                os.path.join(windir, "Fonts", "segoeuib.ttf"),
                os.path.join(windir, "Fonts", "arialbd.ttf"),
                os.path.join(windir, "Fonts", "calibrib.ttf"),
            ])
        else:
            candidates.extend([
                "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
                "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf",
                "/Library/Fonts/Arial Bold.ttf",
            ])

        for font_path in candidates:
            if os.path.isfile(font_path):
                try:
                    return ImageFont.truetype(font_path, font_size)
                except OSError:
                    pass

        # Last-resort fallback: Pillow's built-in font is always available.
        return ImageFont.load_default()

    max_text_w = r_inner * 2 * 0.72  # keep clear margin inside the inner white circle

    # Number: shrink to fit so large totals (3-digit+) never overflow the circle
    num_text = str(total)
    num_size = 46 * SCALE
    while num_size > 14 * SCALE:
        f = load_bold_font(num_size)
        bbox = draw.textbbox((0, 0), num_text, font=f)
        if (bbox[2] - bbox[0]) <= max_text_w:
            break
        num_size -= 2 * SCALE
    font_num = load_bold_font(num_size)
    num_bbox = draw.textbbox((0, 0), num_text, font=font_num)
    num_w, num_h = num_bbox[2] - num_bbox[0], num_bbox[3] - num_bbox[1]

    # Label: shrink similarly, then true-center the (number + label) block as one unit
    label_text = "TEST CASES"
    label_size = 15 * SCALE
    while label_size > 8 * SCALE:
        f = load_bold_font(label_size)
        bbox = draw.textbbox((0, 0), label_text, font=f)
        if (bbox[2] - bbox[0]) <= max_text_w:
            break
        label_size -= 1 * SCALE
    font_label = load_bold_font(label_size)
    label_bbox = draw.textbbox((0, 0), label_text, font=font_label)
    label_w, label_h = label_bbox[2] - label_bbox[0], label_bbox[3] - label_bbox[1]

    gap = 8 * SCALE
    block_h = num_h + gap + label_h
    top = cy - block_h / 2

    draw.text((cx - num_w / 2 - num_bbox[0], top - num_bbox[1]), num_text, fill="#1c2333", font=font_num)
    draw.text((cx - label_w / 2 - label_bbox[0], top + num_h + gap - label_bbox[1]), label_text, fill="#6b7280", font=font_label)

    img = img.resize((220, 220), Image.LANCZOS)
    img.save(out_path)


def b64_of(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()


def email_bar(pct, color, track_color):
    pct = max(0, min(100, pct))
    empty = 100 - pct
    row = f'<td width="{pct}%" style="width:{pct}%;height:8px;line-height:8px;font-size:1px;background:{color};border-radius:4px 0 0 4px;">&nbsp;</td>'
    if empty > 0:
        row += f'<td width="{empty}%" style="width:{empty}%;height:8px;line-height:8px;font-size:1px;background:{track_color};border-radius:0 4px 4px 0;">&nbsp;</td>'
    return f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;"><tr>{row}</tr></table>'


def email_card(header_bg, header_text, body_html):
    return f'''
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;background:{THEME["cardBg"]};border:1px solid {THEME["line"]};border-radius:10px;margin-top:14px;">
  <tr><td style="background:{header_bg};color:#ffffff;font-weight:700;font-size:13.5px;padding:11px 16px;border-radius:9px 9px 0 0;font-family:{FONT};">{header_text}</td></tr>
  <tr><td style="padding:16px;">{body_html}</td></tr>
</table>'''


def build_intro_text(short_title, cr_number_field, rate):
    if rate == 100:
        return (f'Dears,<br><br>Kindly check below {short_title} Report for {cr_number_field} has passed '
                f'testing successfully, TMS updated status &ldquo;Business UAT Sign Off&rdquo;.')
    else:
        return (f'Dears,<br><br>Kindly check below {short_title} and the detailed attached Report for '
                f'{cr_number_field}, TMS updated status &ldquo;Pending Rework&rdquo;.')


def build_email_html(title, gen_date, logo_path, cr_info_rows, status_data, intro_html=None, observations=None):
    logo_src = f"data:image/png;base64,{b64_of(logo_path)}" if logo_path else ""

    import tempfile, os
    with tempfile.TemporaryDirectory() as tmp:
        donut_path = os.path.join(tmp, "donut.png")
        build_donut_png(status_data, donut_path)
        donut_src = f"data:image/png;base64,{b64_of(donut_path)}"

    total = sum(d["count"] for d in status_data)
    passed = next(d["count"] for d in status_data if d["key"] == "pass")
    not_run = next(d["count"] for d in status_data if d["key"] == "notrun")
    executed = total - not_run
    open_issues = sum(d["count"] for d in status_data if d["key"] in ("fail", "blocked"))
    rate = round(passed / total * 100) if total else 0
    completion_pct = (executed / total * 100) if total else 0

    cr_rows_html = "".join(f'''
      <tr>
        <td style="padding:7px 0;font-family:{FONT};font-size:13px;color:{THEME["muted"]};font-weight:600;">{k}</td>
        <td style="padding:7px 0;font-family:{FONT};font-size:13px;color:{THEME["text"]};font-weight:600;text-align:right;">{v}</td>
      </tr>''' for k, v in cr_info_rows)
    cr_card = email_card("#c0208f", "\U0001F4CB Change Request Information",
                          f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0">{cr_rows_html}</table>')

    def kpi_cell(num, label, color):
        return f'''<td width="25%" style="width:25%;background:{THEME["kpiBg"]};border:1px solid {THEME["line"]};border-radius:8px;padding:12px 4px;text-align:center;font-family:{FONT};">
        <div style="font-size:19px;font-weight:800;color:{color};line-height:1.2;">{num}</div>
        <div style="font-size:10px;color:{THEME["muted"]};font-weight:600;margin-top:2px;">{label}</div>
      </td>'''
    kpi_row = f'''<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:separate;border-spacing:4px 0;"><tr>
      {kpi_cell(total, 'Total Cases', THEME["text"])}
      {kpi_cell(passed, 'Passed', '#22a35a')}
      {kpi_cell(open_issues, 'Open Issues', '#e0a52c')}
      {kpi_cell(str(rate) + '%', 'Pass Rate', '#0e6ba8')}
    </tr></table>'''
    summary_body = f'''
      {kpi_row}
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:14px;">
        <tr>
          <td style="font-family:{FONT};font-size:12px;font-weight:600;color:{THEME["text"]};padding-bottom:5px;">Overall Completion</td>
          <td style="font-family:{FONT};font-size:12px;font-weight:600;color:{THEME["muted"]};text-align:right;padding-bottom:5px;">{executed} / {total} executed</td>
        </tr>
      </table>
      {email_bar(completion_pct, '#0e6ba8', THEME["trackBg"])}
    '''
    summary_card = email_card("#1f9d55", "\U0001F4C8 Test Execution Summary", summary_body)

    status_rows_html = ""
    for d in status_data:
        pct = (d["count"] / total * 100) if total else 0
        status_rows_html += f'''
        <tr>
          <td style="padding:9px 10px;border-bottom:1px solid {THEME["line"]};font-family:{FONT};font-size:13px;color:{THEME["text"]};white-space:nowrap;">{d["icon"]} {d["label"]}</td>
          <td style="padding:9px 10px;border-bottom:1px solid {THEME["line"]};font-family:{FONT};font-size:13px;color:{THEME["text"]};">{d["count"]}</td>
          <td style="padding:9px 10px;border-bottom:1px solid {THEME["line"]};font-family:{FONT};font-size:13px;color:{THEME["text"]};">{pct:.1f}%</td>
          <td width="130" style="padding:9px 10px;border-bottom:1px solid {THEME["line"]};width:130px;">{email_bar(pct, d["color"], THEME["trackBg"])}</td>
        </tr>'''
    status_body = f'''
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td style="padding:8px 10px;font-family:{FONT};font-size:11px;color:{THEME["muted"]};font-weight:700;border-bottom:2px solid {THEME["line"]};text-transform:uppercase;">Status</td>
          <td style="padding:8px 10px;font-family:{FONT};font-size:11px;color:{THEME["muted"]};font-weight:700;border-bottom:2px solid {THEME["line"]};text-transform:uppercase;">Count</td>
          <td style="padding:8px 10px;font-family:{FONT};font-size:11px;color:{THEME["muted"]};font-weight:700;border-bottom:2px solid {THEME["line"]};text-transform:uppercase;">%</td>
          <td width="130" style="padding:8px 10px;font-family:{FONT};font-size:11px;color:{THEME["muted"]};font-weight:700;border-bottom:2px solid {THEME["line"]};text-transform:uppercase;width:130px;">Share</td>
        </tr>
        {status_rows_html}
        <tr>
          <td style="padding:9px 10px;font-family:{FONT};font-size:13px;font-weight:700;background:{THEME["totalBg"]};color:{THEME["text"]};">\U0001F4CA Total</td>
          <td style="padding:9px 10px;font-family:{FONT};font-size:13px;font-weight:700;background:{THEME["totalBg"]};color:{THEME["text"]};">{total}</td>
          <td style="padding:9px 10px;font-family:{FONT};font-size:13px;font-weight:700;background:{THEME["totalBg"]};color:{THEME["text"]};">100.0%</td>
          <td style="padding:9px 10px;background:{THEME["totalBg"]};"></td>
        </tr>
      </table>'''
    status_card = email_card("#6a1fd0", "\U0001F4C4 Testing Status Details", status_body)

    legend_cells = []
    for d in status_data:
        pct = f'{(d["count"] / total * 100):.1f}' if total else '0.0'
        legend_cells.append(
            f'<span style="display:inline-block;width:9px;height:9px;border-radius:50%;background:{d["color"]};margin-right:5px;"></span>'
            f'<span style="font-family:{FONT};font-size:12px;color:{THEME["muted"]};font-weight:600;">{d["label"]}: {d["count"]} ({pct}%)</span>')
    legend_table = f'''<table role="presentation" cellpadding="0" cellspacing="0" style="margin:0 auto;"><tr>
        <td style="padding:3px 16px 3px 0;">{legend_cells[0]}</td><td style="padding:3px 0;">{legend_cells[1]}</td></tr>
      <tr><td style="padding:3px 16px 3px 0;">{legend_cells[2]}</td><td style="padding:3px 0;">{legend_cells[3]}</td></tr>
      <tr><td style="padding:3px 16px 3px 0;">{legend_cells[4]}</td><td style="padding:3px 0;">{legend_cells[5]}</td></tr>
    </table>'''
    donut_img_tag = f'<img src="{donut_src}" width="170" height="170" alt="Status distribution chart" style="display:block;margin:0 auto 12px;" />'
    dist_card = email_card("#6a1fd0", "\U0001F550 Testing Status Distribution",
                            f'<div style="text-align:center;">{donut_img_tag}{legend_table}</div>')

    if observations:
        obs_items = "".join(f'<li style="padding:4px 0;">{o}</li>' for o in observations)
        obs_body = f'<ul style="margin:0;padding-left:18px;font-family:{FONT};font-size:13px;color:{THEME["text"]};">{obs_items}</ul>'
    else:
        obs_body = f'<span style="font-family:{FONT};font-size:13px;color:{THEME["muted"]};font-style:italic;">No test observations provided</span>'
    obs_card = email_card("#6a1fd0", "\U0001F441 Test Observations", obs_body)

    cr_num_badge = next((v for k, v in cr_info_rows if "cr number" in k.lower()), "\u2014")
    sys_badge = next((v for k, v in cr_info_rows if k.lower() == "system"), "\u2014")

    def badge_cell(text):
        return f'<td style="background:rgba(255,255,255,.22);border:1px solid rgba(255,255,255,.5);border-radius:999px;padding:6px 14px;color:#ffffff;font-size:12px;font-weight:600;font-family:{FONT};white-space:nowrap;">{text}</td>'
    spacer_cell = '<td style="width:8px;line-height:1px;font-size:1px;">&nbsp;</td>'
    badges_table = f'''
      <table role="presentation" cellpadding="0" cellspacing="0" align="center" style="margin:14px auto 0;border-collapse:separate;border-spacing:0;">
        <tr>
          {badge_cell(cr_num_badge)}
          {spacer_cell}
          {badge_cell(sys_badge)}
          {spacer_cell}
          {badge_cell(str(rate) + '% Pass Rate')}
        </tr>
      </table>'''
    title_block = f'''
      <span style="color:#ffffff;font-size:20px;font-weight:700;font-family:{FONT};">\U0001F4CA {title}</span><br/>
      <span style="color:#ffffff;font-size:12.5px;opacity:.9;font-family:{FONT};">\U0001F4C5 Generated on {gen_date}</span>
      {badges_table}'''
    header_inner = f'''<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;">
          <tr>
            <td width="90" valign="middle" align="left" style="width:90px;padding-right:12px;">
              <img src="{logo_src}" alt="logo" style="display:block;max-height:44px;max-width:90px;background:#fff;padding:5px 8px;border-radius:8px;" />
            </td>
            <td valign="middle" align="center" style="text-align:center;">{title_block}</td>
            <td width="90" style="width:90px;">&nbsp;</td>
          </tr>
        </table>''' if logo_src else f'<div style="text-align:center;">{title_block}</div>'
    header = f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;"><tr><td style="background:#8a2fd6;padding:26px 24px;border-radius:10px;">{header_inner}</td></tr></table>'

    intro_block = ""
    if intro_html:
        intro_block = f'<div style="font-family:{FONT};font-size:14px;color:#1c2333;line-height:1.6;margin-bottom:16px;">{intro_html}</div><div style="height:20px;line-height:20px;font-size:1px;">&nbsp;</div>'

    fragment = f"{intro_block}{header}{cr_card}{summary_card}{status_card}{dist_card}{obs_card}"
    return f'''<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"><title>{title} - Email Ready</title></head>
<body style="margin:0;padding:20px;background:#f4f5f9;">
{fragment}
</body>
</html>
'''


if __name__ == "__main__":
    p = argparse.ArgumentParser()
    p.add_argument("--title", required=True)
    p.add_argument("--gen-date", required=True)
    p.add_argument("--logo", default=None)
    p.add_argument("--system", required=True)
    p.add_argument("--cr-number", required=True)
    p.add_argument("--tester", required=True)
    p.add_argument("--test-date", required=True)
    p.add_argument("--environment", required=True)
    p.add_argument("--status", required=True)
    p.add_argument("--pass", dest="p_pass", type=int, default=0)
    p.add_argument("--fail", type=int, default=0)
    p.add_argument("--blocked", type=int, default=0)
    p.add_argument("--inprogress", type=int, default=0)
    p.add_argument("--invalid", type=int, default=0)
    p.add_argument("--notrun", type=int, default=0)
    p.add_argument("--out", required=True)
    p.add_argument("--observations", nargs="*", default=None,
                    help="One or more observation lines. Omit entirely if there are none.")
    p.add_argument("--short-title", required=True,
                    help="Short form used in the intro sentence, e.g. 'E2E', 'IOT', or a custom value.")
    args = p.parse_args()

    cr_info_rows = [
        ("System", args.system), ("CR Number", args.cr_number),
        ("Tester", args.tester), ("Test Date", args.test_date),
        ("Environment", args.environment), ("Status", args.status),
    ]
    status_data = [
        {"key": "pass", "icon": "\u2705", "label": "Pass", "count": args.p_pass, "color": "#22a35a"},
        {"key": "fail", "icon": "\u274c", "label": "Fail", "count": args.fail, "color": "#e0384c"},
        {"key": "blocked", "icon": "\u26a0\ufe0f", "label": "Blocked", "count": args.blocked, "color": "#e0a52c"},
        {"key": "inprogress", "icon": "\U0001F550", "label": "In Progress", "count": args.inprogress, "color": "#3b8fe0"},
        {"key": "invalid", "icon": "\u2753", "label": "Invalid", "count": args.invalid, "color": "#a15fe0"},
        {"key": "notrun", "icon": "\u2296", "label": "Not Run", "count": args.notrun, "color": "#8b93a3"},
    ]
    total = sum(d["count"] for d in status_data)
    p_pass = args.p_pass
    rate = round(p_pass / total * 100) if total else 0
    intro_html = build_intro_text(args.short_title, args.cr_number, rate)

    subject = f"{args.system}-{args.cr_number}"
    print(f"SUBJECT: {subject}")

    html = build_email_html(args.title, args.gen_date, args.logo, cr_info_rows, status_data, intro_html=intro_html, observations=args.observations)
    with open(args.out, "w") as f:
        f.write(html)
    print(f"Wrote {args.out}")
