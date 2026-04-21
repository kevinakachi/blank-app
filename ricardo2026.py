import streamlit as st
import json
import os
import base64
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import date, timedelta
from pathlib import Path
import io

# ── WeasyPrint for PDF generation ──
try:
    from weasyprint import HTML as WP_HTML
    HAS_WEASYPRINT = True
except Exception:
    HAS_WEASYPRINT = False

# ── Paths ── (works locally and on Streamlit Cloud)
DATA_DIR = Path(__file__).parent / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)
COMPANY_FILE = DATA_DIR / "company.json"
CUSTOMERS_FILE = DATA_DIR / "customers.json"
LOGO_FILE = DATA_DIR / "logo_b64.txt"

# ══════════════════════════════════════
# Data helpers
# ══════════════════════════════════════

def load_json(path, default):
    if path.exists():
        try:
            return json.loads(path.read_text())
        except Exception:
            pass
    return default

def save_json(path, data):
    path.write_text(json.dumps(data, indent=2))

def load_company():
    return load_json(COMPANY_FILE, {
        "name": "Bondi Produce Co. Ltd",
        "tagline": "",
        "address": "188 New Toronto Street",
        "city": "Toronto, ON M8V 2EB",
        "phone": "p. 416-252-7799",
        "email": "",
        "pay_terms": "NET 10 DAYS",
        "sale_terms": "FOB SALE",
        "ship_via": "FOB SALE",
        "carrier": "CUSTOMER'S TRUCK",
        "salesperson": "",
        "smtp_user": "",
        "smtp_pass": "",
    })

def save_company(data):
    save_json(COMPANY_FILE, data)

def load_customers():
    return load_json(CUSTOMERS_FILE, {})

def save_customers(data):
    save_json(CUSTOMERS_FILE, data)

def load_logo_b64():
    if LOGO_FILE.exists():
        return LOGO_FILE.read_text().strip()
    return ""

def save_logo_b64(b64: str):
    LOGO_FILE.write_text(b64)

# ══════════════════════════════════════
# Invoice HTML template
# ══════════════════════════════════════

def build_invoice_html(inv, company, logo_b64):
    logo_html = ""
    if logo_b64:
        logo_html = f'<img class="logo-img" src="data:image/png;base64,{logo_b64}" alt="logo">'
    else:
        logo_html = '<div class="logo-placeholder">Your Logo</div>'

    rows_html = ""
    for item in inv["items"]:
        qty   = item.get("qty", "")
        pack  = item.get("pack", "")
        desc  = item.get("desc", "")
        price = item.get("price", "")
        try:
            price_fmt = f"$ {float(price):,.2f}"
        except Exception:
            price_fmt = price
        rows_html += f"""
        <tr>
          <td class="center">{qty}</td>
          <td class="center">{pack}</td>
          <td>{desc}</td>
          <td class="right">{price_fmt}</td>
        </tr>"""

    # empty filler rows
    filler = max(0, 5 - len(inv["items"]))
    for _ in range(filler):
        rows_html += "<tr class='empty'><td></td><td></td><td></td><td></td></tr>"

    total_qty = sum(int(i.get("qty") or 0) for i in inv["items"])
    try:
        total_price = sum(
            int(i.get("qty") or 0) * float(i.get("price") or 0)
            for i in inv["items"]
        )
        total_price_fmt = f"$ {total_price:,.2f}"
    except Exception:
        total_price_fmt = ""

    co_name    = company.get("name", "")
    co_tagline = company.get("tagline", "")
    co_addr    = company.get("address", "")
    co_city    = company.get("city", "")
    co_phone   = company.get("phone", "")
    co_header  = co_name
    if co_tagline:
        co_header += f"<br><span class='tagline'>{co_tagline}</span>"

    ship_date = inv.get("ship_date", "")
    del_date  = inv.get("del_date", "")

    cust_name  = inv.get("cust_name", "")
    cust_addr  = inv.get("cust_addr", "").replace("\n", "<br>")
    cust_phone = inv.get("cust_phone", "")

    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: Georgia, serif; font-size: 12px; color: #1a1a1a; background: #fff; padding: 32px; }}
  .header {{ display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 14px; }}
  .logo-wrap {{ display: flex; align-items: center; gap: 14px; }}
  .logo-img {{ height: 64px; width: auto; object-fit: contain; }}
  .logo-placeholder {{ width: 80px; height: 64px; background: #eee; border: 1px solid #ccc; border-radius: 4px;
    display: flex; align-items: center; justify-content: center; font-size: 10px; color: #999; text-align: center; padding: 4px; }}
  .co-info {{ font-size: 11px; line-height: 1.6; color: #333; }}
  .co-name {{ font-size: 15px; font-weight: bold; font-family: sans-serif; }}
  .tagline {{ font-size: 11px; color: #666; }}
  .inv-meta {{ text-align: right; font-size: 11px; line-height: 1.8; }}
  .inv-label {{ font-size: 10px; color: #888; text-transform: uppercase; letter-spacing: 0.06em; }}
  .inv-num {{ font-size: 22px; font-weight: bold; font-family: sans-serif; letter-spacing: -0.5px; }}
  .inv-meta td:first-child {{ color: #888; padding-right: 10px; }}
  .inv-meta td {{ padding: 1px 0; }}
  hr.thick {{ border: none; border-top: 2px solid #1a1a1a; margin: 10px 0 8px; }}
  .addr-row {{ display: flex; gap: 20px; margin-bottom: 12px; }}
  .addr-box {{ flex: 1; font-size: 11px; line-height: 1.6; }}
  .addr-label {{ background: #d4831a; color: #fff; font-weight: bold; font-size: 10px;
    padding: 2px 8px; display: inline-block; margin-bottom: 4px; letter-spacing: 0.04em; }}
  .meta-bar {{ display: grid; grid-template-columns: repeat(5, 1fr);
    border: 1px solid #1a1a1a; font-size: 10px; text-align: center; margin-bottom: 1px; }}
  .meta-bar .head {{ background: #1a1a1a; color: #fff; padding: 4px 2px; font-weight: bold; }}
  .meta-bar .cell {{ padding: 4px 2px; border-left: 1px solid #1a1a1a; }}
  table.items {{ width: 100%; border-collapse: collapse; font-size: 11px; margin-top: 8px; }}
  table.items th {{ background: #1a1a1a; color: #fff; padding: 5px 6px; text-align: left;
    font-weight: bold; font-size: 10px; text-transform: uppercase; letter-spacing: 0.04em; }}
  table.items th.right {{ text-align: right; }}
  table.items th.center {{ text-align: center; }}
  table.items td {{ padding: 4px 6px; border-bottom: 0.5px solid #e0e0e0; }}
  table.items tr:nth-child(even) td {{ background: #f9f9f9; }}
  table.items tr.empty td {{ border-bottom: 0.5px solid #ebebeb; height: 18px; }}
  .center {{ text-align: center; }}
  .right {{ text-align: right; }}
  .footer {{ display: flex; justify-content: space-between; align-items: flex-end;
    margin-top: 8px; border-top: 2px solid #1a1a1a; padding-top: 6px; }}
  .total-qty {{ font-size: 20px; font-weight: bold; font-family: sans-serif; }}
  .total-label {{ font-size: 10px; color: #888; text-transform: uppercase; letter-spacing: 0.06em; text-align: right; }}
  .total-amt {{ font-size: 22px; font-weight: bold; font-family: sans-serif; text-align: right; }}
  .datestamp {{ font-size: 9px; color: #bbb; text-align: right; margin-top: 4px; }}
  .message-box {{ margin-top: 18px; padding: 10px 12px; border: 0.5px solid #ddd;
    border-radius: 4px; font-size: 11px; line-height: 1.6; color: #333; white-space: pre-wrap; }}
</style>
</head>
<body>
<div class="header">
  <div class="logo-wrap">
    {logo_html}
    <div class="co-info">
      <div class="co-name">{co_header}</div>
      {co_addr}<br>{co_city}<br>{co_phone}
    </div>
  </div>
  <div class="inv-meta">
    <div class="inv-label">Invoice #</div>
    <div class="inv-num">{inv.get('inv_num','')}</div>
    <table>
      <tr><td>Ship Date:</td><td>{ship_date}</td></tr>
      <tr><td>Delivery Date:</td><td>{del_date}</td></tr>
      <tr><td>Pay Terms:</td><td>{company.get('pay_terms','')}</td></tr>
      <tr><td>Sale Terms:</td><td>{company.get('sale_terms','')}</td></tr>
    </table>
  </div>
</div>

<hr class="thick">

<div class="addr-row">
  <div class="addr-box">
    <span class="addr-label">Bill To:</span><br>
    {cust_name}<br>{cust_addr}<br>{cust_phone}
  </div>
  <div class="addr-box">
    <span class="addr-label">Ship To:</span><br>
    {cust_name}<br>{cust_addr}
  </div>
</div>

<div class="meta-bar">
  <div class="head">Cust PO</div>
  <div class="head cell">Salesperson</div>
  <div class="head cell">Ship Via</div>
  <div class="head cell">Carrier</div>
  <div class="head cell">Trailer / St.</div>
</div>
<div class="meta-bar">
  <div class="cell" style="border-left:none">{inv.get('cust_po','')}</div>
  <div class="cell">{company.get('salesperson','')}</div>
  <div class="cell">{company.get('ship_via','')}</div>
  <div class="cell">{company.get('carrier','')}</div>
  <div class="cell"></div>
</div>

<table class="items">
  <thead>
    <tr>
      <th class="center" style="width:50px">Shipped</th>
      <th class="center" style="width:65px">Pack Size</th>
      <th>Description</th>
      <th class="right" style="width:90px">Price</th>
    </tr>
  </thead>
  <tbody>
    {rows_html}
  </tbody>
</table>

<div class="footer">
  <div class="total-qty">{total_qty}</div>
  <div>
    <div class="total-label">Total</div>
    <div class="total-amt">{total_price_fmt}</div>
  </div>
</div>
<div class="datestamp">{date.today().strftime('%m/%d/%y')}</div>

{"<div class='message-box'>" + inv.get('message','') + "</div>" if inv.get('message','').strip() else ""}

</body>
</html>"""


# ══════════════════════════════════════
# PDF generation
# ══════════════════════════════════════

def html_to_pdf(html: str) -> bytes:
    if HAS_WEASYPRINT:
        buf = io.BytesIO()
        WP_HTML(string=html).write_pdf(buf)
        return buf.getvalue()
    # Fallback: return the HTML as bytes so the user can print-to-PDF from browser
    return html.encode("utf-8")

def pdf_filename(inv_num):
    return f"invoice_{inv_num}.{'pdf' if HAS_WEASYPRINT else 'html'}"

# ══════════════════════════════════════
# Email
# ══════════════════════════════════════

def send_invoice_email(smtp_user, smtp_pass, recipients, subject, body, pdf_bytes, inv_num):
    msg = MIMEMultipart()
    msg["From"]    = smtp_user
    msg["To"]      = ", ".join(recipients)
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain"))

    fname = pdf_filename(inv_num)
    mime_type = "pdf" if HAS_WEASYPRINT else "html"
    att = MIMEApplication(pdf_bytes, _subtype=mime_type)
    att.add_header("Content-Disposition", "attachment", filename=fname)
    msg.attach(att)

    ctx = ssl.create_default_context()
    with smtplib.SMTP("smtp.office365.com", 587) as server:
        server.ehlo()
        server.starttls(context=ctx)
        server.login(smtp_user, smtp_pass)
        server.sendmail(smtp_user, recipients, msg.as_string())


# ══════════════════════════════════════
# Session state init
# ══════════════════════════════════════

def init_state():
    if "items" not in st.session_state or not isinstance(st.session_state.items, list):
        st.session_state.items = [{"qty": "", "pack": "CASE", "desc": "", "price": ""}]
    # Ensure every item is a dict (guard against corrupt state)
    st.session_state.items = [
        i if isinstance(i, dict) else {"qty": "", "pack": "CASE", "desc": "", "price": ""}
        for i in st.session_state.items
    ]
    if "company" not in st.session_state:
        st.session_state.company = load_company()
    if "customers" not in st.session_state:
        st.session_state.customers = load_customers()
    if "logo_b64" not in st.session_state:
        st.session_state.logo_b64 = load_logo_b64()
    if "selected_customer" not in st.session_state:
        st.session_state.selected_customer = ""
    if "email_status" not in st.session_state:
        st.session_state.email_status = None


# ══════════════════════════════════════
# UI helpers
# ══════════════════════════════════════

def company_settings_tab():
    st.subheader("Company information")
    co = st.session_state.company

    col1, col2 = st.columns(2)
    with col1:
        co["name"]      = st.text_input("Company name",    co.get("name",""))
        co["tagline"]   = st.text_input("Tagline / subtitle", co.get("tagline",""))
        co["address"]   = st.text_area("Address", co.get("address",""), height=80)
        co["city"]      = st.text_input("City / Province / Postal", co.get("city",""))
        co["phone"]     = st.text_area("Phone / contact", co.get("phone",""), height=68)
    with col2:
        co["pay_terms"]  = st.text_input("Default pay terms",  co.get("pay_terms","NET 10 DAYS"))
        co["sale_terms"] = st.text_input("Default sale terms", co.get("sale_terms","FOB SALE"))
        co["ship_via"]   = st.text_input("Default ship via",   co.get("ship_via","FOB SALE"))
        co["carrier"]    = st.text_input("Default carrier",    co.get("carrier","CUSTOMER'S TRUCK"))
        co["salesperson"]= st.text_input("Default salesperson",co.get("salesperson",""))

    st.markdown("---")
    st.subheader("Logo")
    uploaded = st.file_uploader("Upload your logo", type=["png","jpg","jpeg","gif","webp"])
    if uploaded:
        b64 = base64.b64encode(uploaded.read()).decode()
        st.session_state.logo_b64 = b64
        save_logo_b64(b64)
        st.success("Logo saved.")
    if st.session_state.logo_b64:
        st.image(base64.b64decode(st.session_state.logo_b64), width=160)
        if st.button("Remove logo"):
            st.session_state.logo_b64 = ""
            if LOGO_FILE.exists(): LOGO_FILE.unlink()

    st.markdown("---")
    st.subheader("Email credentials (Outlook / Office 365)")
    co["smtp_user"] = st.text_input("Your Outlook email address", co.get("smtp_user",""))
    co["smtp_pass"] = st.text_input("App password", co.get("smtp_pass",""), type="password",
        help="Use an app-specific password from your Microsoft account security settings.")

    if st.button("💾  Save company settings", type="primary"):
        save_company(co)
        st.session_state.company = co
        st.success("Company settings saved!")


def customers_tab():
    st.subheader("Saved customers")
    customers = st.session_state.customers

    with st.expander("➕  Add / update a customer", expanded=len(customers)==0):
        c1, c2 = st.columns(2)
        with c1:
            new_name  = st.text_input("Customer name*", key="new_cust_name")
            new_addr  = st.text_area("Address", key="new_cust_addr", height=80,
                                     placeholder="188 New Toronto Street\nToronto, ON M8V 2EB")
        with c2:
            new_phone = st.text_input("Phone", key="new_cust_phone")
            new_email = st.text_input("Email(s) — separate multiple with commas", key="new_cust_email")
            new_po    = st.text_input("Default PO #", key="new_cust_po")
        if st.button("Save customer", type="primary"):
            if not new_name.strip():
                st.error("Customer name is required.")
            else:
                customers[new_name.strip()] = {
                    "name":  new_name.strip(),
                    "addr":  new_addr,
                    "phone": new_phone,
                    "email": new_email,
                    "po":    new_po,
                }
                save_customers(customers)
                st.session_state.customers = customers
                st.success(f"'{new_name.strip()}' saved.")

    if customers:
        st.markdown("---")
        for cname, cdata in sorted(customers.items()):
            with st.expander(cname):
                cc1, cc2 = st.columns([3,1])
                with cc1:
                    st.write(f"**Address:** {cdata.get('addr','').replace(chr(10),', ')}")
                    st.write(f"**Phone:** {cdata.get('phone','')}")
                    st.write(f"**Email:** {cdata.get('email','')}")
                    st.write(f"**Default PO:** {cdata.get('po','')}")
                with cc2:
                    if st.button("🗑 Remove", key=f"del_{cname}"):
                        del customers[cname]
                        save_customers(customers)
                        st.session_state.customers = customers
                        st.rerun()
    else:
        st.info("No customers saved yet.")


def invoice_tab():
    co        = st.session_state.company
    customers = st.session_state.customers
    logo_b64  = st.session_state.logo_b64

    # ── Customer picker ──
    st.subheader("Customer")
    cust_options = ["— type manually —"] + sorted(customers.keys())
    sel = st.selectbox("Select saved customer", cust_options, key="cust_picker")

    if sel != "— type manually —" and sel in customers:
        cd = customers[sel]
        cust_name  = cd.get("name","")
        cust_addr  = cd.get("addr","")
        cust_phone = cd.get("phone","")
        cust_email = cd.get("email","")
        cust_po    = cd.get("po","")
    else:
        cust_name = cust_addr = cust_phone = cust_email = cust_po = ""

    ci1, ci2 = st.columns(2)
    with ci1:
        cust_name  = st.text_input("Customer name",  value=cust_name,  key="inv_cust_name")
        cust_addr  = st.text_area("Address",          value=cust_addr,  key="inv_cust_addr", height=80)
    with ci2:
        cust_phone = st.text_input("Phone",           value=cust_phone, key="inv_cust_phone")
        cust_email = st.text_input("Email(s) for this invoice", value=cust_email, key="inv_cust_email",
                                   help="Separate multiple addresses with commas")

    st.markdown("---")

    # ── Invoice header ──
    st.subheader("Invoice details")
    h1, h2, h3, h4 = st.columns(4)
    with h1:
        inv_num   = st.text_input("Invoice #", value="00000001", key="inv_num")
    with h2:
        cust_po   = st.text_input("Cust PO", value=cust_po, key="inv_po")
    with h3:
        ship_date = st.date_input("Ship date",     value=date.today(),             key="inv_ship")
    with h4:
        del_date  = st.date_input("Delivery date", value=date.today()+timedelta(3), key="inv_del")

    st.markdown("---")

    # ── Line items ──
    st.subheader("Line items")
    header_cols = st.columns([1, 1.4, 4, 1.5, 0.5])
    header_cols[0].markdown("**Qty**")
    header_cols[1].markdown("**Pack**")
    header_cols[2].markdown("**Description**")
    header_cols[3].markdown("**Price / case ($)**")

    items = st.session_state.items
    to_delete = None
    for i, item in enumerate(items):
        c1, c2, c3, c4, c5 = st.columns([1, 1.4, 4, 1.5, 0.5])
        item["qty"]   = c1.text_input("", value=item.get("qty",""),  key=f"qty_{i}",  label_visibility="collapsed")
        item["pack"]  = c2.text_input("", value=item.get("pack","CASE"), key=f"pack_{i}", label_visibility="collapsed")
        item["desc"]  = c3.text_input("", value=item.get("desc",""), key=f"desc_{i}", label_visibility="collapsed")
        item["price"] = c4.text_input("", value=item.get("price",""),key=f"price_{i}",label_visibility="collapsed")
        if c5.button("✕", key=f"del_item_{i}"):
            to_delete = i

    if to_delete is not None:
        items.pop(to_delete)
        st.rerun()

    if st.button("＋  Add line item"):
        items.append({"qty":"","pack":"CASE","desc":"","price":""})
        st.rerun()

    # ── Totals summary ──
    try:
        total_qty   = sum(int(it.get("qty") or 0) for it in items)
        total_price = sum(int(it.get("qty") or 0) * float(it.get("price") or 0) for it in items)
        t1, t2 = st.columns([3,1])
        t1.metric("Total cases", total_qty)
        t2.metric("Invoice total", f"${total_price:,.2f}")
    except Exception:
        pass

    st.markdown("---")

    # ── Optional message ──
    st.subheader("Message (optional)")
    message = st.text_area(
        "Add a note to include on the invoice and email body",
        value="",
        height=90,
        placeholder="e.g. Thank you for your business! Please remit payment within 10 days.",
        key="inv_message"
    )

    # ── Build invoice data dict ──
    inv = {
        "inv_num":    inv_num,
        "cust_po":    cust_po,
        "ship_date":  ship_date.strftime("%m/%d/%y"),
        "del_date":   del_date.strftime("%m/%d/%y"),
        "cust_name":  cust_name,
        "cust_addr":  cust_addr,
        "cust_phone": cust_phone,
        "items":      items,
        "message":    message,
    }

    html = build_invoice_html(inv, co, logo_b64)

    # ── Preview ──
    st.markdown("---")
    st.subheader("Preview")
    st.components.v1.html(html, height=820, scrolling=True)

    # ── Download ──
    pdf_bytes = html_to_pdf(html)
    fname = pdf_filename(inv_num)
    mime  = "application/pdf" if HAS_WEASYPRINT else "text/html"
    st.download_button(
        label=f"⬇️  Download {'PDF' if HAS_WEASYPRINT else 'HTML (print to PDF)'}",
        data=pdf_bytes,
        file_name=fname,
        mime=mime,
    )

    st.markdown("---")

    # ── Send email ──
    st.subheader("Send by email")
    smtp_user = co.get("smtp_user","")
    smtp_pass = co.get("smtp_pass","")

    if not smtp_user or not smtp_pass:
        st.warning("Set your Outlook email and app password in the **Company Settings** tab first.")
    else:
        default_recipients = cust_email
        recipients_raw = st.text_input(
            "Send to (comma-separated email addresses)",
            value=default_recipients,
            key="email_recipients"
        )
        default_subject = f"Invoice #{inv_num} — {cust_name}"
        email_subject = st.text_input("Subject", value=default_subject, key="email_subject")

        default_body = f"Hi,\n\nPlease find attached Invoice #{inv_num}."
        if message.strip():
            default_body += f"\n\n{message.strip()}"
        default_body += f"\n\nTotal: ${total_price:,.2f}\n\nThank you for your business.\n\n{co.get('name','')}"

        email_body = st.text_area("Email body", value=default_body, height=160, key="email_body")

        cc_self = st.checkbox("CC myself", value=True, key="cc_self")

        if st.button("📧  Send invoice", type="primary"):
            recipients = [r.strip() for r in recipients_raw.split(",") if r.strip()]
            if cc_self and smtp_user not in recipients:
                recipients.append(smtp_user)
            if not recipients:
                st.error("Please enter at least one recipient email.")
            else:
                with st.spinner("Sending…"):
                    try:
                        send_invoice_email(
                            smtp_user, smtp_pass,
                            recipients,
                            email_subject, email_body,
                            pdf_bytes, inv_num
                        )
                        st.session_state.email_status = ("success", f"Invoice sent to: {', '.join(recipients)}")
                    except Exception as e:
                        st.session_state.email_status = ("error", str(e))

    if st.session_state.email_status:
        status, msg = st.session_state.email_status
        if status == "success":
            st.success(msg)
        else:
            st.error(f"Failed to send: {msg}")


# ══════════════════════════════════════
# Main
# ══════════════════════════════════════

st.set_page_config(page_title="Invoice Builder", page_icon="🧾", layout="wide")

st.markdown("""
<style>
    .block-container { padding-top: 1.5rem; }
    [data-testid="stMetricValue"] { font-size: 1.4rem; }
</style>
""", unsafe_allow_html=True)

st.title("🧾 Invoice Builder")

init_state()

tab1, tab2, tab3 = st.tabs(["📄  Create Invoice", "👥  Customers", "🏢  Company Settings"])

with tab1:
    invoice_tab()
with tab2:
    customers_tab()
with tab3:
    company_settings_tab()
