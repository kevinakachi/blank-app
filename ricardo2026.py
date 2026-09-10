import streamlit as st
import pandas as pd
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import io

# ============================================================
# BUYER DIRECTORY + LOGIN
# ============================================================
# Buyers live in Streamlit Cloud secrets (Settings -> Secrets), e.g.:
#
# [buyers.SSW]
# name = "Steve S. Williams"
# email = "steve@bondiproduce.com"
# password = "Bondi2026-SSW"
#
# [buyers.KAM]
# name = "Kelly A. Martin"
# email = "kelly@bondiproduce.com"
# password = "Bondi2026-KAM"
#
# Add one [buyers.<INITIALS>] block per person. Initials are the login ID.

buyers = st.secrets["buyers"]

if "logged_in_initials" not in st.session_state:
    st.session_state.logged_in_initials = None

if st.session_state.logged_in_initials is None:
    st.title("📊 Uncommunicated Subs and Cuts Report")
    st.subheader("Log in")
    with st.form("login_form"):
        initials_input = st.text_input("Initials").strip().upper()
        password_input = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Log in")

    if submitted:
        buyer = buyers.get(initials_input)
        if buyer and password_input == buyer["password"]:
            st.session_state.logged_in_initials = initials_input
            st.rerun()
        else:
            st.error("Incorrect initials or password.")
    st.stop()

current_initials = st.session_state.logged_in_initials
current_buyer = buyers[current_initials]

with st.sidebar:
    st.write(f"Logged in as **{current_buyer['name']}** ({current_initials})")
    if st.button("Log out"):
        st.session_state.logged_in_initials = None
        st.rerun()


# ============================================================
# REPORT PROCESSING (unchanged from the original script)
# ============================================================
def process_files(csv_file, excel_file):
    main_df = pd.read_csv(csv_file, dtype=str)
    subs_df = pd.read_excel(excel_file, sheet_name="Short Sheet", dtype=str, header=2, engine="openpyxl")

    main_df.loc[(main_df["VEND"] == "NEWTOR") & (main_df["BUY"].isna() | (main_df["BUY"].str.strip() == "")), "BUY"] = "NTF"

    values_to_remove = ['SUBTOTAL FOR :', 'SUBTOTAL FOR KAM:', 'SUBTOTAL FOR KDM:', 'SUBTOTAL FOR LA:',
                        'SUBTOTAL FOR MD:', 'SUBTOTAL FOR SSW:', 'GRAND TOTAL:', 'SUBTOTAL FOR VGS:',
                        'SUBTOTAL FOR AMC:', 'KK', 'SUBTOTAL FOR KK:', 'RM', 'SUBTOTAL FOR RM:', 'EB',
                        'SUBTOTAL FOR MPG:', 'SUBTOTAL FOR DD:', 'SUBTOTAL FOR KSM:', 'HD', 'SUBTOTAL FOR HD:']
    main_df = main_df[~main_df['BUY'].isin(values_to_remove)]

    remove_reason_code = ['BO']
    main_df = main_df[~main_df['REASON CODE'].isin(remove_reason_code)]

    remove_LT = ['LT']
    main_df = main_df[~main_df['REASON CODE'].isin(remove_LT)]

    remove_OT = ['OT']
    main_df = main_df[~main_df['WH'].isin(remove_OT)]

    main_df["ORIG. PROD"] = pd.to_numeric(main_df["ORIG. PROD"], errors="coerce")
    subs_df["SKU #"] = pd.to_numeric(subs_df["SKU #"], errors="coerce")

    main_df = main_df.dropna(subset=["ORIG. PROD"])
    subs_df = subs_df.dropna(subset=["SKU #"])

    main_df["ORIG. PROD"] = main_df["ORIG. PROD"].astype(int).astype(str)
    subs_df["SKU #"] = subs_df["SKU #"].astype(int).astype(str)

    blank_buy_count = main_df["BUY"].isna().sum()

    mismatched_df = main_df[~main_df["ORIG. PROD"].isin(subs_df["SKU #"])]

    today_date = datetime.datetime.today().strftime("%d-%b-%y")
    result = mismatched_df.groupby("BUY").size().reset_index(name=today_date)

    all_buyers = ["AMC", "VGS", "SSW", "KAM", "DD", "LA", "MD", "KSM", "GEC", "NTF"]
    result = pd.DataFrame({"BUY": all_buyers}).merge(result, on="BUY", how="left").fillna(0)
    result[today_date] = result[today_date].astype(int)

    if blank_buy_count > 0:
        blank_row = pd.DataFrame({"BUY": ["NOT ASSIGNED"], today_date: [blank_buy_count]})
        result = pd.concat([result, blank_row], ignore_index=True)

    missed_products_df = mismatched_df[["BUY", "ORIG. PROD", "ORIG PROD DESC", "CUSTOMER NAME", "SALES ORDER"]].drop_duplicates()

    return result, missed_products_df, today_date


# ============================================================
# EMAIL SENDING (same SMTP account as before, attributed footer)
# ============================================================
def send_email(result, missed_products_df, today_date, excel_file, generated_by_initials, generated_by_name):
    sender_email = st.secrets["smtp"]["sender_email"]
    recipient_email = st.secrets["email"]["recipients"]
    subject = f"Uncommunicated Subs and Cuts Report - {today_date}"

    html_content = f"""
    <html>
    <body>
    <p>Good morning team,</p>
    <p>Uncommunicated subs and cuts from yesterday. Please follow up with the inventory and night team if you see any discrepancy.</p>

    <h3>Summary of Items</h3>
    <table border="1" cellpadding="4" cellspacing="0" style="border-collapse: collapse; width: 50%;">
        <tr style="background-color: #D0D8E8;">
            <th style="text-align: left;">BUYER</th>
            <th>{today_date}</th>
        </tr>
    """
    for _, row in result.iterrows():
        html_content += f"<tr><td>{row['BUY']}</td><td style='text-align: center;'>{row[today_date]}</td></tr>"

    html_content += f"""
    </table>
    <p><b>*Items for <u>THE PRODUCE COUNTER</u> and <u>LT</u> have been excluded from the report.</b></p>
    <p style="color: #888888; font-size: 0.9em;">Report generated by {generated_by_initials} ({generated_by_name})</p>
    </body>
    </html>
    """

    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
        missed_products_df.to_excel(writer, index=False, sheet_name="Summary of Items")
    excel_buffer.seek(0)

    message = MIMEMultipart()
    message["Subject"] = subject
    message["From"] = sender_email
    message["To"] = ", ".join(recipient_email)
    message.attach(MIMEText(html_content, "html"))

    excel_attachment = MIMEBase("application", "octet-stream")
    excel_attachment.set_payload(excel_buffer.getvalue())
    encoders.encode_base64(excel_attachment)
    excel_attachment.add_header("Content-Disposition", f"attachment; filename=SummaryRP90_{today_date}.xlsx")
    message.attach(excel_attachment)

    excel_file.seek(0)
    user_excel_attachment = MIMEBase("application", "octet-stream")
    user_excel_attachment.set_payload(excel_file.read())
    encoders.encode_base64(user_excel_attachment)
    user_excel_attachment.add_header("Content-Disposition", f"attachment; filename={excel_file.name}")
    message.attach(user_excel_attachment)

    try:
        with smtplib.SMTP("smtp.office365.com", 587) as server:
            server.starttls()
            server.login(sender_email, st.secrets["smtp"]["password"])
            server.sendmail(sender_email, recipient_email, message.as_string())
        return f"✅ Email sent successfully! (Report generated by {generated_by_initials})"
    except Exception as e:
        return f"❌ Error sending email: {e}"


# ============================================================
# STREAMLIT UI
# ============================================================
st.title("📊 Uncommunicated Subs and Cuts Report")
st.write("Upload your CSV and Excel files below to process the report.")

csv_file = st.file_uploader("Upload RP90 (CSV format)")
excel_file = st.file_uploader("Upload Subs & Cuts file (XLSM format)")

if "result" not in st.session_state:
    st.session_state.result = None
if "missed_products_df" not in st.session_state:
    st.session_state.missed_products_df = None
if "today_date" not in st.session_state:
    st.session_state.today_date = None

if csv_file and excel_file:
    if not csv_file.name.lower().endswith(".csv"):
        st.error("The uploaded RP90 file must be a .csv file.")
    elif not (excel_file.name.lower().endswith(".xlsx") or excel_file.name.lower().endswith(".xlsm")):
        st.error("The uploaded Subs & Cuts file must be an Excel file (.xlsx or .xlsm).")
    else:
        st.success("Files uploaded successfully!")

        if st.button("Process Files"):
            result, missed_products_df, today_date = process_files(csv_file, excel_file)
            st.session_state.result = result
            st.session_state.missed_products_df = missed_products_df
            st.session_state.today_date = today_date

            st.write("### Mismatched Items Report")
            st.dataframe(result)

            st.write("### Details of Missed Products")
            st.dataframe(missed_products_df)

if st.session_state.result is not None and st.session_state.missed_products_df is not None:
    if st.button("Send the Email"):
        email_status = send_email(
            st.session_state.result,
            st.session_state.missed_products_df,
            st.session_state.today_date,
            excel_file,
            current_initials,
            current_buyer["name"],
        )
        st.success(email_status)
