
import streamlit as st
import pandas as pd
import datetime
import io
import msal
import requests
import base64
import html

st.set_page_config(
    page_title="Uncommunicated Subs and Cuts Report",
    page_icon="📊",
    layout="centered"
)

st.title("📊 Uncommunicated Subs and Cuts Report")

# ============================================================
# CONFIGURATION
# ============================================================

SCOPES = ["User.Read", "Mail.Send"]

GRAPH_URL = "https://graph.microsoft.com/v1.0"

# ============================================================
# BUYER DIRECTORY + LOGIN
# ============================================================

buyers = st.secrets["buyers"]

if "logged_in_initials" not in st.session_state:
    st.session_state.logged_in_initials = None

if st.session_state.logged_in_initials is None:

    st.subheader("Who's running this report?")

    initials_input = st.selectbox(
        "Select your initials",
        options=list(buyers.keys())
    )

    if st.button("Continue"):
        st.session_state.logged_in_initials = initials_input
        st.rerun()

    st.stop()

current_initials = st.session_state.logged_in_initials
current_buyer = buyers[current_initials]

st.success("Buyer login successful!")
st.write("Selected initials:", current_initials)
st.write("Buyer name:", current_buyer["name"])

with st.sidebar:

    st.write(
        f"Running as **{current_buyer['name']}** "
        f"({current_initials})"
    )

    if st.button("Switch user"):
        st.session_state.logged_in_initials = None
        st.rerun()


# ============================================================
# MICROSOFT GRAPH AUTHENTICATION
# ============================================================

def get_msal_client():

    cache = msal.SerializableTokenCache()

    saved_cache = st.session_state.get("msal_cache")

    if saved_cache:
        cache.deserialize(saved_cache)

    client = msal.PublicClientApplication(
        client_id=st.secrets["azure"]["client_id"],
        authority=(
            "https://login.microsoftonline.com/"
            + st.secrets["azure"]["tenant_id"]
        ),
        token_cache=cache
    )

    return client, cache


def connect_microsoft():

    try:

        client, cache = get_msal_client()

        flow = client.initiate_device_flow(
            scopes=SCOPES
        )

        if "user_code" not in flow:
            st.error("Unable to start Microsoft authentication.")
            return

        st.info(
            "Open https://microsoft.com/devicelogin "
            "in another browser tab."
        )

        st.subheader("Your Microsoft login code")

        st.code(flow["user_code"])

        st.warning(
            "Enter the code on Microsoft's website, "
            "sign in with your company account, "
            "and approve the Authenticator request."
        )

        with st.spinner("Waiting for Microsoft authentication..."):

            result = client.acquire_token_by_device_flow(flow)

        if "access_token" in result:

            st.session_state["msal_cache"] = cache.serialize()

            account = result.get("id_token_claims", {})

            st.session_state["msal_account"] = (
                account.get("preferred_username")
                or account.get("upn")
                or "Microsoft account"
            )

            st.success("Microsoft authentication successful!")

            st.rerun()

        else:

            st.error(
                result.get(
                    "error_description",
                    result.get("error", "Authentication failed")
                )
            )

    except Exception as e:

        st.error(f"Authentication error: {e}")


def get_graph_token():

    try:

        client, cache = get_msal_client()

        accounts = client.get_accounts()

        if not accounts:
            return None

        result = client.acquire_token_silent(
            SCOPES,
            account=accounts[0]
        )

        if cache.has_state_changed:
            st.session_state["msal_cache"] = cache.serialize()

        if result and "access_token" in result:
            return result["access_token"]

        return None

    except Exception:
        return None


def logout_microsoft():

    st.session_state.pop("msal_cache", None)
    st.session_state.pop("msal_account", None)

    st.rerun()

# ============================================================
# REPORT PROCESSING
# ============================================================

def process_files(csv_file, excel_file):

    main_df = pd.read_csv(csv_file, dtype=str)

    subs_df = pd.read_excel(
        excel_file,
        sheet_name="Short Sheet",
        dtype=str,
        header=2,
        engine="openpyxl"
    )

    main_df.loc[
        (main_df["VEND"] == "NEWTOR")
        & (
            main_df["BUY"].isna()
            | (main_df["BUY"].str.strip() == "")
        ),
        "BUY"
    ] = "NTF"

    values_to_remove = [
        'SUBTOTAL FOR :',
        'SUBTOTAL FOR KAM:',
        'SUBTOTAL FOR KDM:',
        'SUBTOTAL FOR LA:',
        'SUBTOTAL FOR MD:',
        'SUBTOTAL FOR SSW:',
        'GRAND TOTAL:',
        'SUBTOTAL FOR VGS:',
        'SUBTOTAL FOR AMC:',
        'KK',
        'SUBTOTAL FOR KK:',
        'RM',
        'SUBTOTAL FOR RM:',
        'EB',
        'SUBTOTAL FOR MPG:',
        'SUBTOTAL FOR DD:',
        'SUBTOTAL FOR KSM:',
        'HD',
        'SUBTOTAL FOR HD:'
    ]

    main_df = main_df[
        ~main_df["BUY"].isin(values_to_remove)
    ]

    main_df = main_df[
        ~main_df["REASON CODE"].isin(["BO"])
    ]

    main_df = main_df[
        ~main_df["REASON CODE"].isin(["LT"])
    ]

    main_df = main_df[
        ~main_df["WH"].isin(["OT"])
    ]

    main_df["ORIG. PROD"] = pd.to_numeric(
        main_df["ORIG. PROD"],
        errors="coerce"
    )

    subs_df["SKU #"] = pd.to_numeric(
        subs_df["SKU #"],
        errors="coerce"
    )

    main_df = main_df.dropna(
        subset=["ORIG. PROD"]
    )

    subs_df = subs_df.dropna(
        subset=["SKU #"]
    )

    main_df["ORIG. PROD"] = (
        main_df["ORIG. PROD"].astype(int).astype(str)
    )

    subs_df["SKU #"] = (
        subs_df["SKU #"].astype(int).astype(str)
    )

    blank_buy_count = main_df["BUY"].isna().sum()

    mismatched_df = main_df[
        ~main_df["ORIG. PROD"].isin(subs_df["SKU #"])
    ]

    today_date = datetime.datetime.today().strftime("%d-%b-%y")

    result = (
        mismatched_df.groupby("BUY")
        .size()
        .reset_index(name=today_date)
    )

    all_buyers = [
        "AMC", "VGS", "SSW", "KAM", "DD",
        "LA", "MD", "KSM", "GEC", "NTF"
    ]

    result = (
        pd.DataFrame({"BUY": all_buyers})
        .merge(result, on="BUY", how="left")
        .fillna(0)
    )

    result[today_date] = result[today_date].astype(int)

    if blank_buy_count > 0:

        blank_row = pd.DataFrame({
            "BUY": ["NOT ASSIGNED"],
            today_date: [blank_buy_count]
        })

        result = pd.concat(
            [result, blank_row],
            ignore_index=True
        )

    missed_products_df = mismatched_df[
        [
            "BUY",
            "ORIG. PROD",
            "ORIG PROD DESC",
            "CUSTOMER NAME",
            "SALES ORDER"
        ]
    ].drop_duplicates()

    return result, missed_products_df, today_date


# ============================================================
# EMAIL SENDING - MICROSOFT GRAPH API
# ============================================================

def send_email(
    result,
    missed_products_df,
    today_date,
    excel_file,
    generated_by_initials,
    generated_by_name
):

    try:

        access_token = get_graph_token()

        if not access_token:

            return (
                "❌ Microsoft authentication expired or missing. "
                "Please reconnect your Microsoft account."
            )

        sender_email = st.secrets["smtp"]["sender_email"]

        recipient_email = st.secrets["email"]["recipients"]

        if isinstance(recipient_email, str):

            recipient_email = [
                email.strip()
                for email in recipient_email.split(",")
                if email.strip()
            ]

        if not recipient_email:

            return "❌ No email recipients configured."

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }

        # Verify the signed-in Microsoft account

        me_response = requests.get(
            f"{GRAPH_URL}/me",
            headers=headers,
            timeout=30
        )

        if me_response.status_code != 200:

            return (
                "❌ Could not verify Microsoft account: "
                + me_response.text
            )

        me = me_response.json()

        authenticated_email = (
            me.get("mail")
            or me.get("userPrincipalName")
            or ""
        )

        if authenticated_email.lower() != sender_email.lower():

            return (
                f"❌ You are signed in as {authenticated_email}, "
                f"but the configured sender is {sender_email}. "
                "Please use the correct Microsoft account."
            )

        subject = (
            f"Uncommunicated Subs and Cuts Report - {today_date}"
        )

        # Build HTML email

        html_content = f"""
        <html>
        <body>

        <p>Good morning team,</p>

        <p>
        Uncommunicated subs and cuts from yesterday.
        Please follow up with the inventory and night team
        if you see any discrepancy.
        </p>

        <h3>Summary of Items</h3>

        <table border="1" cellpadding="4" cellspacing="0"
        style="border-collapse: collapse; width: 50%;">

        <tr style="background-color: #D0D8E8;">
            <th style="text-align: left;">BUYER</th>
            <th>{html.escape(str(today_date))}</th>
        </tr>
        """

        for _, row in result.iterrows():

            buyer = html.escape(str(row["BUY"]))
            count = html.escape(str(row[today_date]))

            html_content += (
                f"<tr><td>{buyer}</td>"
                f"<td style='text-align: center;'>{count}</td></tr>"
            )

        html_content += f"""
        </table>

        <p>
        <b>
        *Items for <u>THE PRODUCE COUNTER</u>
        and <u>LT</u> have been excluded from the report.
        </b>
        </p>

        <p style="color: #888888; font-size: 0.9em;">
        Report generated by
        {html.escape(str(generated_by_initials))}
        ({html.escape(str(generated_by_name))})
        </p>

        </body>
        </html>
        """

        # Create summary Excel attachment

        excel_buffer = io.BytesIO()

        with pd.ExcelWriter(
            excel_buffer,
            engine="xlsxwriter"
        ) as writer:

            missed_products_df.to_excel(
                writer,
                index=False,
                sheet_name="Summary of Items"
            )

        excel_buffer.seek(0)

        summary_attachment = base64.b64encode(
            excel_buffer.getvalue()
        ).decode("utf-8")

        # Original uploaded Excel attachment

        original_attachment = base64.b64encode(
            excel_file.getvalue()
        ).decode("utf-8")

        # Prepare email payload

        email_data = {
            "message": {
                "subject": subject,

                "body": {
                    "contentType": "HTML",
                    "content": html_content
                },

                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": email
                        }
                    }
                    for email in recipient_email
                ],

                "attachments": [
                    {
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "name": f"SummaryRP90_{today_date}.xlsx",
                        "contentType": (
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),
                        "contentBytes": summary_attachment
                    },
                    {
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "name": excel_file.name,
                        "contentType": (
                            excel_file.type
                            or "application/octet-stream"
                        ),
                        "contentBytes": original_attachment
                    }
                ]
            },

            "saveToSentItems": True
        }

        # Send email using Microsoft Graph

        response = requests.post(
            f"{GRAPH_URL}/me/sendMail",
            headers=headers,
            json=email_data,
            timeout=60
        )

        if response.status_code == 202:

            return (
                "✅ Email sent successfully! "
                f"(Report generated by {generated_by_initials})"
            )

        if response.status_code == 401:

            return (
                "❌ Microsoft authentication expired. "
                "Please reconnect your Microsoft account."
            )

        if response.status_code == 403:

            return (
                "❌ Microsoft denied permission to send this email. "
                "Check Mail.Send permissions and company security policies. "
                + response.text
            )

        return (
            f"❌ Microsoft Graph error "
            f"{response.status_code}: {response.text}"
        )

    except Exception as e:

        return f"❌ Error sending email: {str(e)}"
