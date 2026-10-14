from datetime import date, datetime, time, timedelta
from pathlib import Path
import os, json, re, hashlib
import calendar
import pandas as pd
import streamlit as st
import altair as alt
from streamlit_webrtc import VideoProcessorBase, webrtc_streamer
st.set_page_config(page_title="Circular Smart Recycling",page_icon="♻️",layout="wide",initial_sidebar_state="expanded")
try:
    ACTIVE_THEME=st.context.theme.type
except Exception:
    ACTIVE_THEME=st.get_option("theme.base") or "light"
IS_DARK=ACTIVE_THEME=="dark"
PALETTE=({
    "page":"radial-gradient(ellipse at 85% 0%,#1b342a,transparent 36%),linear-gradient(145deg,#0c1310,#121d17 55%,#0e1712)","text":"#e7f0e9","muted":"#b1beb6","sidebar":"linear-gradient(180deg,#172920,#101c17)","sidebar_text":"#e5ede7","sidebar_muted":"#9bb5a4",
    "panel":"#16221b","panel2":"#1a2821","border":"#304338","accent":"#68bd87","hover":"#25392d","input":"#111a15","input_border":"#405348",
    "hero":"linear-gradient(115deg,#1b3326,#18231c 68%,#1c2d22)","hero_border":"#35513e","tip":"#192a20","eyebrow":"#91c89f",
    "empty_bin":"#17231d","bin_outline":"#8ba092","calendar":"#18241d","calendar_text":"#e1ebe4","calendar_border":"#34473b",
    "maintenance_bg":"#49391f","maintenance_text":"#f5d894","maintenance_border":"#8e703c","bin_card":"#17231d","bin_id":"#a4b8aa","bin_name":"#e7f0e9",
    "bin_detail":"#b1beb6","bin_value":"#d0dbd2","footer":"#9aada0","active_nav":"linear-gradient(90deg,#254d36,#203a2a)","alert":"#19251e","scheme":"dark"
} if IS_DARK else {
    "page":"radial-gradient(ellipse at 85% 0%,#e1e9de,transparent 32%),linear-gradient(145deg,#f5f5ef,#edf1e9 55%,#f8f7f2)","text":"#34443d","muted":"#68766e","sidebar":"linear-gradient(180deg,#f7faf6,#edf3ee)","sidebar_text":"#30463a","sidebar_muted":"#6c8677",
    "panel":"#ffffffb5","panel2":"#ffffffbb","border":"#dfe6dd","accent":"#28745f","hover":"#edf4ee","input":"#f0f2f6","input_border":"#d4ddd5",
    "hero":"linear-gradient(115deg,#e6ece1,#f4f3eb 68%,#e9eee4)","hero_border":"#d8e2d6","tip":"#eff4ea","eyebrow":"#6c8677",
    "empty_bin":"#ffffff","bin_outline":"#73847a","calendar":"#ffffff","calendar_text":"#34443d","calendar_border":"#e2e8e1",
    "maintenance_bg":"#f7eddb","maintenance_text":"#725421","maintenance_border":"#e4d3ad","bin_card":"#fffefa","bin_id":"#758079","bin_name":"#24483e",
    "bin_detail":"#68766e","bin_value":"#46574e","footer":"#829087","active_nav":"linear-gradient(90deg,#e2eee4,#eef5ee)","alert":"#f4f7f3","scheme":"light"
})
DATA=Path(__file__).parent/"data";DATA.mkdir(exist_ok=True);BOOK=DATA/"circular.xlsx"
TYPES=["Wet","Dry","E-waste"]; PAGES=["Overview","Live sorting","Bin network","Monitor & trends","AI advisor","Pickup calendar","Disposal log"]
NAV_ITEMS={"⌂   Overview":"Overview","◉   Live sorting":"Live sorting","▤   Bin network":"Bin network","◷   Monitor & trends":"Monitor & trends","✦   AI advisor":"AI advisor","▦   Pickup calendar":"Pickup calendar","≡   Disposal log":"Disposal log"}
NAV_ICONS={"Overview":"⌂","Live sorting":"◎","Bin network":"▤","Monitor & trends":"◷","AI advisor":"✦","Pickup calendar":"▦","Disposal log":"☷"}
RULES={"Wet":"Drain liquids and remove plastic packaging before food scraps go into the wet bin.","Dry":"Keep recyclable items clean and dry. Empty containers and flatten cardboard.","E-waste":"Keep batteries and electronics out of regular bins. Use the designated e-waste point."}
DEFAULT_GEMINI_MODEL="gemini-3.5-flash-lite"
THEME_CSS='''<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@500;600;700;800&display=swap');
:root{--app-bg:__PAGE__;--app-text:__TEXT__;--muted:__MUTED__;--sidebar-bg:__SIDEBAR__;--sidebar-text:__SIDEBAR_TEXT__;--sidebar-muted:__SIDEBAR_MUTED__;--panel:__PANEL__;--panel2:__PANEL2__;--border:__BORDER__;--accent:__ACCENT__;--hover:__HOVER__;--input:__INPUT__;--input-border:__INPUT_BORDER__;--hero-bg:__HERO__;--hero-border:__HERO_BORDER__;--tip-bg:__TIP__;--eyebrow:__EYEBROW__;--active-nav:__ACTIVE_NAV__;--alert-bg:__ALERT__;--scheme:__SCHEME__}
.stApp{background:var(--app-bg);font-family:'DM Sans',sans-serif;color:var(--app-text);color-scheme:var(--scheme)}[data-testid="stHeader"]{background:transparent}[data-testid="stAppViewContainer"]{background:transparent}[data-testid="stMainBlockContainer"],.main .block-container{padding-top:1rem!important;margin-top:0!important}
[data-testid="stSidebar"]{background:var(--sidebar-bg);border-right:1px solid var(--border)}[data-testid="stSidebar"] *{color:var(--sidebar-text)}[data-testid="stSidebar"] [data-testid="stSidebarHeader"]{height:2.5rem!important;min-height:2.5rem!important;max-height:2.5rem!important;flex:0 0 2.5rem!important;padding:0!important;margin:0!important}[data-testid="stSidebar"] [data-testid="stSidebarContent"],[data-testid="stSidebar"] [data-testid="stSidebarUserContent"]{padding:0!important;margin:0!important}[data-testid="stSidebar"] [data-testid="stSidebarContent"]>div{padding-top:0!important;margin-top:0!important}
[data-testid="stSidebarCollapsedControl"] svg,[data-testid="stSidebarCollapsedControl"] button svg,[data-testid="stSidebarCollapseButton"] svg,[data-testid="stSidebarCollapseButton"] button svg,button[aria-label="Open sidebar"] svg,button[title="Open sidebar"] svg{visibility:hidden!important}[data-testid="stSidebarCollapsedControl"],button[aria-label="Open sidebar"],button[title="Open sidebar"]{position:relative!important}[data-testid="stSidebarCollapsedControl"]::after,[data-testid="stSidebarCollapsedControl"] button::after,[data-testid="stSidebarCollapseButton"]::after,[data-testid="stSidebarCollapseButton"] button::after,button[aria-label="Open sidebar"]::after,button[title="Open sidebar"]::after{content:'';position:absolute;width:21px;height:17px;left:50%;top:50%;transform:translate(-50%,-50%);background:linear-gradient(to bottom,var(--app-text) 0 3px,transparent 3px 7px,var(--app-text) 7px 10px,transparent 10px 14px,var(--app-text) 14px 17px);pointer-events:none}
[data-testid="stSidebarCollapseButton"]::after,[data-testid="stSidebarCollapseButton"] button::after{content:none!important;display:none!important}
h1,h2,h3{font-family:'Manrope',sans-serif!important;color:var(--app-text);letter-spacing:-.03em}h1{font-size:1.8rem!important;margin:.2rem 0 .15rem!important}h2{font-size:1.3rem!important;margin:.35rem 0 .2rem!important}h3{font-size:1rem!important;margin:.4rem 0 .2rem!important}
p,label,li,[data-testid="stCaptionContainer"]{font-size:.9rem;color:var(--muted)}
div[data-testid="stMetric"]{background:var(--panel2);padding:13px 16px;border:1px solid var(--border);border-radius:15px;box-shadow:0 7px 20px #0002}div[data-testid="stMetricLabel"]{font-size:.75rem;color:var(--muted)}div[data-testid="stMetricValue"]{font-family:'Manrope',sans-serif;color:var(--app-text);font-size:1.55rem}
div[data-testid="stVerticalBlockBorderWrapper"]{background:var(--panel);border:1px solid var(--border);border-radius:17px;box-shadow:0 8px 24px #0002}
div.stButton>button,.stDownloadButton button{background:var(--panel2);border:1px solid var(--border);border-radius:11px;color:var(--app-text);font-weight:650;min-height:40px}div.stButton>button:hover,.stDownloadButton button:hover{background:var(--hover);border-color:var(--accent);color:var(--app-text)}div.stButton>button[kind="primary"]{background:var(--accent);border:0;color:white}
[data-testid="stTextInput"] input,[data-testid="stNumberInput"] input,[data-testid="stDateInput"] input,[data-testid="stTimeInput"] input,[data-testid="stTextArea"] textarea,[data-testid="stSelectbox"] [data-baseweb="select"]>div,[data-testid="stMultiSelect"] [data-baseweb="select"]>div{background:var(--input)!important;color:var(--app-text)!important;border-color:var(--input-border)!important;border-radius:10px}
[data-baseweb="popover"] [role="listbox"],[data-baseweb="menu"],[data-baseweb="select"]{background:var(--panel2)!important;color:var(--app-text)!important}[role="option"]{color:var(--app-text)!important}[role="option"]:hover{background:var(--hover)!important}
[data-testid="stDataFrame"],[data-testid="stTable"],[data-testid="stExpander"]{background:var(--panel);border:1px solid var(--border);border-radius:12px;overflow:hidden}
.hero{border:1px solid var(--hero-border);border-radius:20px;padding:24px 28px;background:var(--hero-bg);margin:6px 0 20px}.hero h2{font-size:25px!important;margin:0 0 5px;color:var(--app-text)}.hero p{color:var(--muted);margin:0}.eyebrow{color:var(--eyebrow);font-size:10px;letter-spacing:.14em;font-weight:700;text-transform:uppercase}.compact-page-header{margin:0 0 10px}.compact-page-header .eyebrow{margin:0 0 2px}.compact-page-header h1{font:800 1.8rem Manrope,sans-serif!important;letter-spacing:-.03em;color:var(--app-text);margin:0!important}.compact-page-header p{font-size:.9rem;color:var(--muted);margin:3px 0 0!important}.tip{background:var(--tip-bg);border:1px solid var(--border);border-left:4px solid var(--accent);padding:14px 17px;border-radius:13px;color:var(--app-text)}
.sidebar-brand{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:5px;width:100%;margin:0 0 18px;padding:0;text-align:center}.sidebar-brand-mark{width:32px;height:32px;display:flex;align-items:center;justify-content:center;color:#28745f;font-size:30px;line-height:1}.sidebar-brand-title{font:800 11px Manrope,sans-serif;letter-spacing:.16em;text-transform:uppercase;color:var(--sidebar-muted)}.main-brand{display:flex;align-items:center;justify-content:flex-start;gap:10px;margin:0 0 14px;padding:0}.main-brand-mark{width:38px;height:38px;display:flex;align-items:center;justify-content:center;color:#28745f;font-size:36px;line-height:1}.main-brand-title{font:800 23px Manrope,sans-serif;letter-spacing:.04em;color:var(--app-text)}
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p{color:var(--sidebar-muted);font-size:10px;letter-spacing:.15em;font-weight:700}[data-testid="stSidebar"] [role="radiogroup"]{gap:7px}[data-testid="stSidebar"] [role="radiogroup"] label{background:transparent;border-radius:11px;padding:7px 10px;transition:background .15s ease}[data-testid="stSidebar"] [role="radiogroup"] label:hover{background:#ffffff12}[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked){background:var(--active-nav);box-shadow:inset 3px 0 var(--accent)}[data-testid="stSidebar"] [role="radiogroup"] label p{color:var(--sidebar-text);font-size:14px;font-weight:600;letter-spacing:0}

.stProgress>div>div>div>div{background:linear-gradient(90deg,#db514d,#ef8176)}[data-testid="stAlert"]{background:var(--alert-bg);border:1px solid var(--border);color:var(--app-text)}hr{border-color:var(--border)}footer{visibility:hidden}</style>'''
for token,key in {'__PAGE__':'page','__TEXT__':'text','__MUTED__':'muted','__SIDEBAR__':'sidebar','__SIDEBAR_TEXT__':'sidebar_text','__SIDEBAR_MUTED__':'sidebar_muted','__PANEL__':'panel','__PANEL2__':'panel2','__BORDER__':'border','__ACCENT__':'accent','__HOVER__':'hover','__INPUT__':'input','__INPUT_BORDER__':'input_border','__HERO__':'hero','__HERO_BORDER__':'hero_border','__TIP__':'tip','__EYEBROW__':'eyebrow','__ACTIVE_NAV__':'active_nav','__ALERT__':'alert','__SCHEME__':'scheme'}.items():
    THEME_CSS = THEME_CSS.replace(token, PALETTE[key])
st.markdown(THEME_CSS, unsafe_allow_html=True)

def load(sheet,cols):
    try:
        x=pd.read_excel(BOOK,sheet_name=sheet)
        for c in cols:
            if c not in x:x[c]=None
        return x[cols]
    except Exception:return pd.DataFrame(columns=cols)
if "bins" not in st.session_state:
    st.session_state.bins=load("Bins",["Bin ID","Name","Zone","Category","Capacity (L)","Fill (L)","Status"])
    if st.session_state.bins.empty:st.session_state.bins=pd.DataFrame([["BIN-001","Library North","Library","Dry",120,48,"Online"],["BIN-002","Cafeteria Wet","Cafeteria","Wet",160,103,"Online"],["BIN-003","Science E-waste","Science","E-waste",80,31,"Online"],["BIN-004","East Gate","East Gate","Dry",100,74,"Online"]],columns=["Bin ID","Name","Zone","Category","Capacity (L)","Fill (L)","Status"])
if "logs" not in st.session_state:st.session_state.logs=load("Disposals",["Timestamp","Object","Category","Volume (L)","Bin ID","Bin name","Guidance"])
if "events" not in st.session_state:st.session_state.events=load("Calendar",["Date","Time","Type","Bin ID","Notes","Status"])
# Normalize older saved event labels so existing calendar entries keep working.
if not st.session_state.events.empty:
    st.session_state.events["Type"]=st.session_state.events["Type"].replace({"Bin pickup":"Bin Cleanup","Campus cleanup":"Bin Cleanup"})
if "first_page_initialized" not in st.session_state:
    st.session_state.page="Live sorting"
    st.session_state.first_page_initialized=True
st.session_state.setdefault("page","Live sorting");st.session_state.setdefault("chat",[{"role":"assistant","content":"Ask where something goes or how to reduce waste."}]);st.session_state.setdefault("classification",{})
def save():
    try:
        with pd.ExcelWriter(BOOK,engine="openpyxl") as w:
            st.session_state.logs.to_excel(w,sheet_name="Disposals",index=False);st.session_state.bins.to_excel(w,sheet_name="Bins",index=False);st.session_state.events.to_excel(w,sheet_name="Calendar",index=False)
        st.toast("Saved to Excel")
    except Exception as ex:st.error(f"Excel save failed (close the workbook if open): {ex}")
def next_bin_id(bins):
    """Return the next sequential BIN number, preserving existing IDs."""
    numbers=[]
    if not bins.empty and "Bin ID" in bins:
        for value in bins["Bin ID"].dropna().astype(str):
            match=re.fullmatch(r"BIN-(\d+)",value.strip(),flags=re.IGNORECASE)
            if match:numbers.append(int(match.group(1)))
    return f"BIN-{max(numbers,default=0)+1:03d}"
def complete_calendar_activity(index):
    """Complete an activity and update the twin and disposal history."""
    event=st.session_state.events.loc[index].copy()
    st.session_state.events.loc[index,"Status"]="Complete"
    event_type=str(event.get("Type",""))
    event_bin=str(event.get("Bin ID",""))
    cleanup_types={"Bin Cleanup","Bin pickup"}
    reset_types=cleanup_types|{"Bin Maintenance","Bin maintenance"}
    if event_type not in reset_types:return None
    if event_bin=="ALL":
        st.session_state.bins.loc[:,"Fill (L)"]=0.0
        if event_type in cleanup_types:
            st.session_state.logs=st.session_state.logs.iloc[0:0].copy()
            return "all bins and disposal records"
        return "all bins"
    if event_bin in {"","Campus-wide","nan","None"}:return None
    matched=st.session_state.bins["Bin ID"].astype(str)==event_bin
    if not matched.any():return None
    st.session_state.bins.loc[matched,"Fill (L)"]=0.0
    bin_name=str(st.session_state.bins.loc[matched,"Name"].iloc[0])
    if event_type in cleanup_types:
        st.session_state.logs=st.session_state.logs[st.session_state.logs["Bin ID"].astype(str)!=event_bin].reset_index(drop=True)
        return f"{bin_name} and its disposal records"
    return bin_name
def schedule_pickup_if_needed(bin_row):
    capacity=float(bin_row["Capacity (L)"] or 0)
    fill=float(bin_row["Fill (L)"] or 0)
    if capacity<=0 or fill/capacity<0.90:return False
    events=st.session_state.events
    existing=events[(events["Status"].astype(str)=="Scheduled") & (events["Type"].astype(str).isin(["Bin Cleanup","Bin pickup"])) & (events["Bin ID"].astype(str)==str(bin_row["Bin ID"]))] if not events.empty else events
    if not existing.empty:return False
    pickup_at=datetime.now()+timedelta(hours=1)
    st.session_state.events.loc[len(events)]={"Date":pickup_at.date().isoformat(),"Time":pickup_at.strftime("%H:%M"),"Type":"Bin Cleanup","Bin ID":str(bin_row["Bin ID"]),"Notes":"Automatic cleanup: bin reached 90% capacity.","Status":"Scheduled"}
    return True
def render_bin_icon(percent):
    p=max(0,min(100,int(percent)));height=30*p/100;y=43-height
    return f"<svg width='42' height='50' viewBox='0 0 42 50' aria-label='Bin {p}% full'><defs><clipPath id='binclip'><path d='M9 13h24l-2 32H11z'/></clipPath></defs><path d='M6 10h30M15 6h12M9 13h24l-2 32H11z' fill='{PALETTE['empty_bin']}' stroke='{PALETTE['bin_outline']}' stroke-width='2.2' stroke-linecap='round' stroke-linejoin='round'/><rect x='9' y='{y:.1f}' width='24' height='{height:.1f}' fill='#db514d' clip-path='url(#binclip)'/><path d='M9 13h24l-2 32H11z' fill='none' stroke='{PALETTE['bin_outline']}' stroke-width='2.2' stroke-linejoin='round'/></svg>"
def header(t,s,k="CIRCULAR CAMPUS",tight=False):
    if tight:
        st.markdown(f"<div class='compact-page-header'><div class='eyebrow'>{k}</div><h1>{t}</h1><p>{s}</p></div>",unsafe_allow_html=True)
    else:
        st.markdown(f"<div class='eyebrow'>{k}</div>",unsafe_allow_html=True);st.title(t);st.caption(s)
def safe_ai_error_detail(exc):
    """Return useful SDK detail while removing any API key before display."""
    detail=str(getattr(exc,"message",None) or exc).strip()
    key=gemini_api_key()
    if key:detail=detail.replace(key,"[API KEY REDACTED]")
    detail=re.sub(r"AIza[0-9A-Za-z_-]{20,}","[API KEY REDACTED]",detail)
    detail=re.sub(r"(?i)(x-goog-api-key|api[_-]?key)\s*[:=]\s*[^\s,;]+",r"\1=[API KEY REDACTED]",detail)
    return detail[:500] or "No error details were provided."

def safe_ai_error(exc, task):
    code=getattr(exc,"status_code",None) or getattr(exc,"code",None)
    error_text=str(exc).lower()
    if isinstance(exc,RuntimeError) and "add your gemini api key" in error_text:
        st.error(f"{task} has no Gemini key saved in this browser session. Enter the key in ‘Connect / replace AI key’ and click ‘Use this key for this session’.")
    elif isinstance(exc,ModuleNotFoundError) and ("google.genai" in str(getattr(exc,"name", "")) or "google" in str(getattr(exc,"name", ""))):
        st.error(f"{task} cannot start because the Gemini SDK is not installed. In the app's project folder, run `pip install -r requirements.txt`, then restart Streamlit.")
    elif code in (401,403) or "api_key_invalid" in error_text or "api key not valid" in error_text:
        st.session_state.ai_key_invalid=True
        st.error(f"{task} paused: Gemini rejected this API key. Replace it in ‘Connect / replace AI key’. The provider's detailed error is hidden so credentials are not echoed.")
    elif code==400 and ("api_key_invalid" in error_text or "api key not valid" in error_text):
        st.session_state.ai_key_invalid=True
        st.error(f"{task} paused: Google says this Gemini API key is invalid. Create a valid Gemini key in Google AI Studio and replace it in ‘Connect / replace AI key’.")
    elif code==429 or "resource_exhausted" in error_text:st.error(f"{task} paused: the Gemini rate limit or account quota was reached. Check Google AI Studio usage and limits.")
    elif code==404:st.error(f"{task} failed because the configured Gemini model was not found. Check GEMINI_MODEL, GEMINI_VISION_MODEL, or GEMINI_TRANSCRIBE_MODEL.")
    elif code==400:st.error(f"{task} was rejected by Gemini (request error 400). Check the selected model and request settings.")
    elif isinstance(exc,(ConnectionError,TimeoutError)) or "connecterror" in type(exc).__name__.lower() or "timeout" in type(exc).__name__.lower():st.error(f"{task} could not connect to Gemini. Check your internet connection, proxy, or firewall, then try again.")
    elif code:st.error(f"{task} failed while contacting Gemini (status {code}). Check the Gemini key, model, and Google AI Studio usage limits.")
    elif isinstance(exc,(RuntimeError,TypeError,ValueError)):st.error(f"{task} failed ({type(exc).__name__}): {safe_ai_error_detail(exc)}")
    else:st.error(f"{task} could not complete due to a local setup or response error ({type(exc).__name__}). Confirm `google-genai` is installed with `pip install -r requirements.txt`, then restart Streamlit. If it still fails, share the Streamlit terminal error without sharing your API key.")

def gemini_api_key():
    return st.session_state.get("gemini_api_key") or os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")

def gemini_client():
    from google import genai
    api_key=gemini_api_key()
    if not api_key:raise RuntimeError("Add your Gemini API key in the Live sorting page under ‘Connect / replace AI key’.")
    # Keep a strong per-session reference. google-genai closes its HTTP client
    # when the Client is garbage-collected, so a temporary `gemini_client().models`
    # can otherwise fail with “Cannot send a request, as the client has been closed.”
    fingerprint=hashlib.sha256(api_key.encode("utf-8")).hexdigest()
    client=st.session_state.get("_gemini_client")
    if client is None or st.session_state.get("_gemini_client_fingerprint")!=fingerprint:
        if client is not None:
            try:client.close()
            except Exception:pass
        client=genai.Client(api_key=api_key)
        st.session_state["_gemini_client"]=client
        st.session_state["_gemini_client_fingerprint"]=fingerprint
    return client

def gemini_generate_content(model,**kwargs):
    """Call the requested model and retry the stable multimodal default on 404."""
    client=gemini_client()
    try:return client.models.generate_content(model=model,**kwargs)
    except Exception as exc:
        if getattr(exc,"code",None)==404 and model!=DEFAULT_GEMINI_MODEL:
            return client.models.generate_content(model=DEFAULT_GEMINI_MODEL,**kwargs)
        raise

def bin_context():
    if st.session_state.bins.empty:return "No bins are registered yet."
    return "\n".join(
        f"{row['Bin ID']}: {row['Name']} ({row['Category']}, {float(row['Fill (L)']):.1f}/{float(row['Capacity (L)']):.0f} L)"
        for _,row in st.session_state.bins.iterrows()
    )

def ask_circular(question,history=None):
    system_instruction=(
        "You are Circular, a practical campus waste-sorting advisor. Answer the user's actual question directly and specifically; "
        "avoid repeating generic tips. For sorting, use only Wet, Dry, or E-waste categories. Be concise, ask a brief clarifying question "
        "when the item is ambiguous, and do not invent local recycling rules. Current bins:\n"+bin_context()
    )
    contents=[]
    if history:
        recent=history[-10:]
        # Gemini 3 requires chat turns to begin with a user turn; the advisor's
        # initial greeting (or a truncated model turn) is display-only context.
        while recent and recent[0].get("role")=="assistant":recent=recent[1:]
        for message in recent:
            role="model" if message.get("role")=="assistant" else "user"
            contents.append({"role":role,"parts":[{"text":str(message.get("content",""))}]})
    contents.append({"role":"user","parts":[{"text":question}]})
    response=gemini_generate_content(
        model=os.getenv("GEMINI_MODEL",DEFAULT_GEMINI_MODEL),contents=contents,
        config={"system_instruction":system_instruction},
    )
    answer=(response.text or "").strip()
    if not answer:raise ValueError("The advisor returned an empty response.")
    return answer

def check_disposal_item(item,selected_category):
    from google.genai import types
    response=gemini_generate_content(
        model=os.getenv("GEMINI_MODEL",DEFAULT_GEMINI_MODEL),
        contents=f"Assess this campus waste item for safe sorting. Return only JSON with actual_category (Wet, Dry, E-waste, or Unclear), is_hazardous (boolean), and reason (short string). Treat batteries, electronics, chemicals, sharps, and unknown hazardous materials carefully.\nItem: {item}\nCategory selected by user: {selected_category}",
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            system_instruction=(
            "Assess a campus waste item for safe sorting. Return only JSON with actual_category (Wet, Dry, E-waste, or Unclear), "
            "is_hazardous (boolean), and reason (short string). Treat batteries, electronics, chemicals, sharps, and unknown hazardous materials carefully."
            ),
        ),
    )
    raw=(response.text or "").strip()
    try:return json.loads(raw)
    except json.JSONDecodeError:
        start,end=raw.find("{"),raw.rfind("}")
        if start>=0 and end>start:return json.loads(raw[start:end+1])
        raise ValueError("The safety check returned an unreadable response.")

def auto_log_camera_item(result):
    """Immediately log a confident, safe camera classification into its matching bin."""
    item=str(result.get("object","")).strip()
    category=result.get("category")
    if not item or item.lower()=="none" or category not in TYPES:return "No clear item was detected; nothing was added."
    if float(result.get("confidence",0) or 0)<0.55:return "The camera result was uncertain, so nothing was added."
    if bool(result.get("is_hazardous",False)):
        return "hazard:"+str(result.get("hazard_reason") or "Potentially hazardous material detected. Keep it out of regular bins.")
    # Workbook values may load as int64 when existing fills are whole litres.
    # Convert to float before adding Gemini's fractional volume estimate.
    bins=st.session_state.bins.copy()
    bins["Fill (L)"]=pd.to_numeric(bins["Fill (L)"],errors="coerce").fillna(0).astype(float)
    bins["Capacity (L)"]=pd.to_numeric(bins["Capacity (L)"],errors="coerce").fillna(0).astype(float)
    matches=bins[bins["Category"].astype(str)==category]
    if matches.empty:return f"No {category} bin is registered; the camera result was not logged."
    target=matches.sort_values(by="Fill (L)",key=lambda values:pd.to_numeric(values,errors="coerce").fillna(0)).iloc[0]
    idx=target.name
    volume=max(0.01,min(100.0,float(result.get("estimated_volume_l",1) or 1)))
    bins.loc[idx,"Fill (L)"]=min(float(target["Capacity (L)"]),float(target["Fill (L)"])+volume)
    st.session_state.bins=bins
    st.session_state.logs.loc[len(st.session_state.logs)]=[
        datetime.now().isoformat(timespec="seconds"),item,category,round(volume,2),str(target["Bin ID"]),str(target["Name"]),str(result.get("guidance") or RULES[category])
    ]
    auto_added=schedule_pickup_if_needed(bins.loc[idx])
    save()
    return f"Added {item} to {target['Bin ID']} · {target['Name']} and updated the digital twin."+(" A pickup was also scheduled at 90% capacity." if auto_added else "")

def classify_waste_frame(rgb_frame):
    from google.genai import types
    from PIL import Image
    import io
    buffer=io.BytesIO();Image.fromarray(rgb_frame).save(buffer,format="JPEG",quality=80)
    image_part=types.Part.from_bytes(data=buffer.getvalue(),mime_type="image/jpeg")
    response=gemini_generate_content(
        model=os.getenv("GEMINI_VISION_MODEL",os.getenv("GEMINI_MODEL",DEFAULT_GEMINI_MODEL)),
        contents=["Identify the most prominent waste item. Return only valid JSON with fields object (short name), category (exactly Wet, Dry, or E-waste), estimated_volume_l (number), confidence (0 to 1), guidance (one concise instruction), is_hazardous (boolean), and hazard_reason (short string). If there is no clear waste item, set object to none and confidence to 0. Mark batteries, electronics, chemicals, sharps, and other unsafe materials hazardous; ordinary clean recyclables are not hazardous.",image_part],
        config=types.GenerateContentConfig(response_mime_type="application/json"),
    )
    raw=(response.text or "").strip()
    try:return json.loads(raw)
    except json.JSONDecodeError:
        start,end=raw.find("{"),raw.rfind("}")
        if start>=0 and end>start:return json.loads(raw[start:end+1])
        raise ValueError("The model response was not valid JSON")

class LatestFrame(VideoProcessorBase):
    def __init__(self): self.latest = None
    def recv(self, frame):
        self.latest = frame.to_ndarray(format="rgb24")
        return frame
incoming_page=st.query_params.get("nav")
if incoming_page in PAGES:
    st.session_state.page=incoming_page
    st.session_state.sidebar_navigation=next((label for label,target in NAV_ITEMS.items() if target==incoming_page),list(NAV_ITEMS)[0])
    del st.query_params["nav"]
with st.sidebar:
    st.markdown("<div class='sidebar-brand'><span class='sidebar-brand-mark'>♻</span><div class='sidebar-brand-title'>SMART BIN NETWORK</div></div>",unsafe_allow_html=True)
    nav_labels=list(NAV_ITEMS)
    current_nav=next((label for label,target in NAV_ITEMS.items() if target==st.session_state.page),nav_labels[0])
    selected_nav=st.radio("Navigation",nav_labels,index=nav_labels.index(current_nav),key="sidebar_navigation",label_visibility="collapsed")
    st.session_state.page=NAV_ITEMS[selected_nav]
    st.divider();st.text_input("Campus","Dayananda Sagar University",label_visibility="collapsed");st.success(f"● {(st.session_state.bins.Status=='Online').sum()} active bins")

st.markdown("<div></div>",unsafe_allow_html=True)
st.markdown("<div class='main-brand'><span class='main-brand-mark'>♻</span><span class='main-brand-title'>CIRCULAR</span></div>",unsafe_allow_html=True)
page=st.session_state.page;b=st.session_state.bins;logs=st.session_state.logs;events=st.session_state.events
if page=="Overview":
    st.markdown("<div class='hero'><div class='eyebrow'>SMARTER SORTING · CLEANER CAMPUS</div><h2>A live view of your circular campus.</h2><p>Track disposal, fullness, pickups, and waste-reduction habits.</p></div>",unsafe_allow_html=True)
    a,c,d,e=st.columns(4);a.metric("Network bins",len(b),f"{(b.Status=='Online').sum()} online");c.metric("Estimated fill",f"{b['Fill (L)'].sum():.0f} L",f"{b['Fill (L)'].sum()/max(1,b['Capacity (L)'].sum()):.0%} capacity");d.metric("Logged disposals",len(logs));e.metric("Scheduled activities",int((events.Status=="Scheduled").sum()) if not events.empty else 0)
    l,r=st.columns([1.2,1],gap="large")
    with l:
        with st.container(border=True):
            st.subheader("Digital twin network")
            for _,x in b.iterrows():
                p=min(100,int(100*x["Fill (L)"]/max(1,x["Capacity (L)"])));st.write(f"**{x['Name']}** · {x['Category']} · {x['Fill (L)']:.1f}/{x['Capacity (L)']} L · {p}%");st.progress(p)
    with r:
        with st.container(border=True):
            st.subheader("Recent disposals");st.dataframe(logs.tail(7).iloc[::-1][["Timestamp","Object","Category","Volume (L)","Bin name"]] if not logs.empty else logs,hide_index=True,use_container_width=True)
    with st.container(border=True):
        st.subheader("Fill by bin")
        if not b.empty:st.altair_chart(alt.Chart(b).mark_bar().encode(y=alt.Y("Name:N",sort="-x",title=None),x=alt.X("Fill (L):Q",title="Estimated litres"),color="Category:N",tooltip=["Bin ID","Name","Category","Fill (L)","Capacity (L)"]).configure(background="transparent",axis=alt.AxisConfig(labelColor=PALETTE["text"],titleColor=PALETTE["muted"],gridColor=PALETTE["border"],domainColor=PALETTE["border"]),legend=alt.LegendConfig(labelColor=PALETTE["text"],titleColor=PALETTE["muted"])),use_container_width=True)
elif page=="Live sorting":
    header("Live waste sorting","Keep the camera running, identify waste, and get immediate disposal guidance.","CAMERA · CLASSIFY · GUIDE",tight=True)
    l,r=st.columns([1.2,1],gap="medium")
    with l:
        with st.container(border=True):
            st.subheader("Live camera")
            with st.expander("Connect / replace AI key"):
                st.caption("Needed for vision and voice. Your key stays in this browser session and is not written to the workbook.")
                entered=st.text_input("Gemini API key",type="password",key="gemini_key_entry")
                # Persist the current widget value on each Streamlit rerun so
                # every Gemini call sees the key even if the submit click is missed.
                if entered.strip():st.session_state.gemini_api_key=entered.strip()
                if st.button("Use this key for this session"):
                    if entered.strip():st.session_state.gemini_api_key=entered.strip();st.session_state.ai_key_invalid=False;st.rerun()
                    else:st.warning("Enter your API key first.")
            ctx=webrtc_streamer(key="circular-waste-camera",video_processor_factory=LatestFrame,media_stream_constraints={"video":True,"audio":False},async_processing=True)
            st.session_state.live_camera_context=ctx
            has_key=bool(gemini_api_key())
            auto=st.toggle("Auto-identify every 5 seconds (uses AI API)",value=False,disabled=(not has_key or st.session_state.get("ai_key_invalid",False)),key="auto_identify_paid",help="This makes repeated paid API calls. Leave it off and use “Classify + add” for one image at a time.")
            if not auto:
                st.session_state.pop("auto_logged_camera_signature",None)
            if auto:
                st.warning("Automatic checks send camera frames to your configured AI provider every few seconds. Each item is added once while it remains in view; remove it and let the camera see an empty view before scanning another item.")
            if has_key and st.button("Scan item and add automatically",type="primary"):
                try:
                    if not ctx.video_processor or ctx.video_processor.latest is None: st.warning("Start the camera and place an item in view first.")
                    else:
                        st.session_state.classification=classify_waste_frame(ctx.video_processor.latest)
                        outcome=auto_log_camera_item(st.session_state.classification)
                        scanned_item=str(st.session_state.classification.get("object") or "").strip()
                        scanned_category=st.session_state.classification.get("category")
                        if scanned_item and scanned_item.lower()!="none" and scanned_category in TYPES:
                            st.session_state.auto_logged_camera_signature=(scanned_item.casefold(),scanned_category)
                        st.session_state.camera_log_notice=("error" if outcome.startswith("hazard:") else "success",outcome.removeprefix("hazard:") if outcome.startswith("hazard:") else outcome)
                        st.rerun()
                except Exception as ex:safe_ai_error(ex,"Frame classification")
    with r:
        if auto:
            @st.fragment(run_every="5s")
            def scanned_waste_panel():
                auto_message=None
                auto_notice_type="success"
                live_ctx=st.session_state.get("live_camera_context")
                if st.session_state.get("ai_key_invalid"):
                    auto_message="Auto-identification is paused because the API key was rejected. Replace the key above."
                    auto_notice_type="error"
                elif live_ctx and live_ctx.video_processor and live_ctx.video_processor.latest is not None:
                    try:
                        detected=classify_waste_frame(live_ctx.video_processor.latest)
                        item=str(detected.get("object") or "").strip()
                        category=detected.get("category")
                        confident=bool(item and item.lower()!="none" and category in TYPES and float(detected.get("confidence",0) or 0)>=0.55)
                        if not confident:
                            st.session_state.pop("auto_logged_camera_signature",None)
                            auto_message="No clear item detected. Show one item to the camera to add it automatically."
                            auto_notice_type="info"
                        else:
                            st.session_state.classification=detected
                            signature=(item.casefold(),category)
                            if st.session_state.get("auto_logged_camera_signature")==signature:
                                auto_message=f"{item} is already logged. Remove it from view before scanning another item."
                                auto_notice_type="info"
                            else:
                                outcome=auto_log_camera_item(detected)
                                st.session_state.auto_logged_camera_signature=signature
                                if outcome.startswith("hazard:"):
                                    auto_message="Potential hazardous waste found: "+outcome.removeprefix("hazard:")
                                    auto_notice_type="error"
                                else:
                                    auto_message=outcome
                                    auto_notice_type="success" if outcome.startswith("Added ") else "warning"
                    except Exception as exc:
                        auto_message=f"Automatic classification failed ({type(exc).__name__}): {safe_ai_error_detail(exc)}"
                        auto_notice_type="error"
                with st.container(border=True):
                    st.subheader("Scanned waste")
                    z=st.session_state.classification
                    if z:
                        st.write(str(z.get("object","Item")))
                        category_col,volume_col=st.columns(2)
                        category_col.metric("Waste category",str(z.get("category","Unclear")))
                        try:estimated_volume=float(z.get("estimated_volume_l",0) or 0)
                        except (TypeError,ValueError):estimated_volume=0.0
                        volume_col.metric("Estimated volume",f"{estimated_volume:.2f} L")
                        guidance=str(z.get("guidance","")).strip()
                        if guidance:st.caption(guidance)
                        latest_logs=st.session_state.logs
                        if not latest_logs.empty:
                            latest=latest_logs.iloc[-1]
                            if str(latest.get("Object",""))==str(z.get("object","")) and str(latest.get("Category",""))==str(z.get("category","")):
                                st.caption(f"Added to {latest.get('Bin ID','')} · {latest.get('Bin name','')}")
                    else:
                        st.caption("The latest camera result will appear here.")
                    if st.session_state.get("camera_log_notice"):
                        notice_type,notice_text=st.session_state.pop("camera_log_notice")
                        (st.error if notice_type=="error" else st.success)(("Potential hazardous waste found: " if notice_type=="error" else "")+notice_text)
                    if auto_message:
                        if auto_notice_type=="error":st.error(auto_message)
                        elif auto_notice_type=="warning":st.warning(auto_message)
                        elif auto_notice_type=="info":st.caption(auto_message)
                        else:st.success(auto_message)
            scanned_waste_panel()
        else:
            with st.container(border=True):
                st.subheader("Scanned waste")
                z=st.session_state.classification
                if z:
                    st.write(str(z.get("object","Item")))
                    category_col,volume_col=st.columns(2)
                    category_col.metric("Waste category",str(z.get("category","Unclear")))
                    try:estimated_volume=float(z.get("estimated_volume_l",0) or 0)
                    except (TypeError,ValueError):estimated_volume=0.0
                    volume_col.metric("Estimated volume",f"{estimated_volume:.2f} L")
                    guidance=str(z.get("guidance","")).strip()
                    if guidance:st.caption(guidance)
                    latest_logs=st.session_state.logs
                    if not latest_logs.empty:
                        latest=latest_logs.iloc[-1]
                        if str(latest.get("Object",""))==str(z.get("object","")) and str(latest.get("Category",""))==str(z.get("category","")):
                            st.caption(f"Added to {latest.get('Bin ID','')} · {latest.get('Bin name','')}")
                else:
                    st.caption("The latest camera result will appear here.")
                if st.session_state.get("camera_log_notice"):
                    notice_type,notice_text=st.session_state.pop("camera_log_notice")
                    (st.error if notice_type=="error" else st.success)(("Potential hazardous waste found: " if notice_type=="error" else "")+notice_text)
        with st.container(border=True):
            st.subheader("Ask verbally where it goes");audio=st.audio_input("Record a question",sample_rate=16000)
            if st.button("Transcribe with AI"):
                if not audio:st.warning("Record a question first.")
                elif not gemini_api_key():st.warning("Connect a Gemini API key in the camera panel to enable speech transcription.")
                else:
                    try:
                        from google.genai import types
                        voice_part=types.Part.from_bytes(data=audio.getvalue(),mime_type="audio/wav")
                        transcription=gemini_generate_content(
                            model=os.getenv("GEMINI_TRANSCRIBE_MODEL",os.getenv("GEMINI_MODEL",DEFAULT_GEMINI_MODEL)),
                            contents=["Transcribe the spoken words exactly. Return only the transcription, with no commentary.",voice_part],
                        )
                        st.session_state.voice=(transcription.text or "").strip()
                        if not st.session_state.voice:raise ValueError("Gemini returned an empty transcription.")
                    except Exception as ex:safe_ai_error(ex,"Voice transcription")
            question=st.text_input("Or Type Your Question",value=st.session_state.get("voice",""),placeholder="Where do I throw this cup?")
            if st.button("Ask Circular",type="primary") and question:
                if not gemini_api_key():
                    st.warning("Add your Gemini API key above to ask Circular.")
                else:
                    try:st.info(ask_circular(question))
                    except Exception as ex:safe_ai_error(ex,"Sorting question")
    with st.container(border=True):
        st.subheader("Manual Disposal Entry");a,c,d,e=st.columns([1.2,1,1,1.5]);obj=a.text_input("Object",value="",placeholder="Item name");cat=c.selectbox("Waste type",TYPES,index=1);volume=d.number_input("Estimated volume (L)",.01,100.,1.0,.1);matches=b[b.Category==cat];options=matches if not matches.empty else b;bid=e.selectbox("Destination bin",options["Bin ID"].astype(str),format_func=lambda k:f"{k} · {options.loc[options['Bin ID'].astype(str)==k,'Name'].iloc[0]}") if not options.empty else None
        if st.button("Save disposal + update digital twin",type="primary",use_container_width=True):
            if bid:
                if not gemini_api_key():
                    st.warning("Connect a Gemini API key in the Live sorting camera panel to run the category and hazard check.")
                else:
                    try:
                        safety=check_disposal_item(obj or "Unidentified item",cat)
                        actual=safety.get("actual_category","Unclear")
                        if safety.get("is_hazardous"):
                            st.error("Potential hazardous waste found: "+str(safety.get("reason") or "Keep this item out of regular bins and follow campus hazardous-waste procedures."))
                        elif actual in TYPES and actual!=cat:
                            st.warning(f"Category mismatch: the AI identifies this as {actual}, but the selected category is {cat}. Nothing was added to the wrong bin. {safety.get('reason','')}")
                        else:
                            i=b.index[b["Bin ID"].astype(str)==bid][0];b.loc[i,"Fill (L)"]=min(float(b.loc[i,"Capacity (L)"]),float(b.loc[i,"Fill (L)"])+volume);row=b.loc[i];auto_added=schedule_pickup_if_needed(row);st.session_state.logs.loc[len(logs)]=[datetime.now().isoformat(timespec="seconds"),obj or "Unidentified item",cat,round(volume,2),bid,row.Name,RULES[cat]];st.session_state.bins=b;save();st.success("AI safety check passed. Saved and bin fill updated.")
                            if auto_added:st.warning(f"{row.Name} reached 90% capacity. A pickup was automatically added to the calendar.")
                    except Exception as ex:safe_ai_error(ex,"Waste safety check")
            else:st.error("Add a bin first.")
elif page=="Bin network":
    header("Bin network","Register, edit, and remove individual tracked bins.","NETWORK MANAGEMENT")
    with st.container(border=True):
        st.subheader(f"Registered bins · {len(b)}");edit=st.data_editor(b,hide_index=True,use_container_width=True,num_rows="fixed",column_config={"Category":st.column_config.SelectboxColumn(options=TYPES),"Capacity (L)":st.column_config.NumberColumn(min_value=1),"Fill (L)":st.column_config.NumberColumn(min_value=0),"Status":st.column_config.SelectboxColumn(options=["Online","Offline","Maintenance"])},key="bin_edit");l,r=st.columns(2)
        if l.button("Save network changes",type="primary",use_container_width=True):
            st.session_state.bins=edit
            triggered=[row["Name"] for _,row in edit.iterrows() if schedule_pickup_if_needed(row)]
            save()
            if triggered:st.toast("Automatic pickups added: "+", ".join(triggered))
            st.rerun()
        remove=r.selectbox("Bin to delete",b["Bin ID"].astype(str)) if not b.empty else None
        if r.button("Delete selected bin") and remove:st.session_state.bins=b[b["Bin ID"].astype(str)!=remove].reset_index(drop=True);save();st.rerun()
    with st.form("addbin"):
        st.subheader("Add a bin");a,c,d,e=st.columns(4);name=a.text_input("Bin name");zone=c.text_input("Zone");cat=d.selectbox("Waste category",TYPES);cap=e.number_input("Capacity (L)",min_value=1.,value=100.)
        if st.form_submit_button("＋ Add to network",use_container_width=True) and name.strip():
            new_id=next_bin_id(st.session_state.bins)
            st.session_state.bins.loc[len(st.session_state.bins)]={"Bin ID":new_id,"Name":name.strip(),"Zone":zone.strip(),"Category":cat,"Capacity (L)":cap,"Fill (L)":0.,"Status":"Online"}
            save();st.success(f"Added {new_id} to the network.");st.rerun()
    st.caption("Fill is estimated from item volumes. Physical sensors are needed for measured real-world fullness.")
elif page=="Monitor & trends":
    header("Monitor & trends","Spot bins needing pickups and understand waste patterns.","OPERATIONS INTELLIGENCE");pct=(b["Fill (L)"].astype(float)/b["Capacity (L)"].astype(float)*100).clip(0,100);a,c,d,e=st.columns(4);a.metric("Bins monitored",len(b));c.metric("Above 75%",int((pct>=75).sum()));d.metric("At capacity",int((pct>=99).sum()));e.metric("Disposals",len(logs))
    l,r=st.columns(2,gap="large")
    with l:
        with st.container(border=True):
            st.subheader("Digital twin fullness")
            for (_,x),p in zip(b.iterrows(),pct):st.write(f"**{x.Name}** · {x.Category} · {p:.0f}%");st.progress(int(p))
    with r:
        with st.container(border=True):
            st.subheader("Disposal trends")
            if logs.empty:st.info("Log disposals to build trends.")
            else:st.bar_chart(logs.groupby("Category").size().reindex(TYPES,fill_value=0),color=PALETTE["accent"])
    with st.container(border=True):
        st.subheader("Pickup watchlist");watch=b.assign(**{"Fill %":pct});watch=watch[watch["Fill %"]>=75].sort_values("Fill %",ascending=False);st.dataframe(watch,hide_index=True,use_container_width=True)
elif page=="AI advisor":
    header("AI advisor","Practical suggestions based on waste habits.","PERSONALIZED CAMPUS COACH");l,r=st.columns([1.2,1],gap="large")
    with l:
        with st.container(border=True):
            st.subheader("Ask Circular")
            for m in st.session_state.chat:
                with st.chat_message(m["role"]):st.write(m["content"])
            q=st.chat_input("Ask about waste or sorting…")
            if q:
                st.session_state.chat.append({"role":"user","content":q})
                api_key=gemini_api_key()
                if not api_key:
                    st.warning("Add your Gemini API key in the Live sorting page under ‘Connect / replace AI key’ before asking the advisor.")
                else:
                    try:
                        answer=ask_circular(q,history=st.session_state.chat[:-1])
                        st.session_state.chat.append({"role":"assistant","content":answer})
                        st.rerun()
                    except Exception as ex:
                        safe_ai_error(ex,"AI advisor")
    with r:
        with st.container(border=True):
            n=int(logs.Object.astype(str).str.lower().str.contains("plastic|bottle|pet",regex=True).sum()) if not logs.empty else 0;st.subheader("Habit insights");st.metric("Plastic/bottle items",n,"from saved records")
            if n>=3:st.markdown("<div class='tip'><b>Try a steel bottle.</b><br>Several bottles were recorded. Refillables can reduce single-use waste.</div>",unsafe_allow_html=True)
            else:st.info("Personal tips improve as your disposal log grows.")
elif page=="Pickup calendar":
    header("Pickup calendar","See bin fullness and scheduled bin cleanups.","CAMPUS OPERATIONS")
    if "calendar_month" not in st.session_state:st.session_state.calendar_month=date.today().replace(day=1)
    st.session_state.setdefault("calendar_selected",date.today())
    top_l,top_r=st.columns([1.2,1],gap="large")
    with top_l:
        with st.container(border=True):
            prev,month_title,nxt=st.columns([1,4,1])
            if prev.button("‹",key="cal_prev",use_container_width=True):
                current=st.session_state.calendar_month;new_month=(current.replace(day=1)-timedelta(days=1)).replace(day=1);st.session_state.calendar_month=new_month;st.session_state.calendar_selected=new_month;st.session_state.calendar_selected_picker=new_month;st.rerun()
            month_start=st.session_state.calendar_month
            month_title.markdown(f"<div style='text-align:center;font:700 20px Manrope;color:{PALETTE['text']};padding:6px'>{month_start.strftime('%B %Y')}</div>",unsafe_allow_html=True)
            if nxt.button("›",key="cal_next",use_container_width=True):
                current=st.session_state.calendar_month;new_month=(current.replace(day=28)+timedelta(days=4)).replace(day=1);st.session_state.calendar_month=new_month;st.session_state.calendar_selected=new_month;st.session_state.calendar_selected_picker=new_month;st.rerun()
            month_events=events.copy()
            event_types={}
            if not month_events.empty:
                month_events["_day"]=pd.to_datetime(month_events["Date"],errors="coerce").dt.date
                for _,event in month_events[month_events["Status"].astype(str)=="Scheduled"].iterrows():
                    day_value=event["_day"]
                    if not pd.isna(day_value) and day_value.year==month_start.year and day_value.month==month_start.month:
                        event_types.setdefault(day_value,set()).add(str(event["Type"]))
            weekday_labels=["SUN","MON","TUE","WED","THU","FRI","SAT"]
            calendar_markup="<div style='display:grid;grid-template-columns:repeat(7,minmax(0,1fr));gap:5px'>"
            for label in weekday_labels:
                calendar_markup+=f"<div style='text-align:center;color:{PALETTE['muted']};font-size:11px;font-weight:800;padding:7px 1px'>{label}</div>"
            weeks=calendar.Calendar(firstweekday=6).monthdayscalendar(month_start.year,month_start.month)
            for week in weeks:
                for day_num in week:
                    if not day_num:
                        calendar_markup+="<div style='height:48px'></div>";continue
                    this_date=date(month_start.year,month_start.month,day_num)
                    types=event_types.get(this_date,set())
                    if "Bin Cleanup" in types or "Bin pickup" in types:background="#d94f4b";color="#ffffff";border="#c43f3b";tag="CLEANUP"
                    elif "Bin Maintenance" in types or "Bin maintenance" in types:background=PALETTE["maintenance_bg"];color=PALETTE["maintenance_text"];border=PALETTE["maintenance_border"];tag="SERVICE"
                    else:background=PALETTE["calendar"];color=PALETTE["calendar_text"];border=PALETTE["calendar_border"];tag=""
                    if this_date==date.today() and "Bin Cleanup" not in types and "Bin pickup" not in types:border="#28745f"
                    detail=f"<div style='font-size:8px;font-weight:800;letter-spacing:.02em;line-height:1.05'>{tag}</div>" if tag else "<div style='height:9px'></div>"
                    calendar_markup+=f"<div style='box-sizing:border-box;min-height:49px;border:2px solid {border};border-radius:10px;background:{background};color:{color};display:flex;flex-direction:column;align-items:center;justify-content:center;padding:3px'><div style='font-size:15px;font-weight:800;line-height:1.1'>{day_num}</div>{detail}</div>"
            calendar_markup+="</div>"
            st.markdown(calendar_markup,unsafe_allow_html=True)
            st.caption("Red = scheduled bin cleanup · Amber = maintenance · Green outline = today")
            selected=st.date_input("Selected date",value=st.session_state.calendar_selected,key="calendar_selected_picker")
            if selected!=st.session_state.calendar_selected:
                st.session_state.calendar_selected=selected
                st.session_state.calendar_month=selected.replace(day=1)
                st.rerun()
            st.markdown(f"**Agenda · {st.session_state.calendar_selected.strftime('%A, %B %d')}**")
            if not month_events.empty:
                agenda=month_events[(month_events["_day"]==st.session_state.calendar_selected) & (month_events["Status"].astype(str)=="Scheduled")].drop(columns="_day")
                if agenda.empty:st.caption("Nothing scheduled for this date.")
                else:st.dataframe(agenda[["Time","Type","Bin ID","Notes","Status"]],hide_index=True,use_container_width=True)
            else:st.caption("Nothing scheduled for this date yet.")
    with top_r:
        with st.container(border=True):
            st.subheader("Schedule an activity")
            with st.form("event"):
                day=st.date_input("Date",st.session_state.calendar_selected);at=st.time_input("Time",time(14));kind=st.selectbox("Activity",["Bin Cleanup","Bin Maintenance"]);bid=st.selectbox("Bin",["ALL"]+b["Bin ID"].astype(str).tolist()) if not b.empty else "ALL";notes=st.text_input("Notes")
                if st.form_submit_button("＋ Schedule",use_container_width=True):
                    st.session_state.events.loc[len(st.session_state.events)]={"Date":str(day),"Time":at.strftime("%H:%M"),"Type":kind,"Bin ID":bid,"Notes":notes,"Status":"Scheduled"};save();st.rerun()
            if not events.empty:
                complete=events[events["Status"].astype(str)=="Scheduled"]
                if not complete.empty:
                    idx=st.selectbox("Mark scheduled activity complete",complete.index,format_func=lambda i:f"{complete.loc[i,'Date']} · {complete.loc[i,'Type']} · {complete.loc[i,'Bin ID']}")
                    if st.button("Mark complete"):
                        emptied_bin=complete_calendar_activity(idx)
                        if emptied_bin:st.toast(f"{emptied_bin} emptied in the digital twin.")
                        save();st.rerun()
    st.subheader("Bin fill levels")
    if b.empty:st.info("No bins in the network yet.")
    else:
        bin_cols=st.columns(3)
        for n,(_,row) in enumerate(b.iterrows()):
            pct=min(100,max(0,int(float(row["Fill (L)"])/max(1,float(row["Capacity (L)"]))*100)))
            bin_id=__import__("html").escape(str(row["Bin ID"]));name=__import__("html").escape(str(row["Name"]));zone=__import__("html").escape(str(row["Zone"]));kind=__import__("html").escape(str(row["Category"]))
            bin_cols[n%3].markdown(f"<div style='display:flex;align-items:center;gap:12px;background:{PALETTE['bin_card']};border:1px solid {PALETTE['border']};border-radius:14px;padding:12px;margin:5px 0'>{render_bin_icon(pct)}<div><div style='font-size:11px;font-weight:800;color:{PALETTE['bin_id']};letter-spacing:.04em'>{bin_id}</div><div style='font-weight:700;color:{PALETTE['bin_name']}'>{name}</div><div style='font-size:12px;color:{PALETTE['bin_detail']}'>{zone} · {kind}</div><div style='font-size:13px;color:{PALETTE['bin_value']}'>{pct}% full · {float(row['Fill (L)']):.1f}/{float(row['Capacity (L)']):.0f} L</div></div></div>",unsafe_allow_html=True)
    st.subheader("All upcoming bin cleanups and maintenance")
    scheduled=events[events["Status"].astype(str)=="Scheduled"].copy() if not events.empty else events
    if scheduled.empty:st.info("No activities scheduled yet. Bins automatically enter this list at 90% capacity.")
    else:
        scheduled["_sort_date"]=pd.to_datetime(scheduled["Date"],errors="coerce")
        scheduled=scheduled.sort_values(["_sort_date","Time"],na_position="last").drop(columns="_sort_date")
        st.dataframe(scheduled,hide_index=True,use_container_width=True)
else:
    header("Disposal log","Search and export timestamp, object, type, volume and destination.","AUDIT TRAIL")
    if logs.empty:st.info("No disposal records yet.")
    else:
        q=st.text_input("Filter records");view=logs[logs.astype(str).apply(lambda col:col.str.contains(q,case=False,na=False)).any(axis=1)] if q else logs;st.dataframe(view.sort_values("Timestamp",ascending=False),hide_index=True,use_container_width=True);st.download_button("Download CSV",view.to_csv(index=False),"circular-disposals.csv","text/csv")
    st.caption(f"Excel workbook: {BOOK}")
st.markdown(f"<div style='text-align:center;color:{PALETTE['footer']};font-size:10px;margin:25px 0'>CIRCULAR · SMART BIN NETWORK</div>",unsafe_allow_html=True)






