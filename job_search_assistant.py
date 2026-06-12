#!/usr/bin/env python3
"""Local-first job description evaluator for a personal job search."""

from __future__ import annotations

import argparse
import csv
import datetime as dt
from difflib import SequenceMatcher
import json
import re
import sqlite3
import sys
import textwrap
import zipfile
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape

from google_drive_sync import (
    DEFAULT_WORKSPACE_SCOPES,
    build_google_sheets_service,
    build_google_services,
    extract_drive_id,
    is_storage_quota_exceeded_error,
    upsert_google_doc_text,
)


ROOT = Path(__file__).resolve().parent
DEFAULT_PROFILE = ROOT / "profile.json"
DEFAULT_DB = ROOT / "data" / "job_results.sqlite"
LOCAL_TIMEZONE = dt.timezone(dt.timedelta(hours=-4))
DEFAULT_SHEETS_WORKBOOK_DIR = ROOT / "exports" / "sheets"
VOICE_PROFILE_MD = ROOT / "voice_application_style.md"
VOICE_PROFILE_JSON = ROOT / "voice_application_style.json"
PACKET_INDEX_JSON = ROOT / "exports" / "application_packets" / "packet_index.json"
APPLICATION_QUESTION_OVERRIDES_JSON = ROOT / "application_question_overrides.json"
GOOGLE_SHEETS_WORKBOOK_URL = ""
GOOGLE_PACKET_FOLDER_URL = ""
GOOGLE_COVER_LETTER_FOLDER_URL = ""
DEFAULT_GOOGLE_SERVICE_ACCOUNT_JSON = ROOT / "secrets" / "google-service-account.json"
DEFAULT_GOOGLE_OAUTH_CLIENT_JSON = ROOT / "secrets" / "google-oauth-client.json"
DEFAULT_GOOGLE_OAUTH_TOKEN_JSON = ROOT / "secrets" / "google-oauth-token.json"
GOOGLE_SHEETS_SYNC_SCOPES = DEFAULT_WORKSPACE_SCOPES

VALID_STATUSES = [
    "discovered",
    "shortlisted",
    "applied",
    "outreach_sent",
    "recruiter_reply",
    "interviewing",
    "offer",
    "rejected",
    "archived",
]

FINAL_STATUSES = {"offer", "rejected", "archived"}
STATUS_COLOR_RULES = {
    "discovered": {"red": 0.95, "green": 0.95, "blue": 0.95},
    "shortlisted": {"red": 0.82, "green": 0.90, "blue": 1.0},
    "applied": {"red": 0.80, "green": 0.94, "blue": 0.80},
    "outreach_sent": {"red": 0.79, "green": 0.92, "blue": 0.90},
    "recruiter_reply": {"red": 0.88, "green": 0.82, "blue": 0.96},
    "interviewing": {"red": 1.0, "green": 0.90, "blue": 0.72},
    "offer": {"red": 0.70, "green": 0.88, "blue": 0.70},
    "rejected": {"red": 0.96, "green": 0.78, "blue": 0.78},
    "archived": {"red": 0.86, "green": 0.86, "blue": 0.86},
}
SHEET_UI_ROW_LIMIT = 1000
STATUS_STAGE_RANK = {
    "discovered": 0,
    "shortlisted": 1,
    "outreach_sent": 2,
    "applied": 3,
    "recruiter_reply": 4,
    "interviewing": 5,
    "offer": 9,
    "rejected": 9,
    "archived": 9,
}
DEFAULT_FOLLOW_UP_DAYS = 7

PRIORITY_VALUES = [
    "P1 Apply Today",
    "P2 Strong",
    "P3 Maybe",
    "Review",
    "Park",
    "Active",
    "Closed",
]

SECTOR_PATTERNS: list[tuple[str, str]] = [
    ("Cybersecurity", r"cybersecurity|IT security|security|fraud|bug bounty|penetration testing|pen[- ]?test"),
    ("AI / Data Platform", r"\bAI\b|\bML\b|machine learning|data platform|analytics|dataiku|data product"),
    ("Robotics / Autonomy", r"robotics?|robotic automation|industrial automation|warehouse automation|autonomy|autonomous|unmanned|UAS|drones?"),
    ("Defense / Dual-Use", r"defense|defence|national security|dual[- ]use|aerospace|DoD|Department of Defense|federal|public sector"),
    ("Healthtech", r"healthcare|healthtech|health tech|medtech|digital health|medical|hospital|pharma|pharmaceutical"),
    ("Fintech", r"fintech|payments?|banking|financial services|financial institutions"),
    ("Web3 / Crypto", r"web3|crypto|blockchain|protocol|wallet|exchange|DeFi|token"),
    ("Gaming / Hardware", r"gaming|eSports?|Twitch|game|computer hardware|hardware"),
    ("SaaS / General", r"SaaS|software|platform|subscription|cloud"),
]

IDENTITY_SECTOR_PATTERNS: list[tuple[str, str]] = [
    ("Cybersecurity", r"flashpoint|inspectiv|security|fraud|bug bounty|penetration testing|pen[- ]?test"),
    ("Healthtech", r"sword health|\bwheel\b|health|medical|pharma"),
    ("Robotics / Autonomy", r"robust ai|robotics?|autonomy|autonomous"),
    ("AI / Data Platform", r"dataiku|render|15five|feathery|\bAI\b|data platform|analytics"),
    ("Defense / Dual-Use", r"onebrief|defense|national security|dual[- ]use|aerospace|DoD"),
    ("Web3 / Crypto", r"tether|trm labs|web3|crypto|blockchain|protocol|wallet|exchange|stablecoin"),
    ("Fintech", r"sardine|fintech|payments?|banking|financial"),
    ("Gaming / Hardware", r"gaming|eSports?|computer hardware|hardware"),
]

ROLE_LANE_PATTERNS: list[tuple[str, str]] = [
    ("Customer Success", r"customer success|client success|customer outcomes|deployments"),
    ("Strategic Account Management", r"strategic account|enterprise account|key account"),
    ("Partnerships / BD", r"partnerships?|partner ecosystem|business development|\bBD\b|ecosystem|expansion manager|blockchain expansion"),
    ("GTM / Commercial Lead", r"\bGTM\b|go-to-market|commercial lead|commercial strategy|director of customer success"),
    ("Sales / Enterprise", r"enterprise sales|account executive|\bAE\b|sales director|sales manager"),
    ("Account Management", r"account manager|account management|client manager"),
    ("Advisor / Ecosystem", r"advisor|advisory|ecosystem lead"),
]

SHEETS_JOB_FIELDNAMES = [
    "ID",
    "Status",
    "Priority",
    "Fit Score",
    "Fit Band",
    "Sector",
    "Role Lane",
    "Company",
    "Title",
    "Source Board",
    "Source URL",
    "Application URL",
    "Packet Status",
    "Packet Updated",
    "Packet Link",
    "Follow Up Date",
    "Last Touch",
    "Next Action",
    "Applied At",
    "Resume Version",
    "Cover Letter Needed",
    "Referral Target",
    "Concerns",
    "Top Fit Signals",
    "Created At",
    "Last Updated At",
]

PACKETS_FIELDNAMES = [
    "Job ID",
    "Company",
    "Title",
    "Priority",
    "Status",
    "Fit Score",
    "Packet Status",
    "Packet Updated",
    "Packet Link",
    "Packet Summary",
]

CONTACTS_FIELDNAMES = [
    "Job ID",
    "Company",
    "Title",
    "Name",
    "Relationship",
    "Role",
    "Telegram Handle",
    "Email",
    "LinkedIn",
    "Notes",
    "Added",
]

CORRESPONDENCE_FIELDNAMES = [
    "Date",
    "Job ID",
    "Company",
    "Title",
    "Contact",
    "Contact Relationship",
    "Channel",
    "Direction",
    "Type",
    "Summary",
    "Follow Up Needed",
    "Follow Up Date",
]

TARGET_COMPANY_FIELDNAMES = [
    "ID",
    "Company",
    "Website",
    "Lane",
    "Description",
    "Funding Date",
    "Funding Amount",
    "Round",
    "Investors",
    "Company Fit Score",
    "Open Roles Found",
    "Best Role Title",
    "Role Fit Score",
    "Role URL",
    "Priority",
    "Target Strategy",
    "Outreach Type",
    "Warm Contact 1",
    "Warm Contact 1 Title",
    "Warm Contact 1 LinkedIn",
    "Warm Contact 2",
    "Warm Contact 2 Title",
    "Warm Contact 2 LinkedIn",
    "Outreach Angle",
    "Outreach Status",
    "Application Status",
    "Notes",
    "Next Action",
    "Last Checked",
    "Source URL",
    "Created At",
    "Last Updated At",
]

RECENT_ACCOUNT_OWNERSHIP_SENTENCE = (
    "In a recent account management role, I managed strategic accounts across a technical product suite."
)

RECENT_ACCOUNT_REVENUE_OUTCOMES_SENTENCE = (
    "I influenced meaningful renewal, recovery, and expansion outcomes across strategic accounts."
)

RECENT_ACCOUNT_REVENUE_BREAKDOWN_SENTENCE = (
    "That included closed, recovered, and post-transition outcomes tied to accounts and opportunities I helped advance."
)

RECENT_ACCOUNT_RECOVERY_SENTENCE = (
    "I also helped reactivate a previously stalled account through targeted stakeholder engagement and commercial follow-through."
)

RECENT_ACCOUNT_CROSS_FUNCTIONAL_SENTENCE = (
    "The role required close coordination with research, product, sales, and client stakeholders so delivery "
    "stayed aligned with customer goals and commercial outcomes."
)


POSITIVE_PATTERNS: list[tuple[str, int, str]] = [
    (r"\brenewals?\b|\brenewal strategy\b", 6, "renewal ownership"),
    (r"\bexpansion\b|\bupsell\b|\bcross-sell\b|\bnet revenue retention\b|\bNRR\b", 7, "expansion / NRR focus"),
    (r"\bgross revenue retention\b|\bGRR\b", 5, "GRR focus"),
    (r"\bstrategic account\b|\benterprise account\b|\btier 1\b|\bkey account\b", 7, "strategic or enterprise account scope"),
    (r"\bnamed accounts?\b|\bbook of business\b|\bportfolio of accounts?\b|\benterprise clients?\b|\bkey customers?\b", 6, "named-account or portfolio ownership"),
    (r"\binstitutional partners?\b|\binstitutional clients?\b|\bcorporate stakeholders?\b|\bgovernment stakeholders?\b", 6, "institutional / multi-stakeholder relationship ownership"),
    (r"\bwallets?\b|\bexchanges?\b|\bpayment processors?\b|\bpayment infrastructure\b|\btokenization\b|\bcustodial solutions?\b|\bstablecoins?\b|\bUSDT\b", 7, "wallet / exchange / stablecoin infrastructure"),
    (r"\bpayments companies\b|\bfinancial institutions\b|\benterprise financial platforms\b|\bfintech companies\b|\bpayments ecosystem\b|\bglobal money movement\b|\bsettlement payments\b", 6, "payments / financial-infrastructure partner network"),
    (r"\bcustomer success\b.{0,80}\b(renewal|expansion|upsell|commercial|revenue|growth)\b|\b(renewal|expansion|upsell|commercial|revenue|growth)\b.{0,80}\bcustomer success\b", 6, "customer success with commercial ownership"),
    (r"\brevenue growth\b|\bpartner revenue\b|\bpartner-led revenue\b|\brevenue-generating partnerships?\b|\bpartnerships? tied to revenue\b", 6, "partnerships tied to revenue"),
    (r"\bpartnership frameworks?\b|\bpilots?\b|\bjoint initiatives?\b|\bcomplex multi[- ]party partnerships?\b|\bstrategic adoption\b|\bpartner-led distribution\b", 6, "strategic partnerships execution"),
    (r"\becosystem management\b|\bblockchain partner ecosystem\b|\bblockchain foundations?\b|\blayer 1\b|\blayer 2\b|\bL1\b|\bL2\b|\breal-world use cases?\b|\basset usage\b", 7, "blockchain ecosystem expansion"),
    (r"\bcross-functional\b|\bproduct\b.{0,40}\bsales\b|\bsales\b.{0,40}\bproduct\b|\bresearch\b.{0,40}\bleadership\b|\bleadership\b.{0,40}\bresearch\b", 5, "cross-functional product/sales/research coordination"),
    (r"\bproject governance\b|\bprogram coordination\b|\bimplementation progress\b|\brisk registers?\b|\bdecision records?\b|\bexecutive reports?\b", 4, "program governance and delivery oversight"),
    (r"\bpartnerships?\b|\bpartner ecosystem\b|\bchannel partners?\b", 6, "partnerships motion"),
    (r"\bexecutive stakeholders?\b|\bC-level\b|\bsenior stakeholders?\b", 5, "executive stakeholder work"),
    (r"\bSaaS\b|\bsoftware\b|\bplatform\b|\bsubscription\b", 5, "SaaS / platform context"),
    (r"\bfintech\b|\bpayments?\b|\bbanking\b|\bfinancial services\b", 5, "fintech domain"),
    (r"\bhealthcare\b|\bhealthtech\b|\bhealth tech\b|\bmedtech\b|\bdigital health\b|\bmedical\b|\bhospital\b|\bpharma\b|\bpharmaceutical\b", 5, "healthcare domain"),
    (r"\bAI\b|\bML\b|\bmachine learning\b|\bdata platform\b|\banalytics\b", 5, "AI/data platform domain"),
    (r"\brobotics?\b|\brobotic automation\b|\bindustrial automation\b|\bwarehouse automation\b|\bphysical automation\b|\bautonomy\b|\bautonomous (?:systems?|vehicles?|robots?|warehouse|platforms?)\b|\bunmanned\b|\bUAS\b|\bdrones?\b", 5, "robotics / autonomy domain"),
    (r"\bdefense tech\b|\bdefence tech\b|\bdefense\b|\bnational security\b|\bdual[- ]use\b|\baerospace\b|\bDoD\b|\bDepartment of Defense\b|\bfederal\b|\bpublic sector\b", 5, "defense / national security domain"),
    (r"\bweb3\b|\bcrypto\b|\bblockchain\b|\bprotocol\b|\bwallet\b|\bexchange\b", 5, "Web3 / crypto domain"),
    (r"\bgaming\b|\besports?\b|\bTwitch\b|\bgame\b|\bcomputer hardware\b|\bhardware\b", 4, "gaming / eSports / hardware domain"),
    (r"\bcybersecurity\b|\bIT security\b|\bsecurity\b|\bfraud\b|\bbug bounty\b|\bpenetration testing\b|\bpen[- ]?test\b", 4, "IT security / cybersecurity domain"),
    (r"\bSalesforce\b|\bHubSpot\b|\bCRM\b", 3, "CRM fluency"),
    (r"\bforecasting\b|\bpipeline\b|\bcommercial strategy\b|\bGTM\b|\bgo-to-market\b", 5, "commercial operating rigor"),
    (r"\bremote\b|\bdistributed\b|\bwork from anywhere\b", 6, "remote-friendly setup"),
]

RED_FLAG_PATTERNS: list[tuple[str, int, str]] = [
    (r"\bon[- ]?site\b|\bin office\b|\bhybrid\b|\bcommute\b", 14, "onsite or hybrid requirement"),
    (r"\bcustomer support\b|\bsupport specialist\b|\bsupport representative\b|\bcustomer service\b|\bsupport desk\b|\bticket queue\b|\bqueue management\b|\bcall center\b", 18, "support/ticketing-first role"),
    (r"\bSDR\b|\bBDR\b|\bsales development\b|\bbusiness development representative\b", 16, "SDR/BDR-first role"),
    (r"\bcold outbound\b|\bcold calling\b|\boutbound prospecting\b|\bhigh-volume prospecting\b", 12, "outbound-heavy acquisition role"),
    (r"\bcommunity manager\b|\bcommunity growth\b|\bcommunity engagement\b|\bdiscord community\b|\btelegram community\b", 16, "community-first role"),
    (r"\bcontent marketing\b|\bcontent strategy\b|\bgrowth marketing\b|\bdemand generation\b|\bbrand marketing\b|\bcopywriter\b|\bSEO\b", 16, "content/marketing role"),
    (r"\bentry[- ]level\b|\bjunior\b|\bassociate account manager\b", 12, "too junior for seniority"),
    (r"\bweapons?\b|\bmilitary targeting\b|\blethal\b|\bsurveillance\b", 18, "weapons/lethal-use/surveillance requires review"),
    (r"\b100% commission\b|\bcommission only\b|\blow base\b", 16, "compensation structure risk"),
    (r"\bimplementation specialist\b|\bonboarding coordinator\b|\btraining coordinator\b", 10, "implementation/onboarding-heavy without commercial ownership"),
]

REGIONAL_TITLE_PATTERNS: list[tuple[str, str, tuple[str, ...]]] = [
    ("Northeast", r"\bnortheast\b|\bnorth east\b", ("new york", "new jersey", "massachusetts", "boston", "philadelphia", "washington", "dc", "connecticut", "northeast")),
    ("East", r"\beast\b", ("new york", "new jersey", "massachusetts", "boston", "philadelphia", "washington", "dc", "east coast", "east")),
    ("West", r"\bwest\b", ("california", "san francisco", "los angeles", "seattle", "west coast", "west")),
    ("Central", r"\bcentral\b", ("texas", "chicago", "central")),
    ("EMEA", r"\bemea\b", ("europe", "middle east", "africa", "emea")),
    ("LATAM", r"\blatam\b", ("latin america", "latam")),
]

ATS_KEYWORD_CANDIDATES = [
    "strategic account management",
    "renewals",
    "expansion",
    "retention",
    "account health",
    "stakeholder management",
    "executive relationships",
    "procurement",
    "commercial negotiation",
    "cross-functional coordination",
    "value realization",
    "executive business reviews",
    "churn mitigation",
    "upsell",
    "customer growth",
    "quota",
    "pipeline",
    "attainment",
    "acv",
    "enterprise accounts",
    "territory planning",
    "account mapping",
    "meddicc",
    "champion mapping",
    "economic buyer",
    "mutual action plan",
    "forecast",
    "sales methodology",
    "cybersecurity",
    "threat intelligence",
    "data security",
    "intelligence",
    "risk",
    "threat actors",
    "security operations",
    "fraud",
    "vulnerability",
    "enterprise risk",
]

DIRECT_DOMAIN_EXPERIENCE_PATTERNS: dict[str, str] = {
    "Cybersecurity": r"cybersecurity|cyber threat|threat intelligence|threat intel|vulnerability intelligence|brand protection|fraud prevention|physical security",
    "AI / Data Platform": r"\bAI\b|machine learning|analytics|data platform|data product",
    "Robotics / Autonomy": r"robotics?|autonomy|autonomous|industrial automation|warehouse automation",
    "Defense / Dual-Use": r"defense|defence|national security|dual[- ]use|aerospace|DoD|Department of Defense|federal",
    "Healthtech": r"healthcare|healthtech|medtech|medical|hospital|pharma|pharmaceutical",
    "Fintech": r"fintech|payments?|banking|financial services|financial institutions",
    "Web3 / Crypto": r"web3|crypto|blockchain|protocol|wallet|exchange|defi|stablecoin",
}

DOMAIN_PREFERENCE_PATTERNS: dict[str, str] = {
    "Cybersecurity": r"ideally in cybersecurity|threat intelligence|cyber and physical threats|security challenges|vulnerability intelligence|brand protection|fraud and brand protection",
    "AI / Data Platform": r"data platform|analytics|AI infrastructure|machine learning",
    "Robotics / Autonomy": r"robotics?|autonomy|autonomous|warehouse automation|industrial automation",
    "Defense / Dual-Use": r"defense|defence|national security|dual[- ]use|aerospace|federal|DoD|Department of Defense",
    "Healthtech": r"healthcare|healthtech|medtech|medical|hospital|pharma|pharmaceutical",
    "Fintech": r"fintech|payments?|banking|financial services|financial institutions",
    "Web3 / Crypto": r"web3|crypto|blockchain|protocol|wallet|exchange|stablecoin",
}

DOMAIN_GAP_PENALTY = 12
REGIONAL_TERRITORY_PENALTY = 5


def load_profile(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as fh:
        return json.load(fh)


def load_voice_profile(path: Path = VOICE_PROFILE_JSON) -> dict[str, Any]:
    if not path.exists():
        return {}
    try:
        with path.open("r", encoding="utf-8") as fh:
            data = json.load(fh)
    except (json.JSONDecodeError, OSError):
        return {}
    return data if isinstance(data, dict) else {}


def load_packet_index(path: Path = PACKET_INDEX_JSON) -> dict[str, Any]:
    if not path.exists():
        return {}
    try:
        with path.open("r", encoding="utf-8") as fh:
            data = json.load(fh)
    except (json.JSONDecodeError, OSError):
        return {}
    return data if isinstance(data, dict) else {}


def load_application_question_overrides(path: Path = APPLICATION_QUESTION_OVERRIDES_JSON) -> dict[str, Any]:
    if not path.exists():
        return {}
    try:
        with path.open("r", encoding="utf-8") as fh:
            data = json.load(fh)
    except (json.JSONDecodeError, OSError):
        return {}
    return data if isinstance(data, dict) else {}


def save_application_question_overrides(overrides: dict[str, Any], path: Path = APPLICATION_QUESTION_OVERRIDES_JSON) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as fh:
        json.dump(overrides, fh, indent=2, sort_keys=True)


def application_question_override_for_job(job: dict[str, Any], path: Path = APPLICATION_QUESTION_OVERRIDES_JSON) -> dict[str, Any]:
    overrides = load_application_question_overrides(path)
    override = overrides.get(str(job.get("id")))
    return override if isinstance(override, dict) else {}


def save_packet_index(index: dict[str, Any], path: Path = PACKET_INDEX_JSON) -> None:
    ensure_packet_index_meta(index)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as fh:
        json.dump(index, fh, indent=2, sort_keys=True)


def find_local_packet_path(job_id: int) -> Path | None:
    index = load_packet_index()
    entry = index.get(str(job_id)) if isinstance(index, dict) else None
    if isinstance(entry, dict):
        local_path = str(entry.get("local_path") or "").strip()
        if local_path and Path(local_path).exists():
            return Path(local_path)
    pattern = f"job-{job_id}-*.md"
    candidates = sorted(PACKET_INDEX_JSON.parent.glob(pattern))
    if candidates:
        return candidates[0]
    bundle_candidates = sorted((ROOT / "job_packets").glob(f"job-{job_id}-*/00_packet_bundle.md"))
    return bundle_candidates[0] if bundle_candidates else None


def infer_cover_letter_docx_path(packet_path: str) -> str:
    clean = str(packet_path or "").strip()
    if not clean:
        return ""
    packet = Path(clean)
    if not packet.exists() or not packet.name:
        return ""
    if packet.name == "00_packet_bundle.md":
        docx_path = packet.with_name("04_cover_letter.docx")
    else:
        docx_path = packet.with_name(f"{packet.stem}-cover-letter.docx")
    return str(docx_path.resolve()) if docx_path.exists() else ""


def ensure_packet_index_meta(index: dict[str, Any]) -> dict[str, Any]:
    meta = index.get("_meta") if isinstance(index.get("_meta"), dict) else {}
    meta["google_drive_folder_url"] = GOOGLE_PACKET_FOLDER_URL
    meta["google_cover_letter_folder_url"] = GOOGLE_COVER_LETTER_FOLDER_URL
    index["_meta"] = meta
    return index


def packet_display_status(base_status: str, question_status: str) -> str:
    clean_base = str(base_status or "").strip()
    clean_question_status = str(question_status or "").strip()
    if not clean_base:
        return ""
    if clean_question_status == "Not Captured":
        return f"{clean_base} - Needs Questions"
    if clean_question_status == "Partially Captured":
        return f"{clean_base} - Needs More Questions"
    return clean_base


def packet_question_prompt(question_status: str) -> str:
    clean_status = str(question_status or "").strip()
    if clean_status == "Not Captured":
        return "Action needed: paste the exact application questions."
    if clean_status == "Partially Captured":
        return "Action needed: add any later-step application questions."
    return ""


def packet_metadata_for_job(row: dict[str, Any], packet_index: dict[str, Any] | None = None) -> dict[str, str]:
    index = packet_index if packet_index is not None else load_packet_index()
    entry = index.get(str(row["id"])) if isinstance(index, dict) else None
    if not isinstance(entry, dict):
        entry = {}

    local_path = str(entry.get("local_path") or "").strip()
    if not local_path:
        discovered = find_local_packet_path(int(row["id"]))
        if discovered is not None:
            local_path = str(discovered.resolve())
    doc_url = str(entry.get("google_doc_url") or "").strip()
    cover_letter_docx_path = str(entry.get("cover_letter_docx_path") or "").strip() or infer_cover_letter_docx_path(local_path)
    cover_letter_doc_url = str(entry.get("google_cover_letter_doc_url") or "").strip()
    updated_at = str(entry.get("updated_at") or "").strip()
    exists_locally = bool(local_path) and Path(local_path).exists()
    if not updated_at and exists_locally:
        updated_at = dt.datetime.fromtimestamp(Path(local_path).stat().st_mtime, tz=dt.timezone.utc).isoformat(timespec="seconds")
    question_status = application_questions_status_payload(application_question_override_for_job(row))["status"]
    if doc_url:
        status = "Linked"
    elif exists_locally:
        status = "Local Only"
    else:
        status = ""
    return {
        "status": packet_display_status(status, question_status),
        "storage_status": status,
        "question_status": question_status,
        "question_prompt": packet_question_prompt(question_status),
        "updated": display_datetime(updated_at),
        "link": doc_url or local_path,
        "doc_url": doc_url,
        "local_path": local_path,
        "cover_letter_docx_path": cover_letter_docx_path,
        "cover_letter_doc_url": cover_letter_doc_url,
    }


def update_packet_index_for_job(
    job: dict[str, Any],
    packet_path: Path,
    *,
    cover_letter_docx_path: Path | None = None,
    doc_url: str | None = None,
    path: Path = PACKET_INDEX_JSON,
) -> None:
    index = load_packet_index(path)
    ensure_packet_index_meta(index)
    key = str(job["id"])
    previous = index.get(key, {}) if isinstance(index.get(key), dict) else {}
    entry = {
        "job_id": int(job["id"]),
        "company": job.get("company") or "",
        "title": job.get("title") or "",
        "local_path": str(packet_path.resolve()),
        "cover_letter_docx_path": (
            str(cover_letter_docx_path.resolve())
            if cover_letter_docx_path is not None
            else str(previous.get("cover_letter_docx_path") or "")
        ),
        "updated_at": now_utc(),
        "google_doc_url": doc_url or previous.get("google_doc_url", ""),
        "google_cover_letter_doc_url": str(previous.get("google_cover_letter_doc_url") or ""),
    }
    index[key] = entry
    save_packet_index(index, path)


def set_packet_doc_url(job_id: int, doc_url: str, *, path: Path = PACKET_INDEX_JSON) -> None:
    index = load_packet_index(path)
    ensure_packet_index_meta(index)
    key = str(job_id)
    previous = index.get(key)
    if not isinstance(previous, dict):
        raise SystemExit(f"No packet index entry found for job #{job_id}. Generate the packet first.")
    previous["google_doc_url"] = doc_url.strip()
    previous["updated_at"] = now_utc()
    index[key] = previous
    save_packet_index(index, path)


def set_cover_letter_doc_url(job_id: int, doc_url: str, *, path: Path = PACKET_INDEX_JSON) -> None:
    index = load_packet_index(path)
    ensure_packet_index_meta(index)
    key = str(job_id)
    previous = index.get(key)
    if not isinstance(previous, dict):
        raise SystemExit(f"No packet index entry found for job #{job_id}. Generate the packet first.")
    previous["google_cover_letter_doc_url"] = doc_url.strip()
    previous["updated_at"] = now_utc()
    index[key] = previous
    save_packet_index(index, path)


def clean_packet_signal(signal: str) -> str:
    clean = normalize(signal)
    replacements = {
        "senior/strategic seniority language": "senior scope",
        "strategic account manager context": "strategic account scope",
        "strategic or enterprise account scope": "enterprise account scope",
        "institutional / multi-stakeholder relationship ownership": "institutional stakeholder work",
        "expansion / NRR focus": "expansion focus",
        "SaaS / platform context": "platform GTM",
        "AI/data platform domain": "data platform domain",
        "wallet / exchange / stablecoin infrastructure": "wallet, exchange, and stablecoin infrastructure",
        "payments / financial-infrastructure partner network": "payments and financial-infrastructure overlap",
        "strategic partnerships execution": "strategic partnerships execution",
        "blockchain ecosystem expansion": "blockchain ecosystem expansion",
        "partnerships / ecosystem expansion title signal": "ecosystem expansion role",
    }
    if clean.lower().startswith("dream title signal:"):
        return ""
    if clean.lower().startswith("adjacent title signal:"):
        return ""
    return replacements.get(clean, clean)


def summarize_packet_signals(signals: list[str]) -> list[str]:
    preferred: list[str] = []
    fallback: list[str] = []
    generic_terms = {"senior scope", "platform GTM"}
    for signal in signals:
        cleaned = clean_packet_signal(signal)
        if not cleaned:
            continue
        if cleaned not in fallback:
            fallback.append(cleaned)
        if cleaned not in generic_terms and cleaned not in preferred:
            preferred.append(cleaned)
    selected = preferred or fallback
    return selected[:3]


def packet_role_hint(role_lane: str) -> str:
    role_lane = normalize(role_lane)
    mapping = {
        "Strategic Account Management": "Strategic AM",
        "Sales / Enterprise": "Enterprise sales",
        "Partnerships / BD": "Partnerships",
        "GTM / Commercial Lead": "GTM",
        "Customer Success": "Customer success",
        "Account Management": "Account management",
        "Advisor / Ecosystem": "Ecosystem",
    }
    return mapping.get(role_lane, role_lane or "Packet")


def join_readable(items: list[str]) -> str:
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    if len(items) == 2:
        return f"{items[0]} and {items[1]}"
    return f"{', '.join(items[:-1])}, and {items[-1]}"


def packet_assets_summary(packet: dict[str, str]) -> str:
    assets: list[str] = []
    if packet.get("local_path"):
        assets.append("local packet")
    if packet.get("cover_letter_docx_path"):
        assets.append("cover-letter docx")
    if packet.get("doc_url"):
        assets.append("mirrored Google Doc")
    if packet.get("cover_letter_doc_url"):
        assets.append("mirrored cover-letter Google Doc")
    return join_readable(assets)


def packet_summary_text(job: dict[str, Any], packet: dict[str, str]) -> str:
    signals = summarize_packet_signals(json_list(job.get("matched_signals")))
    role_hint = packet_role_hint(dashboard_metadata_for_job(job)["role_lane"])
    summary_parts: list[str] = []
    if packet.get("question_prompt"):
        summary_parts.append(str(packet["question_prompt"]).strip())
    if signals:
        summary_parts.append(f"{role_hint} focus: {join_readable(signals)}.")
    assets = packet_assets_summary(packet)
    if assets:
        summary_parts.append(f"Includes {assets}.")
    return " ".join(summary_parts).strip() or f"{role_hint} packet ready."


def packet_hover_note(job: dict[str, Any], packet: dict[str, str]) -> str:
    metadata = dashboard_metadata_for_job(job)
    lines = [f"{job.get('company', '')} - {job.get('title', '')}".strip(" -")]
    lines.append(f"Priority: {metadata['priority']} | Fit: {job.get('fit_score', '')}")
    if packet.get("question_status"):
        lines.append(f"Application Questions: {packet['question_status']}")
    if packet.get("question_prompt"):
        lines.append(str(packet["question_prompt"]).replace("Action needed: ", "Next step: "))
    signals = summarize_packet_signals(json_list(job.get("matched_signals")))
    if signals:
        lines.append(f"Best overlap: {join_readable(signals)}")
    assets = packet_assets_summary(packet)
    if assets:
        lines.append(f"Assets: {assets}")
    if packet.get("updated"):
        lines.append(f"Updated: {packet['updated']}")
    return "\n".join(line for line in lines if line)


def normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def now_utc() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")


def today_local() -> dt.date:
    return dt.datetime.now().date()


def normalize_status(status: str) -> str:
    clean = status.strip().lower().replace(" ", "_").replace("-", "_")
    aliases = {
        "ignore": "archived",
        "ignored": "archived",
        "pass": "archived",
        "outreach": "outreach_sent",
        "reply": "recruiter_reply",
        "recruiter": "recruiter_reply",
        "interview": "interviewing",
    }
    clean = aliases.get(clean, clean)
    if clean not in VALID_STATUSES:
        raise SystemExit(f"Unknown status '{status}'. Use one of: {', '.join(VALID_STATUSES)}")
    return clean


def parse_date(value: str | None, *, field_name: str = "date") -> str | None:
    if not value:
        return None
    clean = value.strip()
    if clean.lower() in {"today", "now"}:
        return today_local().isoformat()
    if clean.lower() == "tomorrow":
        return (today_local() + dt.timedelta(days=1)).isoformat()
    if re.fullmatch(r"\+\d+", clean):
        return (today_local() + dt.timedelta(days=int(clean[1:]))).isoformat()
    try:
        return dt.date.fromisoformat(clean).isoformat()
    except ValueError as exc:
        raise SystemExit(f"Invalid {field_name} '{value}'. Use YYYY-MM-DD, today, tomorrow, or +N.") from exc


def clamp_score(value: int | None) -> int | None:
    if value is None:
        return None
    return max(0, min(100, int(value)))


def company_priority_from_score(score: int | None) -> str:
    value = clamp_score(score) or 0
    if value >= 85:
        return "Tier 1 target"
    if value >= 70:
        return "Tier 2 target"
    if value >= 55:
        return "Monitor"
    return "Skip"


def default_follow_up_date(days: int = DEFAULT_FOLLOW_UP_DAYS) -> str:
    return (today_local() + dt.timedelta(days=days)).isoformat()


def infer_source_board(source: str | None) -> str | None:
    if not source:
        return None
    lowered = source.lower()
    board_map = {
        "jobs.dragonfly.xyz": "Dragonfly Jobs",
        "jobs.a16z.com": "a16z Jobs",
        "jobs.multicoin.capital": "Multicoin Jobs",
        "jobs.panteracapital.com": "Pantera Jobs",
        "paradigm.xyz": "Paradigm Careers",
        "jobs.solana.com": "Solana Jobs",
        "jobs.ashbyhq.com": "Ashby",
        "jobs.lever.co": "Lever",
        "jobs.gem.com": "Gem",
        "boards.greenhouse.io": "Greenhouse",
        "greenhouse.io": "Greenhouse",
        "careers.tether.io": "Tether Careers",
        "tether.recruitee.com": "Tether Careers",
        "wellfound.com": "Wellfound",
        "cryptojobslist.com": "CryptoJobsList",
        "web3.career": "Web3.career",
        "cryptocurrencyjobs.co": "Cryptocurrency Jobs",
        "remote3.co": "Remote3",
        "linkedin.com": "LinkedIn",
    }
    for needle, label in board_map.items():
        if needle in lowered:
            return label
    if re.match(r"https?://", source):
        return "Company/Other"
    return source[:80]


def read_job_description(args: argparse.Namespace) -> str:
    sources = [bool(args.text), bool(args.file), not sys.stdin.isatty()]
    if sum(sources) == 0:
        raise SystemExit("Provide a job description with --text, --file, or stdin.")
    if args.text:
        return args.text
    if args.file:
        return Path(args.file).read_text(encoding="utf-8")
    return sys.stdin.read()


def extract_company_and_title(jd: str, fallback_title: str | None, fallback_company: str | None) -> tuple[str, str]:
    if fallback_title or fallback_company:
        return fallback_title or "Unknown Title", fallback_company or "Unknown Company"

    lines = [line.strip(" -|") for line in jd.splitlines() if line.strip()]
    first_lines = lines[:10]
    title = "Unknown Title"
    company = "Unknown Company"

    title_markers = [
        r"job title:\s*(.+)",
        r"title:\s*(.+)",
        r"position:\s*(.+)",
        r"role:\s*(.+)",
    ]
    company_markers = [
        r"company:\s*(.+)",
        r"organization:\s*(.+)",
        r"employer:\s*(.+)",
    ]

    for line in first_lines:
        for pattern in title_markers:
            match = re.search(pattern, line, flags=re.IGNORECASE)
            if match:
                title = clean_header_value(match.group(1))
        for pattern in company_markers:
            match = re.search(pattern, line, flags=re.IGNORECASE)
            if match:
                company = clean_header_value(match.group(1))

    if title == "Unknown Title" and first_lines:
        title = clean_header_value(first_lines[0])
    if company == "Unknown Company" and len(first_lines) > 1:
        possible = first_lines[1]
        if len(possible.split()) <= 6 and not re.search(r"\b(remote|full-time|job|role)\b", possible, re.I):
            company = clean_header_value(possible)

    return title[:120], company[:120]


def clean_header_value(value: str) -> str:
    value = re.sub(r"\s+[-|]\s+.*$", "", value.strip())
    value = re.sub(r"\(.*?\)", "", value).strip()
    return value or "Unknown"


def find_matches(text: str, patterns: list[tuple[str, int, str]]) -> list[dict[str, Any]]:
    matches: list[dict[str, Any]] = []
    for pattern, weight, label in patterns:
        if re.search(pattern, text, flags=re.IGNORECASE):
            matches.append({"label": label, "weight": weight, "pattern": pattern})
    return matches


def title_score(title: str, jd: str, profile: dict[str, Any]) -> tuple[int, list[str]]:
    haystack = f"{title}\n{jd}".lower()
    dream_titles = profile["target_job_profile"]["dream_titles"]
    adjacent_titles = profile["target_job_profile"]["adjacent_titles"]
    matched: list[str] = []
    score = 0

    for dream in dream_titles:
        if dream.lower() in haystack:
            matched.append(f"dream title signal: {dream}")
            score = max(score, 20)
    for adjacent in adjacent_titles:
        if adjacent.lower() in haystack:
            matched.append(f"adjacent title signal: {adjacent}")
            score = max(score, 14)

    if re.search(r"\baccount manager\b", haystack):
        strategic_context = re.search(
            r"\binstitutional|strategic|enterprise|partner|wallet|exchange|payment|tokenization|stablecoin|blockchain|crypto|fintech|expansion|governance\b",
            haystack,
        )
        if strategic_context:
            matched.append("strategic account manager context")
            score = max(score, 18)
        else:
            matched.append("account manager title signal")
            score = max(score, 12)

    if re.search(r"\b(expansion manager|ecosystem manager|blockchain expansion)\b", haystack):
        matched.append("partnerships / ecosystem expansion title signal")
        score = max(score, 18)

    seniority_terms = ["senior", "lead", "principal", "head", "director", "enterprise", "strategic"]
    if any(term in haystack for term in seniority_terms):
        score += 5
        matched.append("senior/strategic seniority language")

    return min(score, 25), matched


def comp_adjustment(text: str, floor: int, ceiling: int) -> tuple[int, str | None]:
    base_context = re.findall(
        r"(?:base salary|annual base salary|base pay|salary)\D{0,80}((?:\$?\b[1-2]?\d{2},?\d{3}\b\D{0,20}){1,3})",
        text,
        flags=re.IGNORECASE,
    )
    if base_context:
        base_salaries = [
            int(raw.replace(",", ""))
            for chunk in base_context
            for raw in re.findall(r"\$?\b([1-2]?\d{2},?\d{3})\b", chunk)
        ]
        base_salaries = [salary for salary in base_salaries if 50000 <= salary <= 400000]
        if base_salaries:
            high = max(base_salaries)
            low = min(base_salaries)
            if high < floor:
                return -14, f"listed base appears below ${floor:,} target"
            if low < floor <= high:
                return -5, f"listed base range starts below ${floor:,} target"
            if low >= floor and high <= ceiling:
                return 5, f"listed base overlaps target range (${floor:,}-${ceiling:,})"
            return 3, "listed base appears to clear the floor"

    salaries = [int(raw.replace(",", "")) for raw in re.findall(r"\$?\b([1-2]?\d{2},?\d{3})\b", text)]
    salaries = [salary for salary in salaries if 50000 <= salary <= 400000]
    if not salaries:
        return 0, None

    high = max(salaries)
    low = min(salaries)
    if high < floor:
        return -14, f"listed comp appears below ${floor:,} base target"
    if low >= floor and high <= ceiling:
        return 5, f"listed comp overlaps target range (${floor:,}-${ceiling:,})"
    if high >= floor:
        return 3, "listed total comp appears to clear the floor, but base is unclear"
    return 0, None


def profile_professional_text(profile: dict[str, Any]) -> str:
    fragments: list[str] = []
    for key in ("roles", "earlier_relevant_roles"):
        entries = profile.get(key, [])
        if not isinstance(entries, list):
            continue
        for entry in entries:
            if not isinstance(entry, dict):
                continue
            fragments.append(str(entry.get("title") or ""))
            fragments.append(str(entry.get("company") or ""))
            for list_key in ("impact", "skills"):
                values = entry.get(list_key, [])
                if isinstance(values, list):
                    fragments.extend(str(value) for value in values)
    return " ".join(fragment for fragment in fragments if fragment)


def has_direct_professional_domain_experience(profile: dict[str, Any], sector: str) -> bool:
    pattern = DIRECT_DOMAIN_EXPERIENCE_PATTERNS.get(sector)
    if not pattern:
        return True
    return bool(re.search(pattern, profile_professional_text(profile), flags=re.IGNORECASE))


def domain_affinity_key(sector: str) -> str:
    return normalize(sector).lower().replace("/", " ").replace("-", " ").replace(" ", "_")


def personal_domain_affinity_bonus(profile: dict[str, Any], sector: str) -> tuple[int, list[str]]:
    affinities = profile.get("personal_domain_affinity")
    if not isinstance(affinities, dict):
        return 0, []
    affinity = affinities.get(domain_affinity_key(sector))
    if not isinstance(affinity, dict):
        return 0, []
    bonus = affinity.get("bonus_points", 0)
    try:
        bonus_points = int(bonus)
    except (TypeError, ValueError):
        bonus_points = 0
    signals = affinity.get("signals", [])
    if not isinstance(signals, list):
        signals = []
    cleaned_signals = [str(signal).strip() for signal in signals if str(signal).strip()]
    return max(0, min(bonus_points, 4)), cleaned_signals


def job_prefers_specific_domain_experience(jd: str, sector: str) -> bool:
    pattern = DOMAIN_PREFERENCE_PATTERNS.get(sector)
    if not pattern:
        return False
    return bool(re.search(pattern, jd, flags=re.IGNORECASE))


def regional_territory_penalty(title: str, profile: dict[str, Any]) -> tuple[int, str | None]:
    clean_title = normalize(title)
    location = normalize(profile.get("contact", {}).get("location", "")).lower()
    display_location = normalize(profile.get("contact", {}).get("location", "")) or "current location"
    for label, pattern, markers in REGIONAL_TITLE_PATTERNS:
        if not re.search(pattern, clean_title, flags=re.IGNORECASE):
            continue
        if any(marker in location for marker in markers):
            return 0, None
        return REGIONAL_TERRITORY_PENALTY, f"{label} territory may require confirming coverage from {display_location}"
    return 0, None


def domain_gap_concern(sector: str) -> str:
    if sector == "Cybersecurity":
        return "direct cybersecurity or threat-intelligence background is limited"
    return f"direct {sector.lower()} experience appears limited"


def apply_domain_score_guardrail(
    score: int,
    *,
    sector: str,
    domain_preference: bool,
    direct_domain_experience: bool,
    affinity_bonus: int,
) -> int:
    if not domain_preference or direct_domain_experience:
        return score
    cap = 92 if affinity_bonus > 0 else 90
    return min(score, cap)


def evaluate_job(jd: str, profile: dict[str, Any], title: str | None = None, company: str | None = None) -> dict[str, Any]:
    jd = normalize(jd)
    extracted_title, extracted_company = extract_company_and_title(jd, title, company)
    positive_matches = find_matches(jd, POSITIVE_PATTERNS)
    red_flags = find_matches(jd, RED_FLAG_PATTERNS)
    remote_first = re.search(r"\bremote[- ]first\b|\bwork\s*from\s*anywhere\b|\b#WorkFromAnywhere\b|\bfrom home\b.*\banywhere\b", jd, re.I)
    location_ambiguous = False
    if remote_first:
        filtered_flags = []
        for item in red_flags:
            if item["label"] == "onsite or hybrid requirement":
                location_ambiguous = True
                continue
            filtered_flags.append(item)
        red_flags = filtered_flags
    t_score, title_reasons = title_score(extracted_title, jd, profile)
    preliminary_signals = title_reasons + [item["label"] for item in positive_matches]
    preliminary_concerns = [item["label"] for item in red_flags]
    sector = infer_sector_for_job(extracted_company, extracted_title, None, jd, preliminary_signals, preliminary_concerns)
    domain_preference = job_prefers_specific_domain_experience(jd, sector)
    direct_domain_experience = has_direct_professional_domain_experience(profile, sector) if domain_preference else True
    affinity_bonus = 0
    affinity_signals: list[str] = []
    if domain_preference and not direct_domain_experience:
        affinity_bonus, affinity_signals = personal_domain_affinity_bonus(profile, sector)
    territory_penalty, territory_concern = regional_territory_penalty(extracted_title, profile)
    domain_gap_penalty = DOMAIN_GAP_PENALTY if domain_preference and not direct_domain_experience else 0

    positive_score = min(sum(item["weight"] for item in positive_matches), 45)
    red_flag_penalty = min(sum(item["weight"] for item in red_flags), 45)
    comp_delta, comp_reason = comp_adjustment(
        jd,
        profile["target_job_profile"]["comp_floor_base"],
        profile["target_job_profile"]["comp_ceiling_base"],
    )

    base = 30
    raw_score = base + t_score + positive_score + comp_delta + affinity_bonus - red_flag_penalty - domain_gap_penalty - territory_penalty
    score = max(0, min(100, raw_score))
    score = apply_domain_score_guardrail(
        score,
        sector=sector,
        domain_preference=domain_preference,
        direct_domain_experience=direct_domain_experience,
        affinity_bonus=affinity_bonus,
    )

    matched_signals = title_reasons + [item["label"] for item in positive_matches]
    if affinity_bonus > 0 and affinity_signals:
        matched_signals.insert(min(2, len(title_reasons)), f"authentic {sector.lower()} curiosity and motivation")
    if comp_reason and comp_delta > 0:
        matched_signals.append(comp_reason)

    concerns = [item["label"] for item in red_flags]
    if location_ambiguous:
        concerns.append("location wording ambiguous; posting also says remote-first / work from anywhere")
    if comp_reason and comp_delta < 0:
        concerns.append(comp_reason)
    if domain_preference and not direct_domain_experience:
        concerns.append(domain_gap_concern(sector))
    if territory_concern:
        concerns.append(territory_concern)
    concerns.extend(infer_missing_signals(jd))

    fit_band = band_for_score(score)
    outreach = build_outreach(extracted_title, extracted_company, matched_signals, concerns)
    bullets = build_resume_bullets(extracted_title, extracted_company, matched_signals, jd)

    return {
        "title": extracted_title,
        "company": extracted_company,
        "fit_score": score,
        "fit_band": fit_band,
        "matched_signals": unique_keep_order(matched_signals)[:10],
        "concerns": unique_keep_order(concerns)[:8],
        "outreach_message": outreach,
        "resume_bullet_adjustments": bullets,
        "job_description": jd,
    }


def infer_missing_signals(jd: str) -> list[str]:
    concerns: list[str] = []
    if not re.search(r"\bremote\b|\bdistributed\b|\bwork from anywhere\b", jd, re.I):
        concerns.append("remote status not explicit")
    if not re.search(r"\brenewal\b|\bexpansion\b|\bupsell\b|\bretention\b|\bpartnership\b|\brevenue\b|\bdistribution\b|\bcommercial\b|\bbusiness development\b|\bpipeline\b", jd, re.I):
        concerns.append("commercial ownership is not obvious")
    if not re.search(r"\bsalary\b|\bbase\b|\bcompensation\b|\$\b", jd, re.I):
        concerns.append("compensation not listed")
    return concerns


def band_for_score(score: int) -> str:
    if score >= 85:
        return "Excellent"
    if score >= 70:
        return "Strong"
    if score >= 55:
        return "Possible"
    if score >= 40:
        return "Weak"
    return "Pass"


def unique_keep_order(items: list[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for item in items:
        clean = item.strip()
        if clean and clean.lower() not in seen:
            seen.add(clean.lower())
            result.append(clean)
    return result


def voice_phrase(voice: dict[str, Any], index: int, fallback: str) -> str:
    phrases = voice.get("natural_phrases")
    if isinstance(phrases, list) and len(phrases) > index and isinstance(phrases[index], str):
        return phrases[index]
    return fallback


def voice_sample(voice: dict[str, Any], key: str, fallback: str) -> str:
    samples = voice.get("sample_answer_patterns")
    if isinstance(samples, dict):
        value = samples.get(key)
        if isinstance(value, str) and value.strip():
            return value.strip()
    return fallback


def voice_options(voice: dict[str, Any], key: str) -> list[str]:
    value = voice.get(key)
    if isinstance(value, list):
        return [item.strip() for item in value if isinstance(item, str) and item.strip()]
    return []


def voice_section_options(voice: dict[str, Any], section: str, key: str) -> list[str]:
    block = voice.get(section)
    if isinstance(block, dict):
        value = block.get(key)
        if isinstance(value, list):
            return [item.strip() for item in value if isinstance(item, str) and item.strip()]
    return []


def voice_dict_block(voice: dict[str, Any], key: str) -> dict[str, Any]:
    value = voice.get(key)
    return value if isinstance(value, dict) else {}


def stable_seed(*parts: Any) -> int:
    seed = 0
    for part in parts:
        for char in str(part):
            seed += ord(char)
    return seed


def deterministic_choice(options: list[str], seed: int, fallback: str) -> str:
    cleaned = [item.strip() for item in options if item.strip()]
    if cleaned:
        return cleaned[seed % len(cleaned)]
    return fallback


def role_lane_adjustment(voice: dict[str, Any], role_lane: str) -> dict[str, str]:
    adjustments = voice.get("role_lane_personality_adjustments")
    if isinstance(adjustments, dict):
        data = adjustments.get(role_lane)
        if isinstance(data, dict):
            return {k: v for k, v in data.items() if isinstance(v, str)}
    return {}


def role_lane_overlap_sentence(voice: dict[str, Any], role_lane: str) -> str:
    fallback_map = {
        "Strategic Account Management": "The overlap for me here is protecting revenue, growing strategic accounts over time, and handling executive communication without losing trust.",
        "Account Management": "The overlap for me here is relationship ownership, revenue protection, and finding sensible growth paths over time.",
        "Customer Success": "The overlap for me here is retention, executive trust, and making sure customer value turns into durable revenue.",
        "Sales / Enterprise": "The overlap for me here is consultative selling, navigating multi-stakeholder deals, and turning discovery into commercial momentum.",
        "Partnerships / BD": "The overlap for me here is finding the commercial leverage in a relationship, aligning incentives, and helping products reach the right partners or markets.",
        "GTM / Commercial Lead": "The overlap for me here is commercial ownership, market judgment, and cross-functional execution when the path is not fully built yet.",
        "Advisor / Ecosystem": "The overlap for me here is ecosystem judgment, partner strategy, and helping the right relationships turn into real commercial progress.",
    }
    adjustment = role_lane_adjustment(voice, role_lane)
    return adjustment.get("packet_overlap") or fallback_map.get(
        role_lane,
        "The overlap for me here is relationship ownership, commercial judgment, and execution in environments that are not fully polished yet.",
    )


def role_lane_why_anchor(voice: dict[str, Any], role_lane: str, role_focus: str) -> str:
    fallback_map = {
        "Strategic Account Management": "retention, expansion, and long-term stakeholder trust",
        "Account Management": "account stability, commercial follow-through, and trust over time",
        "Customer Success": "customer trust, retention, and commercially useful execution",
        "Sales / Enterprise": "customer discovery, executive conversations, and deal progression that still feels strategic",
        "Partnerships / BD": "ecosystem leverage, market expansion, and commercially meaningful relationships",
        "GTM / Commercial Lead": "GTM judgment, commercial ownership, and messy-market execution",
        "Advisor / Ecosystem": "ecosystem insight, relationship leverage, and strategic positioning",
    }
    adjustment = role_lane_adjustment(voice, role_lane)
    return adjustment.get("why_anchor") or fallback_map.get(role_lane, role_focus)


def technical_positioning_sentence(voice: dict[str, Any], sector: str, seed: int) -> str:
    technical_sectors = {"AI / Data Platform", "Web3 / Crypto", "Robotics / Autonomy", "Defense / Dual-Use"}
    if sector not in technical_sectors:
        return ""
    block = voice.get("technical_product_positioning")
    if isinstance(block, dict):
        variants = block.get("sentence_variants")
        if isinstance(variants, list):
            cleaned = [item.strip() for item in variants if isinstance(item, str) and item.strip()]
            if cleaned:
                return cleaned[seed % len(cleaned)]
    return (
        "I also do well in roles that require working closely with product, data, or research teams "
        "and translating technical value into a clear commercial value proposition."
    )


def format_series(items: list[str]) -> str:
    cleaned = [item.strip() for item in items if item.strip()]
    if not cleaned:
        return ""
    if len(cleaned) == 1:
        return cleaned[0]
    if len(cleaned) == 2:
        return f"{cleaned[0]} and {cleaned[1]}"
    return f"{', '.join(cleaned[:-1])}, and {cleaned[-1]}"


def pick_keywords(text: str, keyword_map: list[tuple[list[str], str]], limit: int) -> list[str]:
    lowered = text.lower()
    picks: list[str] = []
    for keywords, phrase in keyword_map:
        if any(keyword in lowered for keyword in keywords) and phrase not in picks:
            picks.append(phrase)
        if len(picks) >= limit:
            break
    return picks


def normalized_company_key(company: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", (company or "").lower())


def validated_overlap_entries(value: Any) -> list[dict[str, Any]]:
    if not isinstance(value, list):
        return []
    entries: list[dict[str, Any]] = []
    for item in value:
        if not isinstance(item, dict):
            continue
        label = item.get("label")
        if not isinstance(label, str) or not label.strip():
            continue
        story = item.get("story") if isinstance(item.get("story"), str) else ""
        keywords_value = item.get("keywords")
        keywords = []
        if isinstance(keywords_value, list):
            keywords = [keyword.strip().lower() for keyword in keywords_value if isinstance(keyword, str) and keyword.strip()]
        entries.append({"label": label.strip(), "story": story.strip(), "keywords": keywords})
    return entries


def generic_cover_letter_overlap_entries(role_lane: str, sector: str) -> list[dict[str, Any]]:
    recent_account_entry = {
        "label": "recent account management experience across technical and data-heavy clients",
        "story": "RecentAccountRole",
        "keywords": [
            "renewal",
            "retention",
            "expansion",
            "nrr",
            "whitespace",
            "account planning",
            "strategic account",
            "stakeholder",
            "analytics",
            "research",
            "data",
        ],
    }
    enterprise_account_entry = {
        "label": "earlier enterprise account ownership, recurring revenue retention, and contract complexity",
        "story": "EnterpriseAccountRole",
        "keywords": [
            "enterprise",
            "strategic",
            "executive",
            "contract",
            "procurement",
            "institution",
            "institutional",
            "recurring",
            "consultative",
            "customer growth",
        ],
    }
    partnership_entry = {
        "label": "earlier partnerships work across technical ecosystems and commercial motions",
        "story": "PartnershipRole",
        "keywords": [
            "partnership",
            "partner",
            "ecosystem",
            "wallet",
            "exchange",
            "protocol",
            "integration",
            "custody",
            "stablecoin",
            "payments",
            "crypto",
            "web3",
        ],
    }
    technical_entry = {
        "label": "technical-commercial translation around infrastructure, data, and complex products",
        "story": "TechnicalTranslation",
        "keywords": [
            "ai",
            "data",
            "analytics",
            "research",
            "platform",
            "infrastructure",
            "developer",
            "api",
            "technical",
            "autonomy",
            "automation",
        ],
    }
    if sector == "AI / Data Platform":
        return [recent_account_entry, enterprise_account_entry, technical_entry, partnership_entry]
    if sector == "Web3 / Crypto":
        return [recent_account_entry, enterprise_account_entry, partnership_entry, technical_entry]
    if role_lane in {"Strategic Account Management", "Account Management", "Customer Success"}:
        return [recent_account_entry, enterprise_account_entry, technical_entry, partnership_entry]
    if role_lane == "Partnerships / BD":
        return [recent_account_entry, enterprise_account_entry, partnership_entry, technical_entry]
    return [recent_account_entry, enterprise_account_entry, technical_entry, partnership_entry]


def cover_letter_overlap_entries(
    voice: dict[str, Any],
    *,
    company: str,
    role_lane: str,
    sector: str,
    signals: list[str],
    job_description: str,
) -> list[dict[str, Any]]:
    context = " ".join([company, role_lane, sector, job_description, *signals]).lower()
    profiles = voice.get("company_overlap_intelligence")
    company_entries: list[dict[str, Any]] = []
    if isinstance(profiles, dict):
        company_entries = validated_overlap_entries(profiles.get(normalized_company_key(company)))

    selected: list[dict[str, Any]] = []
    seen_labels: set[str] = set()

    def maybe_add(entry: dict[str, Any], *, force: bool = False) -> None:
        label = str(entry.get("label") or "").strip()
        if not label or label.lower() in seen_labels:
            return
        keywords = entry.get("keywords")
        keyword_list = keywords if isinstance(keywords, list) else []
        if force or not keyword_list or any(keyword in context for keyword in keyword_list):
            selected.append(entry)
            seen_labels.add(label.lower())

    for entry in company_entries:
        maybe_add(entry)

    for entry in generic_cover_letter_overlap_entries(role_lane, sector):
        maybe_add(entry)

    if not selected:
        fallback_entries = generic_cover_letter_overlap_entries(role_lane, sector)
        if fallback_entries:
            maybe_add(fallback_entries[0], force=True)

    return selected[:3]


def cover_letter_story_paragraph(story: str) -> str:
    if story == "EnterpriseAccountRole":
        return (
            "Earlier in my career, I owned a large enterprise account portfolio with recurring revenue responsibility across complex customers. "
            "The work included retention, pricing, procurement, and day-to-day relationship management across complex stakeholder groups. "
            "It is still the clearest proof point that I can manage a large book while keeping both revenue and customer relationships on track."
        )
    if story == "PartnershipRole":
        return (
            "Another relevant part of my background is earlier partnerships work across technical ecosystems and cross-functional stakeholders. "
            "I closed multiple partnerships there and spent a lot of time helping a technical product find the right commercial and ecosystem fit. "
            "That experience is especially useful in roles where the product is infrastructure-heavy and relationship quality has a direct impact on adoption."
        )
    return " ".join(
        [
            RECENT_ACCOUNT_OWNERSHIP_SENTENCE,
            RECENT_ACCOUNT_REVENUE_OUTCOMES_SENTENCE,
            RECENT_ACCOUNT_REVENUE_BREAKDOWN_SENTENCE,
            RECENT_ACCOUNT_RECOVERY_SENTENCE,
            RECENT_ACCOUNT_CROSS_FUNCTIONAL_SENTENCE,
        ]
    )


def cover_letter_fit_points(role_lane: str, sector: str, signals: list[str], job_description: str) -> list[str]:
    context = " ".join([role_lane, sector, job_description, *signals]).lower()
    if role_lane == "Strategic Account Management":
        fit_map = [
            (["customer", "account", "relationship"], "managing complex customer relationships over long periods"),
            (["renewal", "retention", "nrr"], "protecting recurring revenue"),
            (["executive", "stakeholder", "strategic"], "navigating executive stakeholders"),
            (["expansion", "growth", "whitespace", "account planning"], "creating growth opportunities within strategic accounts"),
        ]
        fits = pick_keywords(context, fit_map, 4)
        return fits or [
            "managing complex customer relationships over long periods",
            "protecting recurring revenue",
            "navigating executive stakeholders",
            "creating growth opportunities within strategic accounts",
        ]
    if role_lane == "Account Management":
        return pick_keywords(
            context,
            [
                (["mid-market"], "managing a mid-market book"),
                (["retention", "renewal"], "driving retention"),
                (["adoption"], "driving adoption"),
                (["expansion", "upsell", "growth"], "driving expansion"),
                (["account health"], "owning account health"),
                (["executive business review", "business review", "ebr"], "running Executive Business Reviews"),
                (["multi-threaded", "multithreaded"], "building multi-threaded relationships"),
                (["c-level", "key decision makers"], "working with C-level executives and key decision makers"),
                (["launch"], "supporting launch execution"),
                (["cross-functional"], "driving cross-functional execution"),
                (["customer", "account", "relationship"], "owning long-term customer relationships"),
            ],
            4,
        ) or [
            "account management",
            "renewals",
            "expansion",
            "long-term customer ownership",
        ]
    if role_lane == "Partnerships / BD":
        return [
            "building commercially meaningful partner relationships",
            "translating technical products into clear market value",
            "aligning incentives across different stakeholders",
            "turning ecosystem conversations into real traction",
        ]
    if role_lane == "Sales / Enterprise":
        return [
            "running consultative commercial conversations",
            "navigating multi-stakeholder buying processes",
            "translating product value into a clear business case",
            "building momentum inside complex deal cycles",
        ]
    return pick_keywords(
        context,
        [
            (["customer", "relationship"], "managing important relationships over time"),
            (["revenue", "renewal", "retention"], "protecting recurring revenue"),
            (["expansion", "growth"], "creating room for commercial growth"),
            (["executive", "stakeholder"], "working through varied stakeholder environments"),
        ],
        4,
    ) or [
        "managing important relationships over time",
        "protecting recurring revenue",
        "creating room for commercial growth",
    ]


def cover_letter_company_themes(role_lane: str, sector: str, signals: list[str], job_description: str) -> list[str]:
    context = " ".join([role_lane, sector, job_description, *signals]).lower()
    if sector == "Cybersecurity":
        return ["threat intelligence and risk products", "long-term customer growth", "enterprise customer trust"]
    theme_map = [
        (["ai", "data", "analytics", "research"], "AI and data products"),
        (["partnership", "partner"], "enterprise partnerships"),
        (["renewal", "expansion", "growth", "nrr", "account planning"], "long-term customer growth"),
        (["platform", "infrastructure"], "platform products"),
        (["wallet", "protocol", "crypto", "web3"], "commercially important infrastructure"),
    ]
    themes = pick_keywords(context, theme_map, 4)
    if themes:
        return themes
    if sector == "AI / Data Platform":
        return ["AI and data products", "long-term customer growth"]
    if sector == "Web3 / Crypto":
        return ["commercially important infrastructure", "long-term customer growth"]
    return ["long-term customer growth"]


def cover_letter_role_keywords(role_lane: str, job_description: str, signals: list[str]) -> list[str]:
    context = " ".join([role_lane, job_description, *signals]).lower()
    if role_lane == "Account Management":
        return pick_keywords(
            context,
            [
                (["mid-market"], "mid-market customers"),
                (["retention"], "retention"),
                (["adoption"], "adoption"),
                (["expansion", "upsell"], "expansion"),
                (["account health"], "account health"),
                (["c-level"], "C-level executives"),
                (["key decision makers"], "key decision makers"),
                (["multi-threaded", "multithreaded"], "multi-threaded relationships"),
                (["executive business review", "business review", "ebr"], "Executive Business Reviews"),
                (["customer growth"], "customer growth"),
                (["launch"], "launch execution"),
                (["cross-functional"], "cross-functional execution"),
            ],
            6,
        )
    return []


def packet_about_answer(
    voice: dict[str, Any],
    *,
    job_id: int,
    role_lane: str,
    sector: str,
    sector_focus: str,
) -> str:
    opening = deterministic_choice(
        voice_section_options(voice, "answer_openers", "about"),
        job_id,
        "I tend to do best in roles where relationships, strategy, and revenue all overlap.",
    )
    overlap = role_lane_overlap_sentence(voice, role_lane)
    technical = technical_positioning_sentence(voice, sector, job_id + 3)
    base = (
        f"{opening} Most recently, I managed strategic accounts across renewals, expansion, and cross-functional delivery in a technical market. "
        "Before that, I led strategic partnerships in an earlier ecosystem role and built a strong foundation in long-cycle account management. "
        f"{overlap}"
    )
    if technical:
        return f"{base} {technical}"
    return f"{base} I have done that mostly in {sector_focus} environments."


def packet_why_answer(
    voice: dict[str, Any],
    *,
    job_id: int,
    role_lane: str,
    role_focus: str,
    sector_focus: str,
) -> str:
    opening = deterministic_choice(
        voice_section_options(voice, "answer_openers", "why_role"),
        job_id + 5,
        "This role fits because it sits at the intersection of what I have done well: managing strategic relationships, expanding accounts, and turning client context into revenue opportunities.",
    )
    anchor = role_lane_why_anchor(voice, role_lane, role_focus)
    closers = [
        f"The combination of {sector_focus} and {anchor} maps well to the kind of work I have done best.",
        f"I like that it combines {role_focus} with {sector_focus} in a way that still depends on real judgment and follow-through.",
        f"It reads like the kind of role where {anchor} drive the outcome more than surface-level process.",
    ]
    closer = deterministic_choice(closers, job_id + 7, closers[0])
    return f"{opening} {closer}"


def packet_experience_answer(
    voice: dict[str, Any],
    *,
    job_id: int,
    role_lane: str,
    sector: str,
) -> str:
    opening = deterministic_choice(
        voice_section_options(voice, "answer_openers", "experience"),
        job_id + 9,
        "My background is strongest in strategic relationships where trust, commercial judgment, and follow-through matter.",
    )
    role_lane_map = {
        "Strategic Account Management": "That translates well to strategic account work because a lot of it came down to retaining trust, spotting growth opportunities, and keeping complicated relationships on track.",
        "Account Management": "That translates well to account management because most of my work has been around protecting revenue, managing stakeholder expectations, and creating room for sensible growth.",
        "Customer Success": "That translates well to revenue-minded customer success because I am used to balancing retention, expansion, and executive communication.",
        "Sales / Enterprise": "That translates well to enterprise sales because I am comfortable with discovery, multi-stakeholder conversations, and moving technical products through a real commercial process.",
        "Partnerships / BD": "That translates well to partnerships work because I have spent years finding the commercial angle in relationships and getting technical products into the right partner conversations.",
        "GTM / Commercial Lead": "That translates well to GTM roles because I have often been the person connecting customer context, revenue priorities, and cross-functional execution.",
        "Advisor / Ecosystem": "That translates well to ecosystem roles because I have spent a lot of time understanding which relationships are strategic and how to turn them into actual momentum.",
    }
    technical = technical_positioning_sentence(voice, sector, job_id + 11)
    base = (
        f"{opening} In a recent account management role, I worked across renewals, expansion, and cross-functional delivery in a technical market. "
        "In an earlier partnerships role, I closed strategic partnerships across a technical ecosystem. "
        "Earlier in my career, I managed a large recurring-revenue book and worked through complex procurement cycles. "
        f"{role_lane_map.get(role_lane, 'That background has made me comfortable owning important relationships, handling ambiguity, and keeping the commercial side moving.')}"
    )
    if technical:
        return f"{base} {technical}"
    return base


def packet_anything_else_answer(
    voice: dict[str, Any],
    *,
    job_id: int,
    sector: str,
) -> str:
    opening = deterministic_choice(
        voice_section_options(voice, "answer_openers", "anything_else"),
        job_id + 13,
        "I tend to do best in roles where I can explain a technical product clearly, build the right relationships, and turn that into commercial momentum.",
    )
    theme = deterministic_choice(
        voice_options(voice, "strong_commercial_themes"),
        job_id + 15,
        "building trust over time",
    )
    technical = technical_positioning_sentence(voice, sector, job_id + 17)
    closing = "That tends to show up best in roles where trust compounds over time and the work has real revenue consequences."
    if technical:
        closing = f"{technical} That is usually where I can be most useful over time."
    return f"{opening} I tend to be strongest when the work involves {theme}, complicated stakeholders, and clear commercial accountability. {closing}"


def voice_examples_for_job(
    voice: dict[str, Any],
    sector: str,
    role_lane: str,
    *,
    company: str = "",
    product_category: str = "",
    buyer_persona: str = "",
) -> list[tuple[str, str]]:
    examples: list[tuple[str, str]] = []
    account_role_lanes = {"Strategic Account Management", "Account Management", "Customer Success"}
    if role_lane in account_role_lanes:
        challenge = voice_sample(voice, "account_challenge", "")
        skills = voice_sample(voice, "top_account_manager_skills", "")
        project_monitoring = voice_sample(voice, "project_monitoring", "")
        if challenge:
            examples.append(("Account Challenge Example", challenge))
        if skills:
            examples.append(("Top Account-Manager Skills Angle", skills))
        if project_monitoring:
            examples.append(("How You Track Progress and Create Momentum", project_monitoring))
    if sector == "Web3 / Crypto":
        context = " ".join([company, product_category, buyer_persona]).lower()
        if any(token in context for token in ["stablecoin", "usdt", "payments"]):
            why_company = voice_sample(voice, "why_crypto_company", "")
        elif any(token in context for token in ["infrastructure", "staking", "node", "wallet", "custody", "mpc", "api"]):
            why_company = (
                f"Iâ€™m interested in {company} because institutional crypto adoption depends on reliable infrastructure: staking, nodes, APIs, wallet infrastructure, custody-adjacent workflows, and secure access to blockchain networks. "
                "My background includes technical-market relationships, ecosystem stakeholders, and infrastructure-adjacent customer work, so the role feels aligned with the kind of complex customer relationships I know how to manage and grow."
            )
        else:
            why_company = voice_sample(voice, "why_crypto_company", "")
        if why_company:
            examples.append(("Why This Company / Market Angle", why_company))
    return examples


def linkedin_positioning(voice: dict[str, Any]) -> str:
    signal = voice.get("linkedin_about_signal")
    if isinstance(signal, dict):
        text = signal.get("text")
        if isinstance(text, str) and "15+ years" in text:
            return (
                "I've spent 15+ years across partnerships, business development, sales, and account management, "
                "mostly in roles where I need to explain technical products clearly, build trust quickly, and turn that into commercial results."
            )
    return ""


def role_lane_phrase(role_lane: str) -> str:
    mapping = {
        "Strategic Account Management": "strategic account management, renewals, and expansion",
        "Account Management": "account management, relationship ownership, and commercial follow-through",
        "Customer Success": "customer success, retention, and stakeholder management",
        "Partnerships / BD": "partnerships, business development, and strategic relationship management",
        "Sales / Enterprise": "enterprise sales, deal progression, and executive stakeholder work",
        "GTM / Commercial Lead": "commercial ownership, cross-functional execution, and GTM leadership",
        "Advisor / Ecosystem": "ecosystem development, partner strategy, and relationship building",
        "Other": "relationship ownership, commercial judgment, and execution",
    }
    clean = (role_lane or "").strip()
    if clean in mapping:
        return mapping[clean]
    fallback = clean.replace(" / ", ", ").replace("BD", "business development").strip()
    return fallback.lower() if fallback else "relationship ownership and commercial execution"


def sector_phrase(sector: str) -> str:
    mapping = {
        "Web3 / Crypto": "Web3 and crypto",
        "AI / Data Platform": "AI and data infrastructure",
        "Robotics / Autonomy": "robotics and autonomy",
        "Defense / Dual-Use": "defense and dual-use technology",
        "SaaS / General": "SaaS",
        "Gaming / Hardware": "gaming, eSports, and hardware",
    }
    clean = (sector or "").strip()
    if clean in mapping:
        return mapping[clean]
    return clean.lower() if clean else "the market"


def outreach_signal_phrase(signal: str) -> str:
    clean = signal.strip()
    if not clean:
        return ""
    if "title alignment" in clean:
        return ""
    replacements = {
        "strategic partnerships execution": "strategic partnership work",
        "partnerships motion": "partnership-led growth",
        "wallet / exchange / stablecoin infrastructure": "wallet, exchange, and stablecoin infrastructure",
        "payments / financial-infrastructure partner network": "payments and financial-infrastructure partnerships",
        "executive stakeholder work": "executive stakeholder management",
        "senior strategic account ownership": "senior relationship ownership",
        "authentic cybersecurity curiosity and motivation": "a genuine interest in the cybersecurity space",
    }
    if clean in replacements:
        return replacements[clean]
    if clean.endswith(" market fit"):
        return clean.removesuffix(" market fit")
    return clean


def build_outreach(title: str, company: str, matches: list[str], concerns: list[str]) -> str:
    voice = load_voice_profile()
    polished_matches = polish_signals_for_sentence(matches)
    best_matches = select_outreach_signals(polished_matches)
    inferred_lane = infer_role_lane(title, " ".join(matches))
    role_focus = role_lane_phrase(inferred_lane)
    adjustment = role_lane_adjustment(voice, inferred_lane)
    humanized_matches = [item for item in (outreach_signal_phrase(match) for match in best_matches) if item]
    best = ", ".join(humanized_matches[:2]) if humanized_matches else adjustment.get("outreach_focus", role_focus)
    context = " ".join([title, company, *matches]).lower()
    is_web3 = any(token in context for token in ["web3", "crypto", "wallet", "protocol", "blockchain", "stablecoin"])
    if is_web3:
        proof = (
            "In a recent account management role, I managed strategic clients across renewals, expansion, and cross-functional delivery. "
            "Before that, I led partnerships in an earlier technical ecosystem role."
        )
    else:
        proof = (
            "In an earlier enterprise account role, I managed a large recurring-revenue book across complex customers. "
            "More recently, I managed strategic accounts across renewals, expansion, and cross-functional delivery."
        )
    return (
        f"Hey [Name], I saw {company}'s {title} role and wanted to reach out.\n\n"
        f"My background is in {adjustment.get('outreach_focus', role_focus)}, with a strong focus on {best}. "
        f"{proof}\n\n"
        f"The role stood out because it depends on customer ownership, account health, retention, expansion, and technical-commercial translation. "
        f"I would be glad to connect if the team is open to someone with strong account ownership and renewal/expansion experience."
    )


def polish_signals_for_sentence(matches: list[str]) -> list[str]:
    polished: list[str] = []
    for match in matches:
        if match.startswith("dream title signal: "):
            polished.append(f"{match.removeprefix('dream title signal: ')} title alignment")
        elif match.startswith("adjacent title signal: "):
            polished.append(f"{match.removeprefix('adjacent title signal: ')} title alignment")
        elif match == "senior/strategic seniority language":
            polished.append("senior strategic account ownership")
        elif match == "AI/data platform domain":
            polished.append("AI/data platform market fit")
        elif match == "robotics / autonomy domain":
            polished.append("robotics/autonomy market fit")
        elif match == "defense / national security domain":
            polished.append("defense/national-security market fit")
        elif match == "IT security / cybersecurity domain":
            polished.append("cybersecurity market fit")
        elif match.endswith(" domain"):
            polished.append(match.removesuffix(" domain") + " market fit")
        else:
            polished.append(match)
    return unique_keep_order(polished)


def select_outreach_signals(matches: list[str]) -> list[str]:
    title_signals = [item for item in matches if "title alignment" in item or "senior strategic" in item]
    commercial_signals = [
        item
        for item in matches
        if any(term in item.lower() for term in ["renewal", "expansion", "nrr", "partnership", "commercial"])
    ]
    domain_signals = [
        item
        for item in matches
        if any(
            term in item.lower()
            for term in ["ai/", "robotics", "defense", "national security", "healthcare", "fintech", "web3", "gaming", "cybersecurity"]
        )
    ]
    domain_priority = {
        "defense": 0,
        "national security": 0,
        "robotics": 1,
        "autonomy": 1,
        "ai/": 2,
        "cybersecurity": 3,
        "healthcare": 4,
        "fintech": 5,
        "web3": 6,
        "gaming": 7,
    }
    domain_signals.sort(
        key=lambda item: min(
            (rank for term, rank in domain_priority.items() if term in item.lower()),
            default=99,
        )
    )
    selected = unique_keep_order(commercial_signals[:1] + domain_signals[:1])
    if not selected:
        selected = unique_keep_order(title_signals[:1] + commercial_signals[:1] + domain_signals[:1])
    if len(selected) < 3:
        selected = unique_keep_order(selected + matches)
    return selected[:3]


def build_resume_bullets(title: str, company: str, matches: list[str], jd: str) -> list[str]:
    bullets: list[str] = []
    lowered = jd.lower()

    if any(term in lowered for term in ["renewal", "retention", "nrr", "expansion", "upsell"]):
        bullets.append(
            "Emphasize recent strategic-account and revenue-attribution work: managed strategic accounts across a technical product suite, influenced meaningful renewal and expansion outcomes, and helped recover stalled account value through stakeholder coordination."
        )
    if any(term in lowered for term in ["partnership", "partner", "ecosystem", "channel", "integration"]):
        bullets.append(
            "Move earlier partnerships impact higher: highlight ecosystem partnerships, GTM coordination, and onboarding support across technical stakeholders."
        )
    if any(term in lowered for term in ["enterprise", "strategic account", "executive", "stakeholder", "customer success"]):
        bullets.append(
            "Add strategic account language: led senior stakeholder relationships, coordinated cross-functional delivery, and turned account health signals into renewal and expansion plans."
        )
    if any(term in lowered for term in ["health", "healthcare", "procurement", "government"]):
        bullets.append(
            "Surface earlier enterprise-account experience: managed a large recurring-revenue book and negotiated through procurement-heavy customer environments."
        )
    if any(term in lowered for term in ["defense", "national security", "dual-use", "aerospace", "dod", "department of defense", "federal"]):
        bullets.append(
            "Surface defense-adjacent credibility: closed 3 Department of Defense DBPAs totaling $9M over 3 years and held Top Secret clearance for specific DoD-contracted game-development projects."
        )
    if any(term in lowered for term in ["forecast", "pipeline", "salesforce", "crm", "revenue operations", "revops"]):
        bullets.append(
            "Highlight operating cadence: used CRM discipline, pipeline inspection, forecasting, and repeatable account management processes to improve commercial execution."
        )
    if any(term in lowered for term in ["ai", "machine learning", "data platform", "robotics", "automation", "autonomy", "autonomous"]):
        bullets.append(
            "Position technical curiosity and translation strength: explain complex products clearly to partners and customers while tying technical value to adoption, retention, and expansion."
        )

    if not bullets:
        bullets.append(
            f"Tailor the resume summary toward {title} at {company}: strategic account management, partnerships, renewals, expansion pipeline, and executive relationship ownership."
        )

    return bullets[:5]


def extract_ats_keywords(*texts: str) -> list[str]:
    haystack = " ".join(str(text or "") for text in texts).lower()
    return [keyword for keyword in ATS_KEYWORD_CANDIDATES if keyword in haystack]


def ats_support_text(profile: dict[str, Any], signals: list[str], bullets: list[str]) -> str:
    fragments = [profile.get("background_summary", "")]
    for key in ("roles", "earlier_relevant_roles"):
        entries = profile.get(key, [])
        if isinstance(entries, list):
            for entry in entries:
                if not isinstance(entry, dict):
                    continue
                fragments.append(str(entry.get("title") or ""))
                fragments.append(str(entry.get("company") or ""))
                for list_key in ("impact", "skills"):
                    values = entry.get(list_key, [])
                    if isinstance(values, list):
                        fragments.extend(str(value) for value in values)
    fragments.extend(signals)
    fragments.extend(bullets)
    return " ".join(fragment for fragment in fragments if fragment).lower()


def missing_ats_keywords(keywords: list[str], support_text: str) -> list[str]:
    return [keyword for keyword in keywords if keyword.lower() not in support_text]


def application_answer_component(voice: dict[str, Any], key: str, fallback: str = "") -> str:
    components = voice_dict_block(voice, "application_answer_components")
    value = components.get(key)
    if isinstance(value, str) and value.strip():
        return value.strip()
    return fallback


def application_answer_playbook(voice: dict[str, Any], key: str) -> dict[str, Any]:
    playbooks = voice_dict_block(voice, "application_question_playbooks")
    value = playbooks.get(key)
    return value if isinstance(value, dict) else {}


def classify_application_question(question: str) -> str | None:
    clean = normalize(question).lower()
    if not clean:
        return None
    if "account manager" in clean and (" csm" in f" {clean}" or "customer success manager" in clean) and "difference" in clean:
        return "am_vs_csm"
    if "ai" in clean and ("past 3 months" in clean or "learned" in clean or "learning" in clean):
        return "recent_ai_learning"
    if any(term in clean for term in ["ai", "automation", "data insight", "data insights"]) and any(
        term in clean
        for term in ["account", "forecast", "revenue", "expansion", "prioritization", "stakeholder", "renewal"]
    ):
        return "ai_automation"
    if "love for sales" in clean or ("sales" in clean and any(term in clean for term in ["mean to you", "drive results", "used that to drive results"])):
        return "sales_philosophy"
    return None


def draft_sales_philosophy_answer(voice: dict[str, Any]) -> str:
    enterprise_account_scale = application_answer_component(
        voice,
        "enterprise_account_revenue_scale",
        "In an earlier enterprise account role, I managed a large recurring-revenue book across complex customers.",
    )
    enterprise_account_growth = application_answer_component(
        voice,
        "enterprise_account_growth_motion",
        "That work depended on retention, account growth, pricing conversations, and staying close to where revenue was at risk or where growth could develop.",
    )
    enterprise_account_consistency = application_answer_component(
        voice,
        "enterprise_account_consistency",
        "Over time, that showed up in consistent performance against revenue, retention, and account growth goals.",
    )
    recent_account_recent = application_answer_component(
        voice,
        "recent_account_recent_am",
        "In a recent account management role, I managed strategic accounts across a technical product suite.",
    )
    recent_account_metrics = application_answer_component(
        voice,
        "recent_account_recent_am_metrics",
        (
            "I influenced meaningful renewal, recovery, and expansion outcomes across closed, recovered, and "
            "post-transition work. I also helped reactivate a previously stalled account through targeted "
            "stakeholder coordination."
        ),
    )
    trust_compounds = application_answer_component(
        voice,
        "trust_compounds",
        "The part of the work I enjoy most is when trust compounds over time and creates better retention, expansion, and strategic leverage.",
    )
    paragraphs = [
        (
            "To me, loving sales means liking the work of creating value, building trust, and turning relationships into measurable commercial outcomes. "
            "It is not just about the close. It is about understanding what matters to the customer, following through consistently, and keeping momentum alive when budget, timing, or internal alignment gets messy."
        ),
        f"{recent_account_recent} {recent_account_metrics}",
        f"{enterprise_account_scale} {enterprise_account_growth} {enterprise_account_consistency}",
        trust_compounds,
    ]
    return "\n\n".join(paragraphs)


def draft_ai_automation_answer(voice: dict[str, Any]) -> str:
    recent_account_systems = application_answer_component(
        voice,
        "recent_account_systems",
        "In a recent account management role, the workflow was fragmented across multiple systems, so I helped build more structure around renewals, account health, expansion opportunities, and risk tracking.",
    )
    codex_workflow = application_answer_component(
        voice,
        "codex_ai_workflow",
        "More recently, I have been using Codex to build a structured workflow that turns messy information into scored opportunities, flagged risks, and cleaner decision-making.",
    )
    ai_takeaway = application_answer_component(
        voice,
        "ai_judgment_takeaway",
        "My view is that AI is most useful when it improves judgment, prioritization, and operating discipline, not when it replaces commercial ownership.",
    )
    paragraphs = [
        "I have used AI, automation, and structured data mostly to improve visibility, prioritization, and follow-through.",
        (
            f"{recent_account_systems} The goal was to move from scattered account knowledge to a more reliable operating system for renewals, forecasting, and expansion."
        ),
        (
            f"{codex_workflow} That is directly relevant to account management because the same approach can help summarize calls, flag renewal timing, track stakeholder coverage, surface account risk, and identify realistic expansion opportunities earlier."
        ),
        (
            f"For forecasting, I think AI works best when it organizes signals like account health, product usage, stakeholder engagement, sentiment, and pipeline stage while the AM still pressure-tests what is real. {ai_takeaway}"
        ),
    ]
    return "\n\n".join(paragraphs)


def draft_recent_ai_learning_answer(voice: dict[str, Any]) -> str:
    recent_account_notion_story = application_answer_component(
        voice,
        "recent_account_notion_story",
        "I used AI to help structure a renewal and account-management tracker in Notion because it was lightweight, shareable, and realistic for the broader team to adopt quickly.",
    )
    codex_workflow = application_answer_component(
        voice,
        "codex_ai_workflow",
        "More recently, I have been using Codex to build a structured workflow that turns messy information into scored opportunities, flagged risks, and cleaner decision-making.",
    )
    paragraphs = [
        "One of the biggest things I have learned recently about AI is that it becomes much more valuable when it helps turn scattered information into a repeatable operating system.",
        (
            f"In a recent account management role, the workflow was spread across multiple systems, and not every stakeholder had access to the same tools. {recent_account_notion_story} That let me centralize renewal timing, account health, expansion opportunities, risk, next steps, and leadership visibility without waiting for a heavier CRM rebuild."
        ),
        (
            "That experience taught me to think about AI as workflow design, not just faster writing. It can help decide what information matters, reduce ambiguity, and make renewal readiness, forecasting, follow-through, and account planning more consistent."
        ),
        (
            f"{codex_workflow} That has reinforced the same lesson for me: tools like markdown instruction files, reusable context, structured outputs, connectors, and skills matter because AI gets better when the operating workflow is clear. In a Strategic Account Manager role, that should show up as better account visibility, earlier risk detection, and cleaner expansion planning."
        ),
    ]
    return "\n\n".join(paragraphs)


def draft_am_vs_csm_answer() -> str:
    paragraphs = [
        "An Account Manager and a CSM both own the customer relationship, but they are usually measured against different primary outcomes.",
        (
            "A CSM is typically focused on adoption, onboarding, usage, enablement, customer health, and making sure the customer gets value from the product. "
            "A strong CSM reduces friction, improves adoption, and keeps the customer progressing toward the outcomes they bought the product for."
        ),
        (
            "An Account Manager is more directly responsible for the commercial relationship. That includes renewals, expansion, upsell, cross-sell, pricing conversations, procurement, stakeholder mapping, executive alignment, risk mitigation, and forecastable revenue. "
            "A strong AM still needs customer empathy, but the job is to translate customer value into retention and growth."
        ),
        (
            "There is real overlap between the two roles, and in a strong GTM motion they should work closely together. But if there is renewal risk, weak executive alignment, stalled expansion, procurement friction, or unclear budget, the AM is accountable for creating a path forward. "
            "That is the side of the line that fits me best because I want to own the commercial thread, not just the adoption thread."
        ),
    ]
    return "\n\n".join(paragraphs)


def draft_application_answer(
    voice: dict[str, Any],
    *,
    question: str,
) -> tuple[str, dict[str, Any]]:
    category = classify_application_question(question)
    if category == "sales_philosophy":
        return draft_sales_philosophy_answer(voice), application_answer_playbook(voice, category)
    if category == "ai_automation":
        return draft_ai_automation_answer(voice), application_answer_playbook(voice, category)
    if category == "recent_ai_learning":
        return draft_recent_ai_learning_answer(voice), application_answer_playbook(voice, category)
    if category == "am_vs_csm":
        return draft_am_vs_csm_answer(), application_answer_playbook(voice, category)
    return "", {}


def enrich_application_questions_for_packet(
    questions: list[dict[str, Any]],
    *,
    voice: dict[str, Any],
) -> list[dict[str, Any]]:
    enriched_questions: list[dict[str, Any]] = []
    generic_ats_strategy = (
        "Match the wording of the exact prompt, use only defensible experience, and keep the answer tied to the role's actual commercial requirements."
    )
    generic_notes = (
        "If this needs a manual pass, keep it direct, proof-backed, and limited to what the question actually asks."
    )
    for item in questions:
        if not isinstance(item, dict):
            continue
        enriched = dict(item)
        question = str(enriched.get("question") or "").strip()
        if not question:
            continue
        drafted_answer, playbook = draft_application_answer(voice, question=question)
        existing_answer = str(enriched.get("recommended_written_answer") or "").strip()
        if drafted_answer and not existing_answer:
            enriched["recommended_written_answer"] = drafted_answer
        current_answer = str(enriched.get("recommended_written_answer") or "").strip()
        why_it_works = str(enriched.get("why_this_answer_works") or enriched.get("reasoning") or "").strip()
        if not why_it_works:
            why_it_works = str(playbook.get("why_this_works") or "").strip()
        if current_answer and not why_it_works:
            why_it_works = "The answer stays close to the exact prompt, uses real proof, and keeps the framing commercially relevant."
        if why_it_works:
            enriched["why_this_answer_works"] = why_it_works
        ats_strategy = str(enriched.get("ats_strategy") or "").strip()
        if not ats_strategy:
            ats_strategy = str(playbook.get("ats_strategy") or "").strip() or generic_ats_strategy
        enriched["ats_strategy"] = ats_strategy
        recruiter_screen_risk = str(enriched.get("recruiter_screen_risk") or "").strip()
        if not recruiter_screen_risk:
            recruiter_screen_risk = str(playbook.get("recruiter_screen_risk") or "").strip()
        if recruiter_screen_risk:
            enriched["recruiter_screen_risk"] = recruiter_screen_risk
        notes_for_candidate = str(enriched.get("notes_for_candidate") or "").strip()
        if not notes_for_candidate:
            notes_for_candidate = str(playbook.get("notes_for_candidate") or "").strip() or generic_notes
        enriched["notes_for_candidate"] = notes_for_candidate
        recommended_selection = str(enriched.get("recommended_selection") or "").strip()
        if current_answer and recommended_selection in {"", "Needs Codex draft", "Needs candidate input"}:
            enriched["recommended_selection"] = "Use recommended answer below"
        enriched_questions.append(enriched)
    return enriched_questions


def render_application_questions_md(questions: list[dict[str, Any]]) -> str:
    if not questions:
        return """## 6. Application Questions + Tailored Answers

No application questions have been captured for this role yet.

- Add the exact ATS/application questions before using this section.
- Do not generate generic answers for unasked questions.
"""
    sections = ["## 6. Application Questions + Tailored Answers"]
    for index, item in enumerate(questions, start=1):
        question = str(item.get("question") or "").strip()
        answer_type = str(item.get("answer_type") or "unknown").strip()
        risk_level = str(item.get("risk_level") or "").strip() or "unknown"
        fit_gap = "true" if bool(item.get("fit_gap")) else "false"
        likely_knockout = "true" if bool(item.get("likely_knockout")) else "false"
        recommended_selection = str(item.get("recommended_selection") or "").strip() or "Needs candidate input"
        recommended_written_answer = str(item.get("recommended_written_answer") or "").strip() or "No written answer captured."
        ats_strategy = str(item.get("ats_strategy") or "").strip() or "No ATS strategy note captured."
        recruiter_screen_risk = str(item.get("recruiter_screen_risk") or "").strip() or "No recruiter-screen risk note captured."
        why_this_answer_works = (
            str(item.get("why_this_answer_works") or item.get("reasoning") or "").strip() or "No explanation captured."
        )
        notes_for_candidate = str(item.get("notes_for_candidate") or "").strip() or "No extra notes."
        sections.append(
            f"""## Question {index}

### Exact Application Question
> {question}

**Answer Type:** {answer_type}
**Risk Level:** {risk_level}
**Fit Gap:** {fit_gap}
**Likely Knockout:** {likely_knockout}
**Recommended Selection:** {recommended_selection}

### Recommended Answer
{recommended_written_answer}

### Why This Answer Works
{why_this_answer_works}

### ATS Strategy
{ats_strategy}

### Recruiter Screen Risk
{recruiter_screen_risk}

### Notes for the Candidate
{notes_for_candidate}
"""
        )
    return "\n\n".join(sections)


def application_questions_status_payload(override: dict[str, Any]) -> dict[str, str]:
    questions = override.get("questions", []) if isinstance(override.get("questions"), list) else []
    status = str(override.get("capture_status") or "").strip()
    reason = str(override.get("capture_reason") or "").strip()
    next_action = str(override.get("capture_next_action") or "").strip()
    if status == "Captured from Screenshot":
        body = (
            "## 2. Application Questions Status\n\n"
            "Status: Captured from Screenshot\n\n"
            "Notes:\n"
            "- Questions were extracted from the visible screenshot.\n"
            "- Only visible questions were answered.\n"
            "- If additional questions appear later, paste another screenshot or copy the questions directly.\n\n"
            f"Reason:\n- {reason or 'Questions were captured from a screenshot provided by the candidate.'}\n\n"
            f"Next Action:\n- {next_action or 'None'}\n"
        )
        return {"status": "Captured from Screenshot", "body": body}
    if status == "Partially Captured":
        return {
            "status": "Partially Captured",
            "body": (
                "## 2. Application Questions Status\n\n"
                "Status: Partially Captured\n\n"
                "Some application questions were captured, but the form may include additional hidden or later-step questions.\n\n"
                "Generated answers only for captured questions.\n\n"
                "Next Action:\n"
                "If additional questions appear later in the application, paste them here and regenerate only those answers.\n"
            ),
        }
    if status == "Captured" or questions:
        body = (
            "## 2. Application Questions Status\n\n"
            "Status: Captured\n\n"
            f"Reason:\n- {reason or 'Questions were captured and stored for this job.'}\n\n"
            f"Next Action:\n- {next_action or 'None'}\n"
        )
        return {"status": "Captured", "body": body}
    return {
        "status": "Not Captured",
        "body": (
            "## 2. Application Questions Status\n\n"
            "Status: Not Captured\n\n"
            "The public job description was accessible, but the actual application form questions were not captured. They may require clicking Apply, loading JavaScript, passing captcha, logging in, or manually starting the application.\n\n"
            "No application answers were generated because no exact questions were available.\n\n"
            "Next Action:\n"
            "Paste the application questions here and regenerate only the \"Application Questions + Tailored Answers\" section.\n"
        ),
    }


def guess_application_answer_type(question: str) -> str:
    clean = normalize(question).lower()
    if any(term in clean for term in ["authorized to work", "work authorization", "require sponsorship", "visa sponsorship", "sponsorship now or in the future"]):
        return "work_auth"
    if any(term in clean for term in ["compensation", "salary", "ote", "base pay", "base salary", "desired pay"]):
        return "compensation"
    if any(term in clean for term in ["location", "relocate", "relocation", "territory", "remote", "hybrid", "onsite", "on-site"]):
        return "location"
    if clean.startswith(("do you", "are you", "have you", "will you", "can you", "did you")):
        return "yes_no"
    if any(term in clean for term in ["please describe", "please explain", "tell us", "why ", "how ", "list ", "walk us through"]):
        return "long_text"
    return "short_text"


def guess_application_question_risk(question: str) -> tuple[str, bool, bool]:
    clean = normalize(question).lower()
    fit_gap = any(
        term in clean
        for term in [
            "$1m",
            "1m+",
            "quota",
            "cybersecurity",
            "data security",
            "threat intelligence",
            "certification",
            "people manager",
            "years of experience",
            "must be based",
            "territory",
        ]
    )
    likely_knockout = any(
        term in clean
        for term in [
            "$1m",
            "1m+",
            "quota",
            "authorized to work",
            "require sponsorship",
            "must be based",
            "located in",
            "cybersecurity",
            "data security",
            "threat intelligence",
            "certification",
        ]
    )
    risk_level = "high" if likely_knockout or fit_gap else "medium" if len(clean.split()) > 14 else "low"
    return risk_level, fit_gap, likely_knockout


def parse_application_questions_text(text: str) -> list[dict[str, Any]]:
    raw = text.replace("\r\n", "\n").strip()
    if not raw:
        return []
    questions: list[str] = []
    blocks = [block.strip() for block in re.split(r"\n\s*\n", raw) if block.strip()]
    if len(blocks) > 1:
        questions = blocks
    else:
        for line in raw.splitlines():
            clean = line.strip().lstrip("-*0123456789. ").strip()
            if clean:
                questions.append(clean)
    parsed: list[dict[str, Any]] = []
    for question in questions:
        answer_type = guess_application_answer_type(question)
        risk_level, fit_gap, likely_knockout = guess_application_question_risk(question)
        parsed.append(
            {
                "question": question,
                "answer_type": answer_type,
                "risk_level": risk_level,
                "fit_gap": fit_gap,
                "likely_knockout": likely_knockout,
                "recommended_selection": "Needs Codex review" if answer_type == "yes_no" else "Needs Codex draft",
                "recommended_written_answer": "",
                "why_this_answer_works": "",
                "ats_strategy": "",
                "recruiter_screen_risk": "",
                "reasoning": "",
                "notes_for_candidate": "",
            }
        )
    return parsed


def questions_from_input(text: str) -> list[dict[str, Any]]:
    stripped = text.strip()
    if not stripped:
        return []
    try:
        parsed = json.loads(stripped)
    except json.JSONDecodeError:
        return parse_application_questions_text(stripped)
    if isinstance(parsed, dict):
        questions = parsed.get("questions")
        if isinstance(questions, list):
            return [item for item in questions if isinstance(item, dict) and str(item.get("question") or "").strip()]
    if isinstance(parsed, list):
        normalized_questions: list[dict[str, Any]] = []
        for item in parsed:
            if isinstance(item, dict) and str(item.get("question") or "").strip():
                normalized_questions.append(item)
            elif isinstance(item, str) and item.strip():
                normalized_questions.extend(parse_application_questions_text(item))
        return normalized_questions
    return parse_application_questions_text(stripped)


def capture_application_questions(
    db_path: Path,
    *,
    job_id: int,
    text: str | None,
    clear_existing: bool = False,
    capture_status: str | None = None,
    capture_reason: str | None = None,
    capture_next_action: str | None = None,
    path: Path = APPLICATION_QUESTION_OVERRIDES_JSON,
) -> int:
    ensure_schema(db_path)
    overrides = load_application_question_overrides(path)
    key = str(job_id)
    existing = overrides.get(key, {}) if isinstance(overrides.get(key), dict) else {}
    questions = questions_from_input(text or "")
    if not questions and capture_status != "Not Captured":
        raise SystemExit("No application questions were detected. Provide plaintext questions or a JSON list/object.")
    if capture_status == "Not Captured":
        questions = []
    if clear_existing:
        existing["questions"] = questions
    else:
        merged = existing.get("questions", []) if isinstance(existing.get("questions"), list) else []
        merged.extend(questions)
        deduped: list[dict[str, Any]] = []
        seen: set[str] = set()
        for item in merged:
            if not isinstance(item, dict):
                continue
            question = normalize(str(item.get("question") or ""))
            if not question or question.lower() in seen:
                continue
            seen.add(question.lower())
            deduped.append(item)
        existing["questions"] = deduped
    if capture_status:
        existing["capture_status"] = capture_status
    elif questions:
        existing["capture_status"] = "Captured"
    if capture_reason:
        existing["capture_reason"] = capture_reason
    elif questions and not str(existing.get("capture_reason") or "").strip():
        existing["capture_reason"] = "Questions were explicitly captured from the application flow or provided directly by the candidate."
    if capture_next_action:
        existing["capture_next_action"] = capture_next_action
    elif questions and not str(existing.get("capture_next_action") or "").strip():
        existing["capture_next_action"] = "None"
    overrides[key] = existing
    save_application_question_overrides(overrides, path)
    with sqlite3.connect(db_path) as conn:
        require_job(conn, job_id)
        conn.execute("UPDATE job_evaluations SET last_updated_at = ? WHERE id = ?", (now_utc(), job_id))
    stored = overrides.get(key, {}) if isinstance(overrides.get(key), dict) else {}
    return len(stored.get("questions", []))


def ats_risk_assessment_md(
    *,
    profile: dict[str, Any],
    job_description: str,
    signals: list[str],
    bullets: list[str],
    override: dict[str, Any],
) -> str:
    questions = override.get("questions", []) if isinstance(override.get("questions"), list) else []
    question_text = "\n".join(str(item.get("question") or "") for item in questions if isinstance(item, dict))
    keywords = extract_ats_keywords(job_description, question_text)
    support_text = ats_support_text(profile, signals, bullets)
    missing_keywords = missing_ats_keywords(keywords, support_text)
    likely_knockouts = [str(item.get("question") or "").strip() for item in questions if isinstance(item, dict) and item.get("likely_knockout")]
    risky_yes = [
        str(item.get("question") or "").strip()
        for item in questions
        if isinstance(item, dict)
        and str(item.get("recommended_selection") or "").strip().lower().startswith("yes")
        and bool(item.get("fit_gap"))
    ]
    positioning = str(override.get("recommended_positioning") or "").strip()
    if not positioning:
        positioning = "Lead with the strongest truthful bridge between the candidate's renewal and expansion proof, technical-market experience, and the exact buyer or product language in the job."
    lines = ["## 3. ATS Risk Assessment"]
    lines.append("- Likely knockout questions: " + ("; ".join(likely_knockouts) if likely_knockouts else "None captured yet."))
    lines.append("- Required keywords: " + (", ".join(keywords[:15]) if keywords else "No high-signal ATS keywords captured yet."))
    lines.append("- Missing keywords: " + (", ".join(missing_keywords[:8]) if missing_keywords else "No major unsupported keywords jumped out from the captured questions and JD."))
    lines.append(f"- Recommended positioning: {positioning}")
    lines.append("- Any risky questions that need \"Yes with context\": " + ("; ".join(risky_yes) if risky_yes else "None captured yet."))
    return "\n".join(lines)


def ensure_schema(db_path: Path) -> None:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS job_evaluations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                source TEXT,
                title TEXT,
                company TEXT,
                fit_score INTEGER NOT NULL,
                fit_band TEXT NOT NULL,
                matched_signals TEXT NOT NULL,
                concerns TEXT NOT NULL,
                outreach_message TEXT NOT NULL,
                resume_bullet_adjustments TEXT NOT NULL,
                job_description TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'discovered',
                decision TEXT,
                source_board TEXT,
                application_url TEXT,
                applied_at TEXT,
                resume_version TEXT,
                follow_up_date TEXT,
                next_action TEXT,
                archived_reason TEXT,
                last_updated_at TEXT
            )
            """
        )
        ensure_job_evaluation_columns(conn)
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS applications (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                job_id INTEGER NOT NULL,
                created_at TEXT NOT NULL,
                applied_at TEXT NOT NULL,
                application_url TEXT,
                resume_version TEXT,
                cover_note_used TEXT,
                outreach_message TEXT,
                application_status TEXT NOT NULL DEFAULT 'applied',
                follow_up_date TEXT,
                notes TEXT,
                FOREIGN KEY (job_id) REFERENCES job_evaluations(id)
            )
            """
        )
        ensure_application_columns(conn)
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS target_companies (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                company TEXT NOT NULL,
                website TEXT,
                lane TEXT,
                description TEXT,
                funding_date TEXT,
                funding_amount TEXT,
                round TEXT,
                investors TEXT,
                company_fit_score INTEGER,
                open_roles_found TEXT,
                best_role_title TEXT,
                role_fit_score INTEGER,
                role_url TEXT,
                priority TEXT,
                target_strategy TEXT,
                outreach_type TEXT,
                warm_contact_1 TEXT,
                warm_contact_1_title TEXT,
                warm_contact_1_linkedin TEXT,
                warm_contact_2 TEXT,
                warm_contact_2_title TEXT,
                warm_contact_2_linkedin TEXT,
                outreach_angle TEXT,
                outreach_status TEXT,
                application_status TEXT,
                notes TEXT,
                next_action TEXT,
                last_checked TEXT,
                source_url TEXT,
                last_updated_at TEXT
            )
            """
        )
        ensure_target_company_columns(conn)
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS contacts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                job_id INTEGER,
                created_at TEXT NOT NULL,
                name TEXT NOT NULL,
                role TEXT,
                email TEXT,
                linkedin_url TEXT,
                telegram_handle TEXT,
                relationship TEXT,
                notes TEXT,
                FOREIGN KEY (job_id) REFERENCES job_evaluations(id)
            )
            """
        )
        ensure_contacts_columns(conn)
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS correspondence (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                job_id INTEGER NOT NULL,
                contact_id INTEGER,
                created_at TEXT NOT NULL,
                date TEXT NOT NULL,
                channel TEXT NOT NULL,
                direction TEXT NOT NULL,
                type TEXT NOT NULL,
                summary TEXT NOT NULL,
                follow_up_needed INTEGER NOT NULL DEFAULT 0,
                follow_up_date TEXT,
                external_thread_id TEXT,
                external_message_id TEXT,
                FOREIGN KEY (job_id) REFERENCES job_evaluations(id),
                FOREIGN KEY (contact_id) REFERENCES contacts(id)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS notes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                job_id INTEGER NOT NULL,
                created_at TEXT NOT NULL,
                note TEXT NOT NULL,
                FOREIGN KEY (job_id) REFERENCES job_evaluations(id)
            )
            """
        )
        conn.execute("CREATE INDEX IF NOT EXISTS idx_job_status ON job_evaluations(status)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_job_follow_up ON job_evaluations(follow_up_date)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_applications_job ON applications(job_id)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_contacts_job ON contacts(job_id)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_correspondence_job ON correspondence(job_id)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_correspondence_follow_up ON correspondence(follow_up_date)")


def ensure_job_evaluation_columns(conn: sqlite3.Connection) -> None:
    existing = {row[1] for row in conn.execute("PRAGMA table_info(job_evaluations)").fetchall()}
    columns = {
        "status": "TEXT NOT NULL DEFAULT 'discovered'",
        "decision": "TEXT",
        "source_board": "TEXT",
        "application_url": "TEXT",
        "applied_at": "TEXT",
        "resume_version": "TEXT",
        "follow_up_date": "TEXT",
        "next_action": "TEXT",
        "archived_reason": "TEXT",
        "last_updated_at": "TEXT",
        "sector": "TEXT",
        "role_lane": "TEXT",
        "priority": "TEXT",
        "cover_letter_needed": "TEXT",
        "referral_target": "TEXT",
    }
    for name, definition in columns.items():
        if name not in existing:
            conn.execute(f"ALTER TABLE job_evaluations ADD COLUMN {name} {definition}")
    conn.execute("UPDATE job_evaluations SET status = 'discovered' WHERE status IS NULL OR status = ''")
    conn.execute("UPDATE job_evaluations SET last_updated_at = created_at WHERE last_updated_at IS NULL")


def ensure_application_columns(conn: sqlite3.Connection) -> None:
    existing = {row[1] for row in conn.execute("PRAGMA table_info(applications)").fetchall()}
    columns = {
        "questionnaire_completed": "INTEGER NOT NULL DEFAULT 0",
        "video_submitted": "INTEGER NOT NULL DEFAULT 0",
        "submission_summary": "TEXT",
        "response_archive_path": "TEXT",
    }
    for name, definition in columns.items():
        if name not in existing:
            conn.execute(f"ALTER TABLE applications ADD COLUMN {name} {definition}")


def ensure_target_company_columns(conn: sqlite3.Connection) -> None:
    existing = {row[1] for row in conn.execute("PRAGMA table_info(target_companies)").fetchall()}
    columns = {
        "website": "TEXT",
        "lane": "TEXT",
        "description": "TEXT",
        "funding_date": "TEXT",
        "funding_amount": "TEXT",
        "round": "TEXT",
        "investors": "TEXT",
        "company_fit_score": "INTEGER",
        "open_roles_found": "TEXT",
        "best_role_title": "TEXT",
        "role_fit_score": "INTEGER",
        "role_url": "TEXT",
        "priority": "TEXT",
        "target_strategy": "TEXT",
        "outreach_type": "TEXT",
        "warm_contact_1": "TEXT",
        "warm_contact_1_title": "TEXT",
        "warm_contact_1_linkedin": "TEXT",
        "warm_contact_2": "TEXT",
        "warm_contact_2_title": "TEXT",
        "warm_contact_2_linkedin": "TEXT",
        "outreach_angle": "TEXT",
        "outreach_status": "TEXT",
        "application_status": "TEXT",
        "notes": "TEXT",
        "next_action": "TEXT",
        "last_checked": "TEXT",
        "source_url": "TEXT",
        "last_updated_at": "TEXT",
    }
    for name, definition in columns.items():
        if name not in existing:
            conn.execute(f"ALTER TABLE target_companies ADD COLUMN {name} {definition}")


def ensure_contacts_columns(conn: sqlite3.Connection) -> None:
    existing = {row[1] for row in conn.execute("PRAGMA table_info(contacts)").fetchall()}
    columns = {
        "telegram_handle": "TEXT",
    }
    for name, definition in columns.items():
        if name not in existing:
            conn.execute(f"ALTER TABLE contacts ADD COLUMN {name} {definition}")


def save_result(db_path: Path, result: dict[str, Any], source: str | None) -> int:
    ensure_schema(db_path)
    created_at = now_utc()
    metadata = dashboard_metadata_for_result(result, source)
    with sqlite3.connect(db_path) as conn:
        cursor = conn.execute(
            """
            INSERT INTO job_evaluations (
                created_at, source, title, company, fit_score, fit_band,
                matched_signals, concerns, outreach_message,
                resume_bullet_adjustments, job_description, status,
                source_board, last_updated_at, sector, role_lane, priority,
                cover_letter_needed
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                created_at,
                source,
                result["title"],
                result["company"],
                result["fit_score"],
                result["fit_band"],
                json.dumps(result["matched_signals"]),
                json.dumps(result["concerns"]),
                result["outreach_message"],
                json.dumps(result["resume_bullet_adjustments"]),
                result["job_description"],
                "discovered",
                infer_source_board(source),
                created_at,
                metadata["sector"],
                metadata["role_lane"],
                metadata["priority"],
                metadata["cover_letter_needed"],
            ),
        )
        return int(cursor.lastrowid)


def potential_duplicate_jobs(
    db_path: Path,
    *,
    title: str | None,
    company: str | None,
    source: str | None,
) -> list[dict[str, Any]]:
    ensure_schema(db_path)
    clauses: list[str] = []
    params: list[Any] = []
    generic_values = {"unknown title", "unknown company"}
    if title and company and title.strip().lower() not in generic_values and company.strip().lower() not in generic_values:
        clauses.append("(lower(trim(title)) = lower(trim(?)) AND lower(trim(company)) = lower(trim(?)))")
        params.extend([title, company])
    if source and re.match(r"https?://", source.strip(), flags=re.IGNORECASE):
        clauses.append("trim(source) = trim(?)")
        params.append(source.strip())
    if not clauses:
        return []
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            f"""
            SELECT id, company, title, fit_score, status, source
            FROM job_evaluations
            WHERE {' OR '.join(clauses)}
            ORDER BY id DESC
            LIMIT 5
            """,
            params,
        ).fetchall()
    return [dict(row) for row in rows]


def duplicate_jobs_message(rows: list[dict[str, Any]]) -> str:
    lines = ["Potential duplicate job already exists:"]
    for row in rows:
        lines.append(
            f"- #{row['id']} {row['company']} - {row['title']} "
            f"({row['fit_score']}, {row['status']})"
        )
    lines.append("Rerun with --allow-duplicate if this is intentionally a separate posting.")
    return "\n".join(lines)


def first_match(pattern: str, text: str, default: str = "") -> str:
    match = re.search(pattern, text, flags=re.IGNORECASE | re.MULTILINE)
    return normalize(match.group(1)) if match else default


def yamlish_value(text: str, key: str) -> str:
    return first_match(rf"^{re.escape(key)}:\s*(.+?)\s*$", text)


def yamlish_list(text: str, key: str) -> list[str]:
    match = re.search(
        rf"^{re.escape(key)}:\s*\n((?:\s+- .+\n?)+)",
        text,
        flags=re.IGNORECASE | re.MULTILINE,
    )
    if not match:
        return []
    return [normalize(line.strip()[2:]) for line in match.group(1).splitlines() if line.strip().startswith("- ")]


def clean_pasted_chatgpt_text(text: str) -> str:
    clean = text.encode("utf-8", errors="replace").decode("utf-8")
    replacements = {
        "\u2018": "'",
        "\u2019": "'",
        "\u201c": '"',
        "\u201d": '"',
        "\u2013": "-",
        "\u2014": "-",
        "\u2026": "...",
        "\u00e2\u20ac\u0153": '"',
        "\u00e2\u20ac\u009d": '"',
        "\u00e2\u20ac?": '"',
        "\u00e2\u20ac\u02dc": "'",
        "\u00e2\u20ac\u2122": "'",
        "\u00e2\u20ac\u201c": "-",
        "\u00e2\u20ac\u201d": "-",
        "\u00e2\u20ac\u00a6": "...",
        "Ñ‚ÐÐ¨": "'",
        "Ñ‚ÐÐ©": "'",
        "Ñ‚ÐÐ¬": '"',
        "Ñ‚ÐÐ­": '"',
        "Ñ‚ÐÐ£": "-",
        "Ñ‚ÐÐ¤": "-",
        "Ñ‚ÐÐ¶": "...",
        "Ñ‚Ð?": '"',
    }
    for bad, good in replacements.items():
        clean = clean.replace(bad, good)
    return clean


def section_between(text: str, start_pattern: str, end_patterns: list[str]) -> str:
    start = re.search(start_pattern, text, flags=re.IGNORECASE | re.MULTILINE)
    if not start:
        return ""
    section_start = start.end()
    section_end = len(text)
    for pattern in end_patterns:
        end = re.search(pattern, text[section_start:], flags=re.IGNORECASE | re.MULTILINE)
        if end:
            section_end = min(section_end, section_start + end.start())
    return text[section_start:section_end].strip()


def parse_chatgpt_fit_score(text: str) -> int:
    tracker_score = yamlish_value(text, "fit_score")
    if tracker_score:
        try:
            return max(0, min(100, int(float(tracker_score))))
        except ValueError:
            pass
    match = re.search(r"Fit Score:\s*([0-9]+(?:\.[0-9]+)?)\s*/\s*10", text, flags=re.IGNORECASE)
    if match:
        return max(0, min(100, round(float(match.group(1)) * 10)))
    match = re.search(r"Fit Score:\s*([0-9]+)\s*/\s*100", text, flags=re.IGNORECASE)
    if match:
        return max(0, min(100, int(match.group(1))))
    return 70


def priority_from_chatgpt(priority_text: str, fit_score: int) -> str:
    lowered = priority_text.lower()
    if fit_score >= 92 or lowered in {"high", "very high", "highest"}:
        return "P1 Apply Today"
    if fit_score >= 82 or "high" in lowered:
        return "P2 Strong"
    if fit_score >= 70 or "medium" in lowered:
        return "P3 Maybe"
    return "Park"


def role_lane_from_chatgpt(title: str, role_category: str) -> str:
    text = f"{title} {role_category}"
    if re.search(r"customer success|csm|customer growth", text, flags=re.IGNORECASE):
        return "Customer Success"
    if re.search(r"strategic account|account management|account manager|protocol relationships", text, flags=re.IGNORECASE):
        return "Account Management"
    if re.search(r"partnership|business development|\bBD\b", text, flags=re.IGNORECASE):
        return "Partnerships / BD"
    return infer_role_lane(title, role_category)


def sector_from_chatgpt(role_category: str, signals: list[str], full_text: str) -> str:
    text = " ".join([role_category, " ".join(signals), full_text])
    if re.search(r"web3|crypto|protocol|wallet|blockchain|chain|interoperability", text, flags=re.IGNORECASE):
        return "Web3 / Crypto"
    return infer_sector(text)


def parse_chatgpt_job_analysis(text: str) -> dict[str, Any]:
    clean = clean_pasted_chatgpt_text(text).strip()
    if not clean:
        raise SystemExit("Provide ChatGPT job analysis with --text, --file, or stdin.")
    header = re.search(r"^\s*([^,\n]+),\s*(.+?)\s*$", clean, flags=re.MULTILINE)
    if header:
        company = normalize(header.group(1))
        title = normalize(header.group(2))
    else:
        company = yamlish_value(clean, "company") or "Unknown Company"
        title = yamlish_value(clean, "role") or yamlish_value(clean, "title") or "Unknown Title"
    fit_score = parse_chatgpt_fit_score(clean)
    priority_text = yamlish_value(clean, "priority") or first_match(r"^Priority:\s*(.+?)\s*$", clean)
    role_category = yamlish_value(clean, "role_category")
    url = yamlish_value(clean, "url")
    recommendation = yamlish_value(clean, "recommendation") or first_match(r"^Recommendation:\s*(.+?)\s*$", clean)
    positive_signals = yamlish_list(clean, "positive_signals")
    risk_signals = yamlish_list(clean, "risk_signals")
    best_positioning = yamlish_list(clean, "best_positioning")
    next_actions = yamlish_list(clean, "next_action")
    application_status = first_match(r"application_questions_status:\s*\n\s+status:\s*(.+?)\s*$", clean)
    custom_questions = first_match(r"custom_questions_found:\s*(.+?)\s*$", clean)
    question_note = first_match(r"note:\s*\"?(.+?)\"?\s*$", clean)
    cover_letter = section_between(clean, r"^Cover letter angle\s*$", [r"^Outreach message\s*$", r"^Final call:"])
    cover_letter = section_between(cover_letter, r"^Use:\s*$", []) or cover_letter
    outreach = section_between(clean, r"^Outreach message\s*$", [r"^Final call:"])
    positioning_angle = section_between(clean, r"^Best positioning angle\s*$", [r"^Cover letter angle\s*$", r"^Outreach message\s*$"])
    if "Lead with:" in positioning_angle:
        positioning_angle = section_between(positioning_angle, r"Lead with:\s*", [])
    return {
        "company": company,
        "title": title,
        "fit_score": fit_score,
        "priority": priority_from_chatgpt(priority_text, fit_score),
        "priority_text": priority_text,
        "role_category": role_category,
        "role_lane": role_lane_from_chatgpt(title, role_category),
        "sector": sector_from_chatgpt(role_category, positive_signals, clean),
        "url": url,
        "recommendation": recommendation,
        "positive_signals": positive_signals,
        "risk_signals": risk_signals,
        "best_positioning": best_positioning,
        "next_actions": next_actions,
        "application_status": application_status,
        "custom_questions_found": custom_questions.lower() == "true",
        "question_note": question_note,
        "cover_letter": cover_letter,
        "outreach": outreach,
        "positioning_angle": positioning_angle,
        "raw_text": clean,
    }


def result_from_chatgpt_analysis(parsed: dict[str, Any]) -> dict[str, Any]:
    signals = parsed["positive_signals"] or parsed["best_positioning"] or ["ChatGPT recommendation marked this as a plausible fit."]
    concerns = parsed["risk_signals"] or ["Review seniority, compensation, and location fit before applying."]
    bullets = parsed["best_positioning"] or [
        "Use the strongest truthful overlap from the ChatGPT analysis when tailoring resume bullets."
    ]
    outreach = parsed["outreach"] or parsed["recommendation"] or "Draft outreach from the imported ChatGPT job analysis."
    return {
        "title": parsed["title"],
        "company": parsed["company"],
        "fit_score": parsed["fit_score"],
        "fit_band": band_for_score(parsed["fit_score"]),
        "matched_signals": signals,
        "concerns": concerns,
        "outreach_message": outreach,
        "resume_bullet_adjustments": bullets,
        "job_description": parsed["raw_text"],
    }


def save_chatgpt_job_import(db_path: Path, parsed: dict[str, Any], *, allow_duplicate: bool = False) -> int:
    result = result_from_chatgpt_analysis(parsed)
    duplicates = potential_duplicate_jobs(
        db_path,
        title=result["title"],
        company=result["company"],
        source=parsed["url"],
    )
    if duplicates and not allow_duplicate:
        raise SystemExit(duplicate_jobs_message(duplicates))
    job_id = save_result(db_path, result, parsed["url"])
    next_action = "; ".join(parsed["next_actions"]) or parsed["recommendation"] or "Review imported ChatGPT recommendation"
    update_job_details(
        db_path,
        job_id,
        decision=parsed["recommendation"] or None,
        sector=parsed["sector"],
        role_lane=parsed["role_lane"],
        priority=parsed["priority"],
        application_url=parsed["url"] or None,
        cover_letter_needed="Optional" if parsed["cover_letter"] else None,
        next_action=next_action,
        note=chatgpt_import_note(parsed),
    )
    save_chatgpt_application_override(job_id, parsed)
    return job_id


def chatgpt_import_note(parsed: dict[str, Any]) -> str:
    parts = [
        "Imported from ChatGPT job analysis.",
        f"Recommendation: {parsed['recommendation']}" if parsed["recommendation"] else "",
        f"Original priority: {parsed['priority_text']}" if parsed["priority_text"] else "",
        f"Role category: {parsed['role_category']}" if parsed["role_category"] else "",
    ]
    if parsed["positioning_angle"]:
        parts.append(f"Best positioning angle:\n{parsed['positioning_angle']}")
    if parsed["cover_letter"]:
        parts.append(f"Cover letter sample from ChatGPT:\n{parsed['cover_letter']}")
    if parsed["outreach"]:
        parts.append(f"Outreach sample from ChatGPT:\n{parsed['outreach']}")
    return "\n\n".join(part for part in parts if part)


def save_chatgpt_application_override(job_id: int, parsed: dict[str, Any]) -> None:
    overrides = load_application_question_overrides()
    existing = overrides.get(str(job_id), {}) if isinstance(overrides.get(str(job_id)), dict) else {}
    if parsed["application_status"]:
        existing["capture_status"] = parsed["application_status"]
    elif parsed["question_note"]:
        existing["capture_status"] = "Captured"
    if parsed["question_note"]:
        existing["capture_reason"] = parsed["question_note"]
    elif parsed["application_status"]:
        existing["capture_reason"] = "Imported from ChatGPT job analysis."
    if "capture_next_action" not in existing and parsed["application_status"] == "Captured":
        existing["capture_next_action"] = "None"
    positioning = parsed["positioning_angle"] or "; ".join(parsed["best_positioning"])
    if positioning:
        existing["recommended_positioning"] = positioning
    existing["score_override"] = parsed["fit_score"]
    existing["source_chatgpt_analysis"] = parsed["raw_text"]
    existing["suggested_cover_letter_text"] = parsed["cover_letter"]
    overrides[str(job_id)] = existing
    save_application_question_overrides(overrides)


def render_markdown(result: dict[str, Any], row_id: int | None = None) -> str:
    header = f"# Job Fit Evaluation: {result['fit_score']}/100 ({result['fit_band']})"
    if row_id is not None:
        header += f"\n\nStored result ID: {row_id}"

    matches = "\n".join(f"- {item}" for item in result["matched_signals"]) or "- No strong fit signals found"
    concerns = "\n".join(f"- {item}" for item in result["concerns"]) or "- No major concerns found"
    bullets = "\n".join(f"- {item}" for item in result["resume_bullet_adjustments"])

    return f"""{header}

## Role

- Title: {result['title']}
- Company: {result['company']}

## Why It's a Fit

{matches}

## Why It May Not Be a Fit

{concerns}

## Tailored Outreach Message

{result['outreach_message']}

## Resume Bullet Adjustments

{bullets}
"""


def list_results(db_path: Path, limit: int) -> list[dict[str, Any]]:
    ensure_schema(db_path)
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            """
            SELECT id, created_at, title, company, fit_score, fit_band,
                   status, follow_up_date, next_action, source
            FROM job_evaluations
            ORDER BY id DESC
            LIMIT ?
            """,
            (limit,),
        ).fetchall()
    return [dict(row) for row in rows]


def export_csv(db_path: Path, output_path: Path) -> int:
    ensure_schema(db_path)
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            """
            SELECT id, created_at, source, title, company, fit_score, fit_band,
                   matched_signals, concerns, outreach_message,
                   resume_bullet_adjustments, job_description, status, decision,
                   source_board, application_url, applied_at, resume_version,
                   follow_up_date, next_action, archived_reason, last_updated_at,
                   sector, role_lane, priority, cover_letter_needed, referral_target
            FROM job_evaluations
            ORDER BY id ASC
            """
        ).fetchall()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=rows[0].keys() if rows else [
            "id",
            "created_at",
            "source",
            "title",
            "company",
            "fit_score",
            "fit_band",
            "matched_signals",
            "concerns",
            "outreach_message",
            "resume_bullet_adjustments",
            "job_description",
            "status",
            "decision",
            "source_board",
            "application_url",
            "applied_at",
            "resume_version",
            "follow_up_date",
            "next_action",
            "archived_reason",
            "last_updated_at",
            "sector",
            "role_lane",
            "priority",
            "cover_letter_needed",
            "referral_target",
        ])
        writer.writeheader()
        writer.writerows(dict(row) for row in rows)
    return len(rows)


def json_list_to_text(value: str | None, *, limit: int | None = None) -> str:
    if not value:
        return ""
    try:
        parsed = json.loads(value)
    except json.JSONDecodeError:
        return value
    if isinstance(parsed, list):
        items = [str(item) for item in parsed]
        if limit is not None:
            items = items[:limit]
        return " | ".join(items)
    return str(parsed)


def json_list(value: str | None) -> list[str]:
    if not value:
        return []
    try:
        parsed = json.loads(value)
    except json.JSONDecodeError:
        return [value]
    if isinstance(parsed, list):
        return [str(item) for item in parsed]
    return [str(parsed)]


def infer_from_patterns(text: str, patterns: list[tuple[str, str]], fallback: str) -> str:
    for label, pattern in patterns:
        if re.search(pattern, text, flags=re.IGNORECASE):
            return label
    return fallback


def infer_sector(*values: Any) -> str:
    text = " ".join(str(value or "") for value in values)
    return infer_from_patterns(text, SECTOR_PATTERNS, "Other")


def infer_sector_for_job(
    company: str | None,
    title: str | None,
    source: str | None,
    job_description: str | None,
    signals: list[str],
    concerns: list[str],
) -> str:
    identity_text = " ".join([str(company or ""), str(title or ""), str(source or "")])
    identity_sector = infer_from_patterns(identity_text, IDENTITY_SECTOR_PATTERNS, "")
    if identity_sector:
        return identity_sector
    description_sector = infer_sector(job_description)
    if description_sector != "Other":
        return description_sector
    return infer_sector(" ".join(signals), " ".join(concerns))


def infer_role_lane(title: str | None, *values: Any) -> str:
    title_match = infer_from_patterns(str(title or ""), ROLE_LANE_PATTERNS, "")
    if title_match:
        return title_match
    text = " ".join([str(title or ""), *(str(value or "") for value in values)])
    return infer_from_patterns(text, ROLE_LANE_PATTERNS, "Other")


def compute_priority(status: str | None, fit_score: int | None, concerns: list[str]) -> str:
    clean_status = status or "discovered"
    if clean_status in FINAL_STATUSES:
        return "Closed"
    if clean_status in {"recruiter_reply", "interviewing", "offer"}:
        return "Active"
    concern_text = " ".join(concerns).lower()
    if "weapons" in concern_text or "lethal" in concern_text or "surveillance" in concern_text:
        return "Review"
    score = fit_score or 0
    if score >= 92:
        return "P1 Apply Today"
    if score >= 82:
        return "P2 Strong"
    if score >= 70:
        return "P3 Maybe"
    return "Park"


def infer_cover_letter_needed(priority: str, fit_score: int | None, concerns: list[str]) -> str:
    concern_text = " ".join(concerns).lower()
    if "weapons" in concern_text or "lethal" in concern_text or "surveillance" in concern_text:
        return "Review"
    if priority in {"P1 Apply Today", "Active"} or (fit_score or 0) >= 92:
        return "Yes"
    if (fit_score or 0) >= 80:
        return "Optional"
    return "No"


def display_datetime(value: str | None) -> str:
    if not value:
        return ""
    clean = str(value).strip()
    if not clean:
        return ""
    try:
        if "T" in clean:
            stamp = dt.datetime.fromisoformat(clean.replace("Z", "+00:00"))
            if stamp.tzinfo is not None:
                stamp = stamp.astimezone(LOCAL_TIMEZONE)
            hour = stamp.strftime("%I").lstrip("0") or "0"
            minute = stamp.strftime("%M")
            ampm = stamp.strftime("%p")
            weekday = stamp.strftime("%A")
            return f"{hour}:{minute}{ampm}, {weekday} {stamp.month}/{stamp.day}/{stamp.strftime('%y')}"
        parsed = dt.date.fromisoformat(clean[:10])
        return f"{parsed.strftime('%A')} {parsed.month}/{parsed.day}/{parsed.strftime('%y')}"
    except ValueError:
        return clean[:16]


def dashboard_metadata_for_job(row: sqlite3.Row | dict[str, Any]) -> dict[str, str]:
    signals = json_list(row["matched_signals"])
    concerns = json_list(row["concerns"])
    inference_text = " ".join(
        [
            str(row["title"] or ""),
            str(row["company"] or ""),
            str(row["source"] or ""),
            " ".join(signals),
            " ".join(concerns),
            str(row["job_description"] or ""),
        ]
    )
    sector = row["sector"] or infer_sector_for_job(
        row["company"],
        row["title"],
        row["source"],
        row["job_description"],
        signals,
        concerns,
    )
    role_lane = row["role_lane"] or infer_role_lane(row["title"], inference_text)
    priority = row["priority"] or compute_priority(row["status"], row["fit_score"], concerns)
    cover_letter_needed = row["cover_letter_needed"] or infer_cover_letter_needed(priority, row["fit_score"], concerns)
    return {
        "sector": sector,
        "role_lane": role_lane,
        "priority": priority,
        "cover_letter_needed": cover_letter_needed,
    }


def dashboard_metadata_for_result(result: dict[str, Any], source: str | None) -> dict[str, str]:
    signals = [str(item) for item in result.get("matched_signals", [])]
    concerns = [str(item) for item in result.get("concerns", [])]
    inference_text = " ".join(
        [
            str(result.get("title") or ""),
            str(result.get("company") or ""),
            str(source or ""),
            " ".join(signals),
            " ".join(concerns),
            str(result.get("job_description") or ""),
        ]
    )
    sector = infer_sector_for_job(
        result.get("company"),
        result.get("title"),
        source,
        result.get("job_description"),
        signals,
        concerns,
    )
    role_lane = infer_role_lane(result.get("title"), inference_text)
    priority = compute_priority("discovered", result.get("fit_score"), concerns)
    return {
        "sector": sector,
        "role_lane": role_lane,
        "priority": priority,
        "cover_letter_needed": infer_cover_letter_needed(priority, result.get("fit_score"), concerns),
    }


def load_dashboard_jobs(db_path: Path) -> list[dict[str, Any]]:
    ensure_schema(db_path)
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            """
            SELECT id, status, fit_score, fit_band, company, title, source_board,
                   source, application_url, follow_up_date, next_action, applied_at,
                   resume_version, matched_signals, concerns, job_description,
                   sector, role_lane, priority, cover_letter_needed, referral_target,
                   created_at, last_updated_at
            FROM job_evaluations
            ORDER BY
                CASE status
                    WHEN 'interviewing' THEN 1
                    WHEN 'recruiter_reply' THEN 2
                    WHEN 'applied' THEN 3
                    WHEN 'outreach_sent' THEN 4
                    WHEN 'shortlisted' THEN 5
                    WHEN 'discovered' THEN 6
                    ELSE 9
                END,
                fit_score DESC,
                id DESC
            """
        ).fetchall()
    return [dict(row) for row in rows]


def dashboard_row(row: dict[str, Any]) -> dict[str, Any]:
    metadata = dashboard_metadata_for_job(row)
    packet = packet_metadata_for_job(row)
    concerns = json_list(row["concerns"])
    signals = json_list(row["matched_signals"])
    return {
        "ID": row["id"],
        "Status": row["status"],
        "Priority": metadata["priority"],
        "Fit Score": row["fit_score"],
        "Fit Band": row["fit_band"],
        "Sector": metadata["sector"],
        "Role Lane": metadata["role_lane"],
        "Company": row["company"],
        "Title": row["title"],
        "Source Board": row["source_board"] or infer_source_board(row["source"]),
        "Source URL": row["source"],
        "Application URL": row["application_url"],
        "Packet Status": packet["status"],
        "Packet Updated": packet["updated"],
        "Packet Link": packet["link"],
        "Follow Up Date": display_datetime(row["follow_up_date"]),
        "Last Touch": display_datetime(row["last_updated_at"] or row["applied_at"] or row["created_at"]),
        "Next Action": row["next_action"],
        "Applied At": display_datetime(row["applied_at"]),
        "Resume Version": row["resume_version"],
        "Cover Letter Needed": metadata["cover_letter_needed"],
        "Referral Target": row["referral_target"],
        "Concerns": " | ".join(concerns),
        "Top Fit Signals": "; ".join(signals[:3]),
        "Created At": display_datetime(row["created_at"]),
        "Last Updated At": display_datetime(row["last_updated_at"]),
    }


def packet_row_payloads(db_path: Path) -> list[dict[str, Any]]:
    index = load_packet_index()
    rows: list[dict[str, Any]] = []
    for job in load_dashboard_jobs(db_path):
        packet = packet_metadata_for_job(job, index)
        if not packet["status"]:
            continue
        metadata = dashboard_metadata_for_job(job)
        rows.append(
            {
                "sort_key": (metadata["priority"], -int(job["fit_score"] or 0), job["id"]),
                "row": {
                    "Job ID": job["id"],
                    "Company": job["company"],
                    "Title": job["title"],
                    "Priority": metadata["priority"],
                    "Status": job["status"],
                    "Fit Score": job["fit_score"],
                    "Packet Status": packet["status"],
                    "Packet Updated": packet["updated"],
                    "Packet Link": packet["link"],
                    "Packet Summary": packet_summary_text(job, packet),
                },
                "note": packet_hover_note(job, packet),
            }
        )
    rows.sort(key=lambda item: item["sort_key"])
    return rows


def packet_rows(db_path: Path) -> list[dict[str, Any]]:
    return [payload["row"] for payload in packet_row_payloads(db_path)]


def export_sheets_csv(db_path: Path, output_path: Path) -> int:
    """Export the master Jobs dashboard CSV for Google Sheets."""
    rows = [dashboard_row(row) for row in load_dashboard_jobs(db_path)]
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=SHEETS_JOB_FIELDNAMES)
        writer.writeheader()
        writer.writerows(rows)
    return len(rows)


def export_target_companies_csv(db_path: Path, output_path: Path) -> int:
    rows = target_company_rows(db_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=TARGET_COMPANY_FIELDNAMES)
        writer.writeheader()
        writer.writerows(rows)
    return len(rows)


def write_csv_rows(output_path: Path, fieldnames: list[str], rows: list[dict[str, Any]]) -> int:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)
    return len(rows)


def format_sheet_table_rows(table: str, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    date_fields_by_table = {
        "job_evaluations": {"created_at", "applied_at", "follow_up_date", "last_updated_at"},
        "applications": {"created_at", "applied_at", "follow_up_date"},
        "target_companies": {"created_at", "funding_date", "last_checked", "last_updated_at"},
        "contacts": {"created_at"},
        "correspondence": {"created_at", "date", "follow_up_date"},
        "notes": {"created_at"},
    }
    boolean_fields_by_table = {
        "applications": {"questionnaire_completed", "video_submitted"},
    }
    formatted_rows: list[dict[str, Any]] = []
    date_fields = date_fields_by_table.get(table, set())
    boolean_fields = boolean_fields_by_table.get(table, set())
    for row in rows:
        formatted = dict(row)
        for field in date_fields:
            if field in formatted:
                formatted[field] = display_datetime(formatted[field])
        for field in boolean_fields:
            if field in formatted:
                formatted[field] = "Yes" if formatted[field] else ""
        formatted_rows.append(formatted)
    return formatted_rows


def contact_rows(db_path: Path) -> list[dict[str, Any]]:
    ensure_schema(db_path)
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            """
            SELECT c.job_id, j.company, j.title, c.name, c.relationship, c.role,
                   c.telegram_handle, c.email, c.linkedin_url, c.notes, c.created_at
            FROM contacts c
            LEFT JOIN job_evaluations j ON j.id = c.job_id
            ORDER BY lower(COALESCE(j.company, '')), lower(c.name), c.id
            """
        ).fetchall()
    return [
        {
            "Job ID": row["job_id"] or "",
            "Company": row["company"] or "",
            "Title": row["title"] or "",
            "Name": row["name"] or "",
            "Relationship": row["relationship"] or "",
            "Role": row["role"] or "",
            "Telegram Handle": row["telegram_handle"] or "",
            "Email": row["email"] or "",
            "LinkedIn": row["linkedin_url"] or "",
            "Notes": row["notes"] or "",
            "Added": display_datetime(row["created_at"]),
        }
        for row in rows
    ]


def correspondence_rows(db_path: Path) -> list[dict[str, Any]]:
    ensure_schema(db_path)
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            """
            SELECT c.date, c.job_id, j.company, j.title, ct.name AS contact_name,
                   ct.relationship AS contact_relationship, c.channel, c.direction,
                   c.type, c.summary, c.follow_up_needed, c.follow_up_date
            FROM correspondence c
            JOIN job_evaluations j ON j.id = c.job_id
            LEFT JOIN contacts ct ON ct.id = c.contact_id
            ORDER BY c.date DESC, c.id DESC
            """
        ).fetchall()
    return [
        {
            "Date": display_datetime(row["date"]),
            "Job ID": row["job_id"] or "",
            "Company": row["company"] or "",
            "Title": row["title"] or "",
            "Contact": row["contact_name"] or "",
            "Contact Relationship": row["contact_relationship"] or "",
            "Channel": row["channel"] or "",
            "Direction": row["direction"] or "",
            "Type": row["type"] or "",
            "Summary": row["summary"] or "",
            "Follow Up Needed": "Yes" if row["follow_up_needed"] else "",
            "Follow Up Date": display_datetime(row["follow_up_date"]),
        }
        for row in rows
    ]


def application_rows(db_path: Path) -> list[dict[str, Any]]:
    ensure_schema(db_path)
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        rows = [dict(row) for row in conn.execute("SELECT * FROM applications ORDER BY id ASC").fetchall()]
    return format_sheet_table_rows("applications", rows)


def note_rows(db_path: Path) -> list[dict[str, Any]]:
    ensure_schema(db_path)
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        rows = [dict(row) for row in conn.execute("SELECT * FROM notes ORDER BY id ASC").fetchall()]
    return format_sheet_table_rows("notes", rows)


def target_company_rows(db_path: Path) -> list[dict[str, Any]]:
    ensure_schema(db_path)
    return [target_company_row(row) for row in load_target_companies(db_path)]


def workbook_tab_payloads(db_path: Path) -> dict[str, dict[str, Any]]:
    ensure_schema(db_path)
    raw_jobs = load_dashboard_jobs(db_path)
    jobs = [dashboard_row(row) for row in raw_jobs]
    packets = packet_rows(db_path)
    contacts = contact_rows(db_path)
    correspondence = correspondence_rows(db_path)
    applications = application_rows(db_path)
    notes = note_rows(db_path)
    target_companies = target_company_rows(db_path)
    active_statuses = {"discovered", "shortlisted"}
    applied_statuses = {"applied", "outreach_sent", "recruiter_reply", "interviewing", "offer", "rejected"}
    top_today = [
        row for row in jobs
        if row["Status"] in active_statuses and row["Priority"] in {"P1 Apply Today", "P2 Strong", "Review"}
    ][:12]
    applied = [row for row in jobs if row["Status"] in applied_statuses]
    today_plus_three = (today_local() + dt.timedelta(days=3)).isoformat()
    follow_ups = [
        dashboard_row(row) for row in raw_jobs
        if row["follow_up_date"]
        and str(row["follow_up_date"]) <= today_plus_three
        and row["status"] not in FINAL_STATUSES
    ]

    sectors = sorted({str(row["Sector"] or "Other") for row in jobs})
    sector_summary = []
    for sector in sectors:
        sector_jobs = [row for row in jobs if row["Sector"] == sector]
        applied_count = sum(1 for row in sector_jobs if row["Status"] in applied_statuses)
        p1_count = sum(1 for row in sector_jobs if row["Priority"] == "P1 Apply Today")
        avg_fit = round(sum(int(row["Fit Score"] or 0) for row in sector_jobs) / len(sector_jobs), 1)
        best_open = next(
            (
                f"{row['Company']} - {row['Title']}"
                for row in sorted(sector_jobs, key=lambda item: int(item["Fit Score"] or 0), reverse=True)
                if row["Status"] not in FINAL_STATUSES
            ),
            "",
        )
        sector_summary.append(
            {
                "Sector": sector,
                "Jobs": len(sector_jobs),
                "P1 Jobs": p1_count,
                "Applied+": applied_count,
                "Avg Fit Score": avg_fit,
                "Best Open Role": best_open,
            }
        )

    if applications and notes:
        application_fieldnames = list(applications[0].keys())
        note_fieldnames = list(notes[0].keys())
    else:
        with sqlite3.connect(db_path) as conn:
            application_fieldnames = list(applications[0].keys()) if applications else [
                row[1] for row in conn.execute("PRAGMA table_info(applications)").fetchall()
            ]
            note_fieldnames = list(notes[0].keys()) if notes else [
                row[1] for row in conn.execute("PRAGMA table_info(notes)").fetchall()
            ]
    return {
        "Jobs": {"fieldnames": SHEETS_JOB_FIELDNAMES, "rows": jobs},
        "Top Today": {"fieldnames": SHEETS_JOB_FIELDNAMES, "rows": top_today},
        "Applied": {"fieldnames": SHEETS_JOB_FIELDNAMES, "rows": applied},
        "Follow Ups": {"fieldnames": SHEETS_JOB_FIELDNAMES, "rows": follow_ups},
        "Sector Summary": {
            "fieldnames": ["Sector", "Jobs", "P1 Jobs", "Applied+", "Avg Fit Score", "Best Open Role"],
            "rows": sector_summary,
        },
        "Packets": {"fieldnames": PACKETS_FIELDNAMES, "rows": packets},
        "Applications": {"fieldnames": application_fieldnames, "rows": applications},
        "Target Companies": {"fieldnames": TARGET_COMPANY_FIELDNAMES, "rows": target_companies},
        "Contacts": {"fieldnames": CONTACTS_FIELDNAMES, "rows": contacts},
        "Correspondence": {"fieldnames": CORRESPONDENCE_FIELDNAMES, "rows": correspondence},
        "Notes": {"fieldnames": note_fieldnames, "rows": notes},
    }


def export_sheets_workbook(db_path: Path, output_dir: Path) -> dict[str, int]:
    """Export CSVs matching the Google Sheets CRM workbook tabs."""
    payloads = workbook_tab_payloads(db_path)
    counts: dict[str, int] = {}
    for tab_name, payload in payloads.items():
        counts[tab_name] = write_csv_rows(
            output_dir / f"{tab_name}.csv",
            list(payload["fieldnames"]),
            list(payload["rows"]),
        )
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        for table in ["job_evaluations", "applications", "target_companies", "notes"]:
            rows = [dict(row) for row in conn.execute(f"SELECT * FROM {table} ORDER BY id ASC").fetchall()]
            fieldnames = list(rows[0].keys()) if rows else [row[1] for row in conn.execute(f"PRAGMA table_info({table})").fetchall()]
            counts[table] = write_csv_rows(output_dir / f"{table}.csv", fieldnames, format_sheet_table_rows(table, rows))
    return counts


def google_credentials_candidates(
    auth_mode: str,
    *,
    service_account_json: Path,
    oauth_client_json: Path,
    prefer: str,
) -> list[Path]:
    if auth_mode == "oauth":
        return [oauth_client_json]
    if auth_mode == "service-account":
        return [service_account_json]
    ordered = [oauth_client_json, service_account_json] if prefer == "oauth" else [service_account_json, oauth_client_json]
    candidates: list[Path] = []
    for path in ordered:
        if path.exists() and path not in candidates:
            candidates.append(path)
    if candidates:
        return candidates
    return [ordered[0]]


def spreadsheet_column_letter(index: int) -> str:
    value = index
    letters = ""
    while value > 0:
        value, remainder = divmod(value - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def quoted_sheet_name(name: str) -> str:
    return "'" + name.replace("'", "''") + "'"


def workbook_values(fieldnames: list[str], rows: list[dict[str, Any]]) -> list[list[Any]]:
    values: list[list[Any]] = [list(fieldnames)]
    for row in rows:
        values.append(["" if row.get(field) is None else row.get(field) for field in fieldnames])
    return values


def sheet_rows_from_values(values: list[list[Any]]) -> list[dict[str, str]]:
    if not values:
        return []
    headers = [str(value).strip() for value in values[0]]
    rows: list[dict[str, str]] = []
    for values_row in values[1:]:
        row: dict[str, str] = {}
        for index, header in enumerate(headers):
            if not header:
                continue
            row[header] = str(values_row[index]).strip() if index < len(values_row) else ""
        rows.append(row)
    return rows


def parse_sheet_date(value: str) -> str | None:
    clean = normalize(value)
    if not clean:
        return None
    try:
        return parse_date(clean, field_name="sheet date")
    except SystemExit:
        pass
    match = re.search(r"(\d{1,2})/(\d{1,2})/(\d{2,4})", clean)
    if not match:
        return None
    month, day, year = (int(match.group(1)), int(match.group(2)), int(match.group(3)))
    if year < 100:
        year += 2000
    try:
        return dt.date(year, month, day).isoformat()
    except ValueError:
        return None


def parse_sheet_display_datetime(value: str) -> dt.datetime | None:
    clean = normalize(value)
    if not clean:
        return None
    try:
        parsed = dt.datetime.fromisoformat(clean.replace("Z", "+00:00"))
        return parsed if parsed.tzinfo else parsed.replace(tzinfo=dt.timezone.utc)
    except ValueError:
        pass
    match = re.search(r"(\d{1,2}):(\d{2})(AM|PM),\s+\w+\s+(\d{1,2})/(\d{1,2})/(\d{2,4})", clean, flags=re.IGNORECASE)
    if not match:
        return None
    hour = int(match.group(1))
    minute = int(match.group(2))
    ampm = match.group(3).upper()
    month = int(match.group(4))
    day = int(match.group(5))
    year = int(match.group(6))
    if year < 100:
        year += 2000
    if ampm == "PM" and hour != 12:
        hour += 12
    if ampm == "AM" and hour == 12:
        hour = 0
    try:
        return dt.datetime(year, month, day, hour, minute, tzinfo=LOCAL_TIMEZONE)
    except ValueError:
        return None


def read_sheet_rows(sheets_service: Any, spreadsheet_id: str, tab_name: str) -> list[dict[str, str]]:
    response = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{quoted_sheet_name(tab_name)}!A:AZ",
    ).execute()
    values = response.get("values") or []
    return sheet_rows_from_values(values)


def import_jobs_sheet_edits(db_path: Path, sheets_service: Any, spreadsheet_id: str) -> dict[str, Any]:
    ensure_schema(db_path)
    rows = read_sheet_rows(sheets_service, spreadsheet_id, "Jobs")
    imported = 0
    status_updates = 0
    application_records = 0
    warnings: list[str] = []
    editable_text_fields = {
        "Priority": "priority",
        "Sector": "sector",
        "Role Lane": "role_lane",
        "Application URL": "application_url",
        "Next Action": "next_action",
        "Resume Version": "resume_version",
        "Cover Letter Needed": "cover_letter_needed",
        "Referral Target": "referral_target",
    }
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        for row in rows:
            raw_id = row.get("ID", "")
            if not raw_id:
                continue
            try:
                job_id = int(float(raw_id))
            except ValueError:
                warnings.append(f"Skipped row with invalid ID: {raw_id}")
                continue
            current = conn.execute(
                """
                SELECT id, status, priority, sector, role_lane, application_url,
                       next_action, resume_version, cover_letter_needed,
                       referral_target, follow_up_date, applied_at, last_updated_at
                FROM job_evaluations
                WHERE id = ?
                """,
                (job_id,),
            ).fetchone()
            if current is None:
                warnings.append(f"Skipped sheet row for missing local job #{job_id}")
                continue
            sheet_updated = parse_sheet_display_datetime(row.get("Last Updated At", ""))
            local_updated_raw = str(current["last_updated_at"] or "").strip()
            local_updated = parse_sheet_display_datetime(local_updated_raw) if local_updated_raw else None
            if (
                sheet_updated
                and local_updated
                and local_updated.astimezone(dt.timezone.utc)
                > sheet_updated.astimezone(dt.timezone.utc) + dt.timedelta(seconds=90)
            ):
                warnings.append(f"Skipped stale sheet row for job #{job_id}; local CRM is newer than the live sheet row.")
                continue
            updates: dict[str, Any] = {}
            raw_status = row.get("Status", "")
            status_changed = False
            if raw_status:
                try:
                    clean_status = normalize_status(raw_status)
                except SystemExit:
                    warnings.append(f"Skipped invalid status for job #{job_id}: {raw_status}")
                    clean_status = ""
                if clean_status and clean_status != current["status"]:
                    updates["status"] = clean_status
                    status_changed = True
                    status_updates += 1
            if status_changed:
                for sheet_field, db_field in editable_text_fields.items():
                    value = row.get(sheet_field, "")
                    if value and value != str(current[db_field] or ""):
                        updates[db_field] = value
                for sheet_field, db_field in [("Follow Up Date", "follow_up_date"), ("Applied At", "applied_at")]:
                    parsed_date = parse_sheet_date(row.get(sheet_field, ""))
                    if parsed_date and parsed_date != str(current[db_field] or ""):
                        updates[db_field] = parsed_date
            if updates.get("status") == "applied" or (current["status"] == "applied" and "status" not in updates):
                if not updates.get("applied_at") and not current["applied_at"]:
                    updates["applied_at"] = today_local().isoformat()
                if not updates.get("follow_up_date") and not current["follow_up_date"]:
                    updates["follow_up_date"] = default_follow_up_date()
                if not updates.get("next_action") and not current["next_action"]:
                    updates["next_action"] = "Follow up if no response"
            if not updates:
                continue
            assignments = ", ".join(f"{field} = ?" for field in updates)
            conn.execute(
                f"UPDATE job_evaluations SET {assignments}, last_updated_at = ? WHERE id = ?",
                [*updates.values(), now_utc(), job_id],
            )
            imported += 1
            new_status = updates.get("status") or current["status"]
            if new_status == "applied":
                existing_application = conn.execute(
                    "SELECT id FROM applications WHERE job_id = ? LIMIT 1",
                    (job_id,),
                ).fetchone()
                if existing_application is None:
                    conn.execute(
                        """
                        INSERT INTO applications (
                            job_id, created_at, applied_at, application_url,
                            resume_version, cover_note_used, outreach_message,
                            application_status, follow_up_date, notes
                        )
                        SELECT id, ?, COALESCE(applied_at, ?), application_url,
                               resume_version, NULL, outreach_message, 'applied',
                               follow_up_date, 'Created from Jobs tab status import.'
                        FROM job_evaluations
                        WHERE id = ?
                        """,
                        (now_utc(), today_local().isoformat(), job_id),
                    )
                    application_records += 1
    return {
        "rows_imported": imported,
        "status_updates": status_updates,
        "application_records_created": application_records,
        "warnings": warnings,
    }


def parse_sheet_int(value: str) -> int | None:
    clean = normalize(value)
    if not clean:
        return None
    try:
        return clamp_score(int(float(clean)))
    except ValueError:
        return None


def import_target_companies_sheet_edits(db_path: Path, sheets_service: Any, spreadsheet_id: str) -> dict[str, Any]:
    ensure_schema(db_path)
    rows = read_sheet_rows(sheets_service, spreadsheet_id, "Target Companies")
    imported = 0
    warnings: list[str] = []
    text_fields = {
        "Company": "company",
        "Website": "website",
        "Lane": "lane",
        "Description": "description",
        "Funding Amount": "funding_amount",
        "Round": "round",
        "Investors": "investors",
        "Open Roles Found": "open_roles_found",
        "Best Role Title": "best_role_title",
        "Role URL": "role_url",
        "Priority": "priority",
        "Target Strategy": "target_strategy",
        "Outreach Type": "outreach_type",
        "Warm Contact 1": "warm_contact_1",
        "Warm Contact 1 Title": "warm_contact_1_title",
        "Warm Contact 1 LinkedIn": "warm_contact_1_linkedin",
        "Warm Contact 2": "warm_contact_2",
        "Warm Contact 2 Title": "warm_contact_2_title",
        "Warm Contact 2 LinkedIn": "warm_contact_2_linkedin",
        "Outreach Angle": "outreach_angle",
        "Outreach Status": "outreach_status",
        "Application Status": "application_status",
        "Notes": "notes",
        "Next Action": "next_action",
        "Source URL": "source_url",
    }
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        for row in rows:
            raw_id = row.get("ID", "")
            if not raw_id:
                continue
            try:
                target_company_id = int(float(raw_id))
            except ValueError:
                warnings.append(f"Skipped target-company row with invalid ID: {raw_id}")
                continue
            current = conn.execute(
                "SELECT * FROM target_companies WHERE id = ?",
                (target_company_id,),
            ).fetchone()
            if current is None:
                warnings.append(f"Skipped sheet row for missing local target company #{target_company_id}")
                continue
            sheet_updated = parse_sheet_display_datetime(row.get("Last Updated At", ""))
            local_updated_raw = str(current["last_updated_at"] or "").strip()
            local_updated = parse_sheet_display_datetime(local_updated_raw) if local_updated_raw else None
            if (
                sheet_updated
                and local_updated
                and local_updated.astimezone(dt.timezone.utc)
                > sheet_updated.astimezone(dt.timezone.utc) + dt.timedelta(seconds=90)
            ):
                warnings.append(
                    f"Skipped stale sheet row for target company #{target_company_id}; local tracker is newer than the live sheet row."
                )
                continue
            updates: dict[str, Any] = {}
            for sheet_field, db_field in text_fields.items():
                value = row.get(sheet_field, "")
                if value and value != str(current[db_field] or ""):
                    updates[db_field] = value
            for sheet_field, db_field in [("Funding Date", "funding_date"), ("Last Checked", "last_checked")]:
                parsed_date = parse_sheet_date(row.get(sheet_field, ""))
                if parsed_date and parsed_date != str(current[db_field] or ""):
                    updates[db_field] = parsed_date
            for sheet_field, db_field in [("Company Fit Score", "company_fit_score"), ("Role Fit Score", "role_fit_score")]:
                parsed_int = parse_sheet_int(row.get(sheet_field, ""))
                if parsed_int is not None and parsed_int != current[db_field]:
                    updates[db_field] = parsed_int
            if "company_fit_score" in updates and "priority" not in updates:
                updates["priority"] = company_priority_from_score(updates["company_fit_score"])
            if not updates:
                continue
            assignments = ", ".join(f"{field} = ?" for field in updates)
            conn.execute(
                f"UPDATE target_companies SET {assignments}, last_updated_at = ? WHERE id = ?",
                [*updates.values(), now_utc(), target_company_id],
            )
            imported += 1
    return {"rows_imported": imported, "warnings": warnings}


def summarize_google_sheet_sync_error(exc: Exception) -> str:
    message = str(exc)
    lowered = message.lower()
    if "service_disabled" in lowered or "google sheets api has not been used" in lowered:
        return (
            "Google Sheets API is disabled for the selected Google Cloud project. "
            "Enable the Google Sheets API for that project, wait a minute, and rerun the sync."
        )
    if "invalid_scope" in lowered or "insufficient authentication scopes" in lowered:
        return (
            "The current OAuth token does not include Google Sheets access yet. "
            "Rerun the sync and complete the Google browser re-consent so the token is rewritten with Sheets scope."
        )
    if "permission" in lowered or "not have permission" in lowered:
        return (
            "The selected credential does not have edit access to the Job Search CRM sheet. "
            "Share the sheet with the service-account email or rerun with OAuth."
        )
    return message


def sync_packet_sheet_notes(
    sheets_service: Any,
    *,
    spreadsheet_id: str,
    sheet_id: int,
    row_count: int,
    notes: list[str],
) -> None:
    requests: list[dict[str, Any]] = []
    clear_end_row = max(row_count, len(notes) + 1)
    if clear_end_row > 1:
        requests.append(
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 1,
                        "endRowIndex": clear_end_row,
                        "startColumnIndex": 9,
                        "endColumnIndex": 10,
                    },
                    "cell": {"note": ""},
                    "fields": "note",
                }
            }
        )
    if notes:
        requests.append(
            {
                "updateCells": {
                    "start": {
                        "sheetId": sheet_id,
                        "rowIndex": 1,
                        "columnIndex": 9,
                    },
                    "rows": [{"values": [{"note": note}]} for note in notes],
                    "fields": "note",
                }
            }
        )
    if requests:
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": requests},
        ).execute()


def managed_sheet_metadata(sheets_service: Any, spreadsheet_id: str) -> dict[str, dict[str, Any]]:
    metadata = sheets_service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields="sheets(properties(sheetId,title,gridProperties),conditionalFormats)",
    ).execute()
    return {sheet["properties"]["title"]: sheet for sheet in metadata.get("sheets", [])}


def sheet_id_for(sheet_metadata: dict[str, dict[str, Any]], title: str) -> int | None:
    sheet = sheet_metadata.get(title)
    if not sheet:
        return None
    return int(sheet["properties"]["sheetId"])


def status_conditional_format_requests(
    *,
    sheet_id: int,
    status_column_index: int,
    end_column_index: int,
) -> list[dict[str, Any]]:
    requests: list[dict[str, Any]] = []
    status_column_letter = spreadsheet_column_letter(status_column_index + 1)
    for status, background in STATUS_COLOR_RULES.items():
        requests.append(
            {
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [
                            {
                                "sheetId": sheet_id,
                                "startRowIndex": 1,
                                "endRowIndex": SHEET_UI_ROW_LIMIT,
                                "startColumnIndex": 0,
                                "endColumnIndex": end_column_index,
                            }
                        ],
                        "booleanRule": {
                            "condition": {
                                "type": "CUSTOM_FORMULA",
                                "values": [{"userEnteredValue": f'=${status_column_letter}2="{status}"'}],
                            },
                            "format": {"backgroundColor": background},
                        },
                    },
                    "index": 0,
                }
            }
        )
    return requests


def apply_google_sheet_ui_rules(
    sheets_service: Any,
    *,
    spreadsheet_id: str,
    sheet_metadata: dict[str, dict[str, Any]],
) -> None:
    """Apply lightweight workbook UI rules that live inside the Google Sheet."""
    formula_clear_ranges = [
        f"{quoted_sheet_name('Top Today')}!A2:Z",
        f"{quoted_sheet_name('Applied')}!A2:Z",
        f"{quoted_sheet_name('Follow Ups')}!A2:Z",
        f"{quoted_sheet_name('Packets')}!E2:E",
    ]
    sheets_service.spreadsheets().values().batchClear(
        spreadsheetId=spreadsheet_id,
        body={"ranges": formula_clear_ranges},
    ).execute()
    formula_data = [
        {
            "range": f"{quoted_sheet_name('Top Today')}!A2",
            "values": [[
                '=IFERROR(ARRAY_CONSTRAIN(SORT(FILTER(Jobs!A2:Z,'
                '((Jobs!B2:B="discovered")+(Jobs!B2:B="shortlisted"))*'
                '((Jobs!C2:C="P1 Apply Today")+(Jobs!C2:C="P2 Strong")+(Jobs!C2:C="Review"))),4,FALSE),12,26),"")'
            ]],
        },
        {
            "range": f"{quoted_sheet_name('Applied')}!A2",
            "values": [[
                '=IFERROR(SORT(FILTER(Jobs!A2:Z,REGEXMATCH(Jobs!B2:B,'
                '"^(applied|outreach_sent|recruiter_reply|interviewing|offer|rejected)$")),19,FALSE),"")'
            ]],
        },
        {
            "range": f"{quoted_sheet_name('Follow Ups')}!A2",
            "values": [[
                '=IFERROR(SORT(FILTER(Jobs!A2:Z,Jobs!P2:P<>"",Jobs!P2:P<=TODAY()+3,'
                '(Jobs!B2:B<>"offer")*(Jobs!B2:B<>"rejected")*(Jobs!B2:B<>"archived")),16,TRUE),"")'
            ]],
        },
        {
            "range": f"{quoted_sheet_name('Packets')}!E2",
            "values": [['=ARRAYFORMULA(IF(A2:A="",,IFERROR(VLOOKUP(A2:A,Jobs!A:B,2,FALSE),"")))']],
        },
    ]
    sheets_service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"valueInputOption": "USER_ENTERED", "data": formula_data},
    ).execute()

    managed_ranges = {
        "Jobs": {"status_column_index": 1, "end_column_index": len(SHEETS_JOB_FIELDNAMES)},
        "Top Today": {"status_column_index": 1, "end_column_index": len(SHEETS_JOB_FIELDNAMES)},
        "Applied": {"status_column_index": 1, "end_column_index": len(SHEETS_JOB_FIELDNAMES)},
        "Follow Ups": {"status_column_index": 1, "end_column_index": len(SHEETS_JOB_FIELDNAMES)},
        "Packets": {"status_column_index": 4, "end_column_index": len(PACKETS_FIELDNAMES)},
    }
    requests: list[dict[str, Any]] = []
    for title, sheet in sheet_metadata.items():
        if title not in managed_ranges:
            continue
        sheet_id = int(sheet["properties"]["sheetId"])
        for index in reversed(range(len(sheet.get("conditionalFormats") or []))):
            requests.append({"deleteConditionalFormatRule": {"sheetId": sheet_id, "index": index}})

    header_widths = {
        "Jobs": len(SHEETS_JOB_FIELDNAMES),
        "Top Today": len(SHEETS_JOB_FIELDNAMES),
        "Applied": len(SHEETS_JOB_FIELDNAMES),
        "Follow Ups": len(SHEETS_JOB_FIELDNAMES),
        "Sector Summary": 6,
        "Packets": len(PACKETS_FIELDNAMES),
        "Applications": 10,
        "Target Companies": len(TARGET_COMPANY_FIELDNAMES),
        "Contacts": len(CONTACTS_FIELDNAMES),
        "Correspondence": len(CORRESPONDENCE_FIELDNAMES),
        "Notes": 6,
    }
    for title, field_count in header_widths.items():
        sheet_id = sheet_id_for(sheet_metadata, title)
        if sheet_id is None:
            continue
        requests.append(
            {
                "updateSheetProperties": {
                    "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 1}},
                    "fields": "gridProperties.frozenRowCount",
                }
            }
        )
        requests.append(
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": field_count,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {"red": 0.12, "green": 0.16, "blue": 0.22},
                            "horizontalAlignment": "CENTER",
                            "textFormat": {
                                "bold": True,
                                "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                            },
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
                }
            }
        )

    jobs_sheet_id = sheet_id_for(sheet_metadata, "Jobs")
    if jobs_sheet_id is not None:
        requests.append(
            {
                "repeatCell": {
                    "range": {
                        "sheetId": jobs_sheet_id,
                        "startRowIndex": 1,
                        "endRowIndex": SHEET_UI_ROW_LIMIT,
                        "startColumnIndex": 1,
                        "endColumnIndex": 2,
                    },
                    "cell": {
                        "dataValidation": {
                            "condition": {
                                "type": "ONE_OF_LIST",
                                "values": [{"userEnteredValue": status} for status in VALID_STATUSES],
                            },
                            "strict": True,
                            "showCustomUi": True,
                        }
                    },
                    "fields": "dataValidation",
                }
            }
        )

    for title, rule in managed_ranges.items():
        sheet_id = sheet_id_for(sheet_metadata, title)
        if sheet_id is None:
            continue
        requests.extend(
            status_conditional_format_requests(
                sheet_id=sheet_id,
                status_column_index=int(rule["status_column_index"]),
                end_column_index=int(rule["end_column_index"]),
            )
        )

    if requests:
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": requests},
        ).execute()


def sync_google_sheets_workbook(
    db_path: Path,
    *,
    spreadsheet_url: str,
    auth_mode: str,
    service_account_json: Path,
    oauth_client_json: Path,
    oauth_token_json: Path,
) -> dict[str, Any]:
    spreadsheet_id = extract_drive_id(spreadsheet_url)
    prefer = "service-account" if auth_mode == "hybrid" else "oauth"
    attempts: list[str] = []

    for credentials_path in google_credentials_candidates(
        auth_mode,
        service_account_json=service_account_json,
        oauth_client_json=oauth_client_json,
        prefer=prefer,
    ):
        try:
            sheets_service, credentials_type = build_google_sheets_service(
                credentials_path,
                scopes=GOOGLE_SHEETS_SYNC_SCOPES,
                token_json=oauth_token_json,
            )
            metadata = sheets_service.spreadsheets().get(
                spreadsheetId=spreadsheet_id,
                fields="sheets.properties",
            ).execute()
            sheet_map = {sheet["properties"]["title"]: sheet["properties"] for sheet in metadata.get("sheets", [])}
            sheet_import = {"rows_imported": 0, "status_updates": 0, "application_records_created": 0, "warnings": []}
            if "Jobs" in sheet_map:
                sheet_import = import_jobs_sheet_edits(db_path, sheets_service, spreadsheet_id)
            if "Target Companies" in sheet_map:
                target_import = import_target_companies_sheet_edits(db_path, sheets_service, spreadsheet_id)
                sheet_import["target_company_rows_imported"] = target_import.get("rows_imported", 0)
                sheet_import["warnings"] = list(sheet_import.get("warnings") or []) + list(target_import.get("warnings") or [])
            else:
                sheet_import["target_company_rows_imported"] = 0
            payloads = workbook_tab_payloads(db_path)
            packet_notes = [str(payload["note"]) for payload in packet_row_payloads(db_path)]
            missing_tabs = [title for title in payloads if title not in sheet_map]
            if missing_tabs:
                sheets_service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body={"requests": [{"addSheet": {"properties": {"title": title}}} for title in missing_tabs]},
                ).execute()
                metadata = sheets_service.spreadsheets().get(
                    spreadsheetId=spreadsheet_id,
                    fields="sheets.properties",
                ).execute()
                sheet_map = {sheet["properties"]["title"]: sheet["properties"] for sheet in metadata.get("sheets", [])}

            clear_ranges = []
            data = []
            counts: dict[str, int] = {}
            for tab_name, payload in payloads.items():
                fieldnames = list(payload["fieldnames"])
                rows = list(payload["rows"])
                counts[tab_name] = len(rows)
                max_column = spreadsheet_column_letter(max(1, len(fieldnames)))
                clear_ranges.append(f"{quoted_sheet_name(tab_name)}!A:{max_column}")
                data.append(
                    {
                        "range": f"{quoted_sheet_name(tab_name)}!A1",
                        "values": workbook_values(fieldnames, rows),
                    }
                )

            if clear_ranges:
                sheets_service.spreadsheets().values().batchClear(
                    spreadsheetId=spreadsheet_id,
                    body={"ranges": clear_ranges},
                ).execute()
            if data:
                sheets_service.spreadsheets().values().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body={"valueInputOption": "RAW", "data": data},
                ).execute()

            packet_sheet = sheet_map.get("Packets")
            if packet_sheet is not None:
                sync_packet_sheet_notes(
                    sheets_service,
                    spreadsheet_id=spreadsheet_id,
                    sheet_id=int(packet_sheet["sheetId"]),
                    row_count=int(packet_sheet.get("gridProperties", {}).get("rowCount", len(packet_notes) + 1)),
                    notes=packet_notes,
                )

            apply_google_sheet_ui_rules(
                sheets_service,
                spreadsheet_id=spreadsheet_id,
                sheet_metadata=managed_sheet_metadata(sheets_service, spreadsheet_id),
            )

            return {
                "spreadsheet_id": spreadsheet_id,
                "counts": counts,
                "credentials_type": credentials_type,
                "credentials_path": str(credentials_path),
                "sheet_import": sheet_import,
            }
        except Exception as exc:
            attempts.append(f"{credentials_path.name}: {summarize_google_sheet_sync_error(exc)}")

    raise SystemExit(
        "Failed to sync the live Google Sheet. "
        + " | ".join(attempts)
    )


def run_daily_workflow(
    db_path: Path,
    *,
    followup_days: int,
    pipeline_min_score: int,
    pipeline_limit: int,
    spreadsheet_url: str,
    sheet_auth_mode: str,
    service_account_json: Path,
    oauth_client_json: Path,
    oauth_token_json: Path,
    skip_sheet_sync: bool,
) -> dict[str, Any]:
    due_jobs, due_correspondence = followup_rows(db_path, days=followup_days)
    pipeline = pipeline_rows(db_path, status=None, min_score=pipeline_min_score, limit=pipeline_limit)
    counts = {
        "job_results": export_csv(db_path, ROOT / "exports" / "job_results.csv"),
        "google_sheets_jobs": export_sheets_csv(db_path, ROOT / "exports" / "google_sheets_job_tracker.csv"),
        "target_companies": export_target_companies_csv(db_path, ROOT / "exports" / "target_companies.csv"),
    }
    counts["crm"] = export_crm(db_path, ROOT / "exports" / "crm")
    counts["workbook"] = export_sheets_workbook(db_path, DEFAULT_SHEETS_WORKBOOK_DIR)
    sheet_sync = None
    if not skip_sheet_sync:
        sheet_sync = sync_google_sheets_workbook(
            db_path,
            spreadsheet_url=spreadsheet_url,
            auth_mode=sheet_auth_mode,
            service_account_json=service_account_json,
            oauth_client_json=oauth_client_json,
            oauth_token_json=oauth_token_json,
        )
    return {
        "followup_jobs": due_jobs,
        "followup_correspondence": due_correspondence,
        "pipeline": pipeline,
        "exports": counts,
        "sheet_sync": sheet_sync,
    }


def print_table(rows: list[dict[str, Any]]) -> None:
    if not rows:
        print("No stored evaluations yet.")
        return

    widths = {
        "id": 4,
        "score": 7,
        "band": 10,
        "status": 14,
        "follow": 12,
        "title": 30,
        "company": 20,
    }
    print(
        f"{'ID':<{widths['id']}} "
        f"{'Score':<{widths['score']}} "
        f"{'Band':<{widths['band']}} "
        f"{'Status':<{widths['status']}} "
        f"{'Follow-up':<{widths['follow']}} "
        f"{'Title':<{widths['title']}} "
        f"{'Company':<{widths['company']}}"
    )
    print("-" * 107)
    for row in rows:
        title = textwrap.shorten(row["title"] or "", width=widths["title"], placeholder="...")
        company = textwrap.shorten(row["company"] or "", width=widths["company"], placeholder="...")
        status = row.get("status") or "discovered"
        follow = row.get("follow_up_date") or ""
        print(
            f"{row['id']:<{widths['id']}} "
            f"{row['fit_score']:<{widths['score']}} "
            f"{row['fit_band']:<{widths['band']}} "
            f"{status:<{widths['status']}} "
            f"{follow:<{widths['follow']}} "
            f"{title:<{widths['title']}} "
            f"{company:<{widths['company']}}"
        )


def require_job(conn: sqlite3.Connection, job_id: int) -> sqlite3.Row:
    conn.row_factory = sqlite3.Row
    row = conn.execute("SELECT * FROM job_evaluations WHERE id = ?", (job_id,)).fetchone()
    if row is None:
        raise SystemExit(f"No job found with id {job_id}.")
    return row


def add_note(db_path: Path, job_id: int, note: str) -> int:
    ensure_schema(db_path)
    created_at = now_utc()
    with sqlite3.connect(db_path) as conn:
        require_job(conn, job_id)
        cursor = conn.execute(
            "INSERT INTO notes (job_id, created_at, note) VALUES (?, ?, ?)",
            (job_id, created_at, note),
        )
        conn.execute(
            "UPDATE job_evaluations SET last_updated_at = ? WHERE id = ?",
            (created_at, job_id),
        )
        return int(cursor.lastrowid)


def slugify(value: str) -> str:
    clean = re.sub(r"[^A-Za-z0-9]+", "-", value).strip("-").lower()
    return clean[:80] or "job"


def cover_letter_docx_path_for_packet_path(packet_path: Path) -> Path:
    if packet_path.name == "00_packet_bundle.md":
        return packet_path.with_name("04_cover_letter.docx")
    return packet_path.with_name(f"{packet_path.stem}-cover-letter.docx")


def is_crypto_or_web3_sector(sector: str) -> bool:
    normalized = (sector or "").strip().lower()
    return "crypto" in normalized or "web3" in normalized


def cover_letter_signature_lines(sector: str) -> list[str]:
    profile = load_profile(DEFAULT_PROFILE)
    contact = profile.get("contact", {}) if isinstance(profile, dict) else {}
    name = str(profile.get("name") or "Your Name").strip() or "Your Name"
    email = str(contact.get("email") or "you@example.com").strip() or "you@example.com"
    phone = str(contact.get("phone") or "").strip()
    linkedin = str(contact.get("linkedin") or "").strip()
    telegram = str(contact.get("telegram") or "").strip()

    lines = ["Thank you,", name, email]
    if phone:
        lines.append(phone)
    if linkedin:
        lines.append(linkedin.replace("https://", "").replace("http://", "").rstrip("/"))
    if is_crypto_or_web3_sector(sector) and telegram:
        handle = telegram if telegram.startswith("@") else f"@{telegram}"
        lines.append(f"Telegram: {handle}")
    return lines


def packet_cover_letter_text_tether(
    *,
    company: str,
    title: str,
    sector: str,
    variant: int = 0,
) -> str:
    if variant % 2 == 0:
        paragraphs = [
            "Dear Tether Hiring Team,",
            (
                f"{company} operates at a scale in crypto where account management has real commercial weight. "
                f"That is what makes the {title} role interesting to me. It sits in the overlap between customer ownership, "
                "commercial judgment, and infrastructure that already matters in the real world."
            ),
            (
                "Most recently, I managed strategic clients across research, data, and advisory relationships. The job "
                "was not just keeping accounts warm. It meant spotting renewal risk early, coordinating internal teams, and turning "
                "customer context into a clearer path for retention or expansion."
            ),
            (
                "I also bring earlier technical-partnership experience from a prior role, where I worked across ecosystem and "
                "protocol counterparties. That gave me a practical view of how crypto products get adopted, how cross-company work "
                "actually moves, and how much trust and follow-through matter when the market is moving fast."
            ),
            (
                "Earlier in my career, I learned the discipline behind long-cycle account ownership. I managed a large "
                "book of recurring business, worked through pricing and procurement issues, and got comfortable creating structure in situations "
                "that were not neatly defined from the start."
            ),
            (
                f"What I would bring to {company} is a steady account-management approach built around understanding the customer, keeping the "
                "commercial priorities clear, documenting what matters, and making sure momentum does not stall when the work gets messy."
            ),
            (
                f"{company} is compelling to me because the product matters at a scale that very few crypto companies reach. I would welcome "
                "the chance to help strengthen customer trust, support account growth, and keep important relationships moving in the right direction."
            ),
            "\n".join(cover_letter_signature_lines(sector)),
        ]
    else:
        paragraphs = [
            "Dear Tether Hiring Team,",
            (
                f"In crypto, very few companies operate with the reach and commercial importance that {company} does. "
                f"That is why the {title} role is interesting to me. It depends on someone who can keep customers coordinated, confident, and moving forward over time."
            ),
            (
                "In a recent account role, I worked with technical clients where renewal and expansion depended on more than pricing. A lot of the job was understanding "
                "account priorities, keeping internal stakeholders aligned, and making sure customers stayed clear on the value being delivered."
            ),
            (
                "Before that, I led strategic partnerships across ecosystem partners in an earlier role. That experience made me comfortable "
                "working in crypto-native environments where execution is cross-functional, relationship quality matters, and the commercial opportunity is tied to adoption."
            ),
            (
                "I also have a longer account-management background from an earlier enterprise role, where I managed recurring revenue, opened new accounts, and negotiated through "
                "procurement complexity, and learned how to protect revenue without needing ideal conditions or a polished playbook."
            ),
            (
                f"I would be a strong fit for {company} because I bring a practical operating style to customer work: understand the account, keep the next steps clear, "
                "work well across teams, and stay accountable for the commercial outcome."
            ),
            (
                f"Iâ€™d welcome the opportunity to speak further about how I could support {company}'s next stage of account growth and customer trust."
            ),
            "\n".join(cover_letter_signature_lines(sector)),
        ]
    return "\n\n".join(paragraphs).strip()


def packet_cover_letter_text(
    voice: dict[str, Any],
    *,
    company: str,
    title: str,
    role_lane: str,
    sector: str,
    sector_focus: str,
    job_id: int,
    signals: list[str],
    job_description: str,
    variant: int = 0,
) -> str:
    if normalized_company_key(company) == "tether":
        return packet_cover_letter_text_tether(company=company, title=title, sector=sector, variant=variant)
    opening = deterministic_choice(
        voice_options(voice, "cover_letter_openers"),
        stable_seed(company, title, job_id) + (variant * 101),
        "the responsibilities line up closely with the environments where I have historically done my best work",
    )
    anchor = role_lane_why_anchor(voice, role_lane, role_lane_phrase(role_lane))
    technical = technical_positioning_sentence(voice, sector, job_id + 23)
    overlap_entries = cover_letter_overlap_entries(
        voice,
        company=company,
        role_lane=role_lane,
        sector=sector,
        signals=signals,
        job_description=job_description,
    )
    fit_points = cover_letter_fit_points(role_lane, sector, signals, job_description)
    fit_point_sentence = format_series(fit_points)
    role_keywords = cover_letter_role_keywords(role_lane, job_description, signals)
    overlap_stories = unique_keep_order(
        [
            str(entry.get("story") or "").strip()
            for entry in overlap_entries
            if str(entry.get("story") or "").strip() and str(entry.get("story") or "").strip() != "TechnicalTranslation"
        ]
    )
    if not overlap_stories:
        overlap_stories = ["RecentAccountRole"]
    technical_selected = any(str(entry.get("story") or "").strip() == "TechnicalTranslation" for entry in overlap_entries)
    contribution_map = {
        "Strategic Account Management": "protect important customer relationships, retain revenue, and create sensible room for expansion over time",
        "Account Management": "grow and retain its customer base through strong retention, expansion, and cross-functional execution",
        "Customer Success": "turn customer trust into durable retention and commercially useful expansion opportunities",
        "Sales / Enterprise": "run consultative enterprise conversations, navigate multi-stakeholder deals, and build real commercial momentum",
        "Partnerships / BD": "build commercially meaningful partner relationships and help the right ecosystem opportunities turn into traction",
        "GTM / Commercial Lead": "connect market context, customer priorities, and commercial execution in a way that creates real momentum",
        "Advisor / Ecosystem": "help the right relationships turn into strategic leverage and commercially useful progress",
    }
    contribution = contribution_map.get(
        role_lane,
        "bring commercial judgment, relationship discipline, and cross-functional follow-through to important customer or partner work",
    )
    fit_paragraph = deterministic_choice(
        [
            f"That maps well to a role that depends on {anchor}.",
            f"The overlap is straightforward: the role needs {anchor}, and that is where a lot of my strongest work has been.",
            f"That is the kind of account-management motion I know well: {anchor}.",
        ],
        job_id + 27 + (variant * 17),
        f"That maps well to a role that depends on {anchor}.",
    )
    if technical_selected and technical:
        fit_paragraph = f"{technical} {fit_paragraph}"
    company_themes = cover_letter_company_themes(role_lane, sector, signals, job_description)
    theme_text = format_series(company_themes)
    company_interest_sentence = deterministic_choice(
        voice_options(voice, "cover_letter_company_interest_templates"),
        stable_seed(company, sector, job_id) + (variant * 131),
        "I am particularly interested in {company}'s focus on {themes}, which makes this opportunity especially compelling to me.",
    )
    company_interest_sentence = company_interest_sentence.replace("{company}", company).replace("{themes}", theme_text)
    if role_keywords:
        keyword_text = format_series(role_keywords[:6])
        company_interest_sentence = (
            f"What makes {company} especially relevant here is the focus on {keyword_text}. "
            "That is the kind of book I know how to manage well: stay close to the customer, keep decision makers engaged, "
            "spot growth opportunities early, and keep internal teams moving against the same account plan."
        )
    close_sentence = deterministic_choice(
        voice_options(voice, "cover_letter_close_templates"),
        stable_seed(company, title, job_id) + (variant * 151),
        "Iâ€™d welcome the opportunity to discuss how my background can help {company} {contribution}.",
    )
    close_sentence = close_sentence.replace("{company}", company).replace("{contribution}", contribution)
    if role_keywords:
        first_paragraph = (
            f"The {title} role at {company} lines up closely with my background in {fit_point_sentence}."
        )
    else:
        first_paragraph = f"I am applying for the {title} role at {company} because {opening}: {fit_point_sentence}."
    paragraphs = [
        "Hiring Team,",
        first_paragraph,
        *[cover_letter_story_paragraph(story) for story in overlap_stories[:2]],
        f"{fit_paragraph} {company_interest_sentence}",
        close_sentence,
        "\n".join(cover_letter_signature_lines(sector)),
    ]
    return "\n\n".join(paragraphs).strip()


def packet_cover_letter_text_quality_gate(
    *,
    company: str,
    title: str,
    sector: str,
    product_category: str,
    buyer_persona: str,
) -> str:
    if is_crypto_or_web3_sector(sector):
        paragraphs = [
            "Hiring Team,",
            (
                f"The {title} role at {company} sits in the lane where my Web3 background is most useful: strategic account ownership, "
                "partnerships, renewals, expansion, and technical-commercial customer work."
            ),
            (
                f"{RECENT_ACCOUNT_OWNERSHIP_SENTENCE} {RECENT_ACCOUNT_REVENUE_OUTCOMES_SENTENCE} "
                f"{RECENT_ACCOUNT_REVENUE_BREAKDOWN_SENTENCE} {RECENT_ACCOUNT_RECOVERY_SENTENCE} {RECENT_ACCOUNT_CROSS_FUNCTIONAL_SENTENCE}"
            ),
            (
                "Earlier in my career, I managed a large account portfolio with recurring annual revenue responsibility across complex customers. "
                "That work built the account discipline behind long-cycle relationships, procurement, pricing conversations, retention, and expansion across complex stakeholder groups."
            ),
            (
                "Before that, I led strategic partnerships in an earlier ecosystem role, where I closed multiple partnerships across technical stakeholders. "
                "That work required understanding partner goals, supporting integration conversations, building trust across technical and commercial stakeholders, and helping partners see where Web3 infrastructure could create practical value."
            ),
            (
                f"{company}'s work in {product_category.lower()} for {buyer_persona.lower()} is where the role becomes more specific. "
                "My strength is working between customer goals, product context, and commercial outcomes: building multi-threaded relationships, tracking account health, identifying risk, navigating renewal and expansion conversations, and keeping internal teams aligned around the account."
            ),
            (
                f"Iâ€™d welcome the opportunity to discuss how my background in Web3 partnerships, strategic account ownership, renewal strategy, and cross-functional execution can help {company} retain and grow its customer relationships."
            ),
            "\n".join(cover_letter_signature_lines(sector)),
        ]
        return "\n\n".join(paragraphs).strip()
    paragraphs = [
        "Hiring Team,",
        (
            f"The {title} role at {company} is a strong account-management fit because it centers on customer growth, "
            "retention, expansion, stakeholder alignment, and cross-functional execution."
        ),
        (
            "More recently, I managed strategic accounts across a technical product suite. "
            "I influenced meaningful renewal, recovery, and expansion outcomes across closed, recovered, and "
            "post-transition work. I also helped reactivate a previously stalled account through targeted "
            "stakeholder coordination."
        ),
        (
            "In an earlier enterprise account role, I managed a large account portfolio with recurring annual revenue across "
            "healthcare and government customers. That role required long-cycle relationship management, procurement, pricing "
            "conversations, retention, and expansion across complex stakeholder groups."
        ),
        (
            f"{company}'s work in {product_category} for {buyer_persona.lower()} is where the role becomes more specific. "
            "My strength is working between customer goals, product context, and commercial outcomes: tracking "
            "account health, identifying risk, building trust, and helping turn adoption into retention and expansion."
        ),
        (
            f"I would welcome the opportunity to discuss how my background in large-account ownership, renewal strategy, "
            f"customer growth, and cross-functional execution can help {company} support and expand its customer relationships."
        ),
        "\n".join(cover_letter_signature_lines(sector)),
    ]
    return "\n\n".join(paragraphs).strip()


def packet_bundle_directory_name(job_id: int, company: str, title: str) -> str:
    return f"job-{job_id}-{slugify(company)}-{slugify(title)}"


def normalized_sentences(text: str) -> list[str]:
    collapsed = re.sub(r"\s+", " ", (text or "").strip())
    if not collapsed:
        return []
    parts = re.split(r"(?<=[.!?])\s+", collapsed)
    return [normalize(part) for part in parts if normalize(part)]


def anti_copy_meaningful_sentences(sentences: list[str]) -> list[str]:
    ignored = {
        "hiring team",
        "Hiring Team",
        "dear tether hiring team",
        "Dear Tether Hiring Team",
        "your name",
        "Your Name",
        "Thank you, Your Name you@example.com 555-555-5555 linkedin.com/in/your-linkedin",
        "thank you, your name you@example.com 555-555-5555 linkedin.com/in/your-linkedin",
        "Thank you, Your Name you@example.com 555-555-5555 linkedin.com/in/your-linkedin Telegram: @yourhandle",
        "thank you, your name you@example.com 555-555-5555 linkedin.com/in/your-linkedin telegram: @yourhandle",
        "best",
        "Best",
        "thank you",
        "Thank you",
        "sincerely",
        "Sincerely",
        "In a recent account management role, I managed strategic accounts across a technical product suite.",
        "In an earlier partnerships role, I closed strategic partnerships across a technical ecosystem.",
        "In an earlier enterprise account role, I managed a large recurring-revenue book across complex customers.",
    }
    return [
        sentence
        for sentence in sentences
        if sentence not in ignored and len(sentence.split()) >= 6
    ]


def extract_job_requirement_groups(job_description: str, role_lane: str, sector: str) -> tuple[list[str], list[str]]:
    raw_lines = [re.sub(r"^[\-\*\u2022]+\s*", "", line).strip() for line in (job_description or "").splitlines()]
    lines: list[str] = []
    for line in raw_lines:
        if not line:
            continue
        if len(line) > 220:
            lines.extend(part.strip() for part in re.split(r"(?<=[.!?])\s+", line) if len(part.strip()) >= 18)
        else:
            lines.append(line)
    lines = [line for line in lines if len(line) >= 18]
    top: list[str] = []
    nice: list[str] = []
    target = top
    for line in lines:
        lowered = line.lower()
        if any(token in lowered for token in ["nice to have", "bonus", "preferred", "preferred qualifications"]):
            target = nice
            continue
        if any(token in lowered for token in ["responsibilities", "requirements", "qualifications", "about the role", "what you'll do"]):
            continue
        if len(target) < 5:
            target.append(line.rstrip("."))
        if len(top) >= 5 and len(nice) >= 3:
            break
    if not top:
        top = cover_letter_fit_points(role_lane, sector, [], job_description)[:5]
    return unique_keep_order(top)[:5], unique_keep_order(nice)[:3]


def infer_location_remote_status(job_description: str, concerns: list[str]) -> str:
    context = " ".join([job_description, *concerns]).lower()
    if "remote" in context:
        return "Remote or remote-leaning"
    if "hybrid" in context:
        return "Hybrid"
    if "onsite" in context or "on-site" in context or "in office" in context:
        return "Onsite"
    city_match = re.search(r"\b(new york|miami|san francisco|london|singapore|dubai)\b", context)
    if city_match:
        return city_match.group(1).title()
    return "Not clearly stated"


def infer_product_category(job_description: str, sector: str) -> str:
    context = job_description.lower()
    has_crypto_context = any(
        token in context
        for token in ["crypto", "web3", "blockchain", "protocol", "onchain", "stablecoin", "wallet", "defi"]
    )
    if "ramp" in context or any(token in context for token in ["corporate card", "company spend", "spend management", "finance automation"]):
        return "Finance automation / corporate card / spend management platform"
    if "layerzero" in context or "interoperability" in context:
        return "Web3 interoperability infrastructure"
    if "blockdaemon" in context or any(
        token in context
        for token in ["blockchain infrastructure", "staking", "nodes", "mpc", "node infrastructure"]
    ):
        return "Blockchain infrastructure / staking / node operations platform"
    if "flashpoint" in context or any(token in context for token in ["threat intelligence", "vulnerability intelligence", "geopolitical risk", "brand protection"]):
        return "Threat intelligence / cybersecurity risk platform"
    if "15five" in context or any(token in context for token in ["performance management", "engagement surveys", "performance reviews", "manager coaching"]):
        return "AI-powered performance management / HR SaaS platform"
    if "handshake" in context:
        return "Enterprise SaaS / AI-enabled talent and career network"
    if "dataiku" in context or any(
        token in context
        for token in ["machine learning", "data science", "ai orchestration", "ai agents", "enterprise ai"]
    ):
        return "Enterprise AI / analytics / machine learning / data science platform"
    if any(token in context for token in ["stablecoin", "usdt"]) or ("payments" in context and has_crypto_context):
        return "Stablecoin and payments infrastructure"
    if any(token in context for token in ["wallet", "custody", "key management"]):
        return "Wallet, custody, or digital-asset infrastructure"
    if any(token in context for token in ["analytics", "research", "data platform", "dashboard", "onchain data"]) and has_crypto_context:
        return "Crypto data or analytics platform"
    if any(token in context for token in ["analytics", "data platform", "business intelligence"]):
        return "Enterprise analytics or data platform"
    if any(token in context for token in ["aml", "risk", "compliance", "fraud", "kyc"]):
        return "Compliance or risk platform"
    if sector == "Web3 / Crypto":
        return "Crypto infrastructure or platform product"
    if sector == "AI / Data Platform":
        return "Enterprise AI, data, or analytics platform"
    return "B2B software or platform product"


def infer_buyer_persona(job_description: str, sector: str) -> str:
    context = job_description.lower()
    if "ramp" in context or any(token in context for token in ["finance teams", "company spend", "spend management"]):
        return "Mid-market businesses, finance teams, executives, and spend-management stakeholders"
    if "layerzero" in context or "interoperability" in context:
        return "Chains, protocols, and Web3 infrastructure customers"
    if "blockdaemon" in context or any(token in context for token in ["staking", "nodes", "mpc"]):
        return "Exchanges, custodians, crypto platforms, financial institutions, and developer teams"
    if "flashpoint" in context or any(token in context for token in ["threat intelligence", "vulnerability intelligence"]):
        return "Commercial enterprise and government security/risk teams"
    if "15five" in context or any(token in context for token in ["performance management", "engagement surveys", "performance reviews"]):
        return "People leaders, HR teams, executives, and performance-management stakeholders"
    if "handshake" in context:
        return "Enterprise customers and strategic accounts in the talent/career network market"
    if "dataiku" in context or any(
        token in context
        for token in ["financial services", "insurance", "pharmaceuticals", "transportation", "manufacturing"]
    ):
        return "Large enterprise customers across industries"
    if any(token in context for token in ["protocol", "foundation", "dao"]):
        return "Protocol, foundation, or ecosystem teams"
    if any(token in context for token in ["wallet", "exchange", "fintech", "payment"]):
        return "Wallet, exchange, fintech, or payments partners"
    if any(token in context for token in ["enterprise", "institutional", "bank", "hedge fund"]):
        return "Enterprise or institutional customers"
    if any(token in context for token in ["developer", "engineering", "product team"]):
        return "Technical product or developer-facing teams"
    if sector == "Web3 / Crypto":
        return "Crypto-native customers, partners, or institutions"
    return "B2B customers and stakeholders"


def infer_business_model(job_description: str) -> str:
    context = job_description.lower()
    if any(token in context for token in ["ramp", "corporate card", "company spend", "spend management"]):
        return "Fintech / spend-management platform"
    if any(token in context for token in ["15five", "handshake", "performance management", "enterprise software", "saas"]):
        return "Enterprise software / SaaS"
    if "layerzero" in context or "interoperability" in context:
        return "Web3 infrastructure / platform ecosystem"
    if "blockdaemon" in context or any(
        token in context
        for token in ["blockchain infrastructure", "staking", "nodes", "mpc", "api"]
    ):
        return "Web3 infrastructure / platform and usage-based services"
    if "flashpoint" in context or "threat intelligence" in context:
        return "Enterprise intelligence / cybersecurity software and services"
    if "dataiku" in context or any(token in context for token in ["enterprise software", "saas"]):
        return "Enterprise software / SaaS"
    if any(token in context for token in ["subscription", "renewal", "seat", "arr", "nrr"]):
        return "Recurring revenue / subscription"
    if any(token in context for token in ["api", "usage-based", "platform fees"]):
        return "Platform or usage-based revenue"
    if any(token in context for token in ["payments", "stablecoin", "settlement"]):
        return "Infrastructure and transaction-led revenue"
    return "Not clearly stated"


def requirement_match_row(requirement: str, role_lane: str, sector: str) -> tuple[str, str]:
    lowered = requirement.lower()
    if any(token in lowered for token in ["pricing", "scope", "terms", "commercial structuring"]):
        return (
            "Recent account role: influenced meaningful revenue outcomes across strategic accounts; Earlier enterprise account role: managed pricing and procurement conversations across a large recurring-revenue book",
            "Position as commercially disciplined renewal and expansion support, especially around pricing, scope, terms, and stakeholder alignment.",
        )
    if any(token in lowered for token in ["multi-threaded", "business, engineering", "operations", "executive", "executives"]):
        return (
            "Earlier partnerships role: worked across technical ecosystems and integration stakeholders; Recent account role: coordinated client, product, and sales stakeholders",
            "Emphasize technical-commercial translation and stakeholder alignment across business, technical, and executive audiences.",
        )
    if any(token in lowered for token in ["institutional", "global banks", "asset managers", "financial institutions", "market infrastructure"]):
        return (
            "Earlier enterprise account role: managed complex accounts at meaningful recurring-revenue scale; Recent account role: managed strategic technical-account relationships",
            "Use earlier enterprise-account work for long-cycle institutional discipline and recent account work for strategic technical-account ownership.",
        )
    if any(token in lowered for token in ["account plan", "account planning", "account plans"]):
        return (
            "Recent account role: created structure around renewal timing, account health, risk, ownership, next steps, and expansion opportunities",
            "Emphasize account planning, risk tracking, and structured follow-through instead of ad hoc relationship management.",
        )
    if any(token in lowered for token in ["data science", "ai", "machine learning", "technical stakeholder"]):
        return (
            "Recent account work and earlier partnerships work: worked in technical markets with data, research, protocols, integrations, and product feedback",
            "Position as technical-commercial translation, not engineering or data science ownership.",
        )
    if any(token in lowered for token in ["adoption", "value realization", "business use case", "customer outcome"]):
        return (
            "Recent account role: coordinated client needs across research, product, sales, and delivery stakeholders",
            "Frame this as translating customer goals into delivered value, stronger adoption, and measurable business outcomes.",
        )
    if any(token in lowered for token in ["account health", "churn", "downsell", "customer health", "account risk"]):
        return (
            "Recent account role: built renewal, account-health, and risk visibility across fragmented systems",
            "Show ability to identify risk early, create operating structure, and keep renewals or expansion from drifting.",
        )
    if any(token in lowered for token in ["renewal", "retention", "expansion", "account growth", "upsell", "nrr"]):
        return (
            "Recent account role: influenced meaningful revenue outcomes across strategic accounts",
            "Use this as proof of revenue ownership inside existing strategic relationships.",
        )
    if any(token in lowered for token in ["wallet", "exchange", "protocol", "partnership", "ecosystem", "integration"]):
        return (
            "Earlier partnerships role: closed multiple ecosystem partnerships across technical stakeholders",
            "Position this as direct crypto-native partner and ecosystem execution.",
        )
    if any(token in lowered for token in ["enterprise", "institutional", "procurement", "public sector", "complex stakeholder"]):
        return (
            "Earlier enterprise account role: managed a large recurring-revenue book across complex institutions",
            "Use this to show long-cycle account ownership and credibility with complex institutions.",
        )
    if any(token in lowered for token in ["ambiguous", "process", "build", "structure", "cross-functional"]):
        return (
            "Earlier enterprise account work and recent account work: operated in messy environments and improved account visibility and process",
            "Frame this as practical operating judgment, not just polished process adherence.",
        )
    if any(token in lowered for token in ["customer", "account", "stakeholder", "relationship"]):
        return (
            "Recent account work and earlier enterprise account work: managed strategic accounts across renewals, stakeholder alignment, and delivery",
            "Tie this to steady relationship ownership linked directly to revenue outcomes.",
        )
    if sector == "Web3 / Crypto":
        return (
            "Earlier partnerships work and recent account work: relationship ownership across technical ecosystems and strategic clients",
            "Connect it to technical products, customer trust, and commercial follow-through.",
        )
    return (
        "recent account management, earlier partnerships work, and earlier enterprise account work together show durable account and partnership ownership",
        "Keep the positioning focused on relationship-driven revenue and execution discipline.",
    )


def cover_letter_anti_copy_check(
    voice: dict[str, Any],
    *,
    company: str,
    cover_letter_text: str,
) -> dict[str, Any]:
    candidate_sentences = normalized_sentences(cover_letter_text)
    candidate_meaningful_sentences = anti_copy_meaningful_sentences(candidate_sentences)
    first_sentence = candidate_sentences[0] if candidate_sentences else ""
    generic_phrases = [
        "i am excited to apply",
        "i was thrilled to see",
        "i believe my background aligns",
        "unique opportunity",
        "perfect fit",
        "passionate about",
        "leverage my experience",
        "dynamic team",
        "fast paced environment",
        "world class team",
        "what stands out to me",
        "the fit feels practical from my side",
        "commercial thread",
        "stakeholder trust matter in equal measure",
        "sensible expansion opportunities",
        "especially compelling to me",
        "i would welcome the opportunity to learn how your team defines success",
        "the overlap is straightforward",
        "this role lines up well",
        "managed important relationships over time",
        "creating room for commercial growth",
        "commercially useful execution",
        "commercially useful expansion opportunities",
        "platform products",
        "the commercial goal",
    ]
    generic_hits = [phrase for phrase in generic_phrases if phrase in normalize(cover_letter_text)]
    sample_letters = voice.get("sample_cover_letters") if isinstance(voice.get("sample_cover_letters"), dict) else {}
    reused_opening = False
    repeated_starter = False
    shared_sentence = False
    sample_notes: list[str] = []
    first_starter = " ".join(first_sentence.split()[:5])
    for sample_name, sample_text in sample_letters.items():
        sample_sentences = normalized_sentences(str(sample_text))
        sample_meaningful_sentences = anti_copy_meaningful_sentences(sample_sentences)
        if not sample_sentences:
            continue
        similarity = SequenceMatcher(None, first_sentence, sample_sentences[0]).ratio()
        if similarity >= 0.86:
            reused_opening = True
            sample_notes.append(f"opening too close to {sample_name}")
        sample_starter = " ".join(sample_sentences[0].split()[:5])
        if first_starter and sample_starter and first_starter == sample_starter:
            repeated_starter = True
            sample_notes.append(f"opening starter matches {sample_name}")
        if set(candidate_meaningful_sentences) & set(sample_meaningful_sentences):
            shared_sentence = True
            sample_notes.append(f"shared sentence detected with {sample_name}")
    text_normalized = normalize(cover_letter_text).lower()
    text_compact = re.sub(r"\W+", "", text_normalized)
    company_normalized = normalize(company).lower()
    company_specific = company_normalized in text_normalized or normalized_company_key(company) in text_compact
    passed = not reused_opening and not repeated_starter and not shared_sentence and not generic_hits and company_specific
    notes: list[str] = []
    if not company_specific:
        notes.append("Company name did not appear clearly enough in the letter.")
    if generic_hits:
        notes.append(f"Generic phrasing found: {', '.join(generic_hits)}.")
    notes.extend(sample_notes[:3])
    if passed:
        notes.append("No obvious opening reuse, generic filler, or direct sentence carryover was detected.")
    return {
        "passed": passed,
        "notes": notes,
    }


DOCX_CONTENT_TYPES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""

DOCX_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""

DOCX_DOCUMENT_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""

DOCX_STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
</w:styles>
"""

DOCX_APP_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Codex Job Search Assistant</Application>
</Properties>
"""


def docx_core_xml(title: str) -> str:
    created = dt.datetime.now(dt.timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    safe_title = escape(title)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>{safe_title}</dc:title>
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{created}</dcterms:modified>
</cp:coreProperties>
"""


def docx_paragraph_xml(text: str) -> str:
    runs: list[str] = []
    for idx, line in enumerate(text.split("\n")):
        if idx:
            runs.append("<w:r><w:br/></w:r>")
        if line:
            attrs = ' xml:space="preserve"' if line != line.strip() else ""
            runs.append(f"<w:r><w:t{attrs}>{escape(line)}</w:t></w:r>")
    if not runs:
        runs.append("<w:r/>")
    return f"<w:p>{''.join(runs)}</w:p>"


def docx_document_xml(paragraphs: list[str]) -> str:
    body = "".join(docx_paragraph_xml(paragraph) for paragraph in paragraphs)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {body}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
"""


def write_simple_docx(output_path: Path, title: str, body_text: str) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    paragraphs = [paragraph for paragraph in body_text.replace("\r\n", "\n").split("\n\n")]
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", DOCX_CONTENT_TYPES_XML)
        archive.writestr("_rels/.rels", DOCX_RELS_XML)
        archive.writestr("docProps/core.xml", docx_core_xml(title))
        archive.writestr("docProps/app.xml", DOCX_APP_XML)
        archive.writestr("word/document.xml", docx_document_xml(paragraphs))
        archive.writestr("word/styles.xml", DOCX_STYLES_XML)
        archive.writestr("word/_rels/document.xml.rels", DOCX_DOCUMENT_RELS_XML)
    return output_path


def build_application_packet_artifacts(db_path: Path, job_id: int) -> dict[str, Any]:
    ensure_schema(db_path)
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        row = require_job(conn, job_id)
    job = dict(row)
    profile = load_profile(DEFAULT_PROFILE)
    voice = load_voice_profile()
    question_override = application_question_override_for_job(job)
    metadata = dashboard_metadata_for_job(job)
    signals = json_list(job["matched_signals"])
    concerns = json_list(job["concerns"])
    bullets = json_list(job["resume_bullet_adjustments"])
    packet_context = (
        question_override.get("packet_context")
        if isinstance(question_override.get("packet_context"), dict)
        else {}
    )
    company = job["company"] or "the company"
    title = job["title"] or "the role"
    sector = metadata["sector"]
    role_lane = metadata["role_lane"]
    priority = metadata["priority"]
    role_focus = role_lane_phrase(role_lane)
    sector_focus = sector_phrase(sector)
    packet_fit_score = int(question_override.get("score_override") or job["fit_score"])
    fit_score_10 = round((packet_fit_score / 10), 1)
    priority_label = "High" if packet_fit_score >= 85 or str(priority).startswith("P1") else "Medium" if packet_fit_score >= 70 or str(priority).startswith("P2") else "Low"
    concern_text = "\n".join(f"- {item}" for item in concerns[:5]) or "- No major concerns found."
    signal_text = "\n".join(f"- {item}" for item in signals[:8]) or "- No strong fit signals found."
    bullet_text = "\n".join(f"- {item}" for item in bullets[:6]) or "- Use the current strategic account / partnerships resume."
    top_requirements, nice_to_haves = extract_job_requirement_groups(str(job.get("job_description") or ""), role_lane, sector)
    location_status = infer_location_remote_status(str(job.get("job_description") or ""), concerns)
    product_category = infer_product_category(str(job.get("job_description") or ""), sector)
    buyer_persona = infer_buyer_persona(str(job.get("job_description") or ""), sector)
    business_model = infer_business_model(str(job.get("job_description") or ""))
    override_requirements = packet_context.get("role_requirements") if isinstance(packet_context.get("role_requirements"), list) else []
    if override_requirements:
        top_requirements = [str(item).strip() for item in override_requirements if str(item).strip()][:5]
    product_category = str(packet_context.get("product_market") or product_category).strip()
    buyer_persona = str(packet_context.get("buyer") or buyer_persona).strip()
    business_model = str(packet_context.get("business_model") or business_model).strip()
    anti_copy_result: dict[str, Any] | None = None
    cover_letter_text = ""
    provided_cover_letter_text = str(question_override.get("suggested_cover_letter_text") or "").strip()
    if provided_cover_letter_text:
        cover_letter_text = provided_cover_letter_text
        anti_copy_result = cover_letter_anti_copy_check(voice, company=company, cover_letter_text=cover_letter_text)
        anti_copy_result.setdefault("notes", []).append(
            "Used the provided cover letter text from application_question_overrides.json for this exact job."
        )
    else:
        for variant in range(3):
            candidate = packet_cover_letter_text(
                voice,
                company=company,
                title=title,
                role_lane=role_lane,
                sector=sector,
                sector_focus=sector_focus,
                job_id=job_id,
                signals=signals,
                job_description=str(job.get("job_description") or ""),
                variant=variant,
            )
            anti_copy_result = cover_letter_anti_copy_check(voice, company=company, cover_letter_text=candidate)
            cover_letter_text = candidate
            if anti_copy_result.get("passed"):
                break
    if provided_cover_letter_text and not (anti_copy_result or {}).get("passed"):
        anti_copy_result.setdefault("notes", []).append(
            "Preserved the explicit cover letter override for this exact job instead of replacing it with a generic fallback."
        )
    elif not (anti_copy_result or {}).get("passed"):
        cover_letter_text = packet_cover_letter_text_quality_gate(
            company=company,
            title=title,
            sector=sector,
            product_category=product_category,
            buyer_persona=buyer_persona,
        )
        anti_copy_result = cover_letter_anti_copy_check(voice, company=company, cover_letter_text=cover_letter_text)
        anti_copy_result.setdefault("notes", []).append(
            "Rewrote with the quality-gate fallback because an earlier cover letter did not pass anti-copy checks."
        )
    outreach_draft = str(question_override.get("suggested_outreach_text") or "").strip() or build_outreach(title, company, signals, concerns)
    guardrails = voice.get("authenticity_guardrails")
    guardrail_text = "\n".join(f"- {item}" for item in guardrails[:8]) if isinstance(guardrails, list) else ""
    proven_examples = voice_examples_for_job(
        voice,
        sector,
        role_lane,
        company=company,
        product_category=product_category,
        buyer_persona=buyer_persona,
    )
    proven_examples_markdown = "\n\n".join(f"### {heading}\n\n{example}" for heading, example in proven_examples)
    proven_answers_section = (
        f"""
## Proven Answer Angles

{proven_examples_markdown}
"""
        if proven_examples_markdown
        else ""
    )
    why_fit_lines = "\n".join(f"- {item}" for item in signals[:3]) or "- Relationship ownership tied to revenue fits the role.\n- The buyer and product category overlap with prior crypto and account work.\n- The role leans on cross-functional commercial execution."
    possible_concern_lines = "\n".join(f"- {item}" for item in concerns[:2]) or "- No major mismatch jumped out from the current record."
    if any("direct cybersecurity" in item.lower() for item in concerns):
        best_positioning_angle = (
            f"{company} looks strongest when positioned around {role_focus.lower()}, customer trust, and commercial follow-through. "
            "The best angle is that the candidate has operated in technical, risk-sensitive, and procurement-heavy markets where revenue protection, cross-functional coordination, and growth inside existing relationships all mattered at the same time."
        )
    else:
        best_positioning_angle = (
            f"{company} looks strongest when positioned around {role_focus.lower()}, customer trust, and commercial follow-through. "
            f"The best angle is that the candidate has operated in {sector_focus.lower()} environments where revenue protection, cross-functional coordination, and growth inside existing relationships all mattered at the same time."
        )
    provided_positioning = str(question_override.get("recommended_positioning") or "").strip().strip('"')
    if provided_positioning:
        best_positioning_angle = provided_positioning
    role_match_rows = []
    for requirement in top_requirements[:5]:
        proof_point, positioning = requirement_match_row(requirement, role_lane, sector)
        role_match_rows.append(f"| {requirement} | {proof_point} | {positioning} |")
    role_match_table = "\n".join(role_match_rows) or "| Relationship ownership | recent account management, earlier partnerships work, and earlier enterprise account work | Position as strategic relationship ownership tied to revenue. |"
    top_proof_points = [
        "Recent account role: influenced meaningful revenue outcomes across strategic accounts",
        "Earlier partnerships role: closed multiple partnerships across a technical ecosystem",
        "Earlier enterprise account role: managed a large recurring-revenue book across complex customers",
    ]
    override_proof_points = packet_context.get("top_proof_points") if isinstance(packet_context.get("top_proof_points"), list) else []
    if override_proof_points:
        top_proof_points = unique_keep_order([str(item).strip() for item in override_proof_points if str(item).strip()] + top_proof_points)[:3]
    overstatement_risks = [
        "Do not overstate technical architecture or engineering depth.",
        "Do not imply every crypto relationship directly converted to closed revenue unless verified.",
    ]
    if any("compliance" in item.lower() or "aml" in item.lower() for item in top_requirements):
        overstatement_risks.append("Do not position the candidate as an AML or investigations subject-matter expert.")
    interview_points = [
        "How recent account work combined renewal strategy, internal coordination, and commercial judgment.",
        "How earlier partnerships work maps to adoption, ecosystem growth, and partner-led momentum.",
        "How earlier enterprise account work built long-cycle discipline and comfort in messy operating environments.",
    ]
    recruiter_screen_summary = str(packet_context.get("recruiter_screen_summary") or "").strip()
    if not recruiter_screen_summary:
        if is_crypto_or_web3_sector(sector):
            recruiter_screen_summary = (
                f"I am a {role_lane.lower()} operator with long-cycle account-management experience and direct Web3 commercial experience. "
                "Most recently, I managed strategic relationships across renewals and expansion, and before that I led partnerships across a technical ecosystem. "
                "My strongest lane is owning strategic relationships and turning them into retention, growth, or customer momentum."
            )
        else:
            recruiter_screen_summary = (
                f"I am a {role_lane.lower()} operator with long-cycle account-management experience, large-portfolio revenue ownership, and recent strategic account work in technical markets. "
                "My strongest lane is owning customer relationships, creating renewal and expansion structure, and translating customer goals into business outcomes."
            )
    company_one_liner = str(packet_context.get("company_one_liner") or "").strip() or f"{company} is a {product_category.lower()} company serving {buyer_persona.lower()}."
    fit_analysis_md = f"""# Fit Analysis

## Company
{company}

## Role
{title}

## Product / Market
{product_category}

## Buyer
{buyer_persona}

## Why This Could Fit the Candidate
{why_fit_lines}

## Possible Concerns
{possible_concern_lines}

## Best Positioning Angle
{best_positioning_angle}

## Fit Score
{fit_score_10}/10

## Priority
{priority_label}
"""
    role_match_md = f"""# Role Match

| Job Requirement | Candidate Proof Point | How to Position It |
|---|---|---|
{role_match_table}
"""
    positioning_brief_md = f"""# Positioning Brief

## Main Angle
{best_positioning_angle}

## Top 3 Proof Points
1. {top_proof_points[0]}
2. {top_proof_points[1]}
3. {top_proof_points[2]}

## What to Avoid Overstating
- {overstatement_risks[0]}
- {overstatement_risks[1]}
{''.join(f'- {item}\n' for item in overstatement_risks[2:])}
## Interview Talking Points
- {interview_points[0]}
- {interview_points[1]}
- {interview_points[2]}

## Recruiter Screen Summary
{recruiter_screen_summary}
"""
    anti_copy_lines = "\n".join(f"- {item}" for item in (anti_copy_result or {}).get("notes", [])) or "- No anti-copy issues noted."
    cover_letter_md = f"""# Cover Letter

{cover_letter_text}
"""
    cover_letter_notes_md = f"""# Anti-Copy Check

## Passed?
{"Yes" if (anti_copy_result or {}).get("passed") else "No"}

## What Was Checked
- Prior cover letter similarity
- Generic phrasing
- Company specificity
- Reused sentence structure
- Overstated claims

## Notes
{anti_copy_lines}

## Cover Letter Version Notes
- Local DOCX path: __COVER_LETTER_DOCX_PATH__
- Keep the final letter specific to {company}, not just the broader category.
- Use only the two or three strongest proof points for submission.
"""
    interview_prep_md = f"""# Recruiter Prep

## 30-Second Pitch
{recruiter_screen_summary}

## 15-Second Pitch
Iâ€™m a {sector_focus.lower()} account-management and partnerships operator with strong experience in renewals, expansion, and strategic relationship ownership.

## Company One-Liners
- Company: {company_one_liner}

## Best Roles to Prioritize
1. {company} - {title}
2. Strategic account management roles in {sector_focus.lower()}
3. Partnerships or commercial roles with similar buyers and product complexity

## Questions to Ask Recruiter
- How is success measured in the first six to twelve months?
- What kind of buyer or customer relationship owns the most weight in this role?
- Where does the team need the most commercial structure right now?

## Concerns to Clarify
- {concerns[0] if concerns else "Clarify location, buyer, and revenue ownership expectations."}
- {concerns[1] if len(concerns) > 1 else "Confirm whether the role is more relationship ownership or net-new hunting."}

## Compensation Positioning
Iâ€™m flexible depending on role scope, base, OTE, equity, and stage of company. For the right Senior Account Manager, Strategic Partnerships, CSM, or GTM role, Iâ€™d ideally like to be in the $120K+ base range, with upside tied to revenue, expansion, or company growth. But Iâ€™m open to the full package if the company and role are strong.

## Recent Transition Explanation
Customize this section with the candidate's real transition context. Keep it factual, brief, and non-defensive.
"""
    raw_application_questions = question_override.get("questions", []) if isinstance(question_override.get("questions"), list) else []
    application_questions = enrich_application_questions_for_packet(raw_application_questions, voice=voice)
    application_answers_md = render_application_questions_md(application_questions)
    question_status = application_questions_status_payload(question_override)
    ats_section_md = ats_risk_assessment_md(
        profile=profile,
        job_description=str(job.get("job_description") or ""),
        signals=signals,
        bullets=bullets,
        override=question_override,
    )
    score_reason = str(question_override.get("score_reason") or "").strip()
    if not score_reason:
        score_reason = (
            "Strong overlap on role lane, buyer complexity, and commercially relevant proof points, "
            "with any domain, location, or seniority gaps handled in the risk notes."
        )
    next_action = str(job.get("next_action") or "").strip() or ("Apply" if priority_label == "High" else "Needs review")
    bundle_dir_name = packet_bundle_directory_name(job_id, company, title)
    files = {
        "01_fit_analysis.md": fit_analysis_md.strip() + "\n",
        "02_role_match.md": role_match_md.strip() + "\n",
        "03_positioning_brief.md": positioning_brief_md.strip() + "\n",
        "04_cover_letter.md": cover_letter_md.strip() + "\n",
        "05_cover_letter_notes.md": cover_letter_notes_md.strip() + "\n",
        "06_interview_prep.md": interview_prep_md.strip() + "\n",
        "07_application_answers.md": application_answers_md.strip() + "\n",
    }
    files_created_lines = "\n".join(f"- {name}" for name in files)
    resume_bullets_md = "\n".join(f"- {item}" for item in bullets[:5]) or "- Use the current strategic account, renewals, and expansion proof points."
    tracker_yaml = textwrap.dedent(
        f"""\
        company: {company}
        role: {title}
        score: {packet_fit_score}
        priority: {priority_label.lower().replace(' ', '_')}
        action: {next_action}
        positive_signals:
          - {signals[0] if signals else "strategic relationship ownership"}
          - {signals[1] if len(signals) > 1 else "renewal and expansion overlap"}
          - {signals[2] if len(signals) > 2 else "commercial account ownership"}
        risk_signals:
          - {concerns[0] if concerns else "No major risk signal captured."}
          - {concerns[1] if len(concerns) > 1 else "No second major risk signal captured."}
        next_action: {next_action}
        application_url: {job['application_url'] or job['source'] or ''}
        """
    ).strip()
    packet_markdown = f"""# Application Packet: {company} - {title}

## 1. Fit Score + Recommendation
- Score: {fit_score_10}/10
- Apply / Skip / Warm Intro / Monitor: {next_action}
- Priority: {priority_label}
- Reason: {score_reason}

{question_status['body']}

{ats_section_md}

## 4. Best Positioning Angle
{best_positioning_angle}

## 5. Resume Bullets to Emphasize
{resume_bullets_md}

{application_answers_md}

## 7. Cover Letter / Short Note
{cover_letter_text}

## 8. Outreach Message
{outreach_draft}

## 9. Codex Tracker Update
```yaml
{tracker_yaml}
```

## Packet Assets
- Files Created:
{files_created_lines}
- Google Doc: Pending Google sync or existing mirrored doc link in packet index.
- Location / Remote: {location_status}
- Product / Market: {product_category}
- Buyer: {buyer_persona}
- Business Model: {business_model}

---

{fit_analysis_md}

---

{role_match_md}

---

{positioning_brief_md}

---

{cover_letter_notes_md}

---

{interview_prep_md}
{proven_answers_section}

## Voice Guardrails

{guardrail_text or "- Keep the tone direct, proof-backed, and conversational.\n- Avoid sounding overly polished, overly grateful, or inflated."}

## Submission Checklist

- Confirm work authorization and location answers manually.
- Confirm compensation expectations manually.
- Attach the right resume version.
- Review any voluntary demographic questions manually.
- Do not submit until the final page has been reviewed.
"""
    return {
        "bundle_dir_name": bundle_dir_name,
        "packet_markdown": packet_markdown.strip() + "\n",
        "files": files,
        "cover_letter_text": cover_letter_text,
    }


def build_application_packet(db_path: Path, job_id: int) -> tuple[str, str]:
    artifacts = build_application_packet_artifacts(db_path, job_id)
    return "00_packet_bundle.md", artifacts["packet_markdown"]


def write_application_packet(db_path: Path, job_id: int, output_path: Path | None) -> Path:
    artifacts = build_application_packet_artifacts(db_path, job_id)
    if output_path:
        path = output_path
        bundle_dir = output_path.parent
    else:
        bundle_dir = ROOT / "job_packets" / artifacts["bundle_dir_name"]
        path = bundle_dir / "00_packet_bundle.md"
    cover_letter_docx_path = cover_letter_docx_path_for_packet_path(path)
    bundle_dir.mkdir(parents=True, exist_ok=True)
    for filename, content in artifacts["files"].items():
        file_path = bundle_dir / filename
        file_text = content.replace("__COVER_LETTER_DOCX_PATH__", str(cover_letter_docx_path.resolve()))
        file_path.write_text(file_text, encoding="utf-8")
    markdown = artifacts["packet_markdown"].replace("__COVER_LETTER_DOCX_PATH__", str(cover_letter_docx_path.resolve()))
    path.write_text(markdown, encoding="utf-8")
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        job = dict(require_job(conn, job_id))
    write_simple_docx(
        cover_letter_docx_path,
        f"Cover Letter - {job.get('company') or 'Company'} - {job.get('title') or 'Role'}",
        artifacts["cover_letter_text"],
    )
    update_packet_index_for_job(job, path, cover_letter_docx_path=cover_letter_docx_path)
    return path


def packet_google_doc_title(job: dict[str, Any]) -> str:
    return f"Job {job['id']} - {job.get('company') or 'Company'} - {job.get('title') or 'Role'} - Application Packet"


def cover_letter_google_doc_title(job: dict[str, Any]) -> str:
    return f"Job {job['id']} - {job.get('company') or 'Company'} - {job.get('title') or 'Role'} - Cover Letter"


def ensure_local_packet_path(db_path: Path, job_id: int) -> Path:
    discovered = find_local_packet_path(job_id)
    if discovered is not None:
        return discovered
    return write_application_packet(db_path, job_id, None)


def ensure_local_cover_letter_docx(db_path: Path, job: dict[str, Any], packet_path: Path) -> Path:
    cover_letter_path = cover_letter_docx_path_for_packet_path(packet_path)
    if cover_letter_path.exists():
        return cover_letter_path
    cover_letter_text = build_application_packet_artifacts(db_path, int(job["id"]))["cover_letter_text"]
    write_simple_docx(
        cover_letter_path,
        f"Cover Letter - {job.get('company') or 'Company'} - {job.get('title') or 'Role'}",
        cover_letter_text,
    )
    update_packet_index_for_job(job, packet_path, cover_letter_docx_path=cover_letter_path)
    return cover_letter_path


def can_attempt_automatic_google_sync() -> bool:
    oauth_ready = DEFAULT_GOOGLE_OAUTH_CLIENT_JSON.exists() and DEFAULT_GOOGLE_OAUTH_TOKEN_JSON.exists()
    service_account_ready = DEFAULT_GOOGLE_SERVICE_ACCOUNT_JSON.exists()
    return oauth_ready or service_account_ready


def sync_google_drive_docs(
    db_path: Path,
    *,
    job_ids: list[int] | None,
    auth_mode: str,
    service_account_json: Path,
    oauth_client_json: Path,
    oauth_token_json: Path,
    packet_folder_url: str,
    cover_letter_folder_url: str,
) -> list[dict[str, str]]:
    packet_folder_id = extract_drive_id(packet_folder_url)
    cover_letter_folder_id = extract_drive_id(cover_letter_folder_url)
    results: list[dict[str, str]] = []
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        selected_jobs: list[dict[str, Any]] = []
        if job_ids:
            for job_id in job_ids:
                selected_jobs.append(dict(require_job(conn, job_id)))
        else:
            for row in load_dashboard_jobs(db_path):
                packet_path = find_local_packet_path(int(row["id"]))
                if packet_path is not None:
                    selected_jobs.append(row)

    candidate_paths = google_credentials_candidates(
        auth_mode,
        service_account_json=service_account_json,
        oauth_client_json=oauth_client_json,
        prefer="oauth",
    )
    attempts: list[str] = []
    drive = docs = None
    credentials_type = ""
    for candidate_path in candidate_paths:
        try:
            drive, docs, credentials_type = build_google_services(
                candidate_path,
                token_json=oauth_token_json,
            )
            break
        except (FileNotFoundError, RuntimeError, ValueError) as exc:
            attempts.append(f"{candidate_path.name}: {exc}")
    if drive is None or docs is None:
        raise SystemExit("Failed to initialize Google Drive sync credentials. " + " | ".join(attempts))

    for job in selected_jobs:
        job_id = int(job["id"])
        packet_path = ensure_local_packet_path(db_path, job_id)
        packet_text = packet_path.read_text(encoding="utf-8")
        try:
            packet_doc = upsert_google_doc_text(
                drive,
                docs,
                folder_id=packet_folder_id,
                title=packet_google_doc_title(job),
                text=packet_text,
            )
        except Exception as exc:
            if credentials_type == "service_account" and is_storage_quota_exceeded_error(exc):
                raise SystemExit(
                    "Google rejected file creation because service accounts cannot own files in My Drive. "
                    "Use OAuth desktop credentials instead: save an OAuth client JSON at "
                    f"{DEFAULT_GOOGLE_OAUTH_CLIENT_JSON} and rerun "
                    "`python .\\job_search_assistant.py sync-google-drive-docs --auth-mode oauth --id "
                    f"{job_id}`."
                ) from exc
            raise
        set_packet_doc_url(job_id, packet_doc["webViewLink"])

        cover_letter_needed = dashboard_metadata_for_job(job)["cover_letter_needed"]
        cover_letter_url = ""
        if cover_letter_needed != "No":
            cover_letter_path = ensure_local_cover_letter_docx(db_path, job, packet_path)
            cover_letter_md_path = packet_path.with_name("04_cover_letter.md")
            if cover_letter_md_path.exists():
                cover_letter_text = re.sub(
                    r"^# Cover Letter\s*",
                    "",
                    cover_letter_md_path.read_text(encoding="utf-8").strip(),
                ).strip()
            else:
                cover_letter_text = build_application_packet_artifacts(db_path, job_id)["cover_letter_text"]
            try:
                cover_letter_doc = upsert_google_doc_text(
                    drive,
                    docs,
                    folder_id=cover_letter_folder_id,
                    title=cover_letter_google_doc_title(job),
                    text=cover_letter_text,
                )
            except Exception as exc:
                if credentials_type == "service_account" and is_storage_quota_exceeded_error(exc):
                    raise SystemExit(
                        "Google rejected cover-letter creation because service accounts cannot own files in My Drive. "
                        "Use OAuth desktop credentials instead: save an OAuth client JSON at "
                        f"{DEFAULT_GOOGLE_OAUTH_CLIENT_JSON} and rerun "
                        "`python .\\job_search_assistant.py sync-google-drive-docs --auth-mode oauth --id "
                        f"{job_id}`."
                    ) from exc
                raise
            cover_letter_url = cover_letter_doc["webViewLink"]
            set_cover_letter_doc_url(job_id, cover_letter_url)
            update_packet_index_for_job(job, packet_path, cover_letter_docx_path=cover_letter_path)

        results.append(
            {
                "job_id": str(job_id),
                "company": str(job.get("company") or ""),
                "title": str(job.get("title") or ""),
                "packet_url": str(packet_doc.get("webViewLink") or ""),
                "cover_letter_url": cover_letter_url,
            }
        )
    return results


def update_job_status(
    db_path: Path,
    job_id: int,
    status: str,
    *,
    next_action: str | None = None,
    follow_up_date: str | None = None,
    decision: str | None = None,
    archived_reason: str | None = None,
    note: str | None = None,
    sector: str | None = None,
    role_lane: str | None = None,
    priority: str | None = None,
    application_url: str | None = None,
    resume_version: str | None = None,
    cover_letter_needed: str | None = None,
    referral_target: str | None = None,
) -> None:
    ensure_schema(db_path)
    clean_status = normalize_status(status)
    updated_at = now_utc()
    with sqlite3.connect(db_path) as conn:
        require_job(conn, job_id)
        conn.execute(
            """
            UPDATE job_evaluations
            SET status = ?,
                next_action = COALESCE(?, next_action),
                follow_up_date = COALESCE(?, follow_up_date),
                decision = COALESCE(?, decision),
                archived_reason = COALESCE(?, archived_reason),
                sector = COALESCE(?, sector),
                role_lane = COALESCE(?, role_lane),
                priority = COALESCE(?, priority),
                application_url = COALESCE(?, application_url),
                resume_version = COALESCE(?, resume_version),
                cover_letter_needed = COALESCE(?, cover_letter_needed),
                referral_target = COALESCE(?, referral_target),
                last_updated_at = ?
            WHERE id = ?
            """,
            (
                clean_status,
                next_action,
                follow_up_date,
                decision,
                archived_reason,
                sector,
                role_lane,
                priority,
                application_url,
                resume_version,
                cover_letter_needed,
                referral_target,
                updated_at,
                job_id,
            ),
        )
        if note:
            conn.execute(
                "INSERT INTO notes (job_id, created_at, note) VALUES (?, ?, ?)",
                (job_id, updated_at, note),
            )


def update_job_details(
    db_path: Path,
    job_id: int,
    *,
    next_action: str | None = None,
    follow_up_date: str | None = None,
    decision: str | None = None,
    sector: str | None = None,
    role_lane: str | None = None,
    priority: str | None = None,
    application_url: str | None = None,
    resume_version: str | None = None,
    cover_letter_needed: str | None = None,
    referral_target: str | None = None,
    note: str | None = None,
) -> None:
    ensure_schema(db_path)
    updates = {
        "next_action": next_action,
        "follow_up_date": follow_up_date,
        "decision": decision,
        "sector": sector,
        "role_lane": role_lane,
        "priority": priority,
        "application_url": application_url,
        "resume_version": resume_version,
        "cover_letter_needed": cover_letter_needed,
        "referral_target": referral_target,
    }
    selected = {key: value for key, value in updates.items() if value is not None}
    if not selected and not note:
        raise SystemExit("Provide at least one field to update.")
    updated_at = now_utc()
    with sqlite3.connect(db_path) as conn:
        require_job(conn, job_id)
        if selected:
            assignments = ", ".join(f"{key} = ?" for key in selected)
            conn.execute(
                f"UPDATE job_evaluations SET {assignments}, last_updated_at = ? WHERE id = ?",
                [*selected.values(), updated_at, job_id],
            )
        else:
            conn.execute("UPDATE job_evaluations SET last_updated_at = ? WHERE id = ?", (updated_at, job_id))
        if note:
            conn.execute(
                "INSERT INTO notes (job_id, created_at, note) VALUES (?, ?, ?)",
                (job_id, updated_at, note),
            )


def require_target_company(conn: sqlite3.Connection, target_company_id: int) -> sqlite3.Row:
    conn.row_factory = sqlite3.Row
    row = conn.execute("SELECT * FROM target_companies WHERE id = ?", (target_company_id,)).fetchone()
    if row is None:
        raise SystemExit(f"No target company found with id {target_company_id}.")
    return row


def add_target_company(
    db_path: Path,
    *,
    company: str,
    website: str | None,
    lane: str | None,
    description: str | None,
    funding_date: str | None,
    funding_amount: str | None,
    round_name: str | None,
    investors: str | None,
    company_fit_score: int | None,
    open_roles_found: str | None,
    best_role_title: str | None,
    role_fit_score: int | None,
    role_url: str | None,
    priority: str | None,
    target_strategy: str | None,
    outreach_type: str | None,
    warm_contact_1: str | None,
    warm_contact_1_title: str | None,
    warm_contact_1_linkedin: str | None,
    warm_contact_2: str | None,
    warm_contact_2_title: str | None,
    warm_contact_2_linkedin: str | None,
    outreach_angle: str | None,
    outreach_status: str | None,
    application_status: str | None,
    notes: str | None,
    next_action: str | None,
    last_checked: str | None,
    source_url: str | None,
) -> int:
    ensure_schema(db_path)
    created_at = now_utc()
    clean_company_fit = clamp_score(company_fit_score)
    clean_role_fit = clamp_score(role_fit_score)
    resolved_priority = priority or company_priority_from_score(clean_company_fit)
    with sqlite3.connect(db_path) as conn:
        cursor = conn.execute(
            """
            INSERT INTO target_companies (
                created_at, company, website, lane, description, funding_date,
                funding_amount, round, investors, company_fit_score, open_roles_found,
                best_role_title, role_fit_score, role_url, priority, target_strategy,
                outreach_type, warm_contact_1,
                warm_contact_1_title, warm_contact_1_linkedin, warm_contact_2,
                warm_contact_2_title, warm_contact_2_linkedin, outreach_angle,
                outreach_status, application_status, notes, next_action, last_checked,
                source_url, last_updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                created_at,
                company,
                website,
                lane,
                description,
                funding_date,
                funding_amount,
                round_name,
                investors,
                clean_company_fit,
                open_roles_found,
                best_role_title,
                clean_role_fit,
                role_url,
                resolved_priority,
                target_strategy,
                outreach_type,
                warm_contact_1,
                warm_contact_1_title,
                warm_contact_1_linkedin,
                warm_contact_2,
                warm_contact_2_title,
                warm_contact_2_linkedin,
                outreach_angle,
                outreach_status,
                application_status,
                notes,
                next_action,
                last_checked,
                source_url,
                created_at,
            ),
        )
        return int(cursor.lastrowid)


def update_target_company(
    db_path: Path,
    target_company_id: int,
    *,
    company: str | None = None,
    website: str | None = None,
    lane: str | None = None,
    description: str | None = None,
    funding_date: str | None = None,
    funding_amount: str | None = None,
    round_name: str | None = None,
    investors: str | None = None,
    company_fit_score: int | None = None,
    open_roles_found: str | None = None,
    best_role_title: str | None = None,
    role_fit_score: int | None = None,
    role_url: str | None = None,
    priority: str | None = None,
    target_strategy: str | None = None,
    outreach_type: str | None = None,
    warm_contact_1: str | None = None,
    warm_contact_1_title: str | None = None,
    warm_contact_1_linkedin: str | None = None,
    warm_contact_2: str | None = None,
    warm_contact_2_title: str | None = None,
    warm_contact_2_linkedin: str | None = None,
    outreach_angle: str | None = None,
    outreach_status: str | None = None,
    application_status: str | None = None,
    notes: str | None = None,
    next_action: str | None = None,
    last_checked: str | None = None,
    source_url: str | None = None,
) -> None:
    ensure_schema(db_path)
    updates = {
        "company": company,
        "website": website,
        "lane": lane,
        "description": description,
        "funding_date": funding_date,
        "funding_amount": funding_amount,
        "round": round_name,
        "investors": investors,
        "company_fit_score": clamp_score(company_fit_score),
        "open_roles_found": open_roles_found,
        "best_role_title": best_role_title,
        "role_fit_score": clamp_score(role_fit_score),
        "role_url": role_url,
        "priority": priority,
        "target_strategy": target_strategy,
        "outreach_type": outreach_type,
        "warm_contact_1": warm_contact_1,
        "warm_contact_1_title": warm_contact_1_title,
        "warm_contact_1_linkedin": warm_contact_1_linkedin,
        "warm_contact_2": warm_contact_2,
        "warm_contact_2_title": warm_contact_2_title,
        "warm_contact_2_linkedin": warm_contact_2_linkedin,
        "outreach_angle": outreach_angle,
        "outreach_status": outreach_status,
        "application_status": application_status,
        "notes": notes,
        "next_action": next_action,
        "last_checked": last_checked,
        "source_url": source_url,
    }
    selected = {key: value for key, value in updates.items() if value is not None}
    if "company_fit_score" in selected and "priority" not in selected:
        selected["priority"] = company_priority_from_score(selected["company_fit_score"])
    if not selected:
        raise SystemExit("Provide at least one target-company field to update.")
    updated_at = now_utc()
    with sqlite3.connect(db_path) as conn:
        require_target_company(conn, target_company_id)
        assignments = ", ".join(f"{key} = ?" for key in selected)
        conn.execute(
            f"UPDATE target_companies SET {assignments}, last_updated_at = ? WHERE id = ?",
            [*selected.values(), updated_at, target_company_id],
        )


def load_target_companies(db_path: Path) -> list[dict[str, Any]]:
    ensure_schema(db_path)
    priority_rank = (
        "CASE priority "
        "WHEN 'Tier 1 target' THEN 1 "
        "WHEN 'Tier 2 target' THEN 2 "
        "WHEN 'Monitor' THEN 3 "
        "WHEN 'Skip' THEN 4 "
        "ELSE 9 END"
    )
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            f"""
            SELECT *
            FROM target_companies
            ORDER BY
                {priority_rank},
                COALESCE(company_fit_score, 0) DESC,
                COALESCE(last_checked, funding_date, created_at) DESC,
                lower(company) ASC
            """
        ).fetchall()
    return [dict(row) for row in rows]


def target_company_row(row: dict[str, Any]) -> dict[str, Any]:
    company_fit_score = clamp_score(row.get("company_fit_score"))
    priority = row.get("priority") or company_priority_from_score(company_fit_score)
    return {
        "ID": row["id"],
        "Company": row["company"] or "",
        "Website": row["website"] or "",
        "Lane": row["lane"] or "",
        "Description": row["description"] or "",
        "Funding Date": display_datetime(row["funding_date"]),
        "Funding Amount": row["funding_amount"] or "",
        "Round": row["round"] or "",
        "Investors": row["investors"] or "",
        "Company Fit Score": company_fit_score or "",
        "Open Roles Found": row["open_roles_found"] or "",
        "Best Role Title": row["best_role_title"] or "",
        "Role Fit Score": clamp_score(row.get("role_fit_score")) or "",
        "Role URL": row["role_url"] or "",
        "Priority": priority,
        "Target Strategy": row.get("target_strategy") or "",
        "Outreach Type": row.get("outreach_type") or "",
        "Warm Contact 1": row["warm_contact_1"] or "",
        "Warm Contact 1 Title": row["warm_contact_1_title"] or "",
        "Warm Contact 1 LinkedIn": row["warm_contact_1_linkedin"] or "",
        "Warm Contact 2": row["warm_contact_2"] or "",
        "Warm Contact 2 Title": row["warm_contact_2_title"] or "",
        "Warm Contact 2 LinkedIn": row["warm_contact_2_linkedin"] or "",
        "Outreach Angle": row["outreach_angle"] or "",
        "Outreach Status": row["outreach_status"] or "",
        "Application Status": row["application_status"] or "",
        "Notes": row["notes"] or "",
        "Next Action": row["next_action"] or "",
        "Last Checked": display_datetime(row["last_checked"]),
        "Source URL": row["source_url"] or "",
        "Created At": display_datetime(row["created_at"]),
        "Last Updated At": display_datetime(row["last_updated_at"]),
    }


def print_target_company_table(rows: list[dict[str, Any]]) -> None:
    if not rows:
        print("No stored target companies yet.")
        return
    print("ID   Priority        Score  Company                     Best Role")
    print("-" * 80)
    for row in rows:
        company = textwrap.shorten(str(row.get("Company") or ""), width=26, placeholder="...")
        best_role = textwrap.shorten(str(row.get("Best Role Title") or ""), width=30, placeholder="...")
        score = row.get("Company Fit Score") or ""
        priority = textwrap.shorten(str(row.get("Priority") or ""), width=14, placeholder="...")
        print(f"{str(row.get('ID') or ''):<4} {priority:<14} {str(score):<5}  {company:<26} {best_role}")


def apply_to_job(
    db_path: Path,
    job_id: int,
    *,
    applied_at: str,
    application_url: str | None,
    resume_version: str | None,
    cover_note_used: str | None,
    outreach_message: str | None,
    follow_up_date: str | None,
    notes: str | None,
    questionnaire_completed: bool = False,
    video_submitted: bool = False,
    submission_summary: str | None = None,
    response_archive_path: str | None = None,
) -> int:
    ensure_schema(db_path)
    created_at = now_utc()
    follow_up = follow_up_date or default_follow_up_date()
    with sqlite3.connect(db_path) as conn:
        job = require_job(conn, job_id)
        outreach = outreach_message or job["outreach_message"]
        cursor = conn.execute(
            """
            INSERT INTO applications (
                job_id, created_at, applied_at, application_url, resume_version,
                cover_note_used, outreach_message, application_status,
                follow_up_date, notes, questionnaire_completed, video_submitted,
                submission_summary, response_archive_path
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, 'applied', ?, ?, ?, ?, ?, ?)
            """,
            (
                job_id,
                created_at,
                applied_at,
                application_url,
                resume_version,
                cover_note_used,
                outreach,
                follow_up,
                notes,
                1 if questionnaire_completed else 0,
                1 if video_submitted else 0,
                submission_summary,
                response_archive_path,
            ),
        )
        conn.execute(
            """
            UPDATE job_evaluations
            SET status = 'applied',
                application_url = COALESCE(?, application_url),
                applied_at = ?,
                resume_version = COALESCE(?, resume_version),
                follow_up_date = ?,
                next_action = 'Follow up if no response',
                last_updated_at = ?
            WHERE id = ?
            """,
            (application_url, applied_at, resume_version, follow_up, created_at, job_id),
        )
        if notes:
            conn.execute(
                "INSERT INTO notes (job_id, created_at, note) VALUES (?, ?, ?)",
                (job_id, created_at, f"Applied: {notes}"),
            )
        return int(cursor.lastrowid)


def add_contact(
    db_path: Path,
    *,
    job_id: int | None,
    name: str,
    role: str | None,
    email: str | None,
    linkedin_url: str | None,
    telegram_handle: str | None,
    relationship: str | None,
    notes: str | None,
) -> int:
    ensure_schema(db_path)
    created_at = now_utc()
    with sqlite3.connect(db_path) as conn:
        if job_id is not None:
            require_job(conn, job_id)
        cursor = conn.execute(
            """
            INSERT INTO contacts (
                job_id, created_at, name, role, email, linkedin_url,
                telegram_handle, relationship, notes
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (job_id, created_at, name, role, email, linkedin_url, telegram_handle, relationship, notes),
        )
        if job_id is not None:
            conn.execute(
                "UPDATE job_evaluations SET last_updated_at = ? WHERE id = ?",
                (created_at, job_id),
            )
        return int(cursor.lastrowid)


def resolve_contact(
    conn: sqlite3.Connection,
    *,
    job_id: int,
    contact_id: int | None,
    contact_name: str | None,
    contact_email: str | None,
) -> int | None:
    if contact_id is not None:
        contact = conn.execute("SELECT id FROM contacts WHERE id = ?", (contact_id,)).fetchone()
        if contact is None:
            raise SystemExit(f"No contact found with id {contact_id}.")
        return contact_id
    if not contact_name and not contact_email:
        return None

    existing = None
    if contact_email:
        existing = conn.execute(
            "SELECT id FROM contacts WHERE lower(email) = lower(?) AND (job_id = ? OR job_id IS NULL)",
            (contact_email, job_id),
        ).fetchone()
    if existing is None and contact_name:
        existing = conn.execute(
            "SELECT id FROM contacts WHERE lower(name) = lower(?) AND (job_id = ? OR job_id IS NULL)",
            (contact_name, job_id),
        ).fetchone()
    if existing is not None:
        return int(existing[0])

    cursor = conn.execute(
        """
        INSERT INTO contacts (job_id, created_at, name, email, relationship)
        VALUES (?, ?, ?, ?, 'recruiting/contact')
        """,
        (job_id, now_utc(), contact_name or contact_email or "Unknown Contact", contact_email),
    )
    return int(cursor.lastrowid)


def infer_status_from_correspondence(kind: str, direction: str) -> str | None:
    kind_clean = kind.lower().replace(" ", "_").replace("-", "_")
    direction_clean = direction.lower()
    if kind_clean in {"rejection", "not_moving_forward"}:
        return "rejected"
    if kind_clean in {"offer", "verbal_offer"}:
        return "offer"
    if "interview" in kind_clean or kind_clean in {"screen", "onsite", "technical", "hiring_manager"}:
        return "interviewing"
    if direction_clean == "inbound" and kind_clean in {"recruiter_reply", "reply", "email", "screen_request"}:
        return "recruiter_reply"
    if direction_clean == "outbound" and kind_clean in {"outreach", "follow_up", "linkedin"}:
        return "outreach_sent"
    return None


def log_correspondence(
    db_path: Path,
    *,
    job_id: int,
    channel: str,
    direction: str,
    kind: str,
    summary: str,
    date: str,
    follow_up_needed: bool,
    follow_up_date: str | None,
    contact_id: int | None,
    contact_name: str | None,
    contact_email: str | None,
    external_thread_id: str | None,
    external_message_id: str | None,
    status: str | None,
) -> int:
    ensure_schema(db_path)
    created_at = now_utc()
    with sqlite3.connect(db_path) as conn:
        job = require_job(conn, job_id)
        resolved_contact = resolve_contact(
            conn,
            job_id=job_id,
            contact_id=contact_id,
            contact_name=contact_name,
            contact_email=contact_email,
        )
        cursor = conn.execute(
            """
            INSERT INTO correspondence (
                job_id, contact_id, created_at, date, channel, direction, type,
                summary, follow_up_needed, follow_up_date, external_thread_id,
                external_message_id
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                job_id,
                resolved_contact,
                created_at,
                date,
                channel,
                direction,
                kind,
                summary,
                1 if follow_up_needed else 0,
                follow_up_date,
                external_thread_id,
                external_message_id,
            ),
        )
        current_status = normalize_status(job["status"] or "discovered")
        inferred_status = normalize_status(status) if status else infer_status_from_correspondence(kind, direction)
        if inferred_status and not status:
            current_rank = STATUS_STAGE_RANK.get(current_status, 0)
            inferred_rank = STATUS_STAGE_RANK.get(inferred_status, 0)
            if inferred_rank < current_rank:
                inferred_status = None
        next_action = None
        if follow_up_needed:
            next_action = f"Follow up on {kind}"
        conn.execute(
            """
            UPDATE job_evaluations
            SET status = COALESCE(?, status),
                follow_up_date = COALESCE(?, follow_up_date),
                next_action = COALESCE(?, next_action),
                last_updated_at = ?
            WHERE id = ?
            """,
            (inferred_status, follow_up_date, next_action, created_at, job_id),
        )
        return int(cursor.lastrowid)


def pipeline_rows(db_path: Path, *, status: str | None, min_score: int | None, limit: int) -> list[dict[str, Any]]:
    ensure_schema(db_path)
    clauses: list[str] = []
    params: list[Any] = []
    if status:
        clauses.append("status = ?")
        params.append(normalize_status(status))
    if min_score is not None:
        clauses.append("fit_score >= ?")
        params.append(min_score)
    where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
    params.append(limit)
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            f"""
            SELECT id, created_at, title, company, fit_score, fit_band,
                   status, follow_up_date, next_action, source
            FROM job_evaluations
            {where}
            ORDER BY
                CASE status
                    WHEN 'interviewing' THEN 1
                    WHEN 'recruiter_reply' THEN 2
                    WHEN 'applied' THEN 3
                    WHEN 'outreach_sent' THEN 4
                    WHEN 'shortlisted' THEN 5
                    WHEN 'discovered' THEN 6
                    ELSE 9
                END,
                fit_score DESC,
                id DESC
            LIMIT ?
            """,
            params,
        ).fetchall()
    return [dict(row) for row in rows]


def followup_rows(db_path: Path, *, days: int) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    ensure_schema(db_path)
    due_by = (today_local() + dt.timedelta(days=days)).isoformat()
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        jobs = conn.execute(
            """
            SELECT id, title, company, fit_score, status, follow_up_date, next_action, source
            FROM job_evaluations
            WHERE follow_up_date IS NOT NULL
              AND follow_up_date <= ?
              AND status NOT IN ('offer', 'rejected', 'archived')
            ORDER BY follow_up_date ASC, fit_score DESC
            """,
            (due_by,),
        ).fetchall()
        correspondence = conn.execute(
            """
            SELECT c.id, c.job_id, j.title, j.company, c.date, c.channel,
                   c.direction, c.type, c.summary, c.follow_up_date
            FROM correspondence c
            JOIN job_evaluations j ON j.id = c.job_id
            WHERE c.follow_up_needed = 1
              AND c.follow_up_date IS NOT NULL
              AND c.follow_up_date <= ?
              AND j.status NOT IN ('offer', 'rejected', 'archived')
            ORDER BY c.follow_up_date ASC, c.id DESC
            """,
            (due_by,),
        ).fetchall()
    return [dict(row) for row in jobs], [dict(row) for row in correspondence]


def print_followups(jobs: list[dict[str, Any]], correspondence: list[dict[str, Any]]) -> None:
    if not jobs and not correspondence:
        print("No follow-ups due in the selected window.")
        return
    if jobs:
        print("Job follow-ups")
        print("-" * 80)
        for row in jobs:
            action = row["next_action"] or "Follow up"
            print(
                f"{row['follow_up_date']} | #{row['id']} | {row['company']} - "
                f"{row['title']} | {row['status']} | {action}"
            )
    if correspondence:
        if jobs:
            print()
        print("Correspondence follow-ups")
        print("-" * 80)
        for row in correspondence:
            print(
                f"{row['follow_up_date']} | job #{row['job_id']} | {row['company']} - "
                f"{row['type']} via {row['channel']} | {row['summary']}"
            )


def export_table(conn: sqlite3.Connection, table: str, output_dir: Path) -> int:
    conn.row_factory = sqlite3.Row
    rows = conn.execute(f"SELECT * FROM {table} ORDER BY id ASC").fetchall()
    output_path = output_dir / f"{table}.csv"
    fieldnames = rows[0].keys() if rows else [row[1] for row in conn.execute(f"PRAGMA table_info({table})").fetchall()]
    with output_path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(dict(row) for row in rows)
    return len(rows)


def export_crm(db_path: Path, output_dir: Path) -> dict[str, int]:
    ensure_schema(db_path)
    output_dir.mkdir(parents=True, exist_ok=True)
    tables = ["job_evaluations", "applications", "target_companies", "contacts", "correspondence", "notes"]
    with sqlite3.connect(db_path) as conn:
        return {table: export_table(conn, table, output_dir) for table in tables}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Evaluate job descriptions against the local candidate profile."
    )
    parser.add_argument("--profile", default=str(DEFAULT_PROFILE), help="Path to profile JSON.")
    parser.add_argument("--db", default=str(DEFAULT_DB), help="Path to SQLite results database.")

    subparsers = parser.add_subparsers(dest="command", required=True)

    evaluate = subparsers.add_parser("evaluate", help="Score and store a job description.")
    evaluate.add_argument("--text", help="Job description text.")
    evaluate.add_argument("--file", help="Path to a job description text file.")
    evaluate.add_argument("--title", help="Override detected job title.")
    evaluate.add_argument("--company", help="Override detected company name.")
    evaluate.add_argument("--source", help="Source URL, board, or note.")
    evaluate.add_argument("--no-save", action="store_true", help="Print evaluation without saving to SQLite.")

    import_chatgpt = subparsers.add_parser(
        "import-chatgpt-job",
        help="Import a manually reviewed job from ChatGPT's structured analysis output.",
    )
    import_chatgpt.add_argument("--text", help="ChatGPT job analysis text.")
    import_chatgpt.add_argument("--file", help="Path to a text file containing ChatGPT job analysis.")
    import_chatgpt.add_argument("--allow-duplicate", action="store_true", help="Save even if company/title or source URL already exists.")
    import_chatgpt.add_argument("--skip-export-refresh", action="store_true", help="Do not refresh local CSV and workbook exports after importing.")

    listed = subparsers.add_parser("list", help="Show recent stored evaluations.")
    listed.add_argument("--limit", type=int, default=20, help="Number of evaluations to show.")

    pipeline = subparsers.add_parser("pipeline", help="Show jobs by CRM status and score.")
    pipeline.add_argument("--status", help=f"Filter to one status: {', '.join(VALID_STATUSES)}.")
    pipeline.add_argument("--min-score", type=int, help="Filter to jobs with this fit score or higher.")
    pipeline.add_argument("--limit", type=int, default=50, help="Number of jobs to show.")

    add_target = subparsers.add_parser(
        "add-target-company",
        help="Store a recently funded company for outreach, role review, and company-target tracking.",
    )
    add_target.add_argument("--company", required=True, help="Company name.")
    add_target.add_argument("--website", help="Company website.")
    add_target.add_argument("--lane", help="Target lane, such as DeFi / Trading Infrastructure.")
    add_target.add_argument("--description", help="Short company description.")
    add_target.add_argument("--funding-date", help="YYYY-MM-DD, today, tomorrow, or +N.")
    add_target.add_argument("--funding-amount", help="Funding amount text, such as $12M.")
    add_target.add_argument("--round", dest="round_name", help="Round type, such as Seed or Series A.")
    add_target.add_argument("--investors", help="Investor list.")
    add_target.add_argument("--company-fit-score", type=int, help="0-100 company-fit score.")
    add_target.add_argument("--open-roles-found", help='Short status such as "2 relevant roles" or "No relevant roles found as of 2026-06-09".')
    add_target.add_argument("--best-role-title", help="Best open role title.")
    add_target.add_argument("--role-fit-score", type=int, help="0-100 role-fit score.")
    add_target.add_argument("--role-url", help="URL for the best role.")
    add_target.add_argument("--priority", help="Override priority label. Defaults from company-fit score.")
    add_target.add_argument(
        "--target-strategy",
        choices=[
            "application_first",
            "relationship_first",
            "relationship_first_plus_possible_application",
            "selective_relationship_first",
            "monitor_only",
            "skip_for_now",
        ],
        help="Whether this is mainly an application target, relationship target, or monitor.",
    )
    add_target.add_argument(
        "--outreach-type",
        choices=[
            "founder_outreach",
            "gtm_leader_outreach",
            "recruiter_outreach",
            "hiring_manager_outreach",
            "investor_warm_intro",
            "monitor_no_outreach_yet",
        ],
        help="Primary outreach route for this company.",
    )
    add_target.add_argument("--warm-contact-1", help="Primary contact name.")
    add_target.add_argument("--warm-contact-1-title", help="Primary contact title.")
    add_target.add_argument("--warm-contact-1-linkedin", help="Primary contact LinkedIn URL.")
    add_target.add_argument("--warm-contact-2", help="Secondary contact name.")
    add_target.add_argument("--warm-contact-2-title", help="Secondary contact title.")
    add_target.add_argument("--warm-contact-2-linkedin", help="Secondary contact LinkedIn URL.")
    add_target.add_argument("--outreach-angle", help="Suggested outreach angle.")
    add_target.add_argument("--outreach-status", help="Outreach status, such as to_research, drafted, sent.")
    add_target.add_argument("--application-status", help="Application status for the best role, if any.")
    add_target.add_argument("--notes", help="Freeform notes.")
    add_target.add_argument("--next-action", help="Next action.")
    add_target.add_argument("--last-checked", help="YYYY-MM-DD, today, tomorrow, or +N.")
    add_target.add_argument("--source-url", help="Source URL for the raise or company research.")

    update_target = subparsers.add_parser("update-target-company", help="Update a stored target company.")
    update_target.add_argument("--id", type=int, required=True, help="Stored target company ID.")
    update_target.add_argument("--company", help="Company name.")
    update_target.add_argument("--website", help="Company website.")
    update_target.add_argument("--lane", help="Target lane.")
    update_target.add_argument("--description", help="Short company description.")
    update_target.add_argument("--funding-date", help="YYYY-MM-DD, today, tomorrow, or +N.")
    update_target.add_argument("--funding-amount", help="Funding amount text.")
    update_target.add_argument("--round", dest="round_name", help="Round type.")
    update_target.add_argument("--investors", help="Investor list.")
    update_target.add_argument("--company-fit-score", type=int, help="0-100 company-fit score.")
    update_target.add_argument("--open-roles-found", help="Open-role status text.")
    update_target.add_argument("--best-role-title", help="Best open role title.")
    update_target.add_argument("--role-fit-score", type=int, help="0-100 role-fit score.")
    update_target.add_argument("--role-url", help="URL for the best role.")
    update_target.add_argument("--priority", help="Priority label.")
    update_target.add_argument(
        "--target-strategy",
        choices=[
            "application_first",
            "relationship_first",
            "relationship_first_plus_possible_application",
            "selective_relationship_first",
            "monitor_only",
            "skip_for_now",
        ],
        help="Whether this is mainly an application target, relationship target, or monitor.",
    )
    update_target.add_argument(
        "--outreach-type",
        choices=[
            "founder_outreach",
            "gtm_leader_outreach",
            "recruiter_outreach",
            "hiring_manager_outreach",
            "investor_warm_intro",
            "monitor_no_outreach_yet",
        ],
        help="Primary outreach route for this company.",
    )
    update_target.add_argument("--warm-contact-1", help="Primary contact name.")
    update_target.add_argument("--warm-contact-1-title", help="Primary contact title.")
    update_target.add_argument("--warm-contact-1-linkedin", help="Primary contact LinkedIn URL.")
    update_target.add_argument("--warm-contact-2", help="Secondary contact name.")
    update_target.add_argument("--warm-contact-2-title", help="Secondary contact title.")
    update_target.add_argument("--warm-contact-2-linkedin", help="Secondary contact LinkedIn URL.")
    update_target.add_argument("--outreach-angle", help="Suggested outreach angle.")
    update_target.add_argument("--outreach-status", help="Outreach status.")
    update_target.add_argument("--application-status", help="Application status.")
    update_target.add_argument("--notes", help="Freeform notes.")
    update_target.add_argument("--next-action", help="Next action.")
    update_target.add_argument("--last-checked", help="YYYY-MM-DD, today, tomorrow, or +N.")
    update_target.add_argument("--source-url", help="Source URL for the raise or company research.")

    list_target = subparsers.add_parser("list-target-companies", help="Show stored recently funded target companies.")
    list_target.add_argument("--limit", type=int, default=50, help="Number of target companies to show.")

    update = subparsers.add_parser("update-status", help="Update a job's CRM status.")
    update.add_argument("--id", type=int, required=True, help="Stored job evaluation ID.")
    update.add_argument("--status", required=True, help=f"New status: {', '.join(VALID_STATUSES)}.")
    update.add_argument("--next-action", help="Next action to take.")
    update.add_argument("--follow-up-date", help="YYYY-MM-DD, today, tomorrow, or +N.")
    update.add_argument("--decision", help="Short decision note, such as apply, pass, wait, referral.")
    update.add_argument("--archived-reason", help="Reason if archiving.")
    update.add_argument("--note", help="Optional note to add with the status change.")
    update.add_argument("--sector", help="Dashboard sector, such as Cybersecurity or Healthtech.")
    update.add_argument("--role-lane", help="Dashboard role lane, such as Strategic Account Management.")
    update.add_argument("--priority", choices=PRIORITY_VALUES, help="Dashboard priority.")
    update.add_argument("--application-url", help="Application URL for the job.")
    update.add_argument("--resume-version", help="Resume version planned or used.")
    update.add_argument("--cover-letter-needed", choices=["Yes", "Optional", "No", "Review"], help="Cover letter need.")
    update.add_argument("--referral-target", help="Recruiter, employee, or referral target.")

    update_job = subparsers.add_parser("update-job", help="Update job dashboard fields without changing status.")
    update_job.add_argument("--id", type=int, required=True, help="Stored job evaluation ID.")
    update_job.add_argument("--next-action", help="Next action to take.")
    update_job.add_argument("--follow-up-date", help="YYYY-MM-DD, today, tomorrow, or +N.")
    update_job.add_argument("--decision", help="Short decision note, such as apply, pass, wait, referral.")
    update_job.add_argument("--sector", help="Dashboard sector, such as Cybersecurity or Healthtech.")
    update_job.add_argument("--role-lane", help="Dashboard role lane, such as Strategic Account Management.")
    update_job.add_argument("--priority", choices=PRIORITY_VALUES, help="Dashboard priority.")
    update_job.add_argument("--application-url", help="Application URL for the job.")
    update_job.add_argument("--resume-version", help="Resume version planned or used.")
    update_job.add_argument("--cover-letter-needed", choices=["Yes", "Optional", "No", "Review"], help="Cover letter need.")
    update_job.add_argument("--referral-target", help="Recruiter, employee, or referral target.")
    update_job.add_argument("--note", help="Optional note to add.")

    apply_cmd = subparsers.add_parser("apply", help="Mark a job as applied and create an application record.")
    apply_cmd.add_argument("--id", type=int, required=True, help="Stored job evaluation ID.")
    apply_cmd.add_argument("--applied-at", default="today", help="YYYY-MM-DD, today, tomorrow, or +N.")
    apply_cmd.add_argument("--application-url", help="URL used to apply.")
    apply_cmd.add_argument("--resume-version", help="Resume version used.")
    apply_cmd.add_argument("--cover-note", help="Cover note or application note used.")
    apply_cmd.add_argument("--outreach-message", help="Outreach message used; defaults to generated message.")
    apply_cmd.add_argument("--follow-up-date", help="YYYY-MM-DD, today, tomorrow, or +N. Defaults to +7.")
    apply_cmd.add_argument("--note", help="Optional application note.")
    apply_cmd.add_argument("--questionnaire-completed", action="store_true", help="Mark that written questionnaire answers were submitted.")
    apply_cmd.add_argument("--video-submitted", action="store_true", help="Mark that video responses were submitted.")
    apply_cmd.add_argument("--submission-summary", help="Short summary of what was submitted.")
    apply_cmd.add_argument("--response-archive-path", help="Local path to archived submitted answers or transcripts.")

    note = subparsers.add_parser("add-note", help="Add a note to a job.")
    note.add_argument("--id", type=int, required=True, help="Stored job evaluation ID.")
    note.add_argument("--note", required=True, help="Note text.")

    packet = subparsers.add_parser("application-packet", help="Generate a reviewable application packet for a job.")
    packet.add_argument("--id", type=int, required=True, help="Stored job evaluation ID.")
    packet.add_argument("--output", help="Optional Markdown output path.")
    packet.add_argument("--print", action="store_true", help="Also print the packet to stdout.")
    packet.add_argument(
        "--skip-google-sync",
        action="store_true",
        help="Keep packet generation local only even when Google Drive sync credentials are already available.",
    )

    capture_questions = subparsers.add_parser(
        "capture-application-questions",
        help="Store the exact application questions for a job from pasted text, stdin, or a file.",
    )
    capture_questions.add_argument("--id", type=int, required=True, help="Stored job evaluation ID.")
    capture_questions.add_argument("--text", help="Plaintext or JSON question payload.")
    capture_questions.add_argument("--file", help="Path to a plaintext or JSON file containing the exact application questions.")
    capture_questions.add_argument("--clear-existing", action="store_true", help="Replace any existing captured questions for this job instead of merging.")
    capture_questions.add_argument(
        "--status",
        choices=["Captured", "Captured from Screenshot", "Partially Captured", "Not Captured"],
        help="Whether the questions were fully captured, captured from a screenshot, only partially captured, or not captured.",
    )
    capture_questions.add_argument("--reason", help="Short note about how the questions were captured.")
    capture_questions.add_argument("--next-action", help="Optional next action note for the question-capture state.")

    link_packet = subparsers.add_parser("link-packet", help="Attach a Google Doc URL to an existing local packet.")
    link_packet.add_argument("--id", type=int, required=True, help="Stored job evaluation ID.")
    link_packet.add_argument("--doc-url", required=True, help="Google Doc URL for the mirrored application packet.")

    sync_drive = subparsers.add_parser(
        "sync-google-drive-docs",
        help="Sync local packet and cover-letter text into specific Google Drive folders using OAuth or a service account.",
    )
    sync_drive.add_argument("--id", type=int, action="append", help="Specific job ID to sync. Repeat for multiple jobs. Defaults to all local packets.")
    sync_drive.add_argument(
        "--auth-mode",
        choices=["auto", "oauth", "service-account", "hybrid"],
        default="auto",
        help="Authentication mode. Use oauth for personal My Drive folders. auto prefers OAuth if an OAuth client JSON is present. hybrid tries OAuth first, then a service account.",
    )
    sync_drive.add_argument(
        "--service-account-json",
        default=str(DEFAULT_GOOGLE_SERVICE_ACCOUNT_JSON),
        help="Path to a Google service-account JSON key. Best for shared drives or Google Workspace delegation.",
    )
    sync_drive.add_argument(
        "--oauth-client-json",
        default=str(DEFAULT_GOOGLE_OAUTH_CLIENT_JSON),
        help="Path to a Google OAuth desktop client JSON. Best for personal My Drive folders.",
    )
    sync_drive.add_argument(
        "--oauth-token-json",
        default=str(DEFAULT_GOOGLE_OAUTH_TOKEN_JSON),
        help="Path to the cached OAuth user token written after the first browser login.",
    )
    sync_drive.add_argument(
        "--packet-folder-url",
        default=GOOGLE_PACKET_FOLDER_URL,
        help="Google Drive folder URL for packet docs.",
    )
    sync_drive.add_argument(
        "--cover-letter-folder-url",
        default=GOOGLE_COVER_LETTER_FOLDER_URL,
        help="Google Drive folder URL for cover-letter docs.",
    )

    contact = subparsers.add_parser("add-contact", help="Add a recruiter, hiring manager, or referral contact.")
    contact.add_argument("--job-id", type=int, help="Stored job evaluation ID.")
    contact.add_argument("--name", required=True, help="Contact name.")
    contact.add_argument("--role", help="Contact role/title.")
    contact.add_argument("--email", help="Contact email.")
    contact.add_argument("--linkedin-url", help="Contact LinkedIn URL.")
    contact.add_argument("--telegram-handle", help="Contact Telegram handle, for example @username.")
    contact.add_argument("--relationship", help="Recruiter, hiring manager, referral, employee, etc.")
    contact.add_argument("--notes", help="Contact notes.")

    corr = subparsers.add_parser("log-correspondence", help="Log an email, LinkedIn message, recruiter reply, interview, rejection, or offer.")
    corr.add_argument("--job-id", type=int, required=True, help="Stored job evaluation ID.")
    corr.add_argument("--channel", required=True, help="email, linkedin, telegram, ats, phone, interview, referral, etc.")
    corr.add_argument("--direction", required=True, choices=["inbound", "outbound", "internal"], help="Message direction.")
    corr.add_argument("--type", required=True, help="confirmation, outreach, recruiter_reply, follow_up, interview, rejection, offer, etc.")
    corr.add_argument("--summary", required=True, help="Short summary; avoid sensitive personal data.")
    corr.add_argument("--date", default="today", help="YYYY-MM-DD, today, tomorrow, or +N.")
    corr.add_argument("--follow-up-needed", action="store_true", help="Whether this needs a follow-up.")
    corr.add_argument("--follow-up-date", help="YYYY-MM-DD, today, tomorrow, or +N.")
    corr.add_argument("--contact-id", type=int, help="Existing contact ID.")
    corr.add_argument("--contact-name", help="Contact name; creates contact if needed.")
    corr.add_argument("--contact-email", help="Contact email; creates contact if needed.")
    corr.add_argument("--external-thread-id", help="Gmail/LinkedIn/ATS thread ID, if useful.")
    corr.add_argument("--external-message-id", help="Gmail/LinkedIn/ATS message ID, if useful.")
    corr.add_argument("--status", help="Optional explicit job status update.")

    followups = subparsers.add_parser("followups", help="Show follow-ups due soon.")
    followups.add_argument("--days", type=int, default=7, help="Look ahead this many days.")

    export = subparsers.add_parser("export-csv", help="Export stored evaluations to CSV.")
    export.add_argument("--output", default=str(ROOT / "exports" / "job_results.csv"), help="CSV output path.")

    export_sheets = subparsers.add_parser("export-sheets-csv", help="Export a compact Google Sheets-friendly tracker CSV.")
    export_sheets.add_argument(
        "--output",
        default=str(ROOT / "exports" / "google_sheets_job_tracker.csv"),
        help="CSV output path.",
    )

    export_targets = subparsers.add_parser("export-target-companies", help="Export the recently funded target-company tracker to CSV.")
    export_targets.add_argument(
        "--output",
        default=str(ROOT / "exports" / "target_companies.csv"),
        help="CSV output path.",
    )

    export_all = subparsers.add_parser("export-crm", help="Export all CRM tables to CSV files.")
    export_all.add_argument("--output-dir", default=str(ROOT / "exports" / "crm"), help="Directory for CRM CSV exports.")

    export_workbook = subparsers.add_parser("export-sheets-workbook", help="Export CSVs matching the Google Sheets CRM workbook tabs.")
    export_workbook.add_argument("--output-dir", default=str(DEFAULT_SHEETS_WORKBOOK_DIR), help="Directory for workbook tab CSV exports.")

    sync_sheet = subparsers.add_parser(
        "sync-google-sheets-workbook",
        help="Sync the exported workbook tabs into the live Job Search CRM Google Sheet.",
    )
    sync_sheet.add_argument(
        "--spreadsheet-url",
        default=GOOGLE_SHEETS_WORKBOOK_URL,
        help="Google Sheets workbook URL or spreadsheet ID for the live CRM.",
    )
    sync_sheet.add_argument(
        "--auth-mode",
        choices=["auto", "oauth", "service-account", "hybrid"],
        default="hybrid",
        help="Authentication mode. hybrid prefers the service account for Sheets, then falls back to OAuth.",
    )
    sync_sheet.add_argument(
        "--service-account-json",
        default=str(DEFAULT_GOOGLE_SERVICE_ACCOUNT_JSON),
        help="Path to a Google service-account JSON key for unattended sheet sync.",
    )
    sync_sheet.add_argument(
        "--oauth-client-json",
        default=str(DEFAULT_GOOGLE_OAUTH_CLIENT_JSON),
        help="Path to a Google OAuth desktop client JSON for personal fallback access.",
    )
    sync_sheet.add_argument(
        "--oauth-token-json",
        default=str(DEFAULT_GOOGLE_OAUTH_TOKEN_JSON),
        help="Path to the cached OAuth user token written after the first browser login.",
    )

    daily_run = subparsers.add_parser(
        "daily-run",
        help="Run the local daily refresh flow: follow-up check, exports, workbook exports, and optional live sheet sync.",
    )
    daily_run.add_argument("--days", type=int, default=7, help="Look ahead this many days for follow-ups.")
    daily_run.add_argument("--min-score", type=int, default=90, help="Minimum fit score for the printed pipeline snapshot.")
    daily_run.add_argument("--limit", type=int, default=20, help="Maximum number of pipeline rows to summarize.")
    daily_run.add_argument(
        "--skip-sheet-sync",
        action="store_true",
        help="Refresh local exports only and skip the live Google Sheet sync.",
    )
    daily_run.add_argument(
        "--spreadsheet-url",
        default=GOOGLE_SHEETS_WORKBOOK_URL,
        help="Google Sheets workbook URL or spreadsheet ID for the live CRM.",
    )
    daily_run.add_argument(
        "--sheet-auth-mode",
        choices=["auto", "oauth", "service-account", "hybrid"],
        default="hybrid",
        help="Auth mode for live workbook sync. hybrid prefers the service account for Sheets, then falls back to OAuth.",
    )
    daily_run.add_argument(
        "--service-account-json",
        default=str(DEFAULT_GOOGLE_SERVICE_ACCOUNT_JSON),
        help="Path to a Google service-account JSON key for unattended sheet sync.",
    )
    daily_run.add_argument(
        "--oauth-client-json",
        default=str(DEFAULT_GOOGLE_OAUTH_CLIENT_JSON),
        help="Path to a Google OAuth desktop client JSON for personal fallback access.",
    )
    daily_run.add_argument(
        "--oauth-token-json",
        default=str(DEFAULT_GOOGLE_OAUTH_TOKEN_JSON),
        help="Path to the cached OAuth user token written after the first browser login.",
    )

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    db_path = Path(args.db)

    if args.command == "evaluate":
        profile = load_profile(Path(args.profile))
        jd = read_job_description(args)
        result = evaluate_job(jd, profile, args.title, args.company)
        row_id = None if args.no_save else save_result(db_path, result, args.source)
        print(render_markdown(result, row_id))
        return 0

    if args.command == "import-chatgpt-job":
        analysis_text = read_job_description(args)
        parsed = parse_chatgpt_job_analysis(analysis_text)
        job_id = save_chatgpt_job_import(db_path, parsed, allow_duplicate=args.allow_duplicate)
        print(f"Imported ChatGPT-reviewed job #{job_id}: {parsed['company']} - {parsed['title']}")
        print(f"Fit score: {parsed['fit_score']} | Priority: {parsed['priority']} | Sector: {parsed['sector']}")
        if not args.skip_export_refresh:
            csv_count = export_csv(db_path, ROOT / "exports" / "job_results.csv")
            export_crm(db_path, ROOT / "exports" / "crm")
            sheet_count = export_sheets_csv(db_path, ROOT / "exports" / "google_sheets_job_tracker.csv")
            workbook_counts = export_sheets_workbook(db_path, DEFAULT_SHEETS_WORKBOOK_DIR)
            print(
                "Local exports refreshed: "
                f"job_results={csv_count}, sheets_jobs={sheet_count}, "
                f"jobs_tab={workbook_counts.get('Jobs', 0)}, applied_tab={workbook_counts.get('Applied', 0)}"
            )
        return 0

    if args.command == "list":
        print_table(list_results(db_path, args.limit))
        return 0

    if args.command == "pipeline":
        print_table(pipeline_rows(db_path, status=args.status, min_score=args.min_score, limit=args.limit))
        return 0

    if args.command == "add-target-company":
        target_company_id = add_target_company(
            db_path,
            company=args.company,
            website=args.website,
            lane=args.lane,
            description=args.description,
            funding_date=parse_date(args.funding_date, field_name="funding date"),
            funding_amount=args.funding_amount,
            round_name=args.round_name,
            investors=args.investors,
            company_fit_score=args.company_fit_score,
            open_roles_found=args.open_roles_found,
            best_role_title=args.best_role_title,
            role_fit_score=args.role_fit_score,
            role_url=args.role_url,
            priority=args.priority,
            target_strategy=args.target_strategy,
            outreach_type=args.outreach_type,
            warm_contact_1=args.warm_contact_1,
            warm_contact_1_title=args.warm_contact_1_title,
            warm_contact_1_linkedin=args.warm_contact_1_linkedin,
            warm_contact_2=args.warm_contact_2,
            warm_contact_2_title=args.warm_contact_2_title,
            warm_contact_2_linkedin=args.warm_contact_2_linkedin,
            outreach_angle=args.outreach_angle,
            outreach_status=args.outreach_status,
            application_status=args.application_status,
            notes=args.notes,
            next_action=args.next_action,
            last_checked=parse_date(args.last_checked, field_name="last checked"),
            source_url=args.source_url,
        )
        print(f"Added target company #{target_company_id}: {args.company}")
        return 0

    if args.command == "update-target-company":
        update_target_company(
            db_path,
            args.id,
            company=args.company,
            website=args.website,
            lane=args.lane,
            description=args.description,
            funding_date=parse_date(args.funding_date, field_name="funding date") if args.funding_date else None,
            funding_amount=args.funding_amount,
            round_name=args.round_name,
            investors=args.investors,
            company_fit_score=args.company_fit_score,
            open_roles_found=args.open_roles_found,
            best_role_title=args.best_role_title,
            role_fit_score=args.role_fit_score,
            role_url=args.role_url,
            priority=args.priority,
            target_strategy=args.target_strategy,
            outreach_type=args.outreach_type,
            warm_contact_1=args.warm_contact_1,
            warm_contact_1_title=args.warm_contact_1_title,
            warm_contact_1_linkedin=args.warm_contact_1_linkedin,
            warm_contact_2=args.warm_contact_2,
            warm_contact_2_title=args.warm_contact_2_title,
            warm_contact_2_linkedin=args.warm_contact_2_linkedin,
            outreach_angle=args.outreach_angle,
            outreach_status=args.outreach_status,
            application_status=args.application_status,
            notes=args.notes,
            next_action=args.next_action,
            last_checked=parse_date(args.last_checked, field_name="last checked") if args.last_checked else None,
            source_url=args.source_url,
        )
        print(f"Updated target company #{args.id}.")
        return 0

    if args.command == "list-target-companies":
        print_target_company_table(target_company_rows(db_path)[: args.limit])
        return 0

    if args.command == "update-status":
        update_job_status(
            db_path,
            args.id,
            args.status,
            next_action=args.next_action,
            follow_up_date=parse_date(args.follow_up_date, field_name="follow-up date"),
            decision=args.decision,
            archived_reason=args.archived_reason,
            note=args.note,
            sector=args.sector,
            role_lane=args.role_lane,
            priority=args.priority,
            application_url=args.application_url,
            resume_version=args.resume_version,
            cover_letter_needed=args.cover_letter_needed,
            referral_target=args.referral_target,
        )
        print(f"Updated job #{args.id} to {normalize_status(args.status)}.")
        return 0

    if args.command == "update-job":
        update_job_details(
            db_path,
            args.id,
            next_action=args.next_action,
            follow_up_date=parse_date(args.follow_up_date, field_name="follow-up date"),
            decision=args.decision,
            sector=args.sector,
            role_lane=args.role_lane,
            priority=args.priority,
            application_url=args.application_url,
            resume_version=args.resume_version,
            cover_letter_needed=args.cover_letter_needed,
            referral_target=args.referral_target,
            note=args.note,
        )
        print(f"Updated dashboard fields for job #{args.id}.")
        return 0

    if args.command == "apply":
        application_id = apply_to_job(
            db_path,
            args.id,
            applied_at=parse_date(args.applied_at, field_name="applied date") or today_local().isoformat(),
            application_url=args.application_url,
            resume_version=args.resume_version,
            cover_note_used=args.cover_note,
            outreach_message=args.outreach_message,
            follow_up_date=parse_date(args.follow_up_date, field_name="follow-up date") if args.follow_up_date else None,
            notes=args.note,
            questionnaire_completed=args.questionnaire_completed,
            video_submitted=args.video_submitted,
            submission_summary=args.submission_summary,
            response_archive_path=args.response_archive_path,
        )
        print(f"Marked job #{args.id} as applied. Application record #{application_id} created.")
        return 0

    if args.command == "add-note":
        note_id = add_note(db_path, args.id, args.note)
        print(f"Added note #{note_id} to job #{args.id}.")
        return 0

    if args.command == "application-packet":
        output_path = Path(args.output) if args.output else None
        path = write_application_packet(db_path, args.id, output_path)
        cover_letter_path = cover_letter_docx_path_for_packet_path(path)
        if args.print:
            print(path.read_text(encoding="utf-8"))
        print(f"Wrote application packet for job #{args.id}: {path}")
        print(f"Wrote cover letter DOCX for job #{args.id}: {cover_letter_path}")
        if args.skip_google_sync:
            print("Skipped Google Drive sync by request.")
            return 0
        if can_attempt_automatic_google_sync():
            try:
                results = sync_google_drive_docs(
                    db_path,
                    job_ids=[args.id],
                    auth_mode="auto",
                    service_account_json=DEFAULT_GOOGLE_SERVICE_ACCOUNT_JSON,
                    oauth_client_json=DEFAULT_GOOGLE_OAUTH_CLIENT_JSON,
                    oauth_token_json=DEFAULT_GOOGLE_OAUTH_TOKEN_JSON,
                    packet_folder_url=GOOGLE_PACKET_FOLDER_URL,
                    cover_letter_folder_url=GOOGLE_COVER_LETTER_FOLDER_URL,
                )
                for result in results:
                    print(
                        f"Synced Google Drive docs for job #{result['job_id']}\n"
                        f"  Packet Doc: {result['packet_url'] or 'not created'}\n"
                        f"  Cover Letter Doc: {result['cover_letter_url'] or 'not created'}"
                    )
            except SystemExit as exc:
                print(f"Automatic Google Drive sync failed: {exc}")
        else:
            print(
                "Skipped Google Drive sync because no reusable OAuth or service-account credentials were available. "
                "Run `sync-google-drive-docs` after completing Google auth if you want the packet and cover letter mirrored."
            )
        return 0

    if args.command == "capture-application-questions":
        if args.file:
            question_text = Path(args.file).read_text(encoding="utf-8")
        elif args.text:
            question_text = args.text
        elif not sys.stdin.isatty():
            question_text = sys.stdin.read()
        else:
            raise SystemExit("Provide application questions with --text, --file, or stdin.")
        total = capture_application_questions(
            db_path,
            job_id=args.id,
            text=question_text,
            clear_existing=args.clear_existing,
            capture_status=args.status,
            capture_reason=args.reason,
            capture_next_action=args.next_action,
        )
        print(f"Stored {total} application questions for job #{args.id}.")
        return 0

    if args.command == "link-packet":
        set_packet_doc_url(args.id, args.doc_url)
        print(f"Linked packet for job #{args.id} to {args.doc_url}")
        return 0

    if args.command == "sync-google-drive-docs":
        results = sync_google_drive_docs(
            db_path,
            job_ids=args.id,
            auth_mode=args.auth_mode,
            service_account_json=Path(args.service_account_json),
            oauth_client_json=Path(args.oauth_client_json),
            oauth_token_json=Path(args.oauth_token_json),
            packet_folder_url=args.packet_folder_url,
            cover_letter_folder_url=args.cover_letter_folder_url,
        )
        for result in results:
            print(
                f"Synced job #{result['job_id']} ({result['company']} - {result['title']})\n"
                f"  Packet Doc: {result['packet_url'] or 'not created'}\n"
                f"  Cover Letter Doc: {result['cover_letter_url'] or 'not created'}"
        )
        print(f"Synced {len(results)} job packet records to Google Drive.")
        return 0

    if args.command == "sync-google-sheets-workbook":
        result = sync_google_sheets_workbook(
            db_path,
            spreadsheet_url=args.spreadsheet_url,
            auth_mode=args.auth_mode,
            service_account_json=Path(args.service_account_json),
            oauth_client_json=Path(args.oauth_client_json),
            oauth_token_json=Path(args.oauth_token_json),
        )
        print(
            f"Synced workbook tabs to Google Sheets via {result['credentials_type']} credentials: "
            f"{args.spreadsheet_url}"
        )
        sheet_import = result.get("sheet_import") or {}
        print(
            "Imported sheet edits before sync: "
            f"{sheet_import.get('rows_imported', 0)} rows, "
            f"{sheet_import.get('status_updates', 0)} status updates, "
            f"{sheet_import.get('application_records_created', 0)} application records created, "
            f"{sheet_import.get('target_company_rows_imported', 0)} target companies updated"
        )
        for tab_name, count in result["counts"].items():
            print(f"  {tab_name}: {count} rows")
        return 0

    if args.command == "daily-run":
        result = run_daily_workflow(
            db_path,
            followup_days=args.days,
            pipeline_min_score=args.min_score,
            pipeline_limit=args.limit,
            spreadsheet_url=args.spreadsheet_url,
            sheet_auth_mode=args.sheet_auth_mode,
            service_account_json=Path(args.service_account_json),
            oauth_client_json=Path(args.oauth_client_json),
            oauth_token_json=Path(args.oauth_token_json),
            skip_sheet_sync=args.skip_sheet_sync,
        )
        print(
            f"Daily workflow complete. Follow-ups due: "
            f"{len(result['followup_jobs']) + len(result['followup_correspondence'])}"
        )
        if result["pipeline"]:
            top = result["pipeline"][0]
            print(
                f"Top pipeline role: #{top['id']} {top['company']} - {top['title']} "
                f"({top['fit_score']})"
            )
        workbook_counts = result["exports"]["workbook"]
        print(
            "Exports refreshed: "
            f"Jobs={result['exports']['job_results']}, "
            f"Target Companies={result['exports']['target_companies']}, "
            f"Workbook Jobs={workbook_counts.get('Jobs', 0)}, "
            f"Packets={workbook_counts.get('Packets', 0)}"
        )
        if result["sheet_sync"]:
            print(
                "Live Google Sheet sync completed via "
                f"{result['sheet_sync']['credentials_type']} credentials."
            )
            sheet_import = result["sheet_sync"].get("sheet_import") or {}
            print(
                "Imported sheet edits before sync: "
                f"{sheet_import.get('rows_imported', 0)} rows, "
                f"{sheet_import.get('status_updates', 0)} status updates, "
                f"{sheet_import.get('application_records_created', 0)} application records created, "
                f"{sheet_import.get('target_company_rows_imported', 0)} target companies updated"
            )
        else:
            print("Live Google Sheet sync skipped by request.")
        return 0

    if args.command == "add-contact":
        contact_id = add_contact(
            db_path,
            job_id=args.job_id,
            name=args.name,
            role=args.role,
            email=args.email,
            linkedin_url=args.linkedin_url,
            telegram_handle=args.telegram_handle,
            relationship=args.relationship,
            notes=args.notes,
        )
        print(f"Added contact #{contact_id}.")
        return 0

    if args.command == "log-correspondence":
        follow_up_date = parse_date(args.follow_up_date, field_name="follow-up date") if args.follow_up_date else None
        if args.follow_up_needed and follow_up_date is None:
            follow_up_date = default_follow_up_date()
        correspondence_id = log_correspondence(
            db_path,
            job_id=args.job_id,
            channel=args.channel,
            direction=args.direction,
            kind=args.type,
            summary=args.summary,
            date=parse_date(args.date, field_name="correspondence date") or today_local().isoformat(),
            follow_up_needed=args.follow_up_needed,
            follow_up_date=follow_up_date,
            contact_id=args.contact_id,
            contact_name=args.contact_name,
            contact_email=args.contact_email,
            external_thread_id=args.external_thread_id,
            external_message_id=args.external_message_id,
            status=args.status,
        )
        print(f"Logged correspondence #{correspondence_id} for job #{args.job_id}.")
        return 0

    if args.command == "followups":
        jobs, correspondence = followup_rows(db_path, days=args.days)
        print_followups(jobs, correspondence)
        return 0

    if args.command == "export-csv":
        count = export_csv(db_path, Path(args.output))
        print(f"Exported {count} evaluations to {args.output}")
        return 0

    if args.command == "export-sheets-csv":
        count = export_sheets_csv(db_path, Path(args.output))
        print(f"Exported {count} jobs to Google Sheets-friendly CSV: {args.output}")
        return 0

    if args.command == "export-target-companies":
        count = export_target_companies_csv(db_path, Path(args.output))
        print(f"Exported {count} target companies to CSV: {args.output}")
        return 0

    if args.command == "export-crm":
        counts = export_crm(db_path, Path(args.output_dir))
        for table, count in counts.items():
            print(f"Exported {count} rows from {table} to {args.output_dir}")
        return 0

    if args.command == "export-sheets-workbook":
        counts = export_sheets_workbook(db_path, Path(args.output_dir))
        for tab, count in counts.items():
            print(f"Exported {count} rows for {tab} to {args.output_dir}")
        return 0

    parser.print_help()
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
