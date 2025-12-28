#!/usr/bin/env python3
"""
Billing Reports Generator

Generates all the reports Dean needs:
1. Call ratio (interstate vs intrastate) - overall and by customer
2. CDR by customer with costs
3. Phone number counts by customer
4. SMS usage
5. Seat counts

Cost Rates (from Excel formulas):
- Voice: $0.005 per minute (half a cent)
- SMS:   $0.005 per message (half a cent)
"""

import csv
import sys
from pathlib import Path
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Optional
import re

# ============================================================================
# Billing Rates
# ============================================================================
VOICE_RATE_PER_MINUTE = 0.005  # $0.005 per minute
SMS_RATE_PER_MESSAGE = 0.005   # $0.005 per message

# ============================================================================
# Phone Number Treatments - Non-billable (excluded from PVT PhoneNumbers)
# These are fax machines, on-hold numbers, unassigned, etc.
# ============================================================================
NON_BILLABLE_TREATMENTS = {
    'Available Number',  # Unassigned phone numbers
    'FaxSFATA',          # Physical fax machines (SFATA)
    'vFaxSFATA',         # Virtual fax SFATA
    'iFax',              # Internet fax
    'vFax',              # Virtual fax
    'vOn-Hold',          # On-hold music numbers
    'vOffNet',           # Off-network numbers
}

# ============================================================================
# NPA (Area Code) to State mapping
# ============================================================================
NPA_TO_STATE = {
    "205": "AL", "251": "AL", "256": "AL", "334": "AL", "938": "AL",
    "907": "AK",
    "480": "AZ", "520": "AZ", "602": "AZ", "623": "AZ", "928": "AZ",
    "479": "AR", "501": "AR", "870": "AR",
    "209": "CA", "213": "CA", "310": "CA", "323": "CA", "341": "CA",
    "408": "CA", "415": "CA", "424": "CA", "442": "CA", "510": "CA",
    "530": "CA", "559": "CA", "562": "CA", "619": "CA", "626": "CA",
    "628": "CA", "650": "CA", "657": "CA", "661": "CA", "669": "CA",
    "707": "CA", "714": "CA", "747": "CA", "760": "CA", "805": "CA",
    "818": "CA", "820": "CA", "831": "CA", "840": "CA", "858": "CA",
    "909": "CA", "916": "CA", "925": "CA", "949": "CA", "951": "CA",
    "303": "CO", "719": "CO", "720": "CO", "970": "CO", "983": "CO",
    "203": "CT", "475": "CT", "860": "CT", "959": "CT",
    "302": "DE",
    "239": "FL", "305": "FL", "321": "FL", "352": "FL", "386": "FL",
    "407": "FL", "561": "FL", "727": "FL", "754": "FL", "772": "FL",
    "786": "FL", "813": "FL", "850": "FL", "863": "FL", "904": "FL",
    "941": "FL", "954": "FL",
    "229": "GA", "404": "GA", "470": "GA", "478": "GA", "678": "GA",
    "706": "GA", "762": "GA", "770": "GA", "912": "GA", "943": "GA",
    "808": "HI",
    "208": "ID", "986": "ID",
    "217": "IL", "224": "IL", "309": "IL", "312": "IL", "331": "IL",
    "618": "IL", "630": "IL", "708": "IL", "773": "IL", "779": "IL",
    "815": "IL", "847": "IL", "872": "IL",
    "219": "IN", "260": "IN", "317": "IN", "463": "IN", "574": "IN",
    "765": "IN", "812": "IN", "930": "IN",
    "319": "IA", "515": "IA", "563": "IA", "641": "IA", "712": "IA",
    "316": "KS", "620": "KS", "785": "KS", "913": "KS",
    "270": "KY", "364": "KY", "502": "KY", "606": "KY", "859": "KY",
    "225": "LA", "318": "LA", "337": "LA", "504": "LA", "985": "LA",
    "207": "ME",
    "240": "MD", "301": "MD", "410": "MD", "443": "MD", "667": "MD",
    "339": "MA", "351": "MA", "413": "MA", "508": "MA", "617": "MA",
    "774": "MA", "781": "MA", "857": "MA", "978": "MA",
    "231": "MI", "248": "MI", "269": "MI", "313": "MI", "517": "MI",
    "586": "MI", "616": "MI", "734": "MI", "810": "MI", "906": "MI",
    "947": "MI", "989": "MI",
    "218": "MN", "320": "MN", "507": "MN", "612": "MN", "651": "MN",
    "763": "MN", "952": "MN",
    "228": "MS", "601": "MS", "662": "MS", "769": "MS",
    "314": "MO", "417": "MO", "573": "MO", "636": "MO", "660": "MO",
    "816": "MO", "975": "MO",
    "406": "MT",
    "308": "NE", "402": "NE", "531": "NE",
    "702": "NV", "725": "NV", "775": "NV",
    "603": "NH",
    "201": "NJ", "551": "NJ", "609": "NJ", "640": "NJ", "732": "NJ",
    "848": "NJ", "856": "NJ", "862": "NJ", "908": "NJ", "973": "NJ",
    "505": "NM", "575": "NM",
    "212": "NY", "315": "NY", "332": "NY", "347": "NY", "516": "NY",
    "518": "NY", "585": "NY", "607": "NY", "631": "NY", "646": "NY",
    "680": "NY", "716": "NY", "718": "NY", "838": "NY", "845": "NY",
    "914": "NY", "917": "NY", "929": "NY", "934": "NY",
    "252": "NC", "336": "NC", "704": "NC", "743": "NC", "828": "NC",
    "910": "NC", "919": "NC", "980": "NC", "984": "NC",
    "701": "ND",
    "216": "OH", "220": "OH", "234": "OH", "283": "OH", "330": "OH",
    "380": "OH", "419": "OH", "440": "OH", "513": "OH", "567": "OH",
    "614": "OH", "740": "OH", "937": "OH",
    "405": "OK", "539": "OK", "580": "OK", "918": "OK",
    "458": "OR", "503": "OR", "541": "OR", "971": "OR",
    "215": "PA", "223": "PA", "267": "PA", "272": "PA", "412": "PA",
    "445": "PA", "484": "PA", "570": "PA", "582": "PA", "610": "PA",
    "717": "PA", "724": "PA", "814": "PA", "835": "PA", "878": "PA",
    "401": "RI",
    "803": "SC", "839": "SC", "843": "SC", "854": "SC", "864": "SC",
    "605": "SD",
    "423": "TN", "615": "TN", "629": "TN", "731": "TN", "865": "TN",
    "901": "TN", "931": "TN",
    "210": "TX", "214": "TX", "254": "TX", "281": "TX", "325": "TX",
    "346": "TX", "361": "TX", "409": "TX", "430": "TX", "432": "TX",
    "469": "TX", "512": "TX", "682": "TX", "713": "TX", "726": "TX",
    "737": "TX", "806": "TX", "817": "TX", "830": "TX", "832": "TX",
    "903": "TX", "915": "TX", "936": "TX", "940": "TX", "956": "TX",
    "972": "TX", "979": "TX",
    "385": "UT", "435": "UT", "801": "UT",
    "802": "VT",
    "276": "VA", "434": "VA", "540": "VA", "571": "VA", "703": "VA",
    "757": "VA", "804": "VA", "826": "VA", "948": "VA",
    "206": "WA", "253": "WA", "360": "WA", "425": "WA", "509": "WA",
    "564": "WA",
    "202": "DC",
    "304": "WV", "681": "WV",
    "262": "WI", "274": "WI", "414": "WI", "534": "WI", "608": "WI",
    "715": "WI", "920": "WI",
    "307": "WY",
    "340": "VI", "671": "GU", "684": "AS", "787": "PR", "939": "PR", "670": "MP",
}

TOLL_FREE = {"800", "833", "844", "855", "866", "877", "888"}


def normalize_phone(phone: str) -> str:
    """Normalize phone number to 10 digits."""
    digits = ''.join(c for c in str(phone) if c.isdigit())
    if len(digits) == 11 and digits.startswith('1'):
        return digits[1:]
    return digits[:10] if len(digits) >= 10 else digits


def get_area_code(phone: str) -> str:
    """Extract area code from phone number."""
    normalized = normalize_phone(phone)
    return normalized[:3] if len(normalized) >= 3 else ""


def get_state(phone: str) -> str:
    """Get state from phone number."""
    return NPA_TO_STATE.get(get_area_code(phone), "")


def is_toll_free(phone: str) -> bool:
    """Check if toll-free number."""
    return get_area_code(phone) in TOLL_FREE


def classify_call(source: str, dest: str) -> str:
    """Classify call as interstate/intrastate/toll_free/unknown."""
    if is_toll_free(source) or is_toll_free(dest):
        return "toll_free"
    src_state = get_state(source)
    dst_state = get_state(dest)
    if not src_state or not dst_state:
        return "unknown"
    return "intrastate" if src_state == dst_state else "interstate"


def get_customer_name(domain: str) -> str:
    """Extract clean customer name from domain."""
    # "AdamsCoIL.20507.service" -> "AdamsCoIL"
    if not domain:
        return "Unknown"
    parts = domain.split('.')
    return parts[0] if parts else domain


# ============================================================================
# Load phone number to customer mapping
# ============================================================================
def load_phone_mapping(phonenumbers_csv: Path) -> dict:
    """Load phone number to customer (domain) mapping."""
    mapping = {}
    with open(phonenumbers_csv, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            phone = normalize_phone(row.get('Phone Number', ''))
            domain = row.get('Domain', '')
            if phone and domain:
                mapping[phone] = get_customer_name(domain)
    return mapping


# ============================================================================
# Report 1: Call Ratio by Customer
# ============================================================================
@dataclass
class CustomerStats:
    """Statistics for a single customer."""
    name: str
    total_calls: int = 0
    total_seconds: float = 0.0
    total_cost: float = 0.0  # Original cost from CDR
    interstate_calls: int = 0
    interstate_seconds: float = 0.0
    intrastate_calls: int = 0
    intrastate_seconds: float = 0.0
    toll_free_calls: int = 0
    toll_free_seconds: float = 0.0
    unknown_calls: int = 0
    unknown_seconds: float = 0.0
    phone_numbers: set = field(default_factory=set)

    @property
    def total_minutes(self) -> float:
        """Total minutes of calls."""
        return self.total_seconds / 60

    @property
    def billable_cost(self) -> float:
        """Calculate billable cost at $0.005/minute rate."""
        return self.total_minutes * VOICE_RATE_PER_MINUTE

    @property
    def interstate_minutes(self) -> float:
        return self.interstate_seconds / 60

    @property
    def intrastate_minutes(self) -> float:
        return self.intrastate_seconds / 60


def generate_cdr_report(
    cdr_file: Path,
    phone_mapping: dict,
    output_file: Optional[Path] = None
) -> dict[str, CustomerStats]:
    """
    Generate CDR report with call ratios by customer.

    Args:
        cdr_file: Vitelity CDR CSV
        phone_mapping: Phone number to customer mapping
        output_file: Optional CSV output

    Returns:
        Dict of customer name to CustomerStats
    """
    customers: dict[str, CustomerStats] = defaultdict(lambda: CustomerStats(name=""))

    with open(cdr_file, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)

        for row in reader:
            source = row.get('Source', '').strip()
            dest = row.get('Destination', '').strip()
            seconds = float(row.get('Seconds', 0) or 0)
            cost = float(row.get('Cost', 0) or 0)

            if seconds <= 0:
                continue

            # Determine customer from source or destination
            src_normalized = normalize_phone(source)
            dst_normalized = normalize_phone(dest)

            customer = phone_mapping.get(src_normalized) or phone_mapping.get(dst_normalized) or "Unassigned"

            # Initialize customer if needed
            if customers[customer].name == "":
                customers[customer].name = customer

            # Update stats
            stats = customers[customer]
            stats.total_calls += 1
            stats.total_seconds += seconds
            stats.total_cost += cost

            # Track phone numbers
            if src_normalized in phone_mapping:
                stats.phone_numbers.add(src_normalized)
            if dst_normalized in phone_mapping:
                stats.phone_numbers.add(dst_normalized)

            # Classify call
            call_type = classify_call(source, dest)
            if call_type == "interstate":
                stats.interstate_calls += 1
                stats.interstate_seconds += seconds
            elif call_type == "intrastate":
                stats.intrastate_calls += 1
                stats.intrastate_seconds += seconds
            elif call_type == "toll_free":
                stats.toll_free_calls += 1
                stats.toll_free_seconds += seconds
            else:
                stats.unknown_calls += 1
                stats.unknown_seconds += seconds

    # Output CSV if requested
    if output_file:
        with open(output_file, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow([
                'Customer', 'Total Calls', 'Total Minutes', 'CDR Cost', 'Billable Cost',
                'Interstate Calls', 'Interstate Min', 'Intrastate Calls', 'Intrastate Min',
                'Interstate %', 'Intrastate %', 'Phone Numbers'
            ])

            for name, stats in sorted(customers.items(), key=lambda x: x[1].total_seconds, reverse=True):
                jurisdictional = stats.interstate_seconds + stats.intrastate_seconds
                interstate_pct = (stats.interstate_seconds / jurisdictional * 100) if jurisdictional > 0 else 0
                intrastate_pct = (stats.intrastate_seconds / jurisdictional * 100) if jurisdictional > 0 else 0

                writer.writerow([
                    name,
                    stats.total_calls,
                    f"{stats.total_minutes:.2f}",
                    f"{stats.total_cost:.4f}",
                    f"{stats.billable_cost:.4f}",
                    stats.interstate_calls,
                    f"{stats.interstate_minutes:.2f}",
                    stats.intrastate_calls,
                    f"{stats.intrastate_minutes:.2f}",
                    f"{interstate_pct:.2f}%",
                    f"{intrastate_pct:.2f}%",
                    len(stats.phone_numbers)
                ])

    return dict(customers)


# ============================================================================
# Report 2: Phone Number Count by Customer
# ============================================================================
def generate_phone_count_report(
    phonenumbers_csv: Path,
    output_file: Optional[Path] = None,
    excluded_file: Optional[Path] = None
) -> tuple[dict, dict]:
    """
    Generate phone number count by customer, excluding non-billable phones.

    Non-billable phones (fax, on-hold, unassigned) are tracked separately
    and can be output to excluded_file for the InvOther report.

    Returns:
        Tuple of (billable_counts, excluded_counts) dicts
    """
    billable_counts = defaultdict(int)
    excluded_counts = defaultdict(int)
    excluded_rows = []

    with open(phonenumbers_csv, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            domain = row.get('Domain', '')
            treatment = row.get('Treatment', '').strip()
            customer = get_customer_name(domain)

            # Check if this is a non-billable treatment
            is_excluded = treatment in NON_BILLABLE_TREATMENTS

            # Also check for fax/hold keywords in treatment (catch-all)
            treatment_lower = treatment.lower()
            if 'fax' in treatment_lower or 'hold' in treatment_lower:
                is_excluded = True

            if is_excluded:
                excluded_counts[customer] += 1
                excluded_rows.append(row)
            else:
                billable_counts[customer] += 1

    # Output billable phone counts
    if output_file:
        with open(output_file, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Customer', 'Billable Phone Count'])
            for customer, count in sorted(billable_counts.items(), key=lambda x: x[1], reverse=True):
                writer.writerow([customer, count])

    # Output excluded phones (InvOther)
    if excluded_file and excluded_rows:
        with open(excluded_file, 'w', newline='') as f:
            fieldnames = ['Phone Number', 'Domain', 'Treatment', 'Destination', 'Notes', 'Enable']
            writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
            writer.writeheader()
            writer.writerows(excluded_rows)

    return dict(billable_counts), dict(excluded_counts)


# ============================================================================
# Report 3: CallerID Lookup Counts
# ============================================================================
def generate_callerid_report(cdr_file: Path, output_file: Optional[Path] = None) -> dict:
    """Count calls per destination number (for CallerID billing)."""
    counts = defaultdict(int)

    with open(cdr_file, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            dest = normalize_phone(row.get('Destination', ''))
            if dest:
                counts[dest] += 1

    if output_file:
        with open(output_file, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Phone Number', 'Call Count'])
            for phone, count in sorted(counts.items(), key=lambda x: x[1], reverse=True):
                writer.writerow([phone, count])

    return dict(counts)


# ============================================================================
# Report 4: SMS Usage Report
# ============================================================================
@dataclass
class SMSStats:
    """SMS statistics for a customer or overall."""
    total_messages: int = 0
    incoming_messages: int = 0
    outgoing_messages: int = 0
    total_cost: float = 0.0

    @property
    def billable_cost(self) -> float:
        """Calculate billable cost at SMS rate."""
        return self.total_messages * SMS_RATE_PER_MESSAGE


def generate_sms_report(
    sms_file: Path,
    phone_mapping: dict,
    output_file: Optional[Path] = None
) -> tuple[dict[str, SMSStats], SMSStats]:
    """
    Generate SMS usage report by customer.

    Args:
        sms_file: SMS CSV file (Time, Source, Destination, msgDirection, Cost)
        phone_mapping: Phone number to customer mapping
        output_file: Optional CSV output

    Returns:
        Tuple of (customer stats dict, overall stats)
    """
    customers: dict[str, SMSStats] = defaultdict(SMSStats)
    overall = SMSStats()

    with open(sms_file, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)

        for row in reader:
            source = normalize_phone(row.get('Source', ''))
            dest = normalize_phone(row.get('Destination', ''))
            direction = row.get('msgDirection', '').lower()
            try:
                cost = float(row.get('Cost', 0) or 0)
            except ValueError:
                cost = 0.0

            # Determine customer
            customer = phone_mapping.get(source) or phone_mapping.get(dest) or "Unassigned"

            # Update customer stats
            stats = customers[customer]
            stats.total_messages += 1
            stats.total_cost += cost
            if direction == 'incoming':
                stats.incoming_messages += 1
            else:
                stats.outgoing_messages += 1

            # Update overall
            overall.total_messages += 1
            overall.total_cost += cost
            if direction == 'incoming':
                overall.incoming_messages += 1
            else:
                overall.outgoing_messages += 1

    # Output CSV if requested
    if output_file:
        with open(output_file, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow([
                'Customer', 'Total Messages', 'Incoming', 'Outgoing',
                'CDR Cost', 'Billable Cost'
            ])

            for name, stats in sorted(customers.items(), key=lambda x: x[1].total_messages, reverse=True):
                writer.writerow([
                    name,
                    stats.total_messages,
                    stats.incoming_messages,
                    stats.outgoing_messages,
                    f"{stats.total_cost:.4f}",
                    f"{stats.billable_cost:.4f}"
                ])

    return dict(customers), overall


# ============================================================================
# Report 5: Combined CDR (SkySwitch + Vitelity)
# ============================================================================
def extract_skyswitch_cdr(
    master_xlsx: Path,
    sheet_index: int = 26,  # CDR<i> sheet
    output_csv: Optional[Path] = None
) -> list[dict]:
    """
    Extract SkySwitch CDR from master Excel file.

    Args:
        master_xlsx: Path to master Excel file (CDR SS records-xxx.xlsx)
        sheet_index: Sheet number (1-indexed, default 26 for CDR<i>)
        output_csv: Optional path to save extracted CDR as CSV

    Returns:
        List of CDR records as dicts
    """
    import zipfile
    import xml.etree.ElementTree as ET

    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    records = []

    with zipfile.ZipFile(master_xlsx) as z:
        # Get shared strings
        try:
            ss_xml = ET.parse(z.open('xl/sharedStrings.xml')).getroot()
            strings = []
            for si in ss_xml.findall('.//main:si', ns):
                t_el = si.find('.//main:t', ns)
                strings.append(t_el.text if t_el is not None else '')
        except:
            strings = []

        # Read the CDR sheet
        sheet_xml = z.read(f'xl/worksheets/sheet{sheet_index}.xml')
        root = ET.fromstring(sheet_xml)

        rows = root.findall('.//main:row', ns)
        headers = []

        for row in rows:
            row_num = int(row.get('r'))
            row_data = {}

            for cell in row.findall('main:c', ns):
                ref = cell.get('r')
                col = ''.join(c for c in ref if c.isalpha())
                cell_type = cell.get('t')

                val_el = cell.find('main:v', ns)
                val = val_el.text if val_el is not None else ''

                # If shared string, look it up
                if cell_type == 's' and val and strings:
                    idx = int(val)
                    val = strings[idx] if idx < len(strings) else val

                row_data[col] = val

            # First row is headers
            if row_num == 1:
                headers = row_data
                continue

            # Extract relevant fields
            # Column mapping: D=From, E=Dialed, F=To, I=Duration, J=Domain
            source = row_data.get('D', '')
            destination = row_data.get('F', '') or row_data.get('E', '')
            duration_str = row_data.get('I', '0')

            try:
                duration = float(duration_str) if duration_str else 0
            except ValueError:
                duration = 0

            domain = row_data.get('J', '')
            customer = get_customer_name(domain)

            if duration > 0:  # Only include calls with duration
                records.append({
                    'Source': source,
                    'Destination': destination,
                    'Seconds': duration,
                    'Customer': customer,
                    'CDR_Source': 'SkySwitch'
                })

    # Write to CSV if requested
    if output_csv and records:
        with open(output_csv, 'w', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=['Source', 'Destination', 'Seconds', 'Customer', 'CDR_Source'])
            writer.writeheader()
            writer.writerows(records)

    return records


def generate_combined_cdr_report(
    vitelity_cdr: Path,
    skyswitch_xlsx: Path,
    phone_mapping: dict,
    output_file: Optional[Path] = None
) -> dict[str, CustomerStats]:
    """
    Generate combined CDR report from Vitelity CSV and SkySwitch Excel.

    Args:
        vitelity_cdr: Vitelity CDR CSV file
        skyswitch_xlsx: Master Excel file containing SkySwitch CDR
        phone_mapping: Phone number to customer mapping
        output_file: Optional combined CSV output

    Returns:
        Dict of customer name to combined CustomerStats
    """
    customers: dict[str, CustomerStats] = defaultdict(lambda: CustomerStats(name=""))

    # Process Vitelity CDR
    with open(vitelity_cdr, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            source = row.get('Source', '').strip()
            dest = row.get('Destination', '').strip()
            seconds = float(row.get('Seconds', 0) or 0)
            cost = float(row.get('Cost', 0) or 0)

            if seconds <= 0:
                continue

            src_normalized = normalize_phone(source)
            dst_normalized = normalize_phone(dest)
            customer = phone_mapping.get(src_normalized) or phone_mapping.get(dst_normalized) or "Unassigned"

            if customers[customer].name == "":
                customers[customer].name = customer

            stats = customers[customer]
            stats.total_calls += 1
            stats.total_seconds += seconds
            stats.total_cost += cost

            call_type = classify_call(source, dest)
            if call_type == "interstate":
                stats.interstate_calls += 1
                stats.interstate_seconds += seconds
            elif call_type == "intrastate":
                stats.intrastate_calls += 1
                stats.intrastate_seconds += seconds
            elif call_type == "toll_free":
                stats.toll_free_calls += 1
                stats.toll_free_seconds += seconds
            else:
                stats.unknown_calls += 1
                stats.unknown_seconds += seconds

    # Extract and process SkySwitch CDR
    print(f"    Extracting SkySwitch CDR from Excel (this may take a moment)...")
    skyswitch_records = extract_skyswitch_cdr(skyswitch_xlsx)
    print(f"    Extracted {len(skyswitch_records):,} SkySwitch records")

    for record in skyswitch_records:
        source = record['Source']
        dest = record['Destination']
        seconds = record['Seconds']
        customer = record['Customer']

        if not customer or customer == "Unknown":
            src_normalized = normalize_phone(source)
            dst_normalized = normalize_phone(dest)
            customer = phone_mapping.get(src_normalized) or phone_mapping.get(dst_normalized) or "Unassigned"

        if customers[customer].name == "":
            customers[customer].name = customer

        stats = customers[customer]
        stats.total_calls += 1
        stats.total_seconds += seconds
        # SkySwitch cost calculated at our rate
        stats.total_cost += (seconds / 60) * VOICE_RATE_PER_MINUTE

        call_type = classify_call(source, dest)
        if call_type == "interstate":
            stats.interstate_calls += 1
            stats.interstate_seconds += seconds
        elif call_type == "intrastate":
            stats.intrastate_calls += 1
            stats.intrastate_seconds += seconds
        elif call_type == "toll_free":
            stats.toll_free_calls += 1
            stats.toll_free_seconds += seconds
        else:
            stats.unknown_calls += 1
            stats.unknown_seconds += seconds

    # Output CSV if requested
    if output_file:
        with open(output_file, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow([
                'Customer', 'Total Calls', 'Total Minutes', 'Billable Cost',
                'Interstate Calls', 'Interstate Min', 'Intrastate Calls', 'Intrastate Min',
                'Interstate %', 'Intrastate %'
            ])

            for name, stats in sorted(customers.items(), key=lambda x: x[1].total_seconds, reverse=True):
                jurisdictional = stats.interstate_seconds + stats.intrastate_seconds
                interstate_pct = (stats.interstate_seconds / jurisdictional * 100) if jurisdictional > 0 else 0
                intrastate_pct = (stats.intrastate_seconds / jurisdictional * 100) if jurisdictional > 0 else 0

                writer.writerow([
                    name,
                    stats.total_calls,
                    f"{stats.total_minutes:.2f}",
                    f"{stats.billable_cost:.4f}",
                    stats.interstate_calls,
                    f"{stats.interstate_minutes:.2f}",
                    stats.intrastate_calls,
                    f"{stats.intrastate_minutes:.2f}",
                    f"{interstate_pct:.2f}%",
                    f"{intrastate_pct:.2f}%",
                ])

    return dict(customers)


# ============================================================================
# Report 6: Seat Count from Domain Statistics (XLSX)
# ============================================================================
def generate_seat_count_report(
    domain_stats_xlsx: Path,
    output_file: Optional[Path] = None
) -> dict[str, dict]:
    """
    Generate seat count (PBX Users) report from Domain-Statistics XLSX.

    Args:
        domain_stats_xlsx: Path to Domain-Statistics-YYYY-MM-DD.xlsx
        output_file: Optional CSV output

    Returns:
        Dict of customer name to stats dict
    """
    import zipfile
    import xml.etree.ElementTree as ET

    customers = {}
    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

    with zipfile.ZipFile(domain_stats_xlsx) as z:
        sheet_xml = z.read('xl/worksheets/sheet1.xml')
        root = ET.fromstring(sheet_xml)

        rows = root.findall('.//main:row', ns)
        headers = []

        for row in rows:
            row_num = int(row.get('r'))
            row_data = {}

            for cell in row.findall('main:c', ns):
                ref = cell.get('r')
                col = ''.join(c for c in ref if c.isalpha())

                # Get value (inline string or value)
                is_el = cell.find('main:is', ns)
                if is_el is not None:
                    t_el = is_el.find('main:t', ns)
                    val = t_el.text if t_el is not None else ''
                else:
                    val_el = cell.find('main:v', ns)
                    val = val_el.text if val_el is not None else ''

                row_data[col] = val

            # First row is headers
            if row_num == 1:
                headers = row_data
                continue

            # Skip total row
            domain = row_data.get('A', '')
            if domain == 'Total' or not domain:
                continue

            # Extract customer name from domain
            customer = get_customer_name(domain)

            # Get stats
            try:
                pbx_users = int(row_data.get('B', 0) or 0)
                call_center = int(row_data.get('C', 0) or 0)
                call_recording = int(row_data.get('D', 0) or 0)
                sip_trunks = int(row_data.get('E', 0) or 0)
                meeting_rooms = int(row_data.get('F', 0) or 0)
                vm_transcription = int(row_data.get('G', 0) or 0)
                phone_numbers = int(row_data.get('H', 0) or 0)
                teams_connectors = int(row_data.get('I', 0) or 0)
                video_connectors = int(row_data.get('J', 0) or 0)
            except ValueError:
                continue

            customers[customer] = {
                'domain': domain,
                'pbx_users': pbx_users,
                'call_center': call_center,
                'call_recording': call_recording,
                'sip_trunks': sip_trunks,
                'meeting_rooms': meeting_rooms,
                'vm_transcription': vm_transcription,
                'phone_numbers': phone_numbers,
                'teams_connectors': teams_connectors,
                'video_connectors': video_connectors,
            }

    # Output CSV if requested
    if output_file:
        with open(output_file, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow([
                'Customer', 'PBX Users (Seats)', 'Call Center', 'Call Recording',
                'SIP Trunks', 'Meeting Rooms', 'VM Transcription', 'Phone Numbers',
                'Teams Connectors', 'Video Connectors'
            ])

            for name, stats in sorted(customers.items(), key=lambda x: x[1]['pbx_users'], reverse=True):
                writer.writerow([
                    name,
                    stats['pbx_users'],
                    stats['call_center'],
                    stats['call_recording'],
                    stats['sip_trunks'],
                    stats['meeting_rooms'],
                    stats['vm_transcription'],
                    stats['phone_numbers'],
                    stats['teams_connectors'],
                    stats['video_connectors'],
                ])

    return customers


# ============================================================================
# Report 7: Adams County User Summary (Pivot Table)
# ============================================================================
def generate_adams_county_report(
    master_xlsx: Path,
    output_file: Optional[Path] = None
) -> dict:
    """
    Generate Adams County User Summary pivot table from Master Excel.

    Extracts data from 'Copied user_export_AdamsCoIL' sheet (sheet10) and
    creates a pivot table counting extensions by Department and UserType.

    Args:
        master_xlsx: Path to master Excel file (CDR SS records-xxx.xlsx)
        output_file: Optional path to save pivot table as CSV

    Returns:
        Dict with pivot data: {dept: {user_type: count}}
    """
    import zipfile
    import xml.etree.ElementTree as ET

    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

    # Data structure for pivot
    pivot_data = defaultdict(lambda: defaultdict(int))
    user_types = ['u', 'nu', 'nb', 'vm only', 'faxata']

    with zipfile.ZipFile(master_xlsx) as z:
        # Get shared strings
        try:
            ss_xml = ET.parse(z.open('xl/sharedStrings.xml')).getroot()
            strings = []
            for si in ss_xml.findall('.//main:si', ns):
                t_el = si.find('.//main:t', ns)
                strings.append(t_el.text if t_el is not None else '')
        except:
            strings = []

        # Try to find the Adams County user export sheet
        # First check sheet10 (Copied user_export_AdamsCoIL)
        sheet_found = False
        for sheet_num in [10, 11]:  # Try Copied version first, then original
            try:
                sheet_xml = z.read(f'xl/worksheets/sheet{sheet_num}.xml')
                root = ET.fromstring(sheet_xml)
                rows = root.findall('.//main:row', ns)

                if not rows:
                    continue

                # Check if this is the Adams County sheet by looking at header
                header_row = rows[0]
                has_dept = False
                has_usertype = False

                for cell in header_row.findall('main:c', ns):
                    ref = cell.get('r')
                    col = ''.join(c for c in ref if c.isalpha())
                    cell_type = cell.get('t')
                    val_el = cell.find('main:v', ns)
                    val = val_el.text if val_el is not None else ''

                    if cell_type == 's' and val and strings:
                        idx = int(val)
                        val = strings[idx] if idx < len(strings) else val

                    if val.lower() == 'department':
                        has_dept = True
                    if val.lower() == 'usertype':
                        has_usertype = True

                if has_dept:
                    sheet_found = True
                    break
            except:
                continue

        if not sheet_found:
            print("    Warning: Could not find Adams County user export sheet")
            return {}

        # Process rows to build pivot
        for row in rows[1:]:  # Skip header
            row_data = {}
            for cell in row.findall('main:c', ns):
                ref = cell.get('r')
                col = ''.join(c for c in ref if c.isalpha())
                cell_type = cell.get('t')
                val_el = cell.find('main:v', ns)
                val = val_el.text if val_el is not None else ''

                if cell_type == 's' and val and strings:
                    idx = int(val)
                    val = strings[idx] if idx < len(strings) else val

                row_data[col] = val

            # Department is column I, UserType is column AA
            dept = row_data.get('I', '').strip()
            user_type = row_data.get('AA', '').strip()

            if dept:
                pivot_data[dept][user_type] += 1

    # Output CSV if requested
    if output_file and pivot_data:
        with open(output_file, 'w', newline='') as f:
            writer = csv.writer(f)

            # Header row
            writer.writerow(['Department'] + user_types + ['Grand Total'])

            # Data rows sorted by department
            grand_totals = defaultdict(int)
            for dept in sorted(pivot_data.keys()):
                row = [dept]
                row_total = 0
                for ut in user_types:
                    count = pivot_data[dept].get(ut, 0)
                    row.append(count if count > 0 else '')
                    row_total += count
                    grand_totals[ut] += count
                row.append(row_total)
                writer.writerow(row)

            # Grand total row
            total_row = ['Grand Total']
            overall_total = 0
            for ut in user_types:
                total_row.append(grand_totals[ut])
                overall_total += grand_totals[ut]
            total_row.append(overall_total)
            writer.writerow(total_row)

            # Summary calculations
            writer.writerow([])  # Empty row

            # Lines Calculation (billable = u + vm only)
            billable_users = grand_totals.get('u', 0) + grand_totals.get('vm only', 0)
            writer.writerow(['Lines Calculation', '', '', '', billable_users])

            # High Value for Month Users (total active)
            active_users = grand_totals.get('u', 0) + grand_totals.get('nu', 0) + grand_totals.get('vm only', 0)
            writer.writerow(['High Value for Month Users', '', '', '', active_users])

    return dict(pivot_data)


# ============================================================================
# Main
# ============================================================================
def print_summary(customers: dict[str, CustomerStats]):
    """Print summary report to console."""
    print("\n" + "=" * 80)
    print("  BILLING REPORT SUMMARY - BY CUSTOMER")
    print("=" * 80)

    # Calculate totals
    total_calls = sum(c.total_calls for c in customers.values())
    total_minutes = sum(c.total_minutes for c in customers.values())
    total_cdr_cost = sum(c.total_cost for c in customers.values())
    total_billable = sum(c.billable_cost for c in customers.values())
    total_interstate = sum(c.interstate_seconds for c in customers.values())
    total_intrastate = sum(c.intrastate_seconds for c in customers.values())

    print(f"\n  TOTALS")
    print(f"  {'─' * 40}")
    print(f"  Total Calls:     {total_calls:,}")
    print(f"  Total Minutes:   {total_minutes:,.2f}")
    print(f"  CDR Cost:        ${total_cdr_cost:,.4f}")
    print(f"  Billable Cost:   ${total_billable:,.4f}  (@ ${VOICE_RATE_PER_MINUTE}/min)")

    jurisdictional = total_interstate + total_intrastate
    if jurisdictional > 0:
        print(f"\n  ╔═══════════════════════════════════════╗")
        print(f"  ║  OVERALL CALL RATIO                   ║")
        print(f"  ╠═══════════════════════════════════════╣")
        print(f"  ║  Interstate:  {total_interstate/jurisdictional*100:6.2f}%               ║")
        print(f"  ║  Intrastate:  {total_intrastate/jurisdictional*100:6.2f}%               ║")
        print(f"  ╚═══════════════════════════════════════╝")

    print(f"\n  TOP CUSTOMERS BY MINUTES")
    print(f"  {'─' * 90}")
    print(f"  {'Customer':<20} {'Calls':>8} {'Minutes':>10} {'Billable':>10} {'Interstate':>12} {'Intrastate':>12}")
    print(f"  {'─' * 90}")

    sorted_customers = sorted(customers.values(), key=lambda x: x.total_seconds, reverse=True)

    for stats in sorted_customers[:15]:
        jur = stats.interstate_seconds + stats.intrastate_seconds
        inter_pct = f"{stats.interstate_seconds/jur*100:.1f}%" if jur > 0 else "N/A"
        intra_pct = f"{stats.intrastate_seconds/jur*100:.1f}%" if jur > 0 else "N/A"

        print(f"  {stats.name:<20} {stats.total_calls:>8,} {stats.total_minutes:>10,.1f} ${stats.billable_cost:>8,.2f} {inter_pct:>12} {intra_pct:>12}")

    print(f"  {'─' * 90}")
    print()


def main():
    """Main entry point."""
    if len(sys.argv) < 3:
        print("Usage: python billing_reports.py <cdr_file.csv> <phonenumbers.csv> [output_dir] [sms_file.csv] [domain_stats.xlsx] [master_xlsx]")
        print("\nExample:")
        print("  python billing_reports.py http-xxx.csv phonenumbers__xxx.csv ./reports syneteks-xxx.csv Domain-Statistics-xxx.xlsx \"CDR SS records-xxx.xlsx\"")
        print("\nIf master_xlsx is provided, a combined CDR report (SkySwitch + Vitelity) will be generated.")
        sys.exit(1)

    cdr_file = Path(sys.argv[1])
    phonenumbers_file = Path(sys.argv[2])
    output_dir = Path(sys.argv[3]) if len(sys.argv) > 3 and sys.argv[3] else Path("./reports")
    sms_file = Path(sys.argv[4]) if len(sys.argv) > 4 and sys.argv[4] else None
    domain_stats_file = Path(sys.argv[5]) if len(sys.argv) > 5 and sys.argv[5] else None
    master_xlsx_file = Path(sys.argv[6]) if len(sys.argv) > 6 and sys.argv[6] else None

    if not cdr_file.exists():
        print(f"Error: CDR file not found: {cdr_file}")
        sys.exit(1)

    if not phonenumbers_file.exists():
        print(f"Error: Phone numbers file not found: {phonenumbers_file}")
        sys.exit(1)

    # Create output directory
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"Loading phone number mapping from {phonenumbers_file.name}...")
    phone_mapping = load_phone_mapping(phonenumbers_file)
    print(f"  Loaded {len(phone_mapping)} phone numbers")

    print(f"\nProcessing CDR file {cdr_file.name}...")
    customers = generate_cdr_report(
        cdr_file,
        phone_mapping,
        output_dir / "cdr.csv"
    )
    print(f"  Generated: cdr.csv")

    print(f"\nGenerating phone count report...")
    phone_counts, excluded_counts = generate_phone_count_report(
        phonenumbers_file,
        output_dir / "phones.csv",
        output_dir / "phones_excl.csv"
    )
    print(f"  Generated: phones.csv (billable phones)")
    total_billable = sum(phone_counts.values())
    total_excluded = sum(excluded_counts.values())
    if total_excluded > 0:
        print(f"  Generated: phones_excl.csv ({total_excluded} non-billable phones)")
    print(f"  Billable phones: {total_billable}, Excluded (fax/hold/unassigned): {total_excluded}")

    print(f"\nGenerating CallerID report...")
    callerid_counts = generate_callerid_report(
        cdr_file,
        output_dir / "callerid.csv"
    )
    print(f"  Generated: callerid.csv")

    # Generate SMS report if file provided
    sms_overall = None
    if sms_file and sms_file.exists():
        print(f"\nProcessing SMS file {sms_file.name}...")
        sms_customers, sms_overall = generate_sms_report(
            sms_file,
            phone_mapping,
            output_dir / "sms.csv"
        )
        print(f"  Generated: sms.csv")
    elif sms_file:
        print(f"\nWarning: SMS file not found: {sms_file}")

    # Print summary
    print_summary(customers)

    # Print SMS summary if available
    if sms_overall and sms_overall.total_messages > 0:
        print(f"  SMS SUMMARY")
        print(f"  {'─' * 40}")
        print(f"  Total Messages:  {sms_overall.total_messages:,}")
        print(f"  Incoming:        {sms_overall.incoming_messages:,}")
        print(f"  Outgoing:        {sms_overall.outgoing_messages:,}")
        print(f"  CDR Cost:        ${sms_overall.total_cost:,.4f}")
        print(f"  Billable Cost:   ${sms_overall.billable_cost:,.4f}  (@ ${SMS_RATE_PER_MESSAGE}/msg)")
        print()

    # Generate combined CDR report if master xlsx provided
    combined_customers = None
    if master_xlsx_file and master_xlsx_file.exists():
        print(f"\nGenerating Combined CDR Report (SkySwitch + Vitelity)...")
        combined_customers = generate_combined_cdr_report(
            cdr_file,
            master_xlsx_file,
            phone_mapping,
            output_dir / "cdr_combined.csv"
        )
        print(f"  Generated: cdr_combined.csv")

        # Print combined summary
        total_calls = sum(c.total_calls for c in combined_customers.values())
        total_minutes = sum(c.total_minutes for c in combined_customers.values())
        total_billable = sum(c.billable_cost for c in combined_customers.values())
        total_interstate = sum(c.interstate_seconds for c in combined_customers.values())
        total_intrastate = sum(c.intrastate_seconds for c in combined_customers.values())

        print(f"\n  COMBINED CDR SUMMARY (SkySwitch + Vitelity)")
        print(f"  {'─' * 50}")
        print(f"  Total Calls:     {total_calls:,}")
        print(f"  Total Minutes:   {total_minutes:,.2f}")
        print(f"  Billable Cost:   ${total_billable:,.4f}  (@ ${VOICE_RATE_PER_MINUTE}/min)")

        jurisdictional = total_interstate + total_intrastate
        if jurisdictional > 0:
            print(f"\n  ╔═══════════════════════════════════════╗")
            print(f"  ║  COMBINED CALL RATIO                  ║")
            print(f"  ╠═══════════════════════════════════════╣")
            print(f"  ║  Interstate:  {total_interstate/jurisdictional*100:6.2f}%               ║")
            print(f"  ║  Intrastate:  {total_intrastate/jurisdictional*100:6.2f}%               ║")
            print(f"  ╚═══════════════════════════════════════╝")
        print()

        # Generate Adams County User Summary if data exists
        print(f"  Generating Adams County User Summary...")
        adams_pivot = generate_adams_county_report(
            master_xlsx_file,
            output_dir / "adams_co.csv"
        )
        if adams_pivot:
            print(f"  Generated: adams_co.csv")

            # Print Adams County summary
            total_u = sum(d.get('u', 0) for d in adams_pivot.values())
            total_nu = sum(d.get('nu', 0) for d in adams_pivot.values())
            total_nb = sum(d.get('nb', 0) for d in adams_pivot.values())
            total_vm = sum(d.get('vm only', 0) for d in adams_pivot.values())
            total_fax = sum(d.get('faxata', 0) for d in adams_pivot.values())
            total_ext = total_u + total_nu + total_nb + total_vm + total_fax

            print(f"\n  ADAMS COUNTY USER SUMMARY")
            print(f"  {'─' * 50}")
            print(f"  {'User Type':<15} {'Count':>8}")
            print(f"  {'─' * 50}")
            print(f"  {'u (user)':<15} {total_u:>8,}")
            print(f"  {'nu (not used)':<15} {total_nu:>8,}")
            print(f"  {'nb (not billed)':<15} {total_nb:>8,}")
            print(f"  {'vm only':<15} {total_vm:>8,}")
            print(f"  {'faxata':<15} {total_fax:>8,}")
            print(f"  {'─' * 50}")
            print(f"  {'TOTAL':<15} {total_ext:>8,}")
            print(f"  {'─' * 50}")
            print(f"  Billable (u + vm only): {total_u + total_vm:,}")
            print()

    elif master_xlsx_file:
        print(f"\nWarning: Master Excel file not found: {master_xlsx_file}")

    # Generate seat count report if domain stats file provided
    seat_counts = None
    if domain_stats_file and domain_stats_file.exists():
        print(f"\nProcessing Domain Statistics {domain_stats_file.name}...")
        seat_counts = generate_seat_count_report(
            domain_stats_file,
            output_dir / "seats.csv"
        )
        print(f"  Generated: seats.csv")

        # Print seat count summary
        total_seats = sum(c['pbx_users'] for c in seat_counts.values())
        print(f"\n  SEAT COUNT SUMMARY")
        print(f"  {'─' * 40}")
        print(f"  Total PBX Users (Seats): {total_seats:,}")
        print(f"  Customers with seats:    {sum(1 for c in seat_counts.values() if c['pbx_users'] > 0):,}")
        print()

        # Top 10 by seats
        print(f"  TOP 10 CUSTOMERS BY SEAT COUNT")
        print(f"  {'─' * 50}")
        print(f"  {'Customer':<25} {'Seats':>8} {'Phones':>8}")
        print(f"  {'─' * 50}")
        sorted_seats = sorted(seat_counts.items(), key=lambda x: x[1]['pbx_users'], reverse=True)
        for name, stats in sorted_seats[:10]:
            print(f"  {name:<25} {stats['pbx_users']:>8,} {stats['phone_numbers']:>8,}")
        print(f"  {'─' * 50}")
        print()

    elif domain_stats_file:
        print(f"\nWarning: Domain stats file not found: {domain_stats_file}")

    print(f"\nReports saved to: {output_dir.absolute()}")


if __name__ == "__main__":
    main()
