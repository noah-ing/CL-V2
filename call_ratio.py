#!/usr/bin/env python3
"""
Call Ratio Calculator

Simple tool that does ONE thing:
- Reads CDR file
- Classifies each call as Interstate or Intrastate based on area codes
- Calculates the ratio

That's it. No complicated Excel automation.
"""

import csv
import sys
from pathlib import Path
from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal

# NPA (Area Code) to State mapping - covers all US area codes
NPA_TO_STATE = {
    # Alabama
    "205": "AL", "251": "AL", "256": "AL", "334": "AL", "938": "AL",
    # Alaska
    "907": "AK",
    # Arizona
    "480": "AZ", "520": "AZ", "602": "AZ", "623": "AZ", "928": "AZ",
    # Arkansas
    "479": "AR", "501": "AR", "870": "AR",
    # California
    "209": "CA", "213": "CA", "310": "CA", "323": "CA", "341": "CA",
    "408": "CA", "415": "CA", "424": "CA", "442": "CA", "510": "CA",
    "530": "CA", "559": "CA", "562": "CA", "619": "CA", "626": "CA",
    "628": "CA", "650": "CA", "657": "CA", "661": "CA", "669": "CA",
    "707": "CA", "714": "CA", "747": "CA", "760": "CA", "805": "CA",
    "818": "CA", "820": "CA", "831": "CA", "840": "CA", "858": "CA",
    "909": "CA", "916": "CA", "925": "CA", "949": "CA", "951": "CA",
    # Colorado
    "303": "CO", "719": "CO", "720": "CO", "970": "CO", "983": "CO",
    # Connecticut
    "203": "CT", "475": "CT", "860": "CT", "959": "CT",
    # Delaware
    "302": "DE",
    # Florida
    "239": "FL", "305": "FL", "321": "FL", "352": "FL", "386": "FL",
    "407": "FL", "561": "FL", "727": "FL", "754": "FL", "772": "FL",
    "786": "FL", "813": "FL", "850": "FL", "863": "FL", "904": "FL",
    "941": "FL", "954": "FL",
    # Georgia
    "229": "GA", "404": "GA", "470": "GA", "478": "GA", "678": "GA",
    "706": "GA", "762": "GA", "770": "GA", "912": "GA", "943": "GA",
    # Hawaii
    "808": "HI",
    # Idaho
    "208": "ID", "986": "ID",
    # Illinois
    "217": "IL", "224": "IL", "309": "IL", "312": "IL", "331": "IL",
    "618": "IL", "630": "IL", "708": "IL", "773": "IL", "779": "IL",
    "815": "IL", "847": "IL", "872": "IL",
    # Indiana
    "219": "IN", "260": "IN", "317": "IN", "463": "IN", "574": "IN",
    "765": "IN", "812": "IN", "930": "IN",
    # Iowa
    "319": "IA", "515": "IA", "563": "IA", "641": "IA", "712": "IA",
    # Kansas
    "316": "KS", "620": "KS", "785": "KS", "913": "KS",
    # Kentucky
    "270": "KY", "364": "KY", "502": "KY", "606": "KY", "859": "KY",
    # Louisiana
    "225": "LA", "318": "LA", "337": "LA", "504": "LA", "985": "LA",
    # Maine
    "207": "ME",
    # Maryland
    "240": "MD", "301": "MD", "410": "MD", "443": "MD", "667": "MD",
    # Massachusetts
    "339": "MA", "351": "MA", "413": "MA", "508": "MA", "617": "MA",
    "774": "MA", "781": "MA", "857": "MA", "978": "MA",
    # Michigan
    "231": "MI", "248": "MI", "269": "MI", "313": "MI", "517": "MI",
    "586": "MI", "616": "MI", "734": "MI", "810": "MI", "906": "MI",
    "947": "MI", "989": "MI",
    # Minnesota
    "218": "MN", "320": "MN", "507": "MN", "612": "MN", "651": "MN",
    "763": "MN", "952": "MN",
    # Mississippi
    "228": "MS", "601": "MS", "662": "MS", "769": "MS",
    # Missouri
    "314": "MO", "417": "MO", "573": "MO", "636": "MO", "660": "MO",
    "816": "MO", "975": "MO",
    # Montana
    "406": "MT",
    # Nebraska
    "308": "NE", "402": "NE", "531": "NE",
    # Nevada
    "702": "NV", "725": "NV", "775": "NV",
    # New Hampshire
    "603": "NH",
    # New Jersey
    "201": "NJ", "551": "NJ", "609": "NJ", "640": "NJ", "732": "NJ",
    "848": "NJ", "856": "NJ", "862": "NJ", "908": "NJ", "973": "NJ",
    # New Mexico
    "505": "NM", "575": "NM",
    # New York
    "212": "NY", "315": "NY", "332": "NY", "347": "NY", "516": "NY",
    "518": "NY", "585": "NY", "607": "NY", "631": "NY", "646": "NY",
    "680": "NY", "716": "NY", "718": "NY", "838": "NY", "845": "NY",
    "914": "NY", "917": "NY", "929": "NY", "934": "NY",
    # North Carolina
    "252": "NC", "336": "NC", "704": "NC", "743": "NC", "828": "NC",
    "910": "NC", "919": "NC", "980": "NC", "984": "NC",
    # North Dakota
    "701": "ND",
    # Ohio
    "216": "OH", "220": "OH", "234": "OH", "283": "OH", "330": "OH",
    "380": "OH", "419": "OH", "440": "OH", "513": "OH", "567": "OH",
    "614": "OH", "740": "OH", "937": "OH",
    # Oklahoma
    "405": "OK", "539": "OK", "580": "OK", "918": "OK",
    # Oregon
    "458": "OR", "503": "OR", "541": "OR", "971": "OR",
    # Pennsylvania
    "215": "PA", "223": "PA", "267": "PA", "272": "PA", "412": "PA",
    "445": "PA", "484": "PA", "570": "PA", "582": "PA", "610": "PA",
    "717": "PA", "724": "PA", "814": "PA", "835": "PA", "878": "PA",
    # Rhode Island
    "401": "RI",
    # South Carolina
    "803": "SC", "839": "SC", "843": "SC", "854": "SC", "864": "SC",
    # South Dakota
    "605": "SD",
    # Tennessee
    "423": "TN", "615": "TN", "629": "TN", "731": "TN", "865": "TN",
    "901": "TN", "931": "TN",
    # Texas
    "210": "TX", "214": "TX", "254": "TX", "281": "TX", "325": "TX",
    "346": "TX", "361": "TX", "409": "TX", "430": "TX", "432": "TX",
    "469": "TX", "512": "TX", "682": "TX", "713": "TX", "726": "TX",
    "737": "TX", "806": "TX", "817": "TX", "830": "TX", "832": "TX",
    "903": "TX", "915": "TX", "936": "TX", "940": "TX", "956": "TX",
    "972": "TX", "979": "TX",
    # Utah
    "385": "UT", "435": "UT", "801": "UT",
    # Vermont
    "802": "VT",
    # Virginia
    "276": "VA", "434": "VA", "540": "VA", "571": "VA", "703": "VA",
    "757": "VA", "804": "VA", "826": "VA", "948": "VA",
    # Washington
    "206": "WA", "253": "WA", "360": "WA", "425": "WA", "509": "WA",
    "564": "WA",
    # Washington DC
    "202": "DC",
    # West Virginia
    "304": "WV", "681": "WV",
    # Wisconsin
    "262": "WI", "274": "WI", "414": "WI", "534": "WI", "608": "WI",
    "715": "WI", "920": "WI",
    # Wyoming
    "307": "WY",
    # Territories
    "340": "VI", "671": "GU", "684": "AS", "787": "PR", "939": "PR", "670": "MP",
}

# Toll-free prefixes - not interstate or intrastate
TOLL_FREE = {"800", "833", "844", "855", "866", "877", "888"}


@dataclass
class CallRatioResult:
    """Result of call ratio calculation."""
    total_calls: int
    total_seconds: float
    total_minutes: float

    interstate_calls: int
    interstate_seconds: float
    interstate_minutes: float

    intrastate_calls: int
    intrastate_seconds: float
    intrastate_minutes: float

    toll_free_calls: int
    toll_free_seconds: float

    unknown_calls: int
    unknown_seconds: float

    # THE KEY NUMBERS
    interstate_ratio: float  # e.g., 0.65 = 65%
    intrastate_ratio: float  # e.g., 0.35 = 35%


def get_area_code(phone_number: str) -> str:
    """Extract area code from phone number."""
    # Remove any non-digit characters
    digits = ''.join(c for c in str(phone_number) if c.isdigit())

    # Handle different formats
    if len(digits) == 11 and digits.startswith('1'):
        return digits[1:4]  # 1XXXXXXXXXX -> XXX
    elif len(digits) == 10:
        return digits[0:3]  # XXXXXXXXXX -> XXX
    elif len(digits) >= 3:
        return digits[0:3]
    return ""


def get_state(phone_number: str) -> str:
    """Get state from phone number's area code."""
    area_code = get_area_code(phone_number)
    return NPA_TO_STATE.get(area_code, "")


def is_toll_free(phone_number: str) -> bool:
    """Check if number is toll-free."""
    area_code = get_area_code(phone_number)
    return area_code in TOLL_FREE


def classify_call(source: str, destination: str) -> str:
    """
    Classify a call as interstate, intrastate, toll_free, or unknown.

    Returns: 'interstate', 'intrastate', 'toll_free', or 'unknown'
    """
    # Check for toll-free
    if is_toll_free(source) or is_toll_free(destination):
        return "toll_free"

    # Get states
    source_state = get_state(source)
    dest_state = get_state(destination)

    # Can't determine if we don't know the states
    if not source_state or not dest_state:
        return "unknown"

    # Same state = intrastate, different = interstate
    if source_state == dest_state:
        return "intrastate"
    else:
        return "interstate"


def calculate_call_ratio(cdr_file: Path) -> CallRatioResult:
    """
    Calculate the call ratio from a CDR CSV file.

    Expects columns: Source, Destination, Seconds
    (Vitelity format: BillingDate,CallStartDate,Source,Destination,Seconds,CallerID,Disposition,Cost,Peer)
    """
    totals = {
        "interstate": {"calls": 0, "seconds": 0.0},
        "intrastate": {"calls": 0, "seconds": 0.0},
        "toll_free": {"calls": 0, "seconds": 0.0},
        "unknown": {"calls": 0, "seconds": 0.0},
    }

    with open(cdr_file, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)

        for row in reader:
            # Get source, destination, seconds
            # Handle different possible column names
            source = row.get('Source', row.get('source', row.get('src', ''))).strip()
            destination = row.get('Destination', row.get('destination', row.get('dst', ''))).strip()

            seconds_str = row.get('Seconds', row.get('seconds', row.get('duration', '0')))
            try:
                seconds = float(seconds_str.strip() if seconds_str else 0)
            except ValueError:
                seconds = 0.0

            # Skip zero-duration calls
            if seconds <= 0:
                continue

            # Classify the call
            call_type = classify_call(source, destination)

            totals[call_type]["calls"] += 1
            totals[call_type]["seconds"] += seconds

    # Calculate totals
    total_calls = sum(t["calls"] for t in totals.values())
    total_seconds = sum(t["seconds"] for t in totals.values())

    # For ratio calculation, only use interstate + intrastate (exclude toll-free and unknown)
    jurisdictional_seconds = totals["interstate"]["seconds"] + totals["intrastate"]["seconds"]

    if jurisdictional_seconds > 0:
        interstate_ratio = totals["interstate"]["seconds"] / jurisdictional_seconds
        intrastate_ratio = totals["intrastate"]["seconds"] / jurisdictional_seconds
    else:
        interstate_ratio = 0.0
        intrastate_ratio = 0.0

    return CallRatioResult(
        total_calls=total_calls,
        total_seconds=total_seconds,
        total_minutes=total_seconds / 60,

        interstate_calls=totals["interstate"]["calls"],
        interstate_seconds=totals["interstate"]["seconds"],
        interstate_minutes=totals["interstate"]["seconds"] / 60,

        intrastate_calls=totals["intrastate"]["calls"],
        intrastate_seconds=totals["intrastate"]["seconds"],
        intrastate_minutes=totals["intrastate"]["seconds"] / 60,

        toll_free_calls=totals["toll_free"]["calls"],
        toll_free_seconds=totals["toll_free"]["seconds"],

        unknown_calls=totals["unknown"]["calls"],
        unknown_seconds=totals["unknown"]["seconds"],

        interstate_ratio=interstate_ratio,
        intrastate_ratio=intrastate_ratio,
    )


def print_report(result: CallRatioResult, filename: str):
    """Print a nice report of the call ratio."""
    print("\n" + "=" * 60)
    print(f"  CALL RATIO REPORT")
    print(f"  File: {filename}")
    print("=" * 60)

    print(f"\n  SUMMARY")
    print(f"  ─────────────────────────────────────")
    print(f"  Total Calls:    {result.total_calls:,}")
    print(f"  Total Minutes:  {result.total_minutes:,.2f}")

    print(f"\n  JURISDICTION BREAKDOWN")
    print(f"  ─────────────────────────────────────")
    print(f"  Interstate:     {result.interstate_calls:,} calls  │  {result.interstate_minutes:,.2f} min")
    print(f"  Intrastate:     {result.intrastate_calls:,} calls  │  {result.intrastate_minutes:,.2f} min")
    print(f"  Toll-Free:      {result.toll_free_calls:,} calls  │  {result.toll_free_seconds/60:,.2f} min")
    if result.unknown_calls > 0:
        print(f"  Unknown:        {result.unknown_calls:,} calls  │  {result.unknown_seconds/60:,.2f} min")

    print(f"\n  ╔═══════════════════════════════════════╗")
    print(f"  ║  THE CALL RATIO                       ║")
    print(f"  ╠═══════════════════════════════════════╣")
    print(f"  ║  Interstate:  {result.interstate_ratio*100:6.2f}%               ║")
    print(f"  ║  Intrastate:  {result.intrastate_ratio*100:6.2f}%               ║")
    print(f"  ╚═══════════════════════════════════════╝")

    # FCC Safe Harbor comparison
    print(f"\n  FCC Safe Harbor: 64.9% interstate / 35.1% intrastate")
    diff = result.interstate_ratio * 100 - 64.9
    if diff > 0:
        print(f"  Your ratio is {diff:.2f}% MORE interstate than safe harbor")
    else:
        print(f"  Your ratio is {abs(diff):.2f}% LESS interstate than safe harbor")

    print("\n" + "=" * 60)


def main():
    """Main entry point."""
    if len(sys.argv) < 2:
        print("Usage: python call_ratio.py <cdr_file.csv>")
        print("\nExample: python call_ratio.py http-xxx.csv")
        sys.exit(1)

    cdr_file = Path(sys.argv[1])

    if not cdr_file.exists():
        print(f"Error: File not found: {cdr_file}")
        sys.exit(1)

    print(f"Processing {cdr_file.name}...")
    result = calculate_call_ratio(cdr_file)
    print_report(result, cdr_file.name)


if __name__ == "__main__":
    main()
