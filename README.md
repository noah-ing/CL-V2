# Telecom Billing Automation

Automated billing report generator for CirrusLine/Syneteks telecom services. Processes CDR (Call Detail Records) from Vitelity and SkySwitch to generate billing reports with interstate/intrastate call ratios.

## Features

- **Call Ratio Calculation**: Classifies calls as interstate/intrastate based on area codes (NPA-NXX lookup)
- **Per-Customer Breakdown**: CDR analysis broken down by customer/domain
- **Combined CDR Reports**: Merges SkySwitch (from Excel) and Vitelity (CSV) CDR data
- **Cost Calculations**: Applies billing rates ($0.005/min voice, $0.005/msg SMS)
- **Seat Count Reports**: Extracts PBX user counts from Domain Statistics
- **SMS Usage Reports**: Tracks SMS messages by customer
- **CallerID Reports**: Counts calls per destination for CallerID billing

## Installation

```bash
# No dependencies required - uses Python standard library only
python3 --version  # Requires Python 3.9+
```

## Usage

### Quick Start - Call Ratio Only

```bash
python3 call_ratio.py <cdr_file.csv>
```

### Full Billing Reports

```bash
python3 billing_reports.py \
  <vitelity_cdr.csv> \
  <phonenumbers.csv> \
  [output_dir] \
  [sms_file.csv] \
  [Domain-Statistics.xlsx] \
  [master_xlsx_with_skyswitch_cdr]
```

### Example

```bash
python3 billing_reports.py \
  "Working Reports 2025_10/http-lfukktvvuqqvisyogcretfbovva.csv" \
  "Working Reports 2025_10/phonenumbers__20251031_15_25.csv" \
  ./reports \
  "Working Reports 2025_10/syneteks-7832.csv" \
  "Working Reports 2025_10/Domain-Statistics-2025-10-31.xlsx" \
  "Working Reports 2025_10/CDR SS records-25_10_01.xlsx"
```

## Generated Reports

| Report | Description |
|--------|-------------|
| `cdr_by_customer.csv` | Vitelity CDR with interstate/intrastate breakdown + costs |
| `combined_cdr_by_customer.csv` | Combined SkySwitch + Vitelity CDR with call ratios |
| `phone_counts_by_customer.csv` | **Billable** phone numbers per customer (excludes fax/hold/unassigned) |
| `phone_excluded_invother.csv` | Non-billable phones (fax, on-hold, unassigned) for InvOther |
| `callerid_counts.csv` | CallerID lookup counts per number |
| `seat_counts_by_customer.csv` | PBX Users (seats) per customer |
| `sms_by_customer.csv` | SMS usage per customer |
| `adams_county_user_summary.csv` | Adams County User Summary pivot (dept × user type) |

## Phone Number Filtering

Phone numbers are automatically filtered based on the `Treatment` column:

**Billable** (included in `phone_counts_by_customer.csv`):
- `User` - Active phone users
- `Voicemail` - Voicemail boxes
- `Call Queue` - Call queue numbers
- `Conference` - Conference bridges

**Non-Billable** (excluded to `phone_excluded_invother.csv`):
- `Available Number` - Unassigned numbers
- `FaxSFATA`, `vFax`, `iFax`, `vFaxSFATA` - Fax machines
- `vOn-Hold`, `vOffNet` - On-hold/special purpose numbers

## Input File Formats

### Vitelity CDR (CSV)
```
BillingDate,CallStartDate,Source,Destination,Seconds,CallerID,Disposition,Cost,Peer
```

### Phone Numbers (CSV)
```
Phone Number,Domain,Treatment,Destination,Notes,Enable
```

### SMS (CSV)
```
Time,Source,Destination,msgDirection,Cost
```

### Domain Statistics (XLSX)
Excel file with columns: Domain, PBX Users, Call Center, Call Recording, etc.

### Master Excel (CDR SS records)
38-tab workbook containing SkySwitch CDR in the "CDR<i>" sheet (sheet 26).

## Billing Rates

| Service | Rate |
|---------|------|
| Voice | $0.005/minute |
| SMS | $0.005/message |

## Call Classification

Calls are classified as:
- **Interstate**: Source and destination in different states
- **Intrastate**: Source and destination in same state
- **Toll-Free**: 800/833/844/855/866/877/888 numbers
- **Unknown**: Unable to determine state from area code

The call ratio (interstate %) is calculated excluding toll-free and unknown calls.

## FCC Safe Harbor

The FCC Safe Harbor ratio is 64.9% interstate / 35.1% intrastate. Reports compare your actual ratio against this benchmark.

## License

Private - CirrusLine/Syneteks internal use only.
