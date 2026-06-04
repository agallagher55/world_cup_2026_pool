"""
Reads all participant picks from submissions/ and writes a Statistics tab
to WC2026_Pool.xlsx showing how many people picked each team.

Run: python statistics.py
"""

import glob
import os

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

SUBMISSIONS_DIR = "submissions"
MASTER_FILE = "WC2026_Pool.xlsx"
STATS_SHEET = "Statistics"

TIERS = ["Tier 1", "Tier 2", "Tier 3", "Tier 4", "Tier 5", "Tier 6"]
TIER_LABELS = {
    "Tier 1": "Elite Favorites",
    "Tier 2": "Contenders",
    "Tier 3": "Dark Horses",
    "Tier 4": "Long Shots",
    "Tier 5": "Underdogs",
    "Tier 6": "Major Underdogs",
}
TIER_COLORS = {
    "Tier 1": "C00000",
    "Tier 2": "E26B0A",
    "Tier 3": "375623",
    "Tier 4": "17375E",
    "Tier 5": "7030A0",
    "Tier 6": "595959",
}


def read_all_teams() -> list[tuple[str, str, str]]:
    """Return [(tier, team, group), ...] from the master Teams by Tier sheet."""
    wb = openpyxl.load_workbook(MASTER_FILE)
    ws = wb["Teams by Tier"]
    teams = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        tier, team, group = row[0], row[1], row[2]
        if tier and team and str(tier).startswith("Tier"):
            teams.append((str(tier).strip(), str(team).strip(), str(group).strip()))
    return teams


def read_all_picks() -> dict[str, list[str]]:
    """Return {participant_name: [team, ...]} from every submission file."""
    participants: dict[str, list[str]] = {}
    for filepath in sorted(glob.glob(os.path.join(SUBMISSIONS_DIR, "*.xlsx"))):
        wb = openpyxl.load_workbook(filepath)
        if "User Submission Template" not in wb.sheetnames:
            print(f"  WARNING: no 'User Submission Template' in {os.path.basename(filepath)}, skipping.")
            continue
        ws = wb["User Submission Template"]

        name = None
        picks = []
        for row in ws.iter_rows(values_only=True):
            if row[0] and str(row[0]).strip() == "Participant Name:":
                name = str(row[2]).strip() if row[2] else None
            tier_val = str(row[0]).strip() if row[0] else ""
            pick_num = str(row[2]).strip() if row[2] else ""
            team = str(row[3]).strip() if row[3] else ""
            if tier_val in TIERS and pick_num in ("Pick 1", "Pick 2") and team:
                picks.append(team)

        if name:
            participants[name] = picks
        else:
            print(f"  WARNING: participant name missing in {os.path.basename(filepath)}, skipping.")

    return participants


def build_pick_counts(
    teams: list[tuple[str, str, str]],
    participants: dict[str, list[str]],
) -> list[dict]:
    """
    For every team, count how many participants picked them and who.
    Returns a list of dicts sorted by tier then descending pick count.
    """
    total_participants = len(participants)
    rows = []
    for tier, team, group in teams:
        pickers = [name for name, picks in participants.items() if team in picks]
        rows.append({
            "tier": tier,
            "team": team,
            "group": group,
            "count": len(pickers),
            "pct": len(pickers) / total_participants * 100 if total_participants else 0,
            "pickers": pickers,
        })

    rows.sort(key=lambda r: (TIERS.index(r["tier"]), -r["count"], r["team"]))
    return rows


def write_statistics_sheet(rows: list[dict], total_participants: int):
    wb = openpyxl.load_workbook(MASTER_FILE)

    if STATS_SHEET in wb.sheetnames:
        del wb[STATS_SHEET]
    ws = wb.create_sheet(STATS_SHEET)

    # ── shared styles ────────────────────────────────────────────────────────
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="1F4E79")
    white_bold = Font(bold=True, color="FFFFFF")
    bold = Font(bold=True)

    # ── Row 1: title ─────────────────────────────────────────────────────────
    ws.merge_cells("A1:F1")
    t = ws["A1"]
    t.value = f"2026 FIFA World Cup Pool – Team Pick Statistics  ({total_participants} participants)"
    t.font = Font(bold=True, size=13, color="FFFFFF")
    t.fill = header_fill
    t.alignment = center
    ws.row_dimensions[1].height = 22

    # ── Row 2: column headers ────────────────────────────────────────────────
    col_headers = ["Tier", "Team", "WC Group", "# Picked", "% Picked", "Picked By"]
    for col, h in enumerate(col_headers, 1):
        cell = ws.cell(row=2, column=col, value=h)
        cell.font = white_bold
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border
    ws.row_dimensions[2].height = 18

    # ── Data rows ────────────────────────────────────────────────────────────
    alt_fills = {
        tier: PatternFill("solid", fgColor=_lighten(TIER_COLORS[tier]))
        for tier in TIERS
    }
    tier_header_fills = {
        tier: PatternFill("solid", fgColor=TIER_COLORS[tier])
        for tier in TIERS
    }

    current_tier = None
    data_row = 3

    for entry in rows:
        tier = entry["tier"]

        # ── Tier section header ──────────────────────────────────────────────
        if tier != current_tier:
            current_tier = tier
            ws.merge_cells(f"A{data_row}:F{data_row}")
            hdr = ws.cell(
                row=data_row, column=1,
                value=f"  {tier} – {TIER_LABELS[tier]}",
            )
            hdr.font = Font(bold=True, color="FFFFFF", size=11)
            hdr.fill = tier_header_fills[tier]
            hdr.alignment = left
            hdr.border = border
            ws.row_dimensions[data_row].height = 18
            data_row += 1

        # ── Team row ────────────────────────────────────────────────────────
        row_fill = alt_fills[tier]

        values = [
            tier,
            entry["team"],
            entry["group"],
            entry["count"],
            f"{entry['pct']:.0f}%",
            ", ".join(entry["pickers"]) if entry["pickers"] else "—",
        ]
        for col, val in enumerate(values, 1):
            cell = ws.cell(row=data_row, column=col, value=val)
            cell.fill = row_fill
            cell.border = border
            cell.alignment = center if col != 6 else left

        # Bold the count if anyone picked this team
        count_cell = ws.cell(row=data_row, column=4)
        if entry["count"] > 0:
            count_cell.font = Font(bold=True)

        ws.row_dimensions[data_row].height = 16
        data_row += 1

    # ── Column widths ────────────────────────────────────────────────────────
    widths = [10, 20, 12, 10, 10, 50]
    for col, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = width

    ws.freeze_panes = "A3"
    wb.save(MASTER_FILE)


def _lighten(hex_color: str) -> str:
    """Return a lightened pastel version of a hex color for alternating row fills."""
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    r = int(r + (255 - r) * 0.82)
    g = int(g + (255 - g) * 0.82)
    b = int(b + (255 - b) * 0.82)
    return f"{r:02X}{g:02X}{b:02X}"


def main():
    print("Reading teams from master file...")
    teams = read_all_teams()
    print(f"  {len(teams)} teams found.")

    print("Reading participant picks from submissions/...")
    participants = read_all_picks()
    print(f"  {len(participants)} participant(s): {', '.join(participants)}")

    print("Tallying pick counts...")
    rows = build_pick_counts(teams, participants)

    picked = [r for r in rows if r["count"] > 0]
    unpicked = [r for r in rows if r["count"] == 0]
    print(f"  {len(picked)} team(s) picked, {len(unpicked)} team(s) not picked.")

    print("Writing Statistics sheet...")
    write_statistics_sheet(rows, total_participants=len(participants))
    print(f"Done. '{STATS_SHEET}' tab written to {MASTER_FILE}.")

    print("\nPick summary by tier:")
    current_tier = None
    for r in rows:
        if r["tier"] != current_tier:
            current_tier = r["tier"]
            print(f"\n  {r['tier']} – {TIER_LABELS[r['tier']]}:")
        bar = "█" * r["count"]
        pickers = f"  ← {', '.join(r['pickers'])}" if r["pickers"] else ""
        print(f"    {r['team']:<22} {bar:<6} {r['count']}/{len(participants)}{pickers}")


if __name__ == "__main__":
    main()
