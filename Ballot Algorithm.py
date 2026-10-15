#!/usr/bin/env python3
"""Balliol JCR Picture Collection ballot allocator.

Ranked allocation policy:
1. Maximise the number of members who receive one of their three preferences.
2. Among those allocations, maximise first preferences.
3. Then maximise second preferences.
4. Then maximise third preferences.
5. Use the published seed to break any remaining ties.
6. After ranked allocation, maximise the number of eligible fallback recipients.
7. Use the published seed to break equivalent fallback assignments.

Requires: pandas, openpyxl and scipy.
"""

from __future__ import annotations

import argparse
import csv
import random
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable

try:
    import pandas as pd
    from scipy.optimize import linear_sum_assignment
except ModuleNotFoundError as exc:
    print(f"Missing Python package: {exc.name}", file=sys.stderr)
    print("Install pandas, openpyxl and scipy, then run the ballot again.", file=sys.stderr)
    if sys.stdin.isatty():
        try:
            input("\nPress Enter to close.")
        except EOFError:
            pass
    raise SystemExit(1) from exc


CATALOGUE_SHEET = "LEASABLE PICTURE COLLECTION"
RESPONSES_SHEET = "Form responses 1"
DIMENSIONS_COLUMN = "SIZE (CM W X H) (S / M / L)"

DEFAULT_RESPONSES_FILENAME = "JCR Picture Collection Ballot Form (Responses).xlsx"
DEFAULT_CATALOGUE_FILENAME = "Balliol JCR Picture Officer Sheet.xlsx"
DEFAULT_OUTPUT_PREFIX = "ballot_results"

PREFERENCE_COLUMNS = ("Preference 1", "Preference 2", "Preference 3")
BLACKLIST_COLUMNS = ("Blacklist 1", "Blacklist 2")
FALLBACK_COLUMN = "Would you like a random artwork if you do not receive your preference?"
FALLBACK_SIZES_COLUMN = "If yes, what sizes would you be happy to use?"

SMALL_MAX_CM = 35
MEDIUM_MAX_CM = 70

# Excel row numbers (1-based) of the first data row.
CATALOGUE_HEADER_ROW = 6
CATALOGUE_EXCEL_ROW_OFFSET = 8
RESPONSES_EXCEL_ROW_OFFSET = 2

# Tie-break budget keeps the total random contribution strictly below 1, so
# tie-breaks can never override the integer lexicographic preference ordering.
TIE_BREAK_BUDGET = 0.25

# Accept both the full size labels and the single-letter codes, case-insensitively.
SIZE_LABELS = {
    "small": "S",
    "s": "S",
    "medium": "M",
    "m": "M",
    "large": "L",
    "l": "L",
}

AFFIRMATIVE_VALUES = {"yes", "y", "true"}


def _strip_excel_float_suffix(text: str) -> str:
    """Drop a cosmetic '.0' suffix Excel adds to numeric cells.

    ``pd.read_excel(..., dtype=str)`` converts a numeric cell holding 12 to the
    string ``"12.0"``. Only strip the suffix when the leading part is all
    digits, so genuine decimals like ``"12.5"`` are left untouched.
    """
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def parse_artwork_number(value: object) -> str:
    """Return a three-digit artwork number from a form or catalogue value."""
    text = _strip_excel_float_suffix(str(value).strip())
    if not text:
        return ""

    number = text.split("-", 1)[0].strip()
    if not number.isdigit():
        return ""

    number = number.zfill(3)
    return number if len(number) == 3 else ""


def largest_dimension_cm(value: object) -> float:
    """Return the larger dimension from a catalogue value such as '44.5 X 32'."""
    text = str(value).strip().upper().replace("×", "X")
    try:
        width_text, height_text = text.split("X")
        return max(float(width_text.strip()), float(height_text.strip()))
    except ValueError as exc:
        raise ValueError(f"Could not read dimensions {value!r}") from exc


def size_band(largest_cm: float) -> str:
    """Classify an artwork as S, M or L from its largest dimension in cm."""
    if largest_cm <= SMALL_MAX_CM:
        return "S"
    if largest_cm <= MEDIUM_MAX_CM:
        return "M"
    return "L"


def parse_fallback_sizes(value: object) -> list[str]:
    """Parse fallback size preferences into a deduplicated list of S/M/L codes.

    Accepts comma- or semicolon-separated values, full labels or single
    letters, and is case-insensitive.
    """
    text = str(value).strip()
    if not text:
        return []

    sizes: list[str] = []
    for raw_label in text.replace(";", ",").split(","):
        label = raw_label.strip().lower()
        if not label:
            continue
        if label not in SIZE_LABELS:
            raise ValueError(f"Unknown fallback size {raw_label!r}")
        sizes.append(SIZE_LABELS[label])
    return list(dict.fromkeys(sizes))


def is_affirmative(value: object) -> bool:
    """Return True if a form value should be treated as 'yes'."""
    return str(value).strip().lower() in AFFIRMATIVE_VALUES


def _normalise_student_id(value: object) -> str:
    """Return a Student ID as a clean string, dropping Excel's '.0' suffix."""
    return _strip_excel_float_suffix(str(value).strip())


def require_columns(
    table: pd.DataFrame,
    required_columns: Iterable[str],
    source_name: str,
) -> None:
    """Raise a ValueError if any required column is missing from the table."""
    missing = set(required_columns) - set(table.columns)
    if missing:
        raise ValueError(f"{source_name} is missing columns: {sorted(missing)}")


def load_catalogue(path: Path, sheet_name: str = CATALOGUE_SHEET) -> dict[str, dict]:
    """Load the catalogue workbook and return details for unleased artworks."""
    catalogue = pd.read_excel(
        path,
        sheet_name=sheet_name,
        header=CATALOGUE_HEADER_ROW,
        dtype=str,
    ).fillna("")

    require_columns(
        catalogue,
        {"NEW_ID", "LEASE INFORMATION", DIMENSIONS_COLUMN},
        "Catalogue",
    )

    artworks: dict[str, dict] = {}
    seen_numbers: set[str] = set()

    for index, row in catalogue.iterrows():
        new_id = str(row["NEW_ID"]).strip()
        if not new_id:
            continue

        row_number = index + CATALOGUE_EXCEL_ROW_OFFSET
        number = parse_artwork_number(new_id)
        if not number:
            raise ValueError(f"Catalogue row {row_number}: invalid NEW_ID {new_id!r}")
        if number in seen_numbers:
            raise ValueError(f"Catalogue row {row_number}: duplicate artwork {number}")
        seen_numbers.add(number)

        if str(row["LEASE INFORMATION"]).strip().upper() != "UNLEASED":
            continue

        largest_cm = largest_dimension_cm(row[DIMENSIONS_COLUMN])
        artworks[number] = {
            "NEW_ID": new_id,
            "TITLE": str(row.get("NAME / DESCRIPTION", "")).strip(),
            "ARTIST": str(row.get("ARTIST", "")).strip(),
            "SIZE": size_band(largest_cm),
            "LARGEST_DIM_CM": largest_cm,
        }

    if not artworks:
        raise ValueError("No unleased artworks were found")

    return artworks


def parse_preferences(row: pd.Series, row_number: int) -> list[str]:
    """Return the member's preferences with duplicates blanked out."""
    preferences: list[str] = []
    seen: set[str] = set()

    for column in PREFERENCE_COLUMNS:
        raw_value = str(row[column]).strip()
        number = parse_artwork_number(raw_value)

        if raw_value and not number:
            raise ValueError(
                f"Response row {row_number}: invalid {column} value {raw_value!r}"
            )

        if number in seen:
            number = ""
        elif number:
            seen.add(number)

        preferences.append(number)

    return preferences


def parse_blacklist(row: pd.Series, row_number: int) -> set[str]:
    """Return the set of artworks the member has blacklisted."""
    blacklist: set[str] = set()

    for column in BLACKLIST_COLUMNS:
        raw_value = str(row[column]).strip()
        if not raw_value:
            continue

        number = parse_artwork_number(raw_value)
        if not number:
            raise ValueError(
                f"Response row {row_number}: invalid {column} value {raw_value!r}"
            )
        blacklist.add(number)

    return blacklist


def load_responses(path: Path, sheet_name: str = RESPONSES_SHEET) -> list[dict]:
    """Load the Google Form responses workbook into ballot dictionaries."""
    responses = pd.read_excel(path, sheet_name=sheet_name, dtype=str).fillna("")

    required_columns = {
        "First Name",
        "Last Name",
        "Student ID Number",
        "Student Email",
        *PREFERENCE_COLUMNS,
        *BLACKLIST_COLUMNS,
        FALLBACK_COLUMN,
        FALLBACK_SIZES_COLUMN,
    }
    require_columns(responses, required_columns, "Response form")

    ballots: list[dict] = []
    seen_student_ids: set[str] = set()

    for index, row in responses.iterrows():
        row_number = index + RESPONSES_EXCEL_ROW_OFFSET
        student_id = _normalise_student_id(row["Student ID Number"])

        if not student_id:
            raise ValueError(f"Response row {row_number}: Student ID is blank")
        if student_id in seen_student_ids:
            raise ValueError(
                f"Response row {row_number}: duplicate Student ID {student_id}"
            )
        seen_student_ids.add(student_id)

        wants_fallback = is_affirmative(row[FALLBACK_COLUMN])
        fallback_sizes = (
            parse_fallback_sizes(row[FALLBACK_SIZES_COLUMN]) if wants_fallback else []
        )
        if wants_fallback and not fallback_sizes:
            raise ValueError(
                f"Response row {row_number}: fallback is Yes but no fallback size "
                "was selected"
            )

        first_name = str(row["First Name"]).strip()
        last_name = str(row["Last Name"]).strip()

        ballots.append(
            {
                "MEMBER_ID": student_id,
                "MEMBER_NAME": f"{last_name}, {first_name}".strip(", "),
                "EMAIL": str(row["Student Email"]).strip(),
                "PREFS": parse_preferences(row, row_number),
                "EXCLUDE": parse_blacklist(row, row_number),
                "FALLBACK_OK": wants_fallback,
                "SIZE_PREF": fallback_sizes,
            }
        )

    if not ballots:
        raise ValueError("No ballot responses were found")

    return ballots


def warn_about_unavailable_choices(artworks: dict, ballots: list[dict]) -> None:
    """Warn about ranked preferences that cannot be honoured."""
    available = set(artworks)

    for ballot in ballots:
        member_id = ballot["MEMBER_ID"]
        for rank, number in enumerate(ballot["PREFS"], start=1):
            if not number:
                continue

            if number not in available:
                print(
                    f"Warning: {member_id} ranked artwork {number} at Preference "
                    f"{rank}, but it is not currently unleased",
                    file=sys.stderr,
                )

            if number in ballot["EXCLUDE"]:
                print(
                    f"Warning: {member_id} both ranked and blacklisted artwork "
                    f"{number}; the blacklist takes precedence",
                    file=sys.stderr,
                )


def valid_ranked_choices(ballot: dict, artworks: dict) -> Iterable[tuple[int, str]]:
    """Yield (rank, artwork) pairs that could be honoured for this ballot."""
    for rank, artwork in enumerate(ballot["PREFS"], start=1):
        if artwork and artwork in artworks and artwork not in ballot["EXCLUDE"]:
            yield rank, artwork


def ranked_allocation(
    artworks: dict,
    ballots: list[dict],
    seed: int,
) -> tuple[dict[str, str], dict[str, int]]:
    """Find the globally optimal ranked allocation.

    The scoring is lexicographic via the choice of weights:

    * every match beats no match (``priority_base ** 3``),
    * among matches, first preferences beat second, and second beat third,
    * the random tie-break is bounded by ``TIE_BREAK_BUDGET`` in total, so it
      can never override the integer ordering above.

    For JCR-sized ballots (a few hundred members) the weights are well within
    float64's exact-integer range; larger inputs would need an integer or
    two-pass solver.
    """
    members = sorted(ballot["MEMBER_ID"] for ballot in ballots)
    artwork_ids = sorted(artworks)
    ballot_by_member = {ballot["MEMBER_ID"]: ballot for ballot in ballots}

    member_count = len(members)
    priority_base = member_count + 1
    rank_scores = {
        1: priority_base**3 + priority_base**2,
        2: priority_base**3 + priority_base,
        3: priority_base**3 + 1,
    }

    member_rows = {member: index for index, member in enumerate(members)}
    artwork_columns = {artwork: index for index, artwork in enumerate(artwork_ids)}
    real_artwork_count = len(artwork_ids)

    # Dummy columns allow every member to remain unassigned if necessary.
    score_matrix = [
        [-1.0] * real_artwork_count + [0.0] * member_count
        for _ in members
    ]

    ranked_edges: list[tuple[str, str, int]] = []
    for member in members:
        ballot = ballot_by_member[member]
        for rank, artwork in valid_ranked_choices(ballot, artworks):
            ranked_edges.append((member, artwork, rank))

    tie_rng = random.Random(seed)
    tie_scale = TIE_BREAK_BUDGET / max(1, member_count)

    rank_by_edge: dict[tuple[str, str], int] = {}
    for member, artwork, rank in sorted(ranked_edges):
        row = member_rows[member]
        column = artwork_columns[artwork]
        score_matrix[row][column] = rank_scores[rank] + tie_scale * tie_rng.random()
        rank_by_edge[(member, artwork)] = rank

    selected_rows, selected_columns = linear_sum_assignment(
        score_matrix, maximize=True
    )

    assigned: dict[str, str] = {}
    assigned_rank: dict[str, int] = {}

    for row, column in zip(selected_rows, selected_columns):
        if column >= real_artwork_count or score_matrix[row][column] <= 0:
            continue

        member = members[row]
        artwork = artwork_ids[column]
        assigned[member] = artwork
        assigned_rank[member] = rank_by_edge[(member, artwork)]

    return assigned, assigned_rank


def allocate_fallbacks(
    artworks: dict,
    ballots: list[dict],
    ranked_assignments: dict[str, str],
    seed: int,
    verbose: bool = False,
) -> dict[str, str]:
    """Allocate fallbacks while maximising the number of recipients.

    All eligible fallback assignments are considered together. The solver first
    maximises how many members receive a fallback; the published seed then
    breaks ties between equally good fallback matchings.
    """
    assignments = dict(ranked_assignments)
    remaining_artworks = sorted(set(artworks) - set(assignments.values()))
    ballot_by_member = {ballot["MEMBER_ID"]: ballot for ballot in ballots}

    fallback_members = sorted(
        member_id
        for member_id, ballot in ballot_by_member.items()
        if member_id not in assignments and ballot["FALLBACK_OK"]
    )

    if not fallback_members or not remaining_artworks:
        return assignments

    artwork_columns = {
        artwork: index for index, artwork in enumerate(remaining_artworks)
    }
    real_artwork_count = len(remaining_artworks)
    member_count = len(fallback_members)

    # Dummy columns allow members to remain unassigned when necessary.
    score_matrix = [
        [-1.0] * real_artwork_count + [0.0] * member_count
        for _ in fallback_members
    ]

    eligible_edges: list[tuple[str, str]] = []
    for member_id in fallback_members:
        ballot = ballot_by_member[member_id]
        for artwork in remaining_artworks:
            if artwork in ballot["EXCLUDE"]:
                continue
            if artworks[artwork]["SIZE"] not in ballot["SIZE_PREF"]:
                continue
            eligible_edges.append((member_id, artwork))

    tie_rng = random.Random(seed)
    tie_scale = TIE_BREAK_BUDGET / max(1, member_count)
    member_rows = {
        member_id: index for index, member_id in enumerate(fallback_members)
    }

    for member_id, artwork in sorted(eligible_edges):
        row = member_rows[member_id]
        column = artwork_columns[artwork]
        score_matrix[row][column] = 1.0 + tie_scale * tie_rng.random()

    selected_rows, selected_columns = linear_sum_assignment(
        score_matrix, maximize=True
    )

    if verbose:
        print("Fallback allocation:")

    for row, column in zip(selected_rows, selected_columns):
        if column >= real_artwork_count or score_matrix[row][column] <= 0:
            continue

        member_id = fallback_members[row]
        artwork = remaining_artworks[column]
        assignments[member_id] = artwork

        if verbose:
            print(f"  {member_id} -> {artwork} (fallback)")

    return assignments


def summarise_ranked_allocation(assigned_rank: dict[str, int]) -> dict[str, int]:
    """Return counts of ranked matches by preference rank."""
    return {
        "ranked_matches": len(assigned_rank),
        "first_preferences": sum(rank == 1 for rank in assigned_rank.values()),
        "second_preferences": sum(rank == 2 for rank in assigned_rank.values()),
        "third_preferences": sum(rank == 3 for rank in assigned_rank.values()),
    }


def _audit_reason(rank: object, assigned_artwork: object) -> str:
    """Classify a single allocation outcome for the audit log."""
    if rank:
        return f"ranked_pref_{rank}"
    if assigned_artwork:
        return "fallback_random"
    return "no_pick"


def build_audit_rows(
    ballots: list[dict],
    assignments: dict[str, str],
    assigned_rank: dict[str, int],
) -> list[dict]:
    """Build one audit row per ballot describing the outcome."""
    rows: list[dict] = []

    for ballot in ballots:
        member_id = ballot["MEMBER_ID"]
        assigned_artwork = assignments.get(member_id, "")
        rank = assigned_rank.get(member_id, "")
        preference_1, preference_2, preference_3 = ballot["PREFS"]

        rows.append(
            {
                "member_id": member_id,
                "member_name": ballot["MEMBER_NAME"],
                "preference_1": preference_1,
                "preference_2": preference_2,
                "preference_3": preference_3,
                "picked": assigned_artwork,
                "assigned_rank": rank,
                "reason": _audit_reason(rank, assigned_artwork),
            }
        )

    return rows


def allocate(
    artworks: dict,
    ballots: list[dict],
    seed: int,
    verbose: bool = False,
) -> tuple[dict[str, str], list[dict], dict[str, int]]:
    """Run the full ranked-then-fallback allocation."""
    ranked_assignments, assigned_rank = ranked_allocation(artworks, ballots, seed)

    if verbose:
        print("Ranked allocation:")
        for member_id in sorted(ranked_assignments):
            rank = assigned_rank[member_id]
            print(
                f"  {member_id} -> {ranked_assignments[member_id]} "
                f"(Preference {rank})"
            )

    summary = summarise_ranked_allocation(assigned_rank)

    if verbose:
        print(
            "Ranked result: "
            f"{summary['ranked_matches']} members matched "
            f"({summary['first_preferences']} first, "
            f"{summary['second_preferences']} second, "
            f"{summary['third_preferences']} third)"
        )

    assignments = allocate_fallbacks(
        artworks,
        ballots,
        ranked_assignments,
        seed,
        verbose,
    )
    audit_rows = build_audit_rows(ballots, assignments, assigned_rank)

    return assignments, audit_rows, summary


def write_allocations(
    path: Path,
    artworks: dict,
    ballots: list[dict],
    assignments: dict[str, str],
) -> None:
    """Write the per-member allocation CSV."""
    with path.open("w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "MEMBER_ID",
                "MEMBER_NAME",
                "EMAIL",
                "ASSIGNED_SEQ",
                "NEW_ID",
                "TITLE",
                "ARTIST",
                "SIZE",
                "LARGEST_DIM_CM",
            ]
        )

        for ballot in ballots:
            member_id = ballot["MEMBER_ID"]
            artwork_number = assignments.get(member_id, "")
            artwork = artworks.get(artwork_number, {})

            writer.writerow(
                [
                    member_id,
                    ballot["MEMBER_NAME"],
                    ballot["EMAIL"],
                    artwork_number,
                    artwork.get("NEW_ID", ""),
                    artwork.get("TITLE", ""),
                    artwork.get("ARTIST", ""),
                    artwork.get("SIZE", ""),
                    artwork.get("LARGEST_DIM_CM", ""),
                ]
            )


def write_audit_log(
    path: Path,
    audit_rows: list[dict],
    summary: dict[str, int],
    seed: int,
) -> None:
    """Write the audit CSV, including run metadata on every row."""
    run_time = datetime.now(timezone.utc).isoformat()

    with path.open("w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "RUN_TIMESTAMP_UTC",
                "SEED",
                "RANKED_MATCHES",
                "FIRST_PREFERENCES",
                "SECOND_PREFERENCES",
                "THIRD_PREFERENCES",
                "MEMBER_ID",
                "MEMBER_NAME",
                "PREFERENCE_1",
                "PREFERENCE_2",
                "PREFERENCE_3",
                "PICKED_SEQ",
                "ASSIGNED_RANK",
                "REASON",
            ]
        )

        for row in audit_rows:
            writer.writerow(
                [
                    run_time,
                    seed,
                    summary["ranked_matches"],
                    summary["first_preferences"],
                    summary["second_preferences"],
                    summary["third_preferences"],
                    row["member_id"],
                    row["member_name"],
                    row["preference_1"],
                    row["preference_2"],
                    row["preference_3"],
                    row["picked"],
                    row["assigned_rank"],
                    row["reason"],
                ]
            )


def write_outputs(
    artworks: dict,
    ballots: list[dict],
    assignments: dict[str, str],
    audit_rows: list[dict],
    summary: dict[str, int],
    seed: int,
    output_prefix: str,
) -> None:
    """Write both output CSVs, creating parent directories if needed."""
    allocation_path = Path(f"{output_prefix}_allocations.csv")
    audit_path = Path(f"{output_prefix}_audit_log.csv")

    allocation_path.parent.mkdir(parents=True, exist_ok=True)

    write_allocations(allocation_path, artworks, ballots, assignments)
    write_audit_log(audit_path, audit_rows, summary, seed)

    print(f"Allocations written to: {allocation_path}")
    print(f"Audit log written to: {audit_path}")


def prompt_for_seed(seed: int | None) -> int:
    """Return a validated five-digit ballot seed, prompting if necessary."""
    if seed is not None:
        if 10000 <= seed <= 99999:
            return seed
        raise ValueError("Seed must be a five-digit number (10000-99999)")

    green = "\033[92m" if sys.stdout.isatty() else ""
    dim = "\033[2m" if sys.stdout.isatty() else ""
    reset = "\033[0m" if sys.stdout.isatty() else ""

    print()
    print(f"{green}+--------------------------------------+{reset}")
    print(f"{green}| Balliol JCR Picture Collection       |{reset}")
    print(f"{green}| Ballot Algorithm                     |{reset}")
    print(f"{green}+--------------------------------------+{reset}")
    print()
    print(f"{dim}Enter the ballot seed.{reset}")

    while True:
        value = input(f"{green}> {reset}").strip()
        if value.isdigit() and len(value) == 5 and value[0] != "0":
            return int(value)
        print("Invalid seed. Enter exactly five digits (10000-99999).\n")


def parse_arguments() -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description="Balliol JCR Picture Collection ballot allocator"
    )
    parser.add_argument(
        "--responses",
        help=(
            "Path to the Google Form responses workbook. If omitted, the program "
            "looks beside this Python file for the standard response filename."
        ),
    )
    parser.add_argument(
        "--catalogue",
        help=(
            "Path to the picture catalogue workbook. If omitted, the program "
            "looks beside this Python file for the standard catalogue filename."
        ),
    )
    parser.add_argument(
        "--responses-sheet",
        default=RESPONSES_SHEET,
        help="Response worksheet name",
    )
    parser.add_argument(
        "--catalogue-sheet",
        default=CATALOGUE_SHEET,
        help="Catalogue worksheet name",
    )
    parser.add_argument(
        "--seed",
        type=int,
        help="Published five-digit seed used for ties and fallbacks",
    )
    parser.add_argument(
        "--out",
        help="Output file prefix. Defaults to ballot_results beside this Python file.",
    )
    parser.add_argument(
        "--verbose", action="store_true", help="Print allocation details"
    )
    return parser.parse_args()


def _run(args: argparse.Namespace) -> int:
    """Run the allocation for a parsed argument namespace."""
    script_directory = Path(__file__).resolve().parent

    response_path = (
        Path(args.responses)
        if args.responses
        else script_directory / DEFAULT_RESPONSES_FILENAME
    )
    catalogue_path = (
        Path(args.catalogue)
        if args.catalogue
        else script_directory / DEFAULT_CATALOGUE_FILENAME
    )
    output_prefix = args.out or str(script_directory / DEFAULT_OUTPUT_PREFIX)

    if not response_path.exists():
        print(f"Error: Response file not found: {response_path}", file=sys.stderr)
        print(
            f"Place {DEFAULT_RESPONSES_FILENAME!r} in the same folder as this Python file.",
            file=sys.stderr,
        )
        return 1
    if not catalogue_path.exists():
        print(f"Error: Catalogue file not found: {catalogue_path}", file=sys.stderr)
        print(
            f"Place {DEFAULT_CATALOGUE_FILENAME!r} in the same folder as this Python file.",
            file=sys.stderr,
        )
        return 1

    try:
        seed = prompt_for_seed(args.seed)
        artworks = load_catalogue(catalogue_path, args.catalogue_sheet)
        ballots = load_responses(response_path, args.responses_sheet)
        warn_about_unavailable_choices(artworks, ballots)
        assignments, audit_rows, summary = allocate(
            artworks,
            ballots,
            seed=seed,
            verbose=args.verbose,
        )
        write_outputs(
            artworks,
            ballots,
            assignments,
            audit_rows,
            summary,
            seed,
            output_prefix,
        )
    except (KeyError, ValueError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # noqa: BLE001 - surface I/O errors without a traceback
        print(
            f"Unexpected error ({type(exc).__name__}): {exc}",
            file=sys.stderr,
        )
        return 2

    print("\nSummary:")
    print(f"  Ballots: {len(ballots)}")
    print(f"  Available artworks: {len(artworks)}")
    print(f"  Members receiving a ranked choice: {summary['ranked_matches']}")
    print(f"    First preferences: {summary['first_preferences']}")
    print(f"    Second preferences: {summary['second_preferences']}")
    print(f"    Third preferences: {summary['third_preferences']}")
    print(f"  Total artworks assigned after fallbacks: {len(assignments)}")
    print(f"  Artworks remaining: {len(artworks) - len(assignments)}")
    return 0


def pause_before_close() -> None:
    """Keep a double-clicked console window open so the result can be read."""
    if not sys.stdin.isatty():
        return
    try:
        input("\nPress Enter to close.")
    except EOFError:
        pass


def main() -> int:
    """Entry point."""
    double_clicked = len(sys.argv) == 1
    exit_code = _run(parse_arguments())

    if double_clicked:
        pause_before_close()

    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())