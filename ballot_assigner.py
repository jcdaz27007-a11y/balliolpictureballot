import csv
import random
import argparse
from datetime import datetime
from collections import defaultdict
import pandas as pd
import sys
import re
import os

# ---------------------------
# Utilities / Normalization
# ---------------------------

def norm_seq(value):
    """Normalize a sequence ID (integer or string) to a 3-digit zero-padded string.
    Returns empty string for blank/NA values.
    """
    if value is None:
        return ""
    s = str(value).strip()
    if s == "" or s.lower() in ("na", "n/a", "none"):
        return ""
    try:
        n = int(float(s))
        if n < 0:
            return ""
        return str(n).zfill(3)
    except Exception:
        if s.isdigit():
            return s.zfill(3)
        return s.upper()

def parse_list_cell(cell):
    """Parse a CSV cell containing comma-separated values into a list of stripped strings."""
    if cell is None:
        return []
    s = str(cell).strip()
    if s == "":
        return []
    return [p.strip() for p in s.split(",") if p.strip()]

def make_member_id(last_name, first_name):
    """Create a canonical MEMBER_ID from last and first name: LAST_FIRST uppercase, non-alnum -> underscore."""
    ln = str(last_name or "").strip()
    fn = str(first_name or "").strip()
    base = f"{ln}_{fn}".strip()
    if base == "_" or base == "":
        base = "UNKNOWN"
    base = re.sub(r'[^0-9A-Za-z]+', '_', base)
    base = re.sub(r'_{2,}', '_', base)
    base = base.strip('_').upper()
    if base == "":
        base = "UNKNOWN"
    return base

# ---------------------------
# Loaders & Validation (modified ballots loader)
# ---------------------------

def load_items(path):
    """Load items.csv and return dict mapping normalized SEQ->item dict.
       Only include rows with AVAILABLE == 'Y' (case-insensitive).
    """
    df = pd.read_csv(path, dtype=str).fillna("")
    items = {}
    for _, row in df.iterrows():
        seq_raw = row.get("SEQ", "")
        seq = norm_seq(seq_raw)
        if seq == "":
            continue
        size = str(row.get("SIZE", "")).strip().upper()
        available = str(row.get("AVAILABLE", "")).strip().upper()
        if size not in ("S", "M", "L") and size != "":
            print(f"Warning: item {seq} has unexpected SIZE '{size}'. Expected S/M/L.", file=sys.stderr)
        items[seq] = {
            "SEQ": seq,
            "SIZE": size,
            "AVAILABLE": (available == "Y"),
            "meta": row.to_dict()
        }
    # filter to available
    available_items = {k: v for k, v in items.items() if v["AVAILABLE"]}
    if not available_items:
        print("Warning: No available items found in items.csv", file=sys.stderr)
    return available_items

def load_ballots_simple(path):
    df = pd.read_csv(path, dtype=str).fillna("")
    ballots = []
    seen_member_ids = {}
    for row_idx, row in df.iterrows():
        last = row.get("LAST NAME", "").strip()
        first = row.get("FIRST NAME", "").strip()

        member_base = make_member_id(last, first)
        member_id = member_base
        suffix = 1
        while member_id in seen_member_ids:
            suffix += 1
            member_id = f"{member_base}_{suffix}"
        seen_member_ids[member_id] = True

        member_name = f"{last}, {first}".strip(", ")

        raw_size_pref = parse_list_cell(row.get("SIZE_PREF", ""))
        size_pref = []
        for s in raw_size_pref:
            su = s.strip().upper()
            if su in ("S", "M", "L"):
                size_pref.append(su)
            else:
                if su != "":
                    print(f"Warning: row {row_idx+2} invalid SIZE_PREF '{s}' - ignored", file=sys.stderr)

        raw_excl = parse_list_cell(row.get("EXCLUDE", ""))
        if len(raw_excl) > 2:
            print(f"Warning: row {row_idx+2} EXCLUDE has more than 2 values; only first two will be used", file=sys.stderr)
        excl = set()
        for x in raw_excl[:2]:
            nx = norm_seq(x)
            if nx:
                excl.add(nx)

        # FALLBACK: Y/N (default Y)
        fallback = str(row.get("FALLBACK", "Y")).strip().upper() == "Y"

        # PREF1..PREF3
        prefs = []
        for col in ("PREF1", "PREF2", "PREF3"):
            val = row.get(col, "")
            if str(val).strip() != "":
                prefs.append(norm_seq(val))

        # MAX_ITEMS fixed to 1
        max_items = 1

        ballots.append({
            "MEMBER_ID": member_id,
            "MEMBER_NAME": member_name,
            "MAX_ITEMS": max_items,
            "SIZE_PREF": size_pref,
            "EXCLUDE": excl,
            "FALLBACK_OK": fallback,
            "PREFS": prefs
        })

    if not ballots:
        print("Warning: No ballots loaded from ballots.csv", file=sys.stderr)
    return ballots

# ---------------------------
# Allocation core (modified to prioritise ranked prefs over size_pref)
# ---------------------------

def allocate(items_dict, ballots, seed=2026, max_rounds_override=None, verbose=False):
    rnd = random.Random(seed)
    remaining_items = set(items_dict.keys())
    assigned = defaultdict(list)
    audit = []

    if not ballots:
        return dict(assigned), audit

    if max_rounds_override is None:
        max_rounds = max(b["MAX_ITEMS"] for b in ballots)
    else:
        max_rounds = int(max_rounds_override)

    member_ids = [b["MEMBER_ID"] for b in ballots]
    ballot_map = {b["MEMBER_ID"]: b for b in ballots}

    for round_no in range(1, max_rounds + 1):
        rnd.shuffle(member_ids)
        if verbose:
            print(f"Round {round_no}: order: {member_ids}")
        for mid in member_ids:
            b = ballot_map[mid]
            if len(assigned[mid]) >= b["MAX_ITEMS"]:
                if verbose:
                    print(f"  {mid} skipped (capacity reached)")
                continue

            picked = None
            reason = ""
            prefs_tried = []
            size_matched_flag = ""  # will be "Y" or "N" or "" if no pick

            for p in b["PREFS"]:
                if p == "":
                    continue
                prefs_tried.append(p)
                if p in remaining_items and p not in b["EXCLUDE"]:
                    picked = p
                    if b["SIZE_PREF"]:
                        size_matched_flag = "Y" if items_dict.get(p, {}).get("SIZE","") in b["SIZE_PREF"] else "N"
                    else:
                        size_matched_flag = "Y"
                    reason = "ranked_pref"
                    break

            if picked is None and b["FALLBACK_OK"]:
                def eligible_iid(iid):
                    if iid not in remaining_items:
                        return False
                    if iid in b["EXCLUDE"]:
                        return False
                    if b["SIZE_PREF"]:
                        return items_dict.get(iid, {}).get("SIZE","") in b["SIZE_PREF"]
                    return True
                elig_items = [iid for iid in remaining_items if eligible_iid(iid)]
                if elig_items:
                    picked = rnd.choice(elig_items)
                    reason = "fallback_random"
                    size_matched_flag = "Y"  # fallback only selected from size-matching items

            if picked:
                assigned[mid].append(picked)
                remaining_items.remove(picked)
                audit.append({
                    "round": round_no,
                    "member_id": mid,
                    "member_name": b.get("MEMBER_NAME", ""),
                    "prefs_tried": ",".join(prefs_tried[:10]),
                    "picked": picked,
                    "reason": reason,
                    "size_matched": size_matched_flag
                })
                if verbose:
                    print(f"  {mid} -> {picked} ({reason}) size_matched={size_matched_flag}")
            else:
                audit.append({
                    "round": round_no,
                    "member_id": mid,
                    "member_name": b.get("MEMBER_NAME", ""),
                    "prefs_tried": ",".join(prefs_tried[:10]),
                    "picked": "",
                    "reason": "no_pick",
                    "size_matched": ""
                })
                if verbose:
                    print(f"  {mid} -> no pick")

        if not remaining_items:
            if verbose:
                print("No remaining items; stopping early.")
            break
        all_filled = all(len(assigned[b["MEMBER_ID"]]) >= b["MAX_ITEMS"] for b in ballots)
        if all_filled:
            if verbose:
                print("All members reached capacity; stopping early.")
            break

    return dict(assigned), audit

# ---------------------------
# Output writers (include SIZE_MATCHED in audit)
# ---------------------------

def write_outputs(assigned_map, audit_list, seed, out_prefix):
    alloc_path = f"{out_prefix}_allocations.csv"
    with open(alloc_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["MEMBER_ID", "MEMBER_NAME", "ASSIGNED_SEQS"])
        member_name_lookup = {}
        for a in audit_list:
            if a["member_id"] not in member_name_lookup and a.get("member_name"):
                member_name_lookup[a["member_id"]] = a["member_name"]
        for m, seqs in assigned_map.items():
            name = member_name_lookup.get(m, "")
            w.writerow([m, name, ";".join(seqs)])
    audit_path = f"{out_prefix}_audit_log.csv"
    with open(audit_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        run_ts = datetime.utcnow().isoformat()
        w.writerow(["RUN_TIMESTAMP", run_ts])
        w.writerow(["SEED", seed])
        w.writerow([])  # blank line
        w.writerow(["ROUND", "MEMBER_ID", "MEMBER_NAME", "PREFS_TRIED", "PICKED_SEQ", "REASON", "SIZE_MATCHED"])
        for a in audit_list:
            w.writerow([
                a["round"],
                a["member_id"],
                a["member_name"],
                a["prefs_tried"],
                a["picked"],
                a["reason"],
                a["size_matched"]
            ])
    print(f"Allocations written to: {alloc_path}")
    print(f"Audit log written to: {audit_path}")

# ---------------------------
# Main function
# ---------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Preference-based artwork ballot allocator (single-tier, equal-members)."
    )
    parser.add_argument("--items", required=True, help="Path to items.csv")
    parser.add_argument("--ballots", required=True, help="Path to ballots.csv")
    parser.add_argument("--seed", type=int, default=2026, help="Random seed")
    parser.add_argument("--out", default="output", help="Output file prefix")
    parser.add_argument("--verbose", action="store_true", help="Verbose output")
    parser.add_argument("--max-rounds", type=int, help="Override max rounds")
    
    args = parser.parse_args()
    
    if not os.path.exists(args.items):
        print(f"Error: Items file not found: {args.items}", file=sys.stderr)
        sys.exit(1)
    if not os.path.exists(args.ballots):
        print(f"Error: Ballots file not found: {args.ballots}", file=sys.stderr)
        sys.exit(1)
    
    if args.verbose:
        print(f"Loading items from {args.items}...")
    items = load_items(args.items)
    
    if args.verbose:
        print(f"Loading ballots from {args.ballots}...")
    ballots = load_ballots_simple(args.ballots)
    
    if args.verbose:
        print(f"Loaded {len(items)} available items and {len(ballots)} ballots")
        print(f"Using random seed: {args.seed}")
    
    assigned, audit = allocate(items, ballots, seed=args.seed, 
                               max_rounds_override=args.max_rounds, 
                               verbose=args.verbose)
    
    write_outputs(assigned, audit, args.seed, args.out)
    
    total_assigned = sum(len(seqs) for seqs in assigned.values())
    print(f"\nSummary:")
    print(f"  Total ballots: {len(ballots)}")
    print(f"  Total items assigned: {total_assigned}")
    print(f"  Items remaining unassigned: {len(items) - total_assigned}")

if __name__ == "__main__":
    main()