import sys, csv, re, random, os
import pandas as pd


# ─── CONFIG ───────────────────────────────────────────────────────────────────
PREF_FILE = os.path.join(os.path.dirname(__file__), "preferences.txt")
OUTFILE   = "new_schedules.xlsx"

DAYS = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]

# shift windows in minutes since midnight
EARLY_START, EARLY_END = 4*60 + 30, 12*60    # 4:30 A–12 P
HOUR_START,  HOUR_END  = 6*60, 23*60         # 6 A–11 P window

EARLY_LABEL = "4:30–12:00"
HOUR_LABEL  = "6:00–23:00"

DAY_MAP = {
    "M":  "Monday",  "Tu": "Tuesday", "W":  "Wednesday",
    "Th": "Thursday","F":  "Friday",  "Sa": "Saturday",
    "Su": "Sunday"
}

# ─── TIME PARSER ───────────────────────────────────────────────────────────────
def parse_time(t):
    m = re.match(r'^(\d{1,2})(?::(\d{2}))?([AP])$', t.strip(), re.IGNORECASE)
    if not m:
        raise ValueError(f"Invalid time format: {t!r}")
    hour = int(m.group(1))
    minute = int(m.group(2) or 0)
    period = m.group(3).upper()
    if period == 'A':
        if hour == 12: hour = 0
    else:  # 'P'
        if hour != 12: hour += 12
    return hour * 60 + minute

# ─── check for locked Excel file ──────────────────────────────────────────────
if os.path.exists(OUTFILE):
    try:
        # try opening for append+exclusive lock
        with open(OUTFILE, 'a'):
            pass
    except PermissionError:
        print(f"Error: cannot write to '{OUTFILE}' – it looks like the file is open in Excel.\n"
              "Please close it and run this script again.", file=sys.stderr)
        sys.exit(1)

# ─── now safe to proceed with your ExcelWriter─────────────────────────────────


# ─── LOAD PREFERENCES ─────────────────────────────────────────────────────────
employees   = []
preferences = {}

try:
    f = open(PREF_FILE, newline='')
    print("opened", PREF_FILE, "successfully")
    lines = f.readlines()
    print(f"Found {len(lines)} employees")
    f.seek(0)
except FileNotFoundError:
    print(f"Error: preferences file '{PREF_FILE}' not found.", file=sys.stderr)
    sys.exit(1)
except PermissionError:
    print(f"Error: cannot read '{PREF_FILE}' (permission denied).", file=sys.stderr)
    sys.exit(1)

reader = csv.reader(f)
print("Starting schedule generation now...")
for lineno, row in enumerate(reader, start=1):
    if len(row) < 3:
        print(f"Warning: line {lineno} malformed (need ≥3 columns): {row}", file=sys.stderr)
        continue

    name, role, *prefs_raw = [cell.strip() for cell in row]
    role = role.lower()
    if role not in ("cashier","kitchen"):
        print(f"Warning: line {lineno} has invalid role '{role}'", file=sys.stderr)
        continue

    employees.append({"name":name, "role":role})
    prefs = []

    for p in prefs_raw:
        if p.startswith("*"):
            # explicitly non-schedulable
            continue

        parts = p.split()
        if len(parts) != 2:
            print(f"Warning: line {lineno}, can't parse preference '{p}'", file=sys.stderr)
            continue

        day_abbr, timerange = parts
        day = DAY_MAP.get(day_abbr)
        if not day:
            print(f"Warning: line {lineno}, unknown day '{day_abbr}'", file=sys.stderr)
            continue

        try:
            start_s, end_s = timerange.split("-",1)
            start_min = parse_time(start_s)
            end_min   = parse_time(end_s)
        except ValueError as e:
            print(f"Warning: line {lineno}, {e}", file=sys.stderr)
            continue

        prefs.append((day, start_min, end_min))

    preferences[name] = prefs

f.close()

# confirm all parsed times are ints
for name, lst in preferences.items():
    for day, s, e in lst:
        assert isinstance(s, int) and isinstance(e, int), f"Bad entry for {name}: {(day,s,e)}"

# ─── SETUP POOLS & ROW LABELS ─────────────────────────────────────────────────
cashiers = [e["name"] for e in employees if e["role"] == "cashier"]
kitchens = [e["name"] for e in employees if e["role"] == "kitchen"]
row_labels = [f"{e['name']} ({e['role']})" for e in employees]

# ─── SCHEDULE GENERATION ──────────────────────────────────────────────────────
with pd.ExcelWriter(OUTFILE, engine="xlsxwriter") as writer:
    random.seed(42)

    for sched_num in range(1, 4):
        df = pd.DataFrame(index=row_labels, columns=DAYS)
        df[:] = ""
        warnings = []

        for day in DAYS:
            # 1) Early full-day shift coverage
            avail_ce = [c for c in cashiers if any(d==day and s<=EARLY_START and e>=EARLY_END for d,s,e in preferences.get(c,[]))]
            avail_ke = [k for k in kitchens if any(d==day and s<=EARLY_START and e>=EARLY_END for d,s,e in preferences.get(k,[]))]

            if len(avail_ce) < 1:
                warnings.append(f"Sched {sched_num} • {day} early: need ≥1 cashier, got {len(avail_ce)}")
                avail_ce = cashiers.copy()
            if len(avail_ke) < 1:
                warnings.append(f"Sched {sched_num} • {day} early: need ≥1 kitchen, got {len(avail_ke)}")
                avail_ke = kitchens.copy()

            ec = random.choice(avail_ce)
            ek = random.choice(avail_ke)
            df.at[f"{ec} (cashier)", day] = EARLY_LABEL
            df.at[f"{ek} (kitchen)",  day] = EARLY_LABEL

            # 2) Hour-by-hour coverage 6A–11P
            # build a list of all staff still on early shift
            on_early = {ec, ek}

            for hour in range(HOUR_START, HOUR_END, 60):
                # find who can work this hour
                def available(staff_list):
                    return [
                        name for name in staff_list
                        if any(d==day and s<=hour and e>=hour+60 for d,s,e in preferences.get(name,[]))
                    ]

                # combine early-shift folks plus those newly available
                cash_pool = list(on_early & set(cashiers)) + available(cashiers)
                kit_pool  = list(on_early & set(kitchens)) + available(kitchens)

                # enforce minima
                if len(cash_pool) < 2:
                    warnings.append(f"Sched {sched_num} • {day} {hour//60}:00: need ≥2 cashiers, got {len(cash_pool)}")
                    cash_pool = cashiers.copy()
                if len(kit_pool) < 2:
                    warnings.append(f"Sched {sched_num} • {day} {hour//60}:00: need ≥2 kitchens, got {len(kit_pool)}")
                    kit_pool = kitchens.copy()

                # pick distinct pairs
                chosen_c = random.sample(cash_pool, 2)
                chosen_k = random.sample(kit_pool,  2)

                for c in chosen_c:
                    df.at[f"{c} (cashier)", day] = HOUR_LABEL
                for k in chosen_k:
                    df.at[f"{k} (kitchen)", day] = HOUR_LABEL

        # write and emit warnings
        sheet = f"Schedule {sched_num}"
        df.to_excel(writer, sheet_name=sheet)
        if warnings:
            print(f"\n⚠️  Warnings for {sheet}:")
            for w in warnings:
                print("  -", w)

    print(f"\nDone! → {OUTFILE}")
