"""Step D: scrub the SFRA phone slots using the screening workbook beside this script."""
import csv
import os
import re
from collections import defaultdict, deque
from functools import lru_cache
from pathlib import Path
from tempfile import NamedTemporaryFile

import pandas as pd


# True moves each ordinary DNC number to its matching "<phone column> DNC Flag" field.
# TCPA, TROLL, DNC complainer, and NRA matches exclude leads in both modes.
EXPORT_DNC_NUMBERS = False

PHONE_COLUMNS = (
    "First Phone", "Second Phone", "Third Phone",
    "Propstream Phone 1", "Propstream Phone 2", "Propstream Phone 3",
    "Propstream Phone 4", "Propstream Phone 5",
)
REGISTRY_COLUMNS = tuple(f"{slot} DNC Flag" for slot in PHONE_COLUMNS)
LEGACY_REGISTRY_COLUMNS = tuple(f"Registry DNC {n}" for n in range(1, 9))
BORROWER = "Borrower LLC"
OWNER = "Owner Name"
ADDRESS_FIELDS = ("Mailing Address", "Mailing City", "Mailing State", "Mailing Zip")
STATE_NAMES = set((
    "Alabama|Alaska|Arizona|Arkansas|California|Colorado|Connecticut|Delaware|"
    "District of Columbia|Florida|Georgia|Hawaii|Idaho|Illinois|Indiana|Iowa|Kansas|"
    "Kentucky|Louisiana|Maine|Maryland|Massachusetts|Michigan|Minnesota|Mississippi|"
    "Missouri|Montana|Nebraska|Nevada|New Hampshire|New Jersey|New Mexico|New York|"
    "North Carolina|North Dakota|Ohio|Oklahoma|Oregon|Pennsylvania|Rhode Island|"
    "South Carolina|South Dakota|Tennessee|Texas|Utah|Vermont|Virginia|Washington|"
    "West Virginia|Wisconsin|Wyoming|Puerto Rico|Guam|Virgin Islands|US Virgin Islands"
).casefold().split("|"))
MISSING_KEYS = {"", "nan", "none", "null", "<na>", "na", "n/a", "#n/a"}
HARD_TAG_PATTERN = r"tcpa|dnc complainer|(?<![a-z0-9])(?:nra|troll)(?![a-z0-9])"


def match_key(value):
    key = " ".join(value.casefold().split())
    return "" if key in MISSING_KEYS else key


def identity_key(field, value):
    key = match_key(value)
    # Excel can drop the leading zeros of a numeric US ZIP.
    if field == "Mailing Zip" and re.fullmatch(r"[0-9]{1,5}", key):
        return key.zfill(5)
    return key


def normalize_phone(value):
    value = re.sub(r"^(\d+)\.0$", r"\1", value.strip())
    digits = re.sub(r"\D", "", value)
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    return digits if len(digits) == 10 else ""


@lru_cache(maxsize=512)
def classify_status(value):
    # The vendor can combine several tags, including "TCPA troll", in one cell.
    tags = {" ".join(t.casefold().replace("_", " ").replace("-", " ").split())
            for t in value.split("|")}
    if any(re.search(HARD_TAG_PATTERN, tag) for tag in tags):
        return "hard"

    registry = False
    for tag in tags:
        if tag in {"", "clear", "duplicate"}:
            continue
        if tag in {"dnc", "federal dnc", "state dnc"} or (
            tag.endswith(" dnc") and tag[:-4] in STATE_NAMES
        ):
            registry = True
        else:
            raise ValueError(f"Unrecognized screening tag {tag!r}")
    return "registry" if registry else ""


def one_input(folder, extensions):
    folder.mkdir(exist_ok=True)
    paths = [p for p in folder.iterdir() if p.is_file()
             and p.suffix.lower() in extensions and not p.name.startswith("~$")]
    if len(paths) != 1:
        raise ValueError(f"Expected one {'/'.join(extensions)} file in {folder}, found {len(paths)}")
    return paths[0]


def load_screening(path, primary):
    # Preserve duplicate Type headers so each status stays paired with its own phone.
    sheet = pd.read_excel(path, dtype=str, keep_default_na=False, header=None)
    if sheet.empty:
        raise ValueError("The screening workbook is empty")
    headers = [str(h).strip() for h in sheet.iloc[0]]
    data = sheet.iloc[1:]
    borrower = BORROWER if BORROWER in headers else "Borrower"
    fields = {BORROWER: borrower, OWNER: OWNER}
    fields.update({k: k for k in ADDRESS_FIELDS if k in headers and k in primary})
    selected = {}
    for target, header in fields.items():
        if headers.count(header) != 1:
            raise ValueError(f"Expected one {header!r} column in the screening workbook")
        selected[target] = data.iloc[:, headers.index(header)]

    slots = []
    status_positions = set()
    for slot in PHONE_COLUMNS:
        if slot not in headers and slot not in primary:
            continue
        if headers.count(slot) != 1:
            raise ValueError(f"Expected exactly one screening column for {slot!r}")
        i = headers.index(slot)
        if i + 1 >= len(headers) or headers[i + 1].casefold() not in {"type", f"{slot} type".casefold()}:
            raise ValueError(f"The screening Type must be immediately to the right of {slot!r}")
        selected[slot] = data.iloc[:, i]
        selected[slot + " status"] = data.iloc[:, i + 1]
        status_positions.add(i + 1)
        slots.append(slot)

    # Some exports add Type to non-phone fields (e.g. Phone Verified).
    # Never silently skip a flagged phone that this program cannot identify.
    for i, header in enumerate(headers):
        paired_header = i > 0 and header.casefold() == f"{headers[i-1]} type".casefold()
        if (header.casefold() == "type" or paired_header) and i not in status_positions:
            if data.iloc[:, i].str.contains(rf"{HARD_TAG_PATTERN}|dnc|complainer", case=False).any():
                raise ValueError(f"Screening flags found beside unsupported field {headers[i-1]!r}")
    return pd.DataFrame(selected).reset_index(drop=True), slots


def screening_sets(primary, screening, slots):
    identity = [BORROWER, OWNER, *(k for k in ADDRESS_FIELDS if k in screening)]
    source_keys = [{k: identity_key(k, v) for k, v in record.items()}
                   for record in primary[identity].to_dict("records")]
    anchors = [k for k in (BORROWER, OWNER, "Mailing Address") if k in identity]
    index = {k: defaultdict(list) for k in anchors}
    for i, keys in enumerate(source_keys):
        for k in anchors:
            if keys[k]:
                index[k][keys[k]].append(i)

    borrowers, owners, hard_phones, registry_phones, direct = set(), set(), set(), set(), set()
    masks_resolved = 0
    for excel_row, row in enumerate(screening.to_dict("records"), start=2):
        candidates = None
        for slot in slots:
            location = f"Screening row {excel_row}, {slot}"
            try:
                status = classify_status(row[slot + " status"])
            except ValueError as exc:
                raise ValueError(f"{location}: {exc}") from exc
            if not status:
                continue
            raw = row[slot].strip()
            digits = normalize_phone(raw)
            mask = re.fullmatch(r"\*+(\d{2,4})", raw)

            if status == "hard" or mask:
                if candidates is None:
                    keys = {k: identity_key(k, row[k]) for k in identity}
                    anchor = next((k for k in anchors if keys[k]), None)
                    possible = index[anchor][keys[anchor]] if anchor else []
                    candidates = [i for i in possible if all(
                        not val or source_keys[i][k] == val for k, val in keys.items()
                    )]
                    if not candidates:
                        raise ValueError(f"{location}: cannot link flagged lead to the SFRA input")
                if mask:
                    if len(candidates) != 1 or slot not in primary:
                        raise ValueError(f"{location}: masked number needs a unique source lead and phone slot")
                    digits = normalize_phone(primary.at[candidates[0], slot])
                    if not digits.endswith(mask[1]):
                        raise ValueError(f"{location}: source phone does not match the masked ending")
                    masks_resolved += 1
                if status == "hard":
                    if len(candidates) > 1 and not (keys[BORROWER] or keys[OWNER]):
                        raise ValueError(f"{location}: ambiguous flagged lead without an owner or LLC")
                    direct.update(candidates)
                    # Use source names too when the linked workbook row omits one.
                    for field, blocked in ((BORROWER, borrowers), (OWNER, owners)):
                        if keys[field]:
                            blocked.add(keys[field])
                        elif len(candidates) == 1 and source_keys[candidates[0]][field]:
                            blocked.add(source_keys[candidates[0]][field])
            if not digits:
                raise ValueError(f"{location}: flagged phone could not be resolved to ten digits")
            if status == "hard":
                hard_phones.add(digits)
            else:
                registry_phones.add(digits)
    registry_phones.difference_update(hard_phones)
    print(f"Blocked borrower names: {len(borrowers):,}; blocked owner names: {len(owners):,}")
    print(f"Hard-blocked phones: {len(hard_phones):,}; ordinary DNC phones: {len(registry_phones):,}")
    print(f"Masked entries resolved: {masks_resolved:,}")
    excluded = (primary.index.isin(direct) | primary[BORROWER].map(match_key).isin(borrowers)
                | primary[OWNER].map(match_key).isin(owners))
    return excluded, hard_phones, registry_phones


NAME_SUFFIXES = {"jr", "sr", "ii", "iii", "iv", "v", "vi"}
SURNAME_PARTICLES = {
    "al", "bin", "da", "das", "de", "del", "den", "der", "des", "di",
    "do", "dos", "du", "el", "ibn", "la", "las", "le", "los", "saint",
    "san", "santa", "st", "ten", "ter", "van", "von", "zu", "zum", "zur",
}
ENTITY_MARKERS = {
    "corp", "corporation", "dba", "fka", "inc", "incorporated", "irrevocable",
    "llc", "llp", "lp", "ltd", "pllc", "revocable", "trustee",
}
ENTITY_ENDINGS = {
    "association", "company", "enterprises", "estate", "foundation", "fund",
    "holdings", "investments", "partnership", "properties", "realty", "rentals",
    "trust",
}


def is_name_suffix(word):
    return word.casefold().replace(".", "") in NAME_SUFFIXES


def remove_name_suffixes(words):
    while len(words) > 1 and is_name_suffix(words[-1]):
        words = words[:-1]
    return words


def is_person_name_word(word):
    if word.endswith("."):
        stem = word[:-1]
        return (len(stem) == 1 and stem.isalpha()) or stem.casefold() == "st"

    return all(part.isalpha() for part in re.split(r"['’ʼ\-‐‑]", word))


def split_owner_name(value):
    """Return the first given name and surname from an individual owner name."""
    if not isinstance(value, str) or not value.strip():
        return "", ""

    comparable = value.casefold().replace(".", "").replace(",", " ")
    words = comparable.split()
    if (
        not words
        or ENTITY_MARKERS.intersection(words)
        or words[-1] in ENTITY_ENDINGS
        or words[:2] == ["estate", "of"]
        or {"and", "or"}.intersection(words[1:-1])
        or " et al " in f" {' '.join(words)} "
    ):
        return "", ""

    parts = [part.split() for part in value.split(",")]
    if any(not part for part in parts):
        return "", ""

    # Recognize both "Given Surname, Jr." and "Surname, Given, Jr.".
    if len(parts) > 1 and all(is_name_suffix(word) for word in parts[-1]):
        parts.pop()

    if len(parts) == 2:
        surname, given_names = (remove_name_suffixes(part) for part in parts)
    elif len(parts) == 1:
        full_name = remove_name_suffixes(parts[0])
        if len(full_name) < 2:
            return "", ""

        surname_start = len(full_name) - 1
        while (
            surname_start > 1
            and full_name[surname_start - 1].casefold().rstrip(".") in SURNAME_PARTICLES
        ):
            surname_start -= 1

        given_names = full_name[:surname_start]
        surname = full_name[surname_start:]
    else:
        return "", ""

    if not all(is_person_name_word(word) for word in given_names + surname):
        return "", ""

    return given_names[0], " ".join(surname)


def cover_one_more_lead(start, choices, winners, counts):
    """Find the shortest chain of phone transfers that gives this lead a number."""
    queue = deque([start])
    previous = {start: None}
    while queue:
        row = queue.popleft()
        for phone, slot in choices[row]:
            donor, _ = winners[phone]
            if donor in previous:
                continue
            previous[donor] = (row, phone, slot)
            if counts[donor] == 1:
                queue.append(donor)
                continue

            # Only the two endpoints change their number of assigned phones.
            counts[donor] -= 1
            counts[start] += 1
            while donor != start:
                row, phone, slot = previous[donor]
                winners[phone] = (row, slot)
                donor = row
            return


def deduplicate_export_phones(frame):
    """Keep each number in one original slot. An SFRA occurrence always outranks Propstream."""
    counts = [0] * len(frame)
    removed = 0
    # Screening makes the two sets disjoint. Assign regular phones first so that
    # enabling DNC export cannot change which leads keep their regular numbers.
    for columns in (PHONE_COLUMNS, REGISTRY_COLUMNS):
        occurrences = defaultdict(list)
        for slot, column in enumerate(columns):
            if column in frame:
                for row, value in enumerate(frame[column]):
                    if value:
                        phone = normalize_phone(value)
                        occurrences[phone].append((row, slot))

        winners = {}
        choices = defaultdict(list)
        for phone, cells in occurrences.items():
            # Collected in slot order, then original row order. Slots 0–2 are SFRA.
            row, first_slot = cells[0]
            winners[phone] = cells[0]
            counts[row] += 1
            for candidate, slot in cells:
                if first_slot < 3 and slot >= 3:
                    break
                choices[candidate].append((phone, slot))

        for options in choices.values():
            options.sort(key=lambda option: option[1])
        for row in range(len(frame)):
            if counts[row] == 0 and row in choices:
                cover_one_more_lead(row, choices, winners, counts)

        clear = defaultdict(list)
        for phone, cells in occurrences.items():
            for row, slot in cells:
                if (row, slot) != winners[phone]:
                    clear[slot].append(row)
        for slot, rows in clear.items():
            frame.loc[frame.index.take(rows), columns[slot]] = ""
            removed += len(rows)
    return removed, sum(count > 0 for count in counts)


def scrub(primary, screening, slots, export_dnc):
    excluded, hard, registry = screening_sets(primary, screening, slots)
    kept = primary.loc[~excluded].drop(columns=["First Name", "Last Name"], errors="ignore").copy()
    names = [split_owner_name(owner) for owner in kept[OWNER]]
    # The SFRA layout has 23 lead fields before X and Y.
    start = min(23, len(kept.columns))
    kept.insert(start, "First Name", [first for first, _ in names])
    kept.insert(start + 1, "Last Name", [last for _, last in names])

    invalid = 0
    for slot, dedicated in zip(PHONE_COLUMNS, REGISTRY_COLUMNS):
        kept[dedicated] = ""
        if slot not in kept:
            continue
        digits = kept[slot].map(normalize_phone)
        ordinary = digits.isin(registry)
        if export_dnc:
            kept.loc[ordinary, dedicated] = kept.loc[ordinary, slot]
        invalid += int((digits.eq("") & kept[slot].str.strip().ne("")).sum())
        kept.loc[digits.eq("") | digits.isin(hard) | ordinary, slot] = ""

    removed, reachable = deduplicate_export_phones(kept)

    # Verify the finished phone fields before replacing any export.
    seen = set()
    moved = 0
    for slot, dedicated in zip(PHONE_COLUMNS, REGISTRY_COLUMNS):
        if slot in kept and kept[slot].map(normalize_phone).isin(hard | registry).any():
            raise RuntimeError(f"Blocked number in regular output field {slot!r}")
        values = kept.loc[kept[dedicated].ne(""), dedicated].map(normalize_phone)
        if not values.empty and (not export_dnc or not values.isin(registry).all()):
            raise RuntimeError(f"Unexpected number in {dedicated!r}")
        moved += len(values)
        for column in (slot, dedicated):
            if column in kept:
                keys = kept.loc[kept[column].ne(""), column].map(normalize_phone)
                if keys.duplicated().any() or seen.intersection(keys):
                    raise RuntimeError(f"Duplicate number in export field {column!r}")
                seen.update(keys)

    print(f"Leads: {len(primary):,} input, {int(excluded.sum()):,} excluded, {len(kept):,} kept")
    print(f"Leads with an exported phone: {reachable:,}; duplicate phone cells cleared: {removed:,}")
    print(f"DNC phone cells exported: {moved:,}; invalid phone cells omitted: {invalid:,}")
    return kept


def run(base_dir, export_dnc):
    if type(export_dnc) is not bool:
        raise ValueError("Set EXPORT_DNC_NUMBERS to True or False (without quotes)")
    csv_path = one_input(base_dir / "StepD - Input SFRA Leads", (".csv",))
    workbook_path = one_input(base_dir / "StepD - Input TCPA Liti DNC Match", (".xlsx", ".xls"))
    folder = base_dir / "StepD - Output SFRA Leads Scrubbed"
    folder.mkdir(exist_ok=True)
    # pandas silently renames duplicate headers; a second First Phone could bypass screening.
    with csv_path.open(encoding="utf-8-sig", newline="") as stream:
        headers = next(csv.reader(stream), [])
    if len(set(headers)) != len(headers):
        raise ValueError("Duplicate SFRA input headers: each phone slot must be identifiable")
    primary = pd.read_csv(csv_path, dtype=str, keep_default_na=False)
    for field in (BORROWER, OWNER):
        if field not in primary:
            raise ValueError(f"The SFRA input is missing {field!r}")
    if not any(slot in primary for slot in PHONE_COLUMNS):
        raise ValueError("The SFRA input has no recognized phone columns")
    if any(col in primary for col in REGISTRY_COLUMNS + LEGACY_REGISTRY_COLUMNS):
        raise ValueError("Use the original Step C leads as input, not a previous Step D export")
    screening, slots = load_screening(workbook_path, primary)
    result = scrub(primary, screening, slots, export_dnc)
    suffix = "_phones_scrubbed_dnc_included" if export_dnc else "_phones_scrubbed"
    output = folder / (csv_path.stem + suffix + csv_path.suffix)
    temp = None
    try:
        with NamedTemporaryFile(mode="w", encoding="utf-8", newline="", dir=folder,
                                suffix=".tmp", delete=False) as stream:
            temp = Path(stream.name)
            result.to_csv(stream, index=False)
        os.replace(temp, output)
    finally:
        if temp:
            temp.unlink(missing_ok=True)
    print(f"Saved {output}")
    return output


if __name__ == "__main__":
    run(Path(__file__).resolve().parent, EXPORT_DNC_NUMBERS)

