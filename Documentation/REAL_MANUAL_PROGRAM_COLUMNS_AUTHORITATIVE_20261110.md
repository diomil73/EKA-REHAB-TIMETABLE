# Real manual schedule | authoritative operator column semantics (2026-11-10)

**Status:** user-defined interpretation rules, not yet implemented or independently verified against every source cell. The earlier machine-parsed 271-row staging preview is **REJECTED as an import authority** and must not be used to populate patient schedules, slot reservations, or finalized history. The original manually maintained `ΠΡΟΓΡΑΜΜΑΤΑ  ΘΕΡΑΠΕΙΩΝ ΑΣΘΕΝΩΝ 10-11-26.xlsx` remains read-only, separate from disposable H10 synthetic test workbooks.

## Meaning of visible worksheet `ΠΡΟΓΡΑΜΜΑ` columns

| Column | Meaning | Import / UI rule |
| --- | --- | --- |
| A | Clinic and ward (e.g. Α1, Α2), and treating physician for patient on SAME row | Metadata belonging to row's patient; never read as session therapist |
| B | Patient, often surname only | Yellow fill flags infectious patient. Red-letter `διακ` denotes **transferred** status, usually treatment entries crossed out. `διακοπη θερ` denotes **medically stopped treatment program**, a distinct status. Store status and provenance rather than inventing completed sessions |
| C | Physiotherapy (ΦΘ): time, therapist, often weekday codes `ΔΕ-ΤΕ-ΠΑ` | No days marked means **daily** in source schedule. First time + first therapist represent original plan. A subsequent time or therapist for replacement supersedes initial for delivered assignment; preserve original and linked replacement rather than counting two independent appointments. Needs human-validated parser for variable notation |
| D | `Ανακλινόμενο` | Time and weekdays; preserve service category separately from ordinary ΦΘ |
| E | Occupational therapy | Time, day(s), therapist name(s). Even if TWO therapist names occur, **only first** will be optionally retained as primary if entering therapist info; other specialties primarily reserve timeslots by time/day. `(τ)` means technology, not within current import scope. **Different times on different days must be SPLIT into separate day/time reservations** (operator example E104) |
| F | Speech therapy | Time, weekdays, therapist sometimes; for now do not require structured therapist ID, leave **free-text** therapist field for specialty to fill |
| G | Psychologists | Same timeslot-by-day rule; if no explicit time, **no slot can be assigned**. Leave blank rather than guessing |
| H | Ειδική Φυσική Αγωγή (ΕΦΑ) | Time and days; only one colleague currently, so therapist not written in source. Different time/day combinations must be split into separate day/time lines. Do not invent an ID merely from uniqueness without registry confirmation |
| I | Robotics | Time, weekdays, physiotherapist who performs it. Preserve distinct service and mapped slot; physiotherapist identity needs verified alias/registry handling |
| J | Two afternoon physiotherapists | Not included in Diomil account/scheduling yet. **Future policy decision**, not auto-import as ordinary main-shift capacity |

## Global interpretation rules
1. The patient's row is the organizing record for A..J. Preserve source workbook, sheet, cell address, text, formatting (especially yellow/red and strikethrough), original time/day strings, and explicit validation state.
2. When weekday codes are absent in C, source means daily. Do not silently extend that rule to other columns without clarification.
3. The first time/first therapist in C are the **initially scheduled** facts. A following replacement therapist and/or time in parentheses or in a clear continuation is the **replacement** that becomes current after the change; maintain both original and replacement on one linked appointment, NOT an automatically counted extra session. Earlier generic rule about parentheses is subordinate to the user's precise C-column description, and day codes/notes in parentheses are NOT automatically replacements.
4. Split E and H day-specific time combinations: e.g. separate booking rows for time 1 on day group 1 and time 2 on day group 2; preserve single originating cell so edits trace back. Do not multiply an unspecified date into actual dated appointments.
5. Other specialties use time/day as their reservation contract; therapist names are optional operational text where noted. Main ΦΘ therapist assignment rules must not be imposed on specialties.
6. Never reserve timeslot when time is missing; never infer actual attendance or treatment delivery solely from a planning grid. Crossed-out therapy entries under transferred status are not delivered.
7. `διακ` (transferred) and `διακοπη θερ` (medical stop) are semantically distinct and must remain distinct from ordinary absence, replacement, and treatment cancellation.
8. Preserve unrecognized tokens for review and ask user targeted questions. NO writeback to production Diomil, patient status, current schedule, attendance, or histories until a source-grounded day/date mapping, patient/therapist IDs, and strict validation are established.

## Ambiguities to resolve before authoritative converter
- How does C indicate **replacement continuation** when the additional name/time is not parenthesized? Is there a distinctive separator, line break, strikeout, or color, and do replacement and original sometimes share a slot?
- What precisely does **daily** mean in this snapshot: Monday–Friday of an agreed week, or all scheduled operating days, and how is the effective week/date obtained? Filename date alone is not a reliable treatment date.
- How to distinguish yellow infectious fill from other yellow formatting and red status markings; inspect Excel formatting (and conditional formatting if used), not textual content only.
- Whether `Ανακλινόμενο` uses the same shared capacity pool as other specialties. No assumptions until scheduler contract is verified.
- Column J future admission must be discussed separately.

## Design next steps
- Build an **isolated, fail-closed read-only prototype** that expands row+column source cell into explicit day/time/service segments, preserving raw source and flags; blocked rows remain unresolved, not silently accepted.
- Prioritize representative C replacement forms, E104, H split patterns, B transfer/stoppage/cross-out, G missing hours, and J out-of-scope records. Surface each original cell for user review.
- Keep H10 CC frozen as synthetic demo. Real parallel pilot begins only after separate admission and native checks.


## 2026-11-10 — Operator-confirmed concrete replacement examples in column C
- **C33** `10.45 ΓΑΒΡΑΣ (8.30 ΖΗΜΑΡΗΣ)`: initial **10:45 / ΓΑΒΡΑΣ**; replacement **08:30 / ΖΗΜΑΡΗΣ**. The parenthesized time and therapist supersede both original fields for the replacement appointment while retaining original facts.
- **C8** `09.15 ΨΙΑΧΟΥ (ΚΑΛΥΒΙΩΤΗΣ)`: initial **09:15 / ΨΙΑΧΟΥ**; replacement **09:15 / ΚΑΛΥΒΙΩΤΗΣ**. Parentheses change therapist only. **Inherit initial time** when replacement has no explicit time; do not mark hour missing or make a second appointment.
- Parser rule: start with original time+therapist; if a confirmed parenthetical replacement specifies a new therapist but no time, carry original time forward; if it specifies both time+therapist, replace both; always preserve linked original and replacement versions and exact source cell. Applies to these confirmed patterns; unrecognized forms remain HOLD.


## 2026-11-10 — Operator: tolerant reading of hurried handwritten-style Excel entry
Manual data entry is rushed/inconsistent: parentheses may be present or omitted; times may use `.`, `:`, `..` or no separator at all; spaces around names, times and parentheses may be missing. **Do NOT require perfect punctuation for initial parsing.**

Implementation contract:
- Use tolerant lexical normalization on a *working copy of the cell text*: trim/collapse whitespace, handle Greek and Latin punctuation, recognize clock candidates with `:`, `.`, `..` and cautiously recognize compact four-/three-digit clock formats only when context and legal clock bounds support it. Validate hour 0–23 and minute 0–59, do not confuse weekday abbreviations or ward codes with times. Keep original exact text and provenance.
- Tokenize plausible therapist aliases and candidate subsequent therapist/time segments whether or not wrapped in parentheses. If original and candidate replacement are unambiguous against therapist registry and slot/date context, produce a *proposed linked replacement* retaining original. Lack of parentheses alone neither proves nor disproves replacement.
- Do **not** silently turn any second name into a definitive replacement; it may represent another clinician, another service, or a note. When multiple plausible readings remain, mark NEEDS_REVIEW with exact cell, options and reason.
- For clearly confirmed replacement where only therapist changes, inherit original time; where time and therapist change, retain both versions with their correct effective values. Avoid producing two independent treatments from one appointment.
- Be liberal in **text recognition**, conservative in **assignment to a real patient/therapist/time slot**, and strict in conflicts/capacity/provenance checks. No production writeback until validation.
- Regression fixture set should cover `09.15 ΨΙΑΧΟΥ (ΚΑΛΥΒΙΩΤΗΣ)`, `10.45 ΓΑΒΡΑΣ (8.30 ΖΗΜΑΡΗΣ)`, plus equivalent missing-parenthesis/no-space/dot/colon/double-dot/compact-clock variants; do not assume those variant entries actually occur in the source without inspecting.


## 2026-11-10 — First executable column-C read-only parser and actual source result
- Code `Python/tools/parse_real_manual_schedule.py` committed `621e1dd`; tests `Tests/test_real_manual_schedule_parser.py` committed `33df0ca`. **Tests were run locally using a matching working copy: 12 passed**. Post-commit GitHub checkout Windows pytest has NOT yet run.
- Actual uploaded source hash verified unchanged: SHA256 `765264ec8eae95a31ccc57629dc31b1d61f59ea5c992221affa91fcf6a8d2567`.
- Read-only parsing visible worksheet `ΠΡΟΓΡΑΜΜΑ`, column C only: **104 nonempty C cells**; **19 PROPOSED_INITIAL**, **5 PROPOSED_REPLACEMENT**, **80 REVIEW**. Not counts of validated patients or finalized sessions. This is expected with a deliberately restricted therapist alias whitelist (7 names) and unmodeled weekday/notes.
- Known real source cases `C8`, `C25`, `C33`, `C36` produce correct proposed linked replacements, including C8 same hour and C33 time + therapist change. New time separators `.`, `:`, `..`, comma, 3/4-digit compact clock and missing spaces/parentheses are supported where lexical pattern is unambiguous.
- Output in this conversation: `/mnt/data/DIOMIL_COLUMN_C_PROPOSALS_20261110.csv` (column C source cell reference, normalized proposed times and names, status and reason; no copied patient names or free notes). 104 rows.
- **Every auto-parsed result is a PROPOSAL, not an approved appointment**. No weekdays/dates/therapist-ID joins, true timeslot writes, attendance history or operational Diomil writeback. Unknown aliases and ambiguous rows stay REVIEW.
- Next prioritize separate E/H weekday-and-time splitter including real E104 and authorized therapist alias registry, then run fresh Windows tests. Do not restore old rejected 271-row staging assumptions.


## 2026-11-10 — Windows 12/12 C parser tests green; first E/H read-only weekday split
Operator confirmed `12 passed in 1.07s` for `Tests/test_real_manual_schedule_parser.py`.
New source `Python/tools/parse_real_manual_day_time_split.py` commit `b6a0b59`, tests `Tests/test_real_manual_day_time_split.py` commit `cf25a96`.
- E104 literal `11.30(ΔΕ-ΤΕ-ΠΑ)(12.15 ΤΡ-ΠΕ)ΡΑΠΤΗ` => candidate separate weekday/time reservations: ΔΕ/ΤΕ/ΠΑ 11:30, ΤΡ/ΠΕ 12:15; preserve E104 as one originating cell, **not one universal hour**.
- H15 `ΤΡ,ΠΑ11.30,ΤΕ12.15` => ΤΡ/ΠΑ 11:30, ΤΕ 12:15; H107 `ΤΡ 10.00, ΠΕ10.45` => ΤΡ 10:00, ΠΕ 10:45.
- H31 `ΔΕ ΤΡ ΠΕ 11.30, ΠΕ 12.15 SH` => REVIEW due to additional marker and/or duplicated weekday; do not choose a Thursday slot automatically.
- Local source-based check `8 passed in 0.04s`; GitHub-checked-out Windows tests **still need run**. Current uploaded real source SHA256 stayed `765264ec8eae95a31ccc57629dc31b1d61f59ea5c992221affa91fcf6a8d2567`.
- Generated **read-only CSV proposals**, saved `/mnt/data/DIOMIL_E_H_WEEKDAY_TIMES_PROPOSALS_20261110.csv`: 99 nonempty E/H cells (86 E + 13 H), classified **11 PROPOSED_SPLIT, 79 REVIEW, 9 EMPTY/asterisk or headings**, yielding 118 per-day candidate rows total. No actual calendar dates, no patient names, no treatment delivery, no Diomil writeback.
- High REVIEW is deliberate because full E/H free-text therapist-day association has not been validated; limit auto-splitting to explicit day-time groups. Never use this partial split as authority for daily slot reservations.
- Next Windows `python -m pytest Tests/test_real_manual_schedule_parser.py Tests/test_real_manual_day_time_split.py -q`; then improve E/H support for common cases with exactly one time and weekday list, and reconcile overlaps with strict schedule capacity.


## 2026-11-11 — 20/20 Windows tests green; E/H single-time multi-weekday recognition staged
Operator confirmed **`20 passed in 1.33s` / `REAL MANUAL SCHEDULE PARSERS GREEN`** for C and first E/H modules. This confirms previous version only.
- **`01e6023`** expands `Python/tools/parse_real_manual_day_time_split.py`: if one explicit time exists and explicit weekday abbreviations can be found before **or after** the time, propose a reservation on each distinct weekday. In column E only, allow surrounding Greek occupational-therapist name text (one or two names separated by punctuation), and ignore standalone `(Τ)` technology marker. In H, still refuse unexpected non-weekday/clock tokens. Normalize mixed Latin `E/A` in weekday strings for the parser working copy (original source never edited).
- **`8ea8ac2`** adds eight new tests and adjusts one old conservative test. Real source fixtures: `E4` `9:15 ΦΑΝΤΗ ΔΕ-ΤΡΙ-ΠΕ`; `E9` `11:30 ΔΕ - ΤE -ΠA ΦΑΝΤΗ`; `E18` two therapists; `E19` two therapists and `(Τ)`; `H25` `ΤΕ, ΠΑ 8.30`; `H27` `ΔΕ ΤΡ ΠΕ 10.00`. Refuses H therapist name and E cells with no explicit weekdays.
- Direct source inspection in isolated local diagnostic (not the committed parser) found **72** of 99 nonempty E/H cells have an unambiguous *single-time, explicit-weekday* signature under the proposed lexical rule; **11** had no explicit days and **16** were other forms (including two-time), so this is NOT a count of all accepted cells and NOT a verified full export. The raw source has not been modified.
- **Post-change Windows pytest has NOT YET run**. Before claiming a new verified baseline, run `git pull --ff-only origin work/history-statistics-rebuild` followed by `python -m pytest Tests/test_real_manual_schedule_parser.py Tests/test_real_manual_day_time_split.py -q`. Stop on failure. No production booking, no dated slots, no real attendance history writes. Remaining ambiguities include H31 duplicated weekdays, uncertain therapist aliases and week/date mapping; no need for new demonstration XLSM before these gates.
