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
