# v27.1 Excel layout contract

This document records the cell/range structure confirmed against the immutable
`Rehab_Center_System_v27_1.xlsm` baseline. It contains layout metadata only and
must not contain exported patient records.

## MASTER_SCHEDULE

Confirmed used dimension: `A1:L5000`.

Header contract:

| Column | Header | Current role |
|---|---|---|
| A | Μολυσματικός | infectious display/support flag |
| B | Θάλαμος | room |
| C | Ασθενής | patient display name |
| D | ΦΘ | physiotherapy display |
| E | Ρομποτικό | robotic display |
| F | Πισίνα | pool display |
| G | Ανακλινόμενο | reclined display |
| H | Εργο | occupational therapy display |
| I | Λογο | speech therapy display |
| J | ΕΦΑ | EFA display |
| K | Κατάσταση | patient status |

Current cell-level Python dry-run zone: `D2:J5000` only.

`A:C` and `K` are deliberately protected at this stage even though the whole
sheet is a derived view. This prevents a render experiment from damaging the
visible patient identity, room, infectious flag or status columns.

## THERAPIST_DAILY

Confirmed used dimension: `A1:J18`.

The sheet is two therapist blocks with the same seven time rows:

- top headers: `B1:J1`
- top patient grid: `B2:J8`
- top times: `A2:A8`
- bottom headers: `B11:J11`
- bottom patient grid: `B12:J18`
- bottom times: `A12:A18`

Confirmed time sequence in both blocks:

`08:30, 09:15, 10:00, 10:45, 11:30, 12:15, 13:00`

Current Python dry-run zones are only `B2:J8` and `B12:J18`.

Protected areas include:

- therapist header cells;
- time cells in column A;
- the existing `ΑΝΑΝΕΩΣΗ` form-control/button area to the right of the grid;
- any cell outside the two patient grids.

The existing legacy refresh button is still bound to `RefreshAll`.

The resolver does not hard-code therapist names to columns. It reads the two
header rows at runtime and resolves a provider + timeslot to the current cell.
This also allows an active student/provider column to participate in the same
cell mapping.

## REPLACEMENTS

Confirmed used dimension in the baseline snapshot: `A1:V33`.

This sheet is a legacy dynamic canvas. Candidate choices are represented both
as cell text and as DrawingML shape-buttons currently bound to
`ApplyReplacementChoiceV4`.

Because the layout expands horizontally depending on absent therapists and
patients, individual candidate button coordinates are **not** treated as a
stable contract. The current dry-run cell zone is `A1:V5000`; real button
creation/deletion remains disabled until Excel-host automation is introduced.

## CONFLICT_LOG

- title: `A1` = `Συγκρούσεις / Υπερβάσεις`
- output zone: `A2:A5000`

`A1` is protected.

## REPLACEMENT_LOG

Confirmed headers: `A1:H1`:

`Timestamp | Ασθενής | Αρχικός Θεραπευτής | Αρχική Ώρα | Νεος Θεραπευτής | Νέα Ώρα | Status | Note`

This remains a legacy operational sheet and is still protected from Python
write-back.

## THERAPIST_ATTENDANCE

Confirmed headers: `A1:F1`:

`Θεραπευτής | Παρουσία (Ν/Ο) | Slots φορτίου | Καταχωρήσεις | Μολυσματικοί | Σύγκρουση`

This remains legacy operational input/output while the daily absence workflow
is being redesigned. Python does not write to it yet.

## Layout drift gate

Before any future real write-back, `audit_layout()` must pass. It checks:

- confirmed sheet names;
- `MASTER_SCHEDULE` headers;
- both `THERAPIST_DAILY` provider blocks and time rows;
- `REPLACEMENT_LOG` and `THERAPIST_ATTENDANCE` headers;
- `CONFLICT_LOG` title;
- presence of the existing `RefreshAll` and `ApplyReplacementChoiceV4` macro
  bindings as warnings if they change.

A sheet-level permission is no longer enough. Every planned cell change must
also fall inside a confirmed cell-level write zone.
