# Patient-centric timetable cell rendering

## Source rule

The source-of-truth remains one independent assignment per patient. A visual pair/group is derived from assignments that share therapist + time.

## Normal derived pair

```text
ΓΚΑΛΜΑΝΗ [Τρ-Πε]
ΠΕΤΙΡΟΠΟΥΛΟΣ [Δε-Τε-Πα]
```

Each patient remains independently movable.

## Same-patient split schedule

The same patient is never rendered twice merely because they own multiple complementary assignments in the same therapist/time slot.

```text
ΜΟΤΣΙΟΣ
Δε-Τε-Πα ΡΟΜΠ | Τρ-Πε ΦΘ
```

The assignments remain separate internally.

## One-day replacement

Original slot:

```text
ΧΡΙΣΤΟΔΟΥΛΟΥ ΟΛ [Τρ-Πε]
ΠΕΤΙΡΟΠΟΥΛΟΣ [Δε-Τε-Πα]   <- italic / muted
→ Φιλιππούσης 08:30 [Πα 18/09]
```

Destination slot, when its base patient is not active on that date:

```text
ΓΚΑΛΜΑΝΗ [Τρ-Πε]
ΠΕΤΙΡΟΠΟΥΛΟΣ [Πα 18/09]
```

The temporary occurrence is date-labelled. It does not silently become a permanent recurring assignment.

## Destination collision

If the destination therapist/time already has a base assignment active on the replacement date, the preview stops instead of stacking two simultaneous patients. The therapist may still be a candidate, but another free timeslot must be selected.
