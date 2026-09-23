# Provider Policy / Overrides / Leadership

Adds a policy layer above the normal scheduler rules.

Defaults: therapist 6 distinct active timeslots/day, student 5.

Temporary overrides do not rewrite the base patient programme:
- exclude provider from replacement suggestions
- temporary custom maximum load
- reserve/block a specific time
- explicitly open a replacement time for the active Manager/Acting Manager

The active Manager has zero base patient load and is not offered automatically for replacements. An Acting Manager temporarily becomes the active leader. The active leader can accept a replacement only in a time explicitly opened for that purpose.

Patient cross-specialty availability is a separate gate and will be connected next. Both provider policy and patient availability must pass before a candidate is shown.
