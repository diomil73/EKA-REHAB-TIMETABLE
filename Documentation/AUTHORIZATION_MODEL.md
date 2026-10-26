# Authorization and Login Model

Status: agreed architecture direction, implementation pending.

## Core principle

Access is not determined only by which worksheet or button is visible.

Every protected action must be authorized at execution time using:

`authenticated user -> role -> specialty scope -> capability -> action authorization`

UI visibility follows the same authorization result, but hiding a button is never treated as sufficient security.

## Login

The application must request login at startup before the normal operational UI is exposed.

- Usernames are created/assigned by the administrator.
- Each user sets their own password.
- Passwords must never be stored or displayed in readable form.
- Passwords are stored only as secure salted password hashes.
- Even `SUPER_USER` does not read another user's password.
- `SUPER_USER` may create/disable accounts, change roles/scopes and reset credentials.

## User model

A user record must support at least:

- `UserID`
- `Username`
- password hash / password metadata
- active/inactive state
- role
- optional specialty scope
- administrative flags where required

The model must support role plus specialty scope rather than creating a different hard-coded role for every specialty.

Example:

`DEPARTMENT_HEAD + LOGOTHERAPY`

and

`DEPARTMENT_HEAD + OCCUPATIONAL_THERAPY`

share the same role semantics but operate only inside their own specialty scope.

## Agreed specialty authorization scopes

The authoritative specialty scopes are:

- `PHYSIOTHERAPY`
- `OCCUPATIONAL_THERAPY`
- `SPEECH_THERAPY`
- `PSYCHOLOGY`
- `SPECIAL_PHYSICAL_EDUCATION`

### Physiotherapy ownership

`PHYSIOTHERAPY` is one authorization scope and includes all of the following operational areas:

- standard Physiotherapy;
- Robotics;
- Tilt Table / reclining-table treatment (`Ανακλινόμενο`);
- Pool / hydrotherapy (`Πισίνα`).

Robotics, Tilt Table and Pool are **not separate authorization scopes**. They are Physiotherapy-owned treatment areas and inherit Physiotherapy permissions.

Therefore, a user authorized for `PHYSIOTHERAPY` may act on those Physiotherapy sub-areas according to the capabilities granted by their role. A user without `PHYSIOTHERAPY` scope does not gain access to those areas merely because they appear in separate MASTER columns.

### Non-Physiotherapy aggregate scope

`ALL_NON_PHYSIOTHERAPY` is a derived scope used by `PARAMEDICAL_HEAD` and means exactly:

- `OCCUPATIONAL_THERAPY`
- `SPEECH_THERAPY`
- `PSYCHOLOGY`
- `SPECIAL_PHYSICAL_EDUCATION`

It explicitly excludes `PHYSIOTHERAPY` and all Physiotherapy-owned sub-areas.

## Agreed operational roles

### 1. SUPER_USER / ADMIN

Intended for the Physiotherapy Head / system administrator.

Access:

- unrestricted application access;
- all specialties;
- all patient therapy scheduling/editing functions;
- all administrative and operational screens;
- all registries and settings;
- therapist workload and management information;
- daily change/restore functions;
- user and permission administration;
- ability to intervene in all permitted operational data.

`SUPER_USER` is an authorization superset of all other application roles, except that passwords themselves remain non-readable by design.

### 2. PARAMEDICAL_HEAD

Head of paramedical services.

Access:

- may intervene in patient therapy scheduling for all specialties **except Physiotherapy**;
- may change the patient names/entries and therapy times that belong to those specialties, according to the operational editing rules;
- has no authority to alter Physiotherapy scheduling solely by virtue of this role;
- administrative/system privileges are not implied.

Specialty scope:

`ALL_NON_PHYSIOTHERAPY`

This expands to Occupational Therapy, Speech Therapy, Psychology and Special Physical Education only.

### 3. DEPARTMENT_HEAD

Heads/responsible persons of the remaining specialties.

Access:

- same type of patient-therapy intervention as `PARAMEDICAL_HEAD`;
- restricted strictly to the user's own specialty;
- may change the patient names/entries and therapy times belonging to that specialty;
- may not modify other specialties or Physiotherapy unless separately granted another role/scope.

Specialty scope is mandatory for this role and is one of:

- `OCCUPATIONAL_THERAPY`
- `SPEECH_THERAPY`
- `PSYCHOLOGY`
- `SPECIAL_PHYSICAL_EDUCATION`

### 4. MEDICAL_VIEWER

Includes Director of Medical Service, Coordinating Directors and other physicians.

Access:

- read-only access to `MASTER_SCHEDULE`;
- link/action to the patient's electronic record;
- link/action to the patient progress view (the view using green/grey/red face indicators);
- no scheduling/editing authority is implied by this role.

### 5. MASTER_VIEWER

Includes nursing staff and patient transport/orderly staff.

Access:

- read-only access to `MASTER_SCHEDULE` only;
- no therapist workload;
- no management metrics;
- no patient scheduling/edit actions;
- no registries/settings.

## Therapist workload privacy rule

Therapist workload is not a general MASTER feature.

- workload badges/metrics/button must be removed from the normal `MASTER_SCHEDULE` surface;
- seeing MASTER does not imply permission to see therapist workload;
- workload belongs to a management/supervisory surface;
- authorization for workload must be explicit (for example a future `VIEW_THERAPIST_WORKLOAD` capability);
- the final supervisor-entry location can be decided separately from this contract.

Until the role-aware UI is implemented, the safe default is not to expose workload from MASTER.

## Capability direction

Implementation should use capabilities instead of scattering role-name comparisons through VBA/Python.

Initial capability vocabulary should include at least:

- `VIEW_MASTER`
- `VIEW_PATIENT_ELECTRONIC_RECORD_LINK`
- `VIEW_PATIENT_PROGRESS`
- `EDIT_PATIENT_THERAPY`
- `EDIT_PHYSIOTHERAPY_THERAPY`
- `EDIT_SPECIALTY_THERAPY`
- `VIEW_THERAPIST_WORKLOAD`
- `RESTORE_DAILY_CHANGE`
- `MANAGE_REGISTRIES`
- `MANAGE_SETTINGS`
- `MANAGE_USERS`
- `ADMIN_ALL`

Capabilities that modify specialty-owned data must also evaluate the user's specialty scope.

## Specialty ownership rule

Therapy cells/actions must have an authoritative specialty owner.

Authorization to edit a patient therapy entry is therefore evaluated as:

1. identify the therapy/schedule specialty;
2. normalize any Physiotherapy sub-area (Robotics, Tilt Table, Pool) to `PHYSIOTHERAPY` ownership;
3. check whether the logged-in user's role grants editing capability;
4. check whether the user's specialty scope includes that specialty;
5. only then expose/execute the edit.

This prevents a department head from editing another department merely by reaching the same worksheet or macro.

## UI implications

The application shell must eventually be role-aware after login.

Examples:

- `MASTER_VIEWER`: MASTER only.
- `MEDICAL_VIEWER`: MASTER plus electronic-record and patient-progress actions.
- `DEPARTMENT_HEAD`: operational editing controls only for owned specialty.
- `PARAMEDICAL_HEAD`: operational editing controls for all non-Physiotherapy specialties.
- `SUPER_USER`: complete application/navigation surface.

The same action must still perform an authorization check even if the UI normally hides it from unauthorized users.

## Audit requirement

When authentication is implemented, state-changing actions should record at least:

- user identity;
- timestamp;
- action type;
- affected patient/entity;
- affected specialty;
- before/after or event reference where applicable.

This is especially important for schedule edits, daily restores, registry changes and administrative interventions.

## Open implementation decisions

The following are intentionally not fixed yet:

- exact login-screen visual design;
- password reset workflow;
- where user administration appears in the UI;
- final supervisor-management navigation surface;
- session timeout / lock policy;
- whether some users may hold multiple roles/scopes simultaneously.

These decisions should be made before implementation of the authentication layer, but they do not block removal of therapist workload from MASTER now.
