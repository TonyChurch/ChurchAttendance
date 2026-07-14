# Member Management — Design (FEATURE-004)

**Feature:** FEATURE-004
**Date:** 2026-07-14
**Status:** Design only — no VBA written, `ChurchAttendance.xlsm` untouched.
**Companion doc:** `docs/MEMBER_MANAGEMENT_ANALYSIS.md` (current-state audit).

This document specifies the complete Member (Servant) Management subsystem so
implementation can follow without further design decisions.

---

## 1. Member Fields

Canonical field set for a member record.

| # | Field (EN) | Field (AR) | Type | Max / Format | Required | Notes |
|---|------------|------------|------|--------------|----------|-------|
| 1 | Member ID | الكود | Integer | auto, unique | Yes | Primary key, see §3 |
| 2 | Full Name | الاسم الكامل | Text | 80 chars | Yes | Display name + lookup key |
| 3 | Nick Name | الاسم المتداول | Text | 40 chars | No | Short/calling name |
| 4 | Gender | النوع | Enum | Male / Female | Yes | Dropdown |
| 5 | Birth Date | تاريخ الميلاد | Date | yyyy-mm-dd | No | Enables age stats later |
| 6 | Mobile | المحمول | Text | 11–15 digits | No* | Unique when present |
| 7 | Father Name | اسم الأب | Text | 80 chars | No | |
| 8 | Mother Name | اسم الأم | Text | 80 chars | No | |
| 9 | Address | العنوان | Text | 200 chars | No | |
| 10 | Service Stage | المرحلة | Enum | from Stage list | No | Links to attendance event "Stage" |
| 11 | Servant | الخادم المسؤول | Text/Enum | servant list | No | Responsible servant |
| 12 | Notes | ملاحظات | Text | 255 chars | No | Free text |
| 13 | Status | الحالة | Enum | Active / Stopped | Yes | Soft-delete flag, default Active |

\* Mobile is optional but, if provided, must be unique across active members.

### Field grouping (for forms/reports)

- **Identity:** Member ID, Full Name, Nick Name, Gender, Birth Date
- **Contact / Family:** Mobile, Father Name, Mother Name, Address
- **Church role:** Service Stage, Servant, Status
- **Meta:** Notes

---

## 2. Validation Rules

Applied at two layers: **worksheet data validation** (first line, always on)
and **VBA form validation** (before save).

### 2.1 Worksheet-level (data validation on the `الخدام` sheet)

| Field | Rule |
|-------|------|
| Member ID | Whole number > 0; no blanks; custom formula enforces uniqueness across used range |
| Full Name | Not blank; length ≤ 80 |
| Gender | List: `Male,Female` |
| Birth Date | Date ≤ today; blank allowed |
| Mobile | Optional; if not blank → digits only, length 11–15; custom formula enforces uniqueness among non-blank |
| Service Stage | List sourced from a `Settings` stage list |
| Servant | List sourced from the Full Name column of active members (or a settings list) |
| Status | List: `Active,Stopped`; default `Active` |
| Others | Length caps only (no format rules) |

### 2.2 Form-level (VBA, pre-save)

1. Full Name required and trimmed; reject if empty.
2. Gender required (radio/combobox), must be `Male` or `Female`.
3. Member ID required, positive integer, and **unique** (excluding the row being edited).
4. Mobile: if present, digits only, 11–15 chars, and unique among other members
   (case-insensitive, ignoring spaces/dashes).
5. Birth Date: if present, must parse as a date and not be in the future.
6. Service Stage / Servant / Status: must be one of the allowed list values.
7. On any failure → show the first offending field and focus it; do **not** write the row.
8. On success → write row, then re-apply/refresh data validation to keep the
   sheet self-consistent.

### 2.3 Cross-field / business rules

- A member cannot be set to `Stopped` if they have future dated attendance
  records (warn, allow override with confirmation).
- Editing a Member ID is discouraged; if changed, the attendance join must be
  remapped (see §8).

---

## 3. Member ID Generation Strategy

**Goal:** stable, unique, human-readable, gap-tolerant primary key.

- **Storage of counter:** a single cell in the `الاعدادات` (Settings) sheet,
  e.g. key `LastMemberID` → value `6`. Avoids scanning the whole sheet on every add.
- **On Add (new member):**
  1. Read `LastMemberID` from Settings.
  2. `NewID = LastMemberID + 1`.
  3. Verify `NewID` is not already present in column A of `الخدام`
     (safety against manual edits); if collision, scan forward to next free ID.
  4. Write the row, then update `LastMemberID = NewID`.
- **On Edit:** keep the existing ID; do **not** regenerate.
- **On Stop (soft-delete):** keep the ID; set Status = `Stopped`. IDs are never reused.
- **Uniqueness guarantee:** combined counter + in-sheet uniqueness check +
  worksheet validation formula.
- **Why not `Max+1` scan alone:** the counter is O(1) and survives row deletions;
  the scan is only a fallback safety net.
- **Future-proofing:** ID is an opaque integer; later it can be prefixed
  (e.g. `S-0001`) without breaking attendance joins, since attendance stores the
  same ID value.

---

## 4. Worksheet Layout

### 4.1 `الخدام` (Servants / Members) — data store

Header row = row 1; data starts at row 2. One member per row.

| Col | Header (AR) | Header (EN) | Field | Validation |
|-----|-------------|-------------|-------|-----------|
| A | الكود | Member ID | Member ID | whole >0, unique |
| B | الاسم الكامل | Full Name | Full Name | required, ≤80 |
| C | الاسم المتداول | Nick Name | Nick Name | ≤40 |
| D | النوع | Gender | Gender | list Male/Female |
| E | تاريخ الميلاد | Birth Date | Birth Date | date ≤ today |
| F | المحمول | Mobile | Mobile | digits 11–15, unique if present |
| G | اسم الأب | Father Name | Father Name | ≤80 |
| H | اسم الأم | Mother Name | Mother Name | ≤80 |
| I | العنوان | Address | Address | ≤200 |
| J | المرحلة | Service Stage | Service Stage | list (stages) |
| K | الخادم المسؤول | Servant | Servant | list (servants) |
| L | الملاحظات | Notes | Notes | ≤255 |
| M | الحالة | Status | Status | list Active/Stopped, default Active |

(13 columns — extends the current 8-column layout; existing columns A–H map
directly, new columns I–M appended. See migration note in §8.)

### 4.2 `الاعدادات` (Settings) — additions

| Key (AR) | Key (EN) | Value |
|----------|----------|-------|
| اخر كود خادم | LastMemberID | integer (current max) |
| قائمة المراحل | StageList | comma/newline list for dropdown |
| قائمة الخدام | ServantList | list of servant names for dropdown |

Stage/Servant lists may alternatively live in a hidden `Lists` area; Settings
key approach is preferred for single-source config.

### 4.3 Named ranges (recommended)

- `tblMembers` → `الخدام!$A$1:$M$<lastRow>` (for stable VBA/validation references)
- `rngMemberID`, `rngFullName`, … per-column named ranges, or a single
  `MemberColumns` config map in VBA constants.

This removes the hard-coded `"الخدام"` string and `(1–4)` column indices found
in the current `cmdLoad` (see analysis doc §3.5/§3.6).

---

## 5. UserForm Layout

### 5.1 `frmServants` — main management form

Reached from `modMain.cmdServants_Click` (currently a stub `MsgBox`).

**Layout (top → bottom):**

```
┌──────────────────────────────────────────────────────────────┐
│  إدارة الخدام (Member Management)                [Search box ▾] │
├───────────────────────────────────────┬──────────────────────┤
│  List box: lstMembers                  │  Detail panel:       │
│   columns: ID | Full Name | Stage      │   Member ID  [txtID] │
│   | Status (filtered by search)        │   Full Name  [txtName]*│
│                                        │   Nick Name  [txtNick]│
│                                        │   Gender  ( ) Male    │
│                                        │           ( ) Female  │
│                                        │   Birth Date [dtpBirth]│
│                                        │   Mobile    [txtMobile]│
│                                        │   Father     [txtFather]│
│                                        │   Mother     [txtMother]│
│                                        │   Address    [txtAddress]│
│                                        │   Stage    [cboStage] ▾│
│                                        │   Servant  [cboServant]▾│
│                                        │   Status   [cboStatus] ▾│
│                                        │   Notes    [txtNotes]  │
├───────────────────────────────────────┴──────────────────────┤
│  [Add] [Edit] [Save] [Stop] [Delete?] [Clear]  Status: <msg>  │
└──────────────────────────────────────────────────────────────┘
```

- **Left:** read-only list (`lstMembers`) showing ID, Full Name, Stage, Status;
  double-click or "Edit" loads the selected row into the detail panel.
- **Right:** bound input controls. Required fields marked with `*`.
- **Buttons:**
  - `cmdAdd` — prepare a blank detail panel, generate a preview Member ID.
  - `cmdSave` — validate (§2) then create or update the row.
  - `cmdEdit` — load selected member into the panel (edit mode).
  - `cmdStop` — soft-delete: set Status = `Stopped` (confirm).
  - `cmdClear` — reset the panel.
  - (Hard delete intentionally omitted; use Stop. A separate admin "purge"
    can be added later if needed.)
- **Status label** instead of multiple `MsgBox`es (fixes current double-popup).

### 5.2 Reuse in `frmAttendance`

`frmAttendance.cmdLoad_Click` should be refactored to read from `tblMembers`
(via `clsMember`) instead of hard-coded sheet/columns, and to show all relevant
columns. Keeps attendance lookup consistent with the new schema.

---

## 6. CRUD Operations

All CRUD goes through `clsMember` (properties + `Load`/`Save`/`Validate`) and a
small `Members` module (the placeholder `Members.bas` becomes the real module).
No direct cell writes from forms except via these helpers.

| Op | Trigger | Logic |
|----|---------|-------|
| **Create** | `cmdAdd` → `cmdSave` | Generate ID (§3); validate; append row to `الخدام`; update `LastMemberID`; refresh list. |
| **Read** | Form open / search / attendance load | `clsMember.Load(id)` reads the row by ID; list loads all (or filtered) rows. |
| **Update** | `cmdEdit` → `cmdSave` | Validate; overwrite the row matching the (unchanged) Member ID; refresh list. |
| **Stop (soft Delete)** | `cmdStop` | Set `Status = Stopped`; keep row & ID; refresh list; optionally hide from active views. |
| **(Hard Delete)** | — | Not provided by default; preserves history & attendance integrity. |

**Concurrency / safety:**
- Operate on the workbook in read-only-of-others context; single-user app, but
  always read `LastRow` fresh via `End(xlUp)` and write the exact target row.
- After any write, re-apply worksheet data validation to the touched range.

---

## 7. Search Options

Search box above `lstMembers` filters the list live (on `Change`/`AfterUpdate`).

| Filter | Behavior |
|--------|----------|
| By Member ID | exact / prefix match on column A |
| By Full/Nick Name | case-insensitive substring on B/C |
| By Service Stage | dropdown equals on column J |
| By Status | toggle All / Active / Stopped (column M) |
| By Mobile | substring on column F |
| Combined | AND of all active filters |

Implementation: build a filtered array in `clsMember`/`Members` and rebind
`lstMembers.List`. Keep the underlying sheet as the source of truth; search
never alters data.

---

## 8. Future Compatibility with Attendance

Design choices that keep Attendance (FEATURE-00x) clean:

1. **Stable join key.** Attendance records reference members by **Member ID**
   (column A of `الخدام`), exactly as the current `الحضور` sheet already plans
   (`B$1=الكود`). No name-based joins.
2. **Service Stage alignment.** `Service Stage` (col J) uses the same stage list
   as Attendance events' "Stage" field (from `Settings`), so filtering
   attendance by stage reuses the member stage vocabulary.
3. **Status awareness.** Attendance load should exclude `Stopped` members by
   default (or flag them), preventing marking attendance for inactive servants.
4. **`clsMember` as shared model.** Both `frmServants` and `frmAttendance` use
   `clsMember.Load(id)` — single source of member truth.
5. **Schema migration (non-breaking).** Current `الخدام` has cols A–H; new layout
   appends I–M. Existing rows keep A–H values; new columns default to blank/
   `Active`. No column reordering → existing `cmdLoad` (cols 1–4) still works
   until refactored.
6. **Statistics hook.** `الاحصائيات` can join on Member ID + Stage to compute
   per-member attendance %, reusing the same ID and stage taxonomy.
7. **No ID reuse** (§3) guarantees historical attendance always resolves to the
   correct member even after a Stop.

---

## 9. Implementation Order (hand-off to FEATURE-005+)

1. Add `Settings` keys: `LastMemberID`, `StageList`, `ServantList`.
2. Extend `الخدام` to 13 columns (I–M appended) + named range `tblMembers`.
3. Implement `clsMember` (fields, `Load`/`Save`/`Validate`).
4. Implement real `Members` module (CRUD + search helpers).
5. Build `frmServants` (layout §5.1) and wire `modMain.cmdServants_Click`.
6. Apply worksheet data validation (§2.1).
7. Refactor `frmAttendance.cmdLoad` to use `tblMembers`/`clsMember`.
8. Add Status filtering to attendance load + statistics join.
9. Update `ROADMAP.md` (Servants → In progress/Partial) and inventories.

See `docs/MEMBER_MANAGEMENT_ANALYSIS.md` §6 for the safety-ordered plan that
this design fulfills.
