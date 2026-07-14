# Member Management Analysis — ChurchAttendance

**Feature:** FEATURE-003
**Date:** 2026-07-14
**Author:** Analysis task (read-only inspection)
**Scope:** Member (Servant) management workflow, registration, Member ID, and
attendance lookup.
**Constraint honored:** `ChurchAttendance.xlsm` and all VBA were inspected
**read-only**; nothing was modified.

---

## 1. How the Analysis Was Performed

The `.xlsm` is a binary. For analysis only (no modification), the workbook was
opened read-only via Excel COM and every VBA component was exported to a temp
folder. The findings below reflect the **actual** workbook contents, cross-checked
with `docs/WORKBOOK_ANALYSIS.md` (worksheet schema).

> Note: the repository `src/` folder contains `Members.bas`, `Attendance.bas`,
> `Reports.bas`, and `clsMember.cls`. These are **placeholders** — they do **not**
> exist inside the workbook's VBA project (confirmed via export). The live VBA
> project contains only: `ThisWorkbook`, `Worksheet____1..5`, `frmAttendance`,
> and `modMain`.

### Actual VBA Project Inventory

| Component | Role | Member-management relevance |
|-----------|------|-----------------------------|
| `modMain` (UserForm) | Main menu | Has a "إدارة الخدام" (Manage Servants) button |
| `frmAttendance` | Weekly attendance screen | Loads servant list from the `الخدام` sheet |
| `ThisWorkbook` | Workbook module | No event handlers |
| `Worksheet____1` (`الخدام`) | Servants data sheet | **Primary data store** |
| `Worksheet____2` (`الحضور`) | Attendance sheet | Headers only, not yet used |
| `Worksheet____3` (`الاحصائيات`) | Statistics sheet | Headers only, not yet used |
| `Worksheet____4` (`الاعدادات`) | Settings sheet | Headers only, not yet used |
| `Worksheet____5` (`لوحة التحكم`) | Dashboard sheet | Title cell only |

---

## 2. Current Behavior

### 2.1 Data Source — the `الخدام` (Servants) sheet

Members are stored as plain worksheet rows. Columns:

| Col | Header (AR) | Header (EN) | Observed |
|-----|-------------|-------------|----------|
| A | الكود | Code / Member ID | Sequential integers `1..6` (manually entered) |
| B | اسم الخادم | Servant Name | Filled (6 names) |
| C | المرحلة | Stage | Empty |
| D | الخدمة | Service | Empty |
| E | الهاتف | Phone | Empty |
| F | الحالة | Status (active/stopped) | Empty |
| G | تاريخ البداية | Start Date | Empty |
| H | ملاحظات | Notes | Empty |

There are 6 servant rows. Most attribute columns (C–H) are empty, including
**Status**, so there is currently no concept of an active vs. stopped member.

### 2.2 Registration

**No registration form exists.** Member data is entered by directly typing into
the `الخدام` worksheet cells. There is no add/edit/delete routine in VBA.

### 2.3 Member ID

The Member ID lives in column A (`الكود`). It is a manually assigned, sequential
integer (`1, 2, 3, 4, 5, 6`). There is **no code** that generates, validates
uniqueness, or auto-increments the ID.

### 2.4 Attendance Lookup

The only member-aware code is `frmAttendance.cmdLoad_Click()`:

```vb
Set ws = ThisWorkbook.Sheets("الخدام")      ' hard-coded sheet name
lstServants.Clear
lstServants.ColumnCount = 4
lstServants.ColumnWidths = "40 pt;180 pt;120 pt;120 pt"
LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
For i = 2 To LastRow
    .AddItem ws.Cells(i, 1).Value            ' code
    .List(.ListCount - 1, 1) = ws.Cells(i, 2).Value   ' name
    .List(.ListCount - 1, 2) = ws.Cells(i, 3).Value   ' stage
    .List(.ListCount - 1, 3) = ws.Cells(i, 4).Value   ' service
Next i
MsgBox "..." & lstServants.ListCount & " ..."   ' (shows loaded count)
MsgBox "تم تحميل " & lstServants.ListCount & " خادم."  ' "Loaded N servants"
```

It reads columns 1–4 (code, name, stage, service) from `الخدام` into a list box.
It does **not** record attendance, does **not** write to the `الحضور` sheet, and
ignores columns E–H (phone, status, start date, notes).

### 2.5 Menu Entry Point

`modMain` is the main menu UserForm. Its "إدارة الخدام" (Manage Servants) button
(`cmdServants_Click`) only shows a `MsgBox "إدارة الخدام"` — there is **no**
servant management screen behind it.

---

## 3. Problems Found

1. **Servant management is not implemented.** `Doc/ROADMAP.md` marks
   "Servants — [x] Existing", but in reality only a data sheet and a
   load-to-listbox routine exist. Add / Edit / Stop / Search are absent.
2. **Manual Member ID.** IDs are hand-typed sequential integers with no
   auto-generation or uniqueness enforcement → risk of duplicates and gaps.
3. **No input form / no registration flow.** Data is typed directly into cells,
   which is error-prone and untraceable.
4. **No validation rules.** No worksheet data validation and no VBA checks
   (required name, unique phone, valid stage/service/status, ID uniqueness).
5. **Hard-coded sheet name** `"الخدام"`. Renaming the sheet breaks `cmdLoad`
   silently.
6. **Hard-coded column indices** (1–4) and list-box column widths. Inserting or
   reordering columns, or reading columns E–H, breaks or silently drops data.
7. **Duplicate/garbled status feedback.** `cmdLoad` fires two consecutive
   `MsgBox` calls (one appears to be a leftover/garbled string).
8. **No search/filter** capability for members.
9. **Status column unused.** There is no soft-delete / "stopped" handling, so
   removing a member would mean deleting history.
10. **Source export mojibake.** Arabic UI strings stored as ANSI in VBA render as
    mojibake in the repo `src/` text export, hurting diffs/version control
    readability (the workbook itself is fine at runtime).
11. **Misleading placeholders.** `src/` contains `Members.bas`, `Attendance.bas`,
    `Reports.bas`, `clsMember.cls` that are not part of the compiled workbook,
    giving a false impression of existing structure.

---

## 4. Risks

| Risk | Impact |
|------|--------|
| Manual entry + no validation | Inconsistent / duplicate / missing member records |
| No ID uniqueness | Collisions corrupt attendance joins and statistics |
| Hard-coded names/columns | Silent breakage on rename/insert; attendance lookup fails |
| Roadmap marked "done" but not implemented | False confidence; missing safety net before changes |
| No soft-delete | Deleting a member destroys history and breaks past attendance |
| Data lives only inside `.xlsm` | No external DB as the roadmap intends; single point of failure |
| Placeholder `src/` modules out of sync with workbook | Maintainers may edit code that isn't actually compiled |
| Lock file present (`~$ChurchAttendance.xlsm`) | Indicates the workbook may be open; must never be committed |

---

## 5. Recommended Improvements

1. **Real Servant Management form (`frmServants`)** with Add / Edit / Stop
   (soft-delete) / Search, opened from `modMain.cmdServants_Click`.
2. **Auto-generated Member ID** (e.g. `Max(code)+1` or a counter in
   `الاعدادات`), with a uniqueness check before save.
3. **Centralized references** — constants or a config block in `الاعدادات` for
   sheet names and column map, replacing hard-coded strings/indices.
4. **Validation rules** — required name, unique phone, and dropdown lists
   (stage / service / status) via worksheet data validation or form combos.
5. **Use `clsMember` properly** — give it properties (Code, Name, Stage, Service,
   Phone, Status, StartDate, Notes) plus `Load` / `Save` / `Validate` methods.
6. **Search / filter** by name, code, or stage; preserve the list-box binding.
7. **Single, clear status feedback** (one message / status label) instead of two
   `MsgBox` calls.
8. **Soft-delete (`الحالة`)** — mark stopped instead of deleting rows, preserving
   attendance history and statistics.
9. **UTF-8-safe source handling** for UI strings to keep repo diffs readable.
10. **Reconcile the roadmap** — mark Servants as "Partial / In progress" until
    functional; either implement or remove the placeholder `src/` modules so the
    repo reflects the true compiled project.
11. **Never commit the workbook lock file** (`~$ChurchAttendance.xlsm`).

---

## 6. Implementation Order (safe, incremental)

Each step should be its own feature branch + commit per `docs/PROJECT_RULES.md`.

1. **Freeze & document** — this analysis; keep `.xlsm` untouched. *(done)*
2. **Centralize config** — sheet-name + column-map constants / `الاعدادات`
   config; remove hard-coded `"الخدام"` and `(1–4)` indices.
3. **Implement `clsMember`** — properties + `Load`/`Save`/`Validate`. (No UI yet.)
4. **Build `frmServants`** — Add / Edit / Search over the `الخدام` sheet.
5. **Wire the menu** — `modMain.cmdServants_Click` opens `frmServants`.
6. **Member ID auto-generation** + uniqueness check on save.
7. **Worksheet data validation** — dropdowns for stage / service / status.
8. **Soft-delete ("Stop")** handling + hook into statistics later.
9. **UX cleanup** — single status message; remove duplicate `MsgBox`.
10. **Docs reconciliation** — update `ROADMAP.md` status, `VBA_INVENTORY.md`,
    `SOURCE_AUDIT.md`; align or remove placeholder `src/` modules.

---

## 7. File / Component Map (for the improvement work)

| Concern | Current location | Proposed owner |
|---------|-----------------|----------------|
| Member data | `Worksheet____1` (`الخدام`) cols A–H | Keep sheet; add validation |
| Load servant list | `frmAttendance.cmdLoad_Click` | Refactor to use `clsMember` |
| Manage servants entry | `modMain.cmdServants_Click` | Open `frmServants` |
| Member entity logic | *(none)* / placeholder `clsMember.cls` | Implement `clsMember` |
| Registration form | *(missing)* | Add `frmServants` |
| Member ID | Column A, manual | Auto-gen in `clsMember`/config |
| Settings/config | `Worksheet____4` (`الاعدادات`) | Store sheet/column config + ID counter |
