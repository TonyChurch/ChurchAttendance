# Project Rules — ChurchAttendance

1. **Never modify production code without commit**
   - All changes to `src/` or `ChurchAttendance.xlsm` must go through Git.
   - No direct edits on the production workbook without a corresponding commit.

2. **One feature = one commit**
   - Keep commits focused. A single commit should represent one logical change.
   - If a feature spans multiple unrelated changes, split them.

3. **Update documentation after every feature**
   - Adjust `docs/ROADMAP.md` to reflect new status.
   - Update `docs/VBA_INVENTORY.md` when modules are added or removed.
   - Update `docs/SOURCE_AUDIT.md` when code structure changes.

4. **Keep backward compatibility**
   - Do not break existing worksheet layouts or named ranges without migration.
   - Preserve existing procedure signatures unless a breaking change is explicitly planned.

5. **Preserve workbook functionality**
   - `ChurchAttendance.xlsm` must remain openable and runnable after every commit.
   - Do not introduce code that prevents the workbook from loading in Excel.

6. **Test before push**
   - Run the workbook and verify the changed feature manually.
   - Verify that unrelated features are not affected.
   - Only push after the local test passes.
