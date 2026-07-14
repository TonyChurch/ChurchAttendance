# Contributing — ChurchAttendance

## Branch Strategy

- `main` — protected production branch. Do not push directly.
- `feature/*` — one branch per feature (e.g. `feature/events`, `feature/attendance`).
- `bugfix/*` — one branch per bugfix (e.g. `bugfix/sort-order`).
- `release/*` — release preparation branches (e.g. `release/v1.1`).

Workflow:
1. Create a branch from `main`.
2. Implement the change.
3. Open a Pull Request back to `main`.
4. Merge after review.

## Commit Message Convention

Format:
```
<type>: <short description>
```

Types:
- `feat` — new feature
- `fix` — bug fix
- `docs` — documentation only
- `refactor` — code change that neither fixes a bug nor adds a feature
- `chore` — build, tooling, or maintenance

Rules:
- Imperative mood ("Add" not "Added").
- First line <= 72 characters.
- Body optional; use it to explain *why*, not *what*.

## Feature Workflow

1. Create branch: `feature/<feature-name>`
2. Implement one feature only.
3. Update `docs/ROADMAP.md` status.
4. Update relevant inventory/audit docs if code structure changes.
5. Commit: `feat: <feature-name>`
6. Push and open PR to `main`.

## Bugfix Workflow

1. Create branch: `bugfix/<bug-name>`
2. Reproduce and fix the bug.
3. Commit: `fix: <bug-name>`
4. Push and open PR to `main`.

## Release Workflow

1. Create branch: `release/<version>`
2. Bump version in documentation.
3. Final regression check.
4. Merge to `main`.
5. Tag the commit: `v<version>`.
