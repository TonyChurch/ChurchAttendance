# Church Attendance System

## Vision

A professional Church Attendance Management System for servants.

The system must be simple, fast, stable, and capable of generating accurate statistics at any time.

---

# Phase 1

## Servants

- Manage servants
- Add
- Edit
- Stop
- Search

Status:
- [x] Existing

---

## Events

Every attendance belongs to an Event.

Each Event contains:

- Event ID
- Date
- Day
- Event Name
- Stage
- Notes

Examples

- Friday Mass
- Friday Service
- Preparation Meeting
- Visitation
- Saturday Servants Meeting
- Family Meeting
- Conference
- Trip
- Spiritual Day

Status

- [ ] Pending

---

## Attendance

Select Event

↓

Load Servants

↓

Mark Present / Absent

↓

Save

Status

- [ ] Pending

---

## Statistics

For each servant

- Total Events
- Present
- Absent
- Attendance %
- Last Attendance
- Last Absence

Status

- [ ] Pending

---

## Reports

- By Date
- By Event
- By Stage
- By Servant

Status

- [ ] Pending

---

# Database

Servants

Events

Attendance

Statistics

Settings

---

# Git Workflow

One Feature

↓

One Commit

↓

Push

No direct editing on production without commit.
