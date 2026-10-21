# Changelog

This project follows semantic versioning (MAJOR.MINOR.PATCH).

## [3.2.0] - 2026-10-21
### Added
- Row order per tab in Settings (`row_order_orders`, `row_order_drafts`, `row_order_abandoned`): `top`, `sorted`, `resort` or `bottom`.
- Sync schedule in Settings (`sync_every`, `sync_hour`): 5 minutes to daily. Changes apply on the next run or via Re-install trigger.
- Menu items: Sort newest first, Remove my trigger.
- A sync that runs out of time continues a minute later from where it stopped.
- Missing Settings rows are added automatically, with dropdowns.
- Warning when a second person would install a duplicate sync trigger.
### Changed
- New rows now go at the top, newest first (was: appended at the bottom).
- Default sync interval is 1 hour (was: 5 minutes).
- Recovered checkouts wait at least two sync intervals before `no match`.
- New permission `userinfo.email`; run `setup` once after upgrading.

## [3.1.0] - 2026-10-17
### Added
- Drafts and Abandoned checkouts tabs; Remarks column on every tab.
- Offline test suite (test_code.js), 15 tests.
### Changed
- SPEC.md aligned with the shipped code.
