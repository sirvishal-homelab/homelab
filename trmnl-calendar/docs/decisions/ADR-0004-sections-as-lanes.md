# ADR-0004: Sections are weighted horizontal lanes stored in the database
Status: Proposed
Date: 2026-10-20

## Context
The user wants each calendar cell divided into named sections (events, work, lunch, dinner,
and others they define), configurable and persistent.

## Decision
A `sections` table (name, position, weight 1-4, visible). Each event belongs to exactly one
section. Week and two-week layouts draw sections as labelled rows; month draws one line per
section per date cell with a legend. Empty sections can be hidden and heights reflow.

## Consequences
Good: simple model; works in 1-bit; easy to reorder and hide.
Bad: an event cannot be in two sections; month view gets cramped past about four sections;
no per-section styling yet.

## Alternatives considered
Tags/categories with colours (no colour on 1-bit); free-form per-cell text boxes (harder to
structure and render).

## Addendum 2026-10-25
Sections are called "categories" in the user interface. A section may be placed as a `lane` (default)
or as a `banner` strip at the top of the day (for holidays and similar whole-day items). See `03`.
