# loading-dock-signin
Digital visitor sign-in system for a loading dock, built in Excel and VBA

# Visitor Sign-In System

A digital replacement for the pen-and-paper visitor log at a loading dock,
built in Excel with VBA. Captures visitor information, stamps date and
check-out time automatically, maintains a contact list of returning
visitors, and flags duplicate phone numbers.

## Background

The loading dock previously used a paper sign-in sheet. I was asked to
move it to Excel, but a direct one-to-one digital version was noticeably
slower than writing by hand because every entry required clicking between
cells, formatting dates manually, and there was no way to recognize
returning visitors. This project rebuilt the workflow around the actual
goal of tracking who came through, and not just digitizing the form.

## Features

- Automatic date stamp when a visitor name is entered
- Row highlighting for active (checked-in) visitors
- One-click check-out with automatic timestamp
- Returning-visitor checkbox that adds the person to a persistent contact list
- Duplicate phone number detection with visual flagging
- Error handling that prevents Excel events from getting stuck disabled

## How it works

The workbook has two sheets: a daily sign-in sheet and a `ContactList`
sheet backed by a structured table (`tblContactList`). A
`Worksheet_Change` event handler watches columns B (name), F (new
visitor checkbox), and H (check-out checkbox) and dispatches to the
right behavior based on which column was edited.

## Files

- `LoadingDockSign.xlsm` — the workbook
- `module1.bas` — exported VBA source for the change handler

## Architecture

1. **Event Listener** - Monitors the sheet for specific input (Worksheet_Change) and pauses standard Excel events to prevent recursive loops.
2. **Routing & Validation** - Identifies input context and routes logic accordingly.
3. **Database Query** - In the event of a new user, queries the ContactList table to flag duplicates.
4. **State Execution** - Updates UI and appends timestamps automatically.

## What I learned

This was my first VBA project and my first time writing event-driven
code. The biggest lesson was that digitizing a process is not the same
as improving it because the original Excel version I built was slower than
paper because I'd copied the form without thinking about the workflow.
Rewriting it around what the user actually needed to do (sign in fast,
check out fast, recognize repeat visitors) was where the real work was.

I also learned why `Application.EnableEvents = False` and
a `CleanUp` error handler matter: an uncaught error in an event handler
can leave events disabled and make the whole sheet stop responding,
with no obvious cause.
