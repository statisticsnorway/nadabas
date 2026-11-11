# Designing NADABAS for its users

Maintainer review draft, 2026-11-11. These principles guide product changes
throughout the repository. They describe intended behaviour. Verify current
implementation and deployment constraints before claiming a principle is satisfied.

## Start with the work

NADABAS helps statistical staff compile, inspect and maintain national accounts.
Users should understand which data they are working with, what an action changes,
and whether it succeeded. Preserve the visibility of calculations in Excel.

For a new flow, identify the person, their task, their starting context, and a
clear completion condition. Use familiar statistical and Excel terminology.
Reduce repeated entry, unnecessary setup, and decisions the system can resolve
reliably. Keep data quality and control visible as functionality grows.

## People, responsibilities and authority

These responsibilities are a proposed model for design review. Map them to the
actual deployment's permissions before implementation. One person may carry
several responsibilities, each with its own scope.

| Responsibility | Typical work | Scope to make explicit |
|---|---|---|
| Statistical user | Open workbooks, inspect calculations, load/save permitted data, run permitted batches | Database, workbook, period, reservation and write permission |
| NADABAS administrator | Maintain shared settings, workbook definitions, administrator membership and application permissions | Selected NADABAS database and affected resources |
| Database administrator | Manage database access, backup, restore and schema operations | Database engine, database and deployment |
| IT operator | Deploy software, manage runtime prerequisites and update policy | Device, user installation or organizational policy |
| Release maintainer | Review source, build and approve signed releases | Repository, artifacts and signing authority |

Identify the relevant administrator in labels and support instructions. Use
**NADABAS administrator** or **IT support** where the distinction affects the
next action. Windows elevation, repository ownership and a NADABAS administrator
flag each require their own authorization context.

Display effective identity and role where they affect a decision, particularly
when editing shared settings or access. Make the active database and workbook
easy to identify. Invalidate stale role state when the context changes.
Provide an explicit recovery route for administrator lockout.

### Existing behaviour requiring care

In [cmdAdministrators.bas](vba/modules/cmdAdministrators.bas), `TestAdministrator`
sets `isAdministrator = True` when the current database's administrator
collection is empty. Otherwise it checks membership using `get_NTUserName`.
This is a source observation of legacy bootstrap behaviour.

A role redesign must address first-database setup, intentional bootstrap,
missing or failed administrator-list loading, existing installations and
recovery. Specify which trusted actor may establish the first administrator.
Avoid silently broadening access or locking out existing operators.
Runtime exposure and backend enforcement need a dedicated assessment.

## Authorization throughout the action

Base permissions on the actor, action and specific resource. Check privileged
operations through their underlying command path as well as the ribbon.
Unknown authorization state should prevent privileged writes and offer a useful
recovery message. Define any bootstrap exception explicitly.

Use the database or service boundary for enforcement where available. Document
the guarantees provided by each backend and deployment. Shared Access files
require particular attention to file access and the ability to bypass the UI.
Keep appropriate operation logs and protect their contents. These principles
follow least privilege and operation-level authorization guidance.
[OWASP authorization guidance](https://cheatsheetseries.owasp.org/cheatsheets/Authorization_Cheat_Sheet.html)

## Keep everyday work simple

- Give frequent actions clear placement and concise labels.
- Reveal optional tools when the task or content makes them relevant. Keep
  essential settings, language selection and help discoverable.
- Use role-appropriate administration views. Explain unavailable actions when
  the user needs to understand a missing prerequisite or permission.
- Keep ribbon groups stable and readable. Check longer translations and optional
  features so additional functionality does not crowd everyday actions.
- Preserve workbook, selection and task context across dialogs where practical.
  Perform background work without repeatedly taking focus.

For Exchange tools, consider existing links, capability availability and user
permission together. Define how an authorized user creates the first link when
tools depend on existing content.

## Consequences and recovery

Show the relevant target and effect before destructive or broad changes,
including delete, overwrite, restore, schema changes and permission edits.
Choose a useful preview, validation, confirmation or reversible operation
according to risk. Routine safe actions should stay quick.

Show useful progress for long operations and whether cancellation remains
possible. Report partial completion accurately. Before retrying a write,
determine whether it already committed. Keep user edits and reservation conflicts
visible so concurrent work can be resolved deliberately.

Messages should explain what happened, the affected task and the next available
action. A verification failure should offer useful details or support information.
Put technical traces behind an accessible details action and avoid revealing
credentials or sensitive database content.

Use review and recovery patterns to prevent costly data mistakes. W3C guidance
informs these principles for Excel and desktop UI.
[Error prevention guidance](https://www.w3.org/WAI/WCAG22/Understanding/error-prevention-legal-financial-data.html)

## Language and accessibility

Offer language selection during installation or first use and an accessible way
to change it later. Preserve established preferences. Use consistent terminology
across VBA, .NET, documentation and support material. Include English, French,
Portuguese and Bahasa Indonesia when validating shared language behaviour.

Support keyboard navigation, visible focus, readable text at larger display
scales, descriptive control labels and status understandable without colour alone.
Check long translations and screen-reader exposure where the technology supports
it. Record limitations and actual accessibility test coverage.

Announce progress and completion without unnecessarily moving keyboard focus.
This is informed by W3C status-message guidance.
[Status messages](https://www.w3.org/WAI/WCAG22/Understanding/status-messages.html)

## Installation and updates

Make installation, activation, language choice and recovery understandable for
the person using the product. Apply IT policy consistently and explain relevant
restrictions with an appropriate support route.

Approved updates may download in the background. Let users choose when to restart
Excel and preserve normal save prompts. Use labels that describe the consequence,
such as **Restart and update** and **Later**. Make cancellation's effect on future
offers clear. Expose component versions through diagnostics when support needs them.

## Review a changed experience

Use the questions relevant to the change:

1. Can the intended user find and complete the task with their actual permissions?
2. Are identity, database, workbook and effects clear where they matter?
3. Does permission enforcement hold when the UI is bypassed or context changes?
4. What happens after interruption, partial completion, concurrent edits or retry?
5. Can the user recover, cancel or obtain useful help without losing work?
6. Does the flow remain understandable with keyboard use, larger text and translation?

Role mappings, bootstrap policy, recovery ownership and accessibility support
must be agreed and tested for the relevant deployment.
