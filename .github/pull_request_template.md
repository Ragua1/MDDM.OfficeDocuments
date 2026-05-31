# Pull Request Checklist

## Summary

- What changed?
- Why is the change needed?

## Scope Classification

- [ ] Core feature or core maintenance
- [ ] Advanced feature or advanced maintenance

Explain briefly why this belongs in the selected scope.

## Product Fit

- [ ] The change targets XML Office formats only (`.xlsx` / `.docx`)
- [ ] The change does not add legacy `.xls` / `.doc` scope
- [ ] The change supports simple document creation, data reading, or efficient work with data

## Implementation Notes

- [ ] `DocumentFormat.OpenXml` is used as the default base
- [ ] Or: a different targeted mechanism/library is used for a specific part because it is simpler, safer, or more efficient

If not using only `DocumentFormat.OpenXml`, explain why the chosen approach is better and why it still fits the XML-only minimal-core direction.

## API Review

- [ ] No unnecessary OpenXml detail leaks into the public API
- [ ] Public API remains small and task-oriented
- [ ] If API surface grows, the reason is explicitly justified

## Dependencies

- [ ] No new dependency added
- [ ] New dependency added and justified below

If a new dependency is introduced, explain:

- why it is needed
- why the cost is acceptable
- whether the feature should instead live in advanced scope or a separate library

## Validation

- [ ] Focused tests added or updated
- [ ] Relevant documentation updated
- [ ] README updated when public behavior changed

## Risks

- Main technical risks:
- Main compatibility risks:
- Follow-up work, if any:
