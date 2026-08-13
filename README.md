# PowerPoint Quote Linker (Python)

Python implementation for adding Google Books/search hyperlinks to quoted text in PowerPoint presentations while preserving formatting and supporting repeatable automated tests.

## Goals

- Keep the Python implementation separate from the VBA implementation.
- Process `.pptx` / `.pptm` files deterministically.
- Preserve existing text formatting and unrelated presentation content.
- Make repeated runs idempotent.
- Verify behavior with automated tests locally and in GitHub Actions.

## Development workflow

1. Add representative PowerPoint fixtures under `tests/fixtures/`.
2. Implement transformation logic under `src/`.
3. Run the same pytest suite locally and in GitHub Actions.
4. Perform a final acceptance check in Microsoft PowerPoint on Windows.
