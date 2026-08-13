# Requirements

## R1 — Quote Hyperlink Creation

The program shall identify quoted text in PowerPoint text boxes and create hyperlinks based on bold text contained within each quotation.

1. The complete text enclosed by quotation marks shall be the hyperlink display range.
2. Bold text within the quotation shall define the search text used to construct the hyperlink.
3. The hyperlink shall be applied to the entire quotation, including the opening and closing quotation marks.
4. Bold text shall remain bold after processing.
5. All other existing text and formatting shall be preserved.
6. Underlining has no special semantic meaning.
7. If a quotation contains no bold text, no hyperlink shall be created.
8. Text outside the quotation shall not be modified.
9. Reprocessing an already correctly processed presentation shall not create duplicate hyperlinks or otherwise alter it.

## Acceptance Principle

The written requirements define the general behavior.

Approved example input/output PowerPoint files define the expected application of these requirements and serve as acceptance-test cases.

Where a written requirement is ambiguous, the approved example case shall clarify the intended behavior.
