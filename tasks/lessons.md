# Lessons

- 2026-06-12: For timeline scroll bugs, measure both `scrollTop` and the scroll container geometry. Preserving `scrollTop` is not enough when panels above the scroll area change height; stabilize the layout source first.
- 2026-06-12: When restoring scroll during render, restore after all sibling UI updates too. A render-time restore before `updateControls()` can still leave a small scroll shift when the selected editor content changes.
