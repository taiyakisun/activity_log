# Palette drop scroll fix

- [x] Inspect drag/create rendering and scroll preservation.
- [x] Preserve the timeline viewport scroll when a palette item is dropped.
- [x] Verify the change with static checks and a local browser run if practical.
- [x] Stabilize the selected activity panel height when activity selection or summary changes.
- [x] Verify both palette drop and Delete removal keep the timeline viewport stable.
- [x] Preserve the pages scroll after the full controls update during Delete removal.

## Review

- `node --check app.js` passed.
- Browser verification passed: after three palette drops at `scrollTop=1400`, activity count increased to 3 and scroll stayed at 1400 after each drop.
- Follow-up verification passed: with the target day selected, three palette drops changed the day summary from 1 to 4 rows and Delete reduced it to 3 rows, while the editor panel height stayed at 292px and `pages-container.scrollTop` stayed at 1400.
- Viewport check passed at 1280x720: the timeline area remains visible with the stabilized selected activity panel height.
- Delete verification passed after final restoration fix: before Delete, immediately after Delete, and two animation frames later, `pages-container.scrollTop` stayed at 1400 while the selected activity editor reset to the unselected text.
