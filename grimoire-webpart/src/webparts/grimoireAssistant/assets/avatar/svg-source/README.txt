Avatar SVG Source Pack
================================

Files:
1) space_ghost_particles.svg
   - Groups: ghost, bg_particles, face_particles, eyes (left_eye/right_eye), pupils, mouth

2) grimoire_mascot_contract.svg
   - Groups: grimoire_mascot, halo, covers, pages, page_stack, page_lines, brows, left_eye, right_eye, mouth

3) anony_mousse.svg
   - Groups: grimoire_mascot, halo, brows, left_eye, right_eye, mouth

4) robot.svg
   - Groups: pixel_ai, brows, left_eye, right_eye, mouth

Renderer usage:
- `SvgAvatar.tsx` consumes these source files directly.
- `grimoire_mascot_contract.svg` is used for `classic` (label: `GriMoire`, default).
- `anony_mousse.svg` is used for `anonyMousse` (label: `AnonyMousse`).
- `robot.svg` is used for `robot` (label: `Robot`).
- Group IDs are used for runtime transforms: `left_eye`, `right_eye`, `brows`, `mouth`, plus accent groups such as `halo` and `cheeks`.
- For SVGs without `left_eye`/`right_eye`, the renderer falls back to animating `eyes` as one group.

All SVGs have transparent backgrounds and are original (not copies of referenced designs).
