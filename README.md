# NetworkDiagram

VBA (Excel host + Visio automation) that generates **Layer 1 network schematics**
from a cable-schedule workbook.

Given a workbook of cable-schedule sheets, the macro draws one Visio page per
source device, laying destination rooms/ODFs out as swim-lane sections and
joining device ports with cable-type-styled links.

## Files

| File                     | VBA component | Purpose |
|--------------------------|---------------|---------|
| `SchematicGenerator.bas` | Standard module | Main logic. Run `CreateSchematic`. |
| `clsCircuit.cls`         | Class module    | One cable/circuit segment. |
| `clsNode.cls`            | Class module    | One device endpoint. |

## Expected workbook layout

Sheets read by name (see the `SHT_*` / `CC_*` constants at the top of
`SchematicGenerator.bas`):

- `Cable Schedule - Fibres`, `Cable Schedule - DAC`, `Cable Schedule - Copper`
  — the circuit rows.
- `DeviceNames` — column A lists device-name stems that mark a row's End-A as a
  *source* device (one Visio page is produced per source device).
- `Front Sheet` — site banner text: `B9` acronym, `C9` site code, `D9` site name.

No site-, customer-, or project-specific values are hard-coded in the scripts —
all site text is read from the `Front Sheet` at run time.

## Usage

1. In the Excel VBE (`Alt+F11`), import all three files
   (File ▸ Import File…), or paste them into a module / two class modules.
2. Add a reference to **Microsoft Visio** is *not* required — Visio is bound
   late via `CreateObject`. Visio must be installed.
3. Run `CreateSchematic`.
4. `TestColumns` is a diagnostic that dumps the parsed columns of the first
   circuit row, useful if your schedule's column positions differ.

## v16 layout notes

- **One column per hop depth.** Every device the same "hop distance" from the
  source shares a single vertical column. So all the ODFs a source reaches in a
  room line up in one column instead of each getting its own — which is what
  previously fanned the links out into diagonals.
- **Endpoints pinned.** The source (A-end) is column 0, pinned to the far left;
  the deepest hop (terminal B-end) is pinned to the far right.
- **Dynamic spacing.** Column spacing fills the page width: one hop-column ⇒ wide,
  centred spacing; more hops ⇒ tighter, evenly spaced. Page width grows with the
  number of hop columns.
- **Rooms as tiled zones.** Each room is a background band spanning the hop
  columns of its devices, tiled edge-to-edge with a dashed divider at each
  boundary (Room 1 | Room 2 | …).
- **Clean links.** Because every hop of a chain shares its row Y and columns
  increase left-to-right, links are straight horizontal lines. (`DrawCktLine`
  still falls back to an orthogonal L-shape if two ends ever differ in height.)

Example — a two-room circuit renders like:

```
                Room 1              |            Room 2
core-router-1 >> ODF1 T1 P1 >>>>>>>>|>> ODF2 T1 P1 >>>>> core-router-2
core-router-1 >> ODF1 T1 P1 >>>>>>>>|>>>>>>>>>>>>>>>>>>> core-router-2
```

Both rows share the same ODF1 column and the same core-router-2 column, so the
devices stay aligned and the links stay horizontal.

### Known limitation

A device reached at different hop depths on different circuits is drawn once, at
its deepest column; a link into it on a shorter path is still horizontal but
spans the intermediate (empty) columns. Rooms whose hop ranges interleave can
produce overlapping bands — normal linear circuits (room 1 then room 2) tile
cleanly.
