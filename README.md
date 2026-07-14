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

## v15 layout notes

- **Hop-ordered placement.** Each device is one box, positioned left-to-right by
  its "hop distance" from the source. The A-end (source) is pinned to the far
  left; the terminal B-end is pinned to the far right.
- **Equal room zones.** Rooms divide the page into equal-width zones, dynamically:
  one room ⇒ hops are spread wide and centred; N rooms ⇒ the page is split into
  N zones with a dashed divider between them. Page width grows with the number of
  rooms and the number of hops per room.
- **Clean links.** Every hop of a chain shares the same row Y, so links are
  straight horizontal lines joining device edges. (`DrawCktLine` still falls back
  to an orthogonal L-shape if two ends ever differ in height.)
- Every circuit segment is drawn, including multi-room continuation (CAL-chain)
  and NIS implied-tie links.

Example — a two-room circuit renders like:

```
            Room 1                    |            Room 2
core-router-1 >> ODF1 T1 P1 >>>>>>>>>>|>> ODF2 T1 P1 >>>>> core-router-2
```

### Known limitation

Within a single room, a device that is both a pass-through's incoming and
outgoing end is shown as one box; if a room contains many hops the zone widens
to fit them. Very dense single-room chains can therefore make a page quite wide.
