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

## v14 layout notes

- Every circuit segment is drawn, including multi-room continuation (CAL-chain)
  and NIS implied-tie links that route right-to-left. Earlier versions only
  drew a link when the destination sat strictly to the right of its source end,
  which left cross-room circuits unconnected.
- Links are joined at the exact port-row anchor on both device edges. When the
  two ends differ in height the link is routed orthogonally (L-shaped) instead
  of as a diagonal.
- Page width grows with the number of room sections in the circuit
  (`same-room group + MID_GAP + other-room group`), so wide, multi-ODF circuits
  get a wider page automatically.
