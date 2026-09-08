# Template Row Manual

The macros never hard-code a column layout. Each output sheet reads one row of the `template` sheet and builds
itself from what it finds there — which fields appear, in what order, and which of them get totalled. This is how to
write that row.

```
      A       B       C       D       E        F         G          H        I         J        K          L      M         N
   ┌───────┬───────┬───────┬───────┬───────┬──────────┬──────────┬─────────┬────────┬────────┬───────────┬───────┬────────┬─────────┐
 9 │  NO   │ ZONE  │ BLOK  │ LEVL  │ TYPE  │ BEDTYPE  │ GIFA Σ   │ minAREA │ BEDS Σ │ PERS Σ │ DUAL Σ%   │ PAS Σ │ minPAS │ min10 % │
   └───────┴───────┴───────┴───────┴───────┴──────────┴──────────┴─────────┴────────┴────────┴───────────┴───────┴────────┴─────────┘
```

Row 9 of the `template` sheet, as the Long schedule reads it. Every name is a column; the markers say what the
summary rows do with it.

---

## Which row you are writing

Each output sheet has its own row. They work identically — the only difference is which macro reads them and which
printed header block sits above them.

| Row | Defines | Run by | Printed block |
| --- | --- | --- | --- |
| **9** | Long schedule — one row per unit | `GenerateUnitSchedule` | rows 1–8 |
| **18** | Short schedule — totals only, no unit rows | `GenerateUnitShort` | rows 10–17 |
| **29** | Unit types block — one row per unique type | `UnitTypes` | rows 20–28 |

The mapping row itself is **read, never copied**. Nothing you write in it reaches the printed sheet, so it can carry
markers and abbreviations that would look wrong on a drawing issue. The visible column headings come from the printed
block above it — if you add a column here, add its heading there too.

The Short schedule reads row 9 as well: it builds a working copy of the Long schedule first, then condenses it into
the layout named in row 18. A field can only reach row 18 if it is also in row 9.

## How a cell becomes a column

- **The name is the link to the drawing.** It must match the attribute tag exported from AutoCAD, as it appears in
  the header line of the `.txt`. Matching ignores case and surrounding spaces, so `gifa`, `GIFA` and `GIFA ` are the
  same field.
- **Position is the layout.** A field lands in whichever column of the output sheet its name occupies here. To
  reorder the schedule, reorder these cells — nothing else needs changing.
- **Omitting a name omits the column.** A tag that exists in the export but is not named here is simply not brought
  across.
- **An unrecognised name is harmless.** If nothing in the export matches and no macro fills it, you get an empty
  column under that heading.
- **Blanks are fine, duplicates are not.** Empty cells are skipped, so you can leave a gap. If the same name appears
  twice the leftmost wins and the other column stays empty.
- **Only columns A to CV are scanned** — the first 100. Anything beyond that is invisible to the macros.

> **Leave TYPE alone.** The `TYPE` column is imported as text on purpose, so a unit type like `1E4` stays `1E4`
> instead of being read as scientific notation and stored as 10000. Don't reformat that column to General or Number
> on the output sheet.

## Sigma and percent

A trailing marker on the cell tells the summary rows what to do with that column. The marker is stripped before the
name is matched, so `GIFA Σ` is still the field `GIFA`.

| Marker | Effect |
| --- | --- |
| **Σ** | Adds a `=SUM()` for the column on every level, block and zone summary row, and on the scheme total. |
| **%** | Adds a percentage-of-units row directly beneath each summary row, formatted `0%`. |

| Cell contains | Column name | Behaviour |
| --- | --- | --- |
| `GIFA` | `GIFA` | plain column, no summary figure |
| `GIFA Σ` | `GIFA` | totalled |
| `DUAL %` | `DUAL` | percentage row |
| `DUAL Σ%` | `DUAL` | both — either order works |
| `%` | `%` | a column literally named `%`, not a marker |

That last line is deliberate: the types block has a column whose name really is `%`. A cell that is *nothing but*
markers is always treated as a name.

### Typing the sigma

Insert › Symbol, or put `=UNICHAR(931)` in a spare cell and paste the result as a value. Both the Greek sigma
(Σ U+03A3) and the maths summation sign (∑ U+2211) are accepted, as is a lowercase σ, so it does not matter which one
you land on.

### The two fixed behaviours

- `NO` is always the unit count. It gets a `=COUNTA()` rather than a sum, whether or not you mark it — that is what
  produces the number of units per level and block.
- The bedroom tally columns (`1 BED`, `2 BED`…) are generated at run time from the bed counts actually found, so
  there is no cell to mark. They are always totalled and always given a percentage.

> **Migrating an old template.** If a mapping row carries no markers at all, the macro falls back to the list it has
> always used — `GIFA`, `minAREA`, `BEDS`, `PERS`, `DUAL`, `minPAS`, `PAS`, `minCAS`, `min10` summed, and `min10`
> plus `DUAL` as percentages.
>
> The moment you mark *one* cell, the fallback switches off and the row is taken at its word. Mark all of them in one
> pass, not one at a time.

## Field reference

### Read from the drawing

These have to exist as attribute tags on the unit stamp block, or the column comes through empty.

| Name | Holds | Used for |
| --- | --- | --- |
| `NO` | Unit number | unit count; sort key; `XX` drops the row |
| `ZONE` | Zone code | outermost grouping and summary break |
| `BLOK` | Block code | middle grouping |
| `LEVL` | Level code | innermost grouping |
| `TYPE` | Unit type reference | the key the types block groups on |
| `BEDTYPE` | Apartment / Duplex / House | picks which lookup table applies |
| `BEDS` | Bedroom count | lookup key; tally columns |
| `PERS` | Person count | lookup key; tally columns |
| `GIFA` | Floor area | checked against `minAREA` |
| `PAS` | Private amenity space | checked against `minPAS` |
| `DUAL` | Dual aspect, 1 or 0 | counted and shown as a percentage |
| `BED1`…`BED5` | Individual bedroom areas | added up into `AGBED` |

Any other tag in the export can be listed too — it is carried across and printed like the rest, it just has no
special behaviour. `NO` is also matched as `NO.`, `NUM`, `UNIT NO` or `UNIT NO.`, so an older export still lines up.

### Filled in by the macro

These need no attribute in the drawing. Name one in the mapping row and it is calculated; leave it out and the
calculation is skipped.

| Name | Filled with |
| --- | --- |
| `minAREA` `minPAS` `minCAS` | the minimum for this unit, from the lookup table |
| `minAGBED` `minLVNG` `minSTOR` `minMAIN` | further minima, same lookup, same row |
| `minBED1`…`minBED5` | per-bedroom minima, same lookup |
| `AGBED` | sum of `BED1` to `BED5` |
| `min10` | 1 when `GIFA` exceeds `minAREA` by more than 10%, otherwise 0 |
| `MIX` | Short schedule only — where the bed/type tally block is inserted, widening as needed |
| `TMIX` | Short schedule only — optional total-mix block, titled "Total Mix" |
| `%` | Types block only — that type's share of all units |

A `min…` column only fills if the same name exists in the lookup table's own header. The pairing is by name, in both
directions.

## Lookup tables

The minima and the row colours come from three small tables further down the template sheet. Each is found by a cell
in **column A** reading exactly `Apartment Lookup`, `House Lookup` or `Duplex Lookup`; that cell's row is the table's
header row and the ten rows beneath it are the table.

- Which table applies comes from `BEDTYPE` containing *apartment*/*apt*, *house*, or *duplex*/*dup*.
- The row is found by the key `<beds>b <pers>p` — a 2 bed 4 person unit looks for `2b 4p`. Write the keys in that
  exact form.
- A `COLOUR` column tints the whole unit row on the schedule with that cell's fill.
- If `BEDTYPE` matches nothing, or the `2b 4p` key is missing, **the whole unit row turns red**. That red is the
  signal to fix the table or the drawing, not a formatting choice.

## Sorting and dropped rows

Before any of the layout work happens the schedule is sorted `ZONE` › `BLOK` › `LEVL` › `NO`, and summary rows are
broken wherever one of those changes. Leave a grouping field out of the mapping row and you lose both its sort level
and its subtotals.

Any unit whose `ZONE`, `BLOK`, `LEVL` or `NO` reads `XX` is deleted before the schedule is built. That is the
intended way to keep a stamp in the drawing but out of the schedule — cores, plant, unnumbered shells.

## Other cells that matter

| Cell | Holds |
| --- | --- |
| `AA5` | Full path to the exported `.txt` or `.csv`. Surrounding quotes are stripped, so a path pasted from Explorer works as-is. Every macro reads this one cell. |
| `X3` | Pattern used to reduce a `TYPE` value to the key the types block groups on. Empty means `.*` — the whole value. |
| `X2` | Read by the Long schedule but currently unused by it. Changing it has no effect today. |

## Before you run it

- [ ] Every name in the row matches a header in the export, or is one of the macro-filled names above.
- [ ] The printed header block above the mapping row has a heading over each column you named.
- [ ] Markers are on the mapping row, not on the printed heading.
- [ ] If you marked one cell, you marked all of them — the fallback list is off now.
- [ ] The lookup tables have a `<beds>b <pers>p` row for every combination in the scheme.
- [ ] `AA5` points at the export you just wrote out of AutoCAD.

---

Column layouts are read by `BuildHeaderMap`; the markers by `BuildSumColumns` and `BuildPercentColumns`, all in
`Common Utilities.bas`. A condensed version of these notes can be stamped into the workbook itself as cell notes —
run `WriteTemplateMemo` from `Template Memo.bas`.
