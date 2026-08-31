# GTFS trip_headsign Apple Manager

A lightweight Python tool that converts Google-ready GTFS feeds into Apple Maps-ready GTFS feeds while keeping the original GTFS structure intact.

The tool is designed for workflows where the same set of GTFS feeds is regularly regenerated for Google Maps and then needs a second, Apple-specific export with a different `trip_headsign` convention.

## What it does

The manager:

- reads all GTFS ZIP files from `Base GTFS`;
- discovers the current `trip_headsign` values across the feeds;
- maintains a persistent Excel mapping for Apple `trip_headsign` values;
- generates a smart default Apple headsign when no manual override is present;
- keeps manual Apple overrides across future feed replacements;
- identifies the most frequent terminal stop for each headsign by looking at the last `stop_sequence` of each trip;
- creates Google Maps links using coordinates only;
- applies Italian vowel/accent normalization to both `trip_headsign` and `stop_name`;
- preserves the original `stop_name` casing unless an Apple override is explicitly provided;
- writes the Apple-ready feeds into `Apple`.

The generator changes only the fields that are intended to be Apple-specific and copies the rest of the GTFS contents unchanged.

## Folder structure

The working directory is intentionally minimal:

```text
GTFS_Apple_HeadSign_Manager/
├── Base GTFS/
│   ├── feed01.zip
│   ├── feed02.zip
│   └── ...
├── Apple/
├── Mappa_Headsign.xlsx
├── headsign_manager.py
└── Avvia_Manager.bat
```

Put the Google-ready GTFS ZIP files in `Base GTFS`.

The generated Apple-ready ZIP files appear in `Apple`.

No additional data folders are required for normal use.

## Workflow

### 1. Replace the GTFS feeds

Whenever a new Google-ready release is available, replace the ZIP files inside:

```text
Base GTFS/
```

The feeds may be renamed or replaced. Existing bindings are not tied to the physical filename or spreadsheet row position.

### 2. Open the Manager

Run:

```text
Avvia_Manager.bat
```

### 3. Update the mapping

Click:

```text
AGGIORNA MAPPATURA
```

The manager scans the current GTFS feeds and updates `Mappa_Headsign.xlsx`.

New headsigns are added automatically.

Previously known headsigns keep their existing Apple mapping.

HeadSigns that disappear from the current feeds are retained as historical/inactive entries rather than being deleted.

### 4. Review the Excel mapping

The main sheet is organized by feed, with visual separation between feeds.

For headSigns:

- `Google trip_headsign` contains the source value.
- `Apple trip_headsign (Override)` is the manual Apple value.
- `Default Apple` contains the automatically generated value.
- the override wins whenever it is filled;
- if the override is empty, the default is used.

For stop information, the workbook also provides terminal-stop information, coordinates and Maps links.

### 5. Close Excel

Before generating the feeds, save and completely close `Mappa_Headsign.xlsx`.

### 6. Generate the Apple feeds

Click:

```text
GENERA APPLE
```

The manager creates the Apple-ready ZIP files in:

```text
Apple/
```

## Persistent bindings

Bindings are intentionally persistent.

They are not based on:

- the ZIP filename;
- the order of the feeds;
- the Excel row number;
- the position of a headsign in the current import.

This means that normal feed replacement does not require rebuilding the mappings from scratch.

A headsign that returns after being absent can recover its previous mapping when its stable key can be resolved again.

## Smart `trip_headsign` formatting

The headsign formatter is deliberately conservative.

The goal is not to rewrite the whole string blindly. It recognizes the patterns that should remain uppercase or structurally unchanged and converts normal uppercase text into readable Title Case.

Typical examples:

```text
VARESE FN/FS Aut., Kennedy
→ Varese FN/FS Aut., Kennedy

REMEDELLO SOPRA ITAS
→ Remedello Sopra ITAS

RIVA DEL GARDA Autostazione
→ Riva del Garda Autostazione

ROE' VOLCIANO M.te Covolo
→ Roè Volciano M.te Covolo

CANTU' CERM. FS, Stazione
→ Cantù Cerm. FS, Stazione
```

The engine preserves patterns such as:

- known acronyms such as `FN`, `FS`, `IIS`, `ISIS`, `ITS`, `CAI`, `ITIS`, `ITC`, `ICS`, etc.;
- road codes such as `SS35`, `SP11`, `SP31/SP79`;
- Roman numerals such as `IV`, `XXIV`, `XXV`;
- initial-style abbreviations such as `A.`, `O.`, `M.`, `C.`, etc.;
- `fr.` as lowercase;
- punctuation and already meaningful mixed-case text.

The formatter is intended to be general rather than dependent on a hard-coded list of individual municipalities.

## Accent normalization

Italian apostrophe-based accented vowels are normalized in both `trip_headsign` and `stop_name`.

Examples:

```text
CANTU' → CANTÙ
Cantu' → Cantù
ROE' → ROÈ
```

For stop names, this normalization does not apply the headsign Title Case engine. The original stop-name casing is otherwise preserved.

## Terminal stop logic

For every `trip_headsign`, the manager determines the terminal stop as follows:

1. group trips by `trip_headsign`;
2. inspect each individual trip;
3. select the row with the highest `stop_sequence`;
4. treat that stop as the actual terminal stop of the trip;
5. count terminal stops across all trips sharing the headsign;
6. select the most frequent terminal stop.

This avoids confusing "most visited stop" with "terminal stop".

## Google Maps links

The `APRI MAPS` links are generated from coordinates only.

Example:

```text
https://www.google.com/maps/search/?api=1&query=45.123456,10.654321
```

The stop name is intentionally not included in the URL.

## Apple export

The Apple generator keeps the GTFS structure intact and applies the Apple-specific changes needed by the manager.

The current workflow is intentionally limited to:

- `trip_headsign`;
- `stop_name` accent normalization and explicit stop-name overrides.

Other GTFS files and fields are copied unchanged.

## Requirements

The application is designed for Windows and uses Python together with the libraries installed by:

```text
install_dependencies.bat
```

Typical dependencies include Excel/COM integration and ZIP/CSV processing support.

## Running from source

```text
1. Put the feeds in Base GTFS/
2. Run install_dependencies.bat (first setup only)
3. Run Avvia_Manager.bat
4. Click AGGIORNA MAPPATURA
5. Edit the Excel mapping when necessary
6. Save and close Excel
7. Click GENERA APPLE
```
