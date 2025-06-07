# Google Sheets Workout Formatter

A Google Apps Script that reads your raw workout “Plan” sheet and generates a styled, per-block workout table in a separate “Block X” sheet—complete with headers, banners, exercise grouping, and E1RM formulas.

## Features

- **Automatic table generation**  
  Reads your “Plan” sheet and writes a formatted output with:  
    - Client name, week number, workout number  
    - One row per set, with reps, RPE, TrueRPE/TrueWeight, E1RM formula, and comments  
- **Customizable blocks & weeks**  
  Adjust `blockNumber` and `week` in `generateWeek()` to target any block/week.  
- **Styled layouts**  
    - Colored header row  
    - “Block X” and “Week Y” banners  
    - Workout section headers with borders  
    - Exercise-group borders and highlights  
- **E1RM formula injection**  
  Inserts an `IF(...VLOOKUP...)` formula in column K for each set to compute estimated 1RM using your “E1RM” lookup table.

---

## Input Sheet Format

Your **input** sheet must be named exactly:

    Block <blockNumber> Plan

and have these columns (A–K):

| Col | Header      | Description                                                  |
|:---:|:------------|:-------------------------------------------------------------|
| A   | Client name | Name of the trainee                                          |
| B   | Week        | Week number (integer)                                        |
| C   | Workout     | Workout session number within the block                     |
| D   | Exercise    | Exercise name (e.g. “Squat”)                                 |
| E   | Sets        | Number of sets (integer)                                     |
| F   | Reps        | Comma-separated reps per set (e.g. `5,5,5`) or single value |
| G   | RPE/%       | Comma-separated RPE or % per set (e.g. `7,7,7`)              |
| H   | TrueRPE     | *(Optional)* manual RPE override                             |
| I   | TrueWeight  | *(Optional)* manual weight override                          |
| J   | E1RM        | *(Output only)* formula will be inserted here                |
| K   | Comment     | `text;setNumber` (e.g. `10–12 reps;1`)                       |

![Example input sheet](docs/input-sheet.png)

---

## E1RM Lookup Sheet

You need a sheet named exactly:

    E1RM

with two columns, populated in rows A2:B100:

    Repmax   %
    1        1.00
    1.5      0.98
    2        0.96
    2.5      0.94
    3        0.92
    …        …
    49.5     0.38
    50       0.38

The script uses this formula in `estimatedMax()`:

    formulaCell.setFormula(
      '=IF(AND(NOT(ISBLANK($I' + i + ')); NOT(ISBLANK($J' + i + '))), ' +
      '$J' + i + '/VLOOKUP($G' + i + ' + (10 - $I' + i + '), E1RM!$A$2:$B$100, 2, TRUE), "")'
    );

You can import the lookup as a CSV:

1. In your repo, there is `data/E1RM.csv`  
2. In Google Sheets: **File → Import → Upload** → select `data/E1RM.csv`  
3. Choose “Insert new sheet” and name it `E1RM`

Your script will then reference `E1RM!$A$2:$B$100` as before.

---

## Installation

### 1. Manual Copy-Paste

1. Open your Google Sheet  
2. Go to **Extensions → Apps Script**  
3. Delete any placeholder code and paste the contents of `Code.gs`  
4. Save and authorize the script

### 2. Using `clasp` (recommended)

    npm install -g @google/clasp
    clasp login
    clasp create --title "Workout Formatter" --type sheets

1. Copy `Code.gs` into your local project directory  
2. Run `clasp push`  
3. In your sheet, open **Extensions → Apps Script** and authorize

---

## Configuration

At the top of `generateWeek()`:

    var blockNumber = 1;    // Block sheet to write to
    var week        = 1;    // Week of data to process
    var clear       = false; // true → wipe existing output first

---

## Usage

1. Populate your **Block X Plan** sheet (cols A–K)  
2. Ensure you have an **E1RM** sheet or imported CSV with Repmax/% lookup  
3. Create or open the **Block X** sheet  
4. In the Apps Script editor, run `generateWeek()`

---

## Contributing

1. Fork this repo  
2. Create a branch: `git checkout -b feat/your-feature`  
3. Commit: `git commit -m "feat: …"`  
4. Push and open a Pull Request

---

## License

This project is licensed under the MIT License. See `LICENSE` for details.
