# SMT Weekly Deck – V2 (new code)

This folder holds only the new code. The old code and notebooks are in the `weekly auto` folder.

## Files
| File | What it is |
|---|---|
| `SMT_Weekly_Deck.ipynb` | **Run this.** Two steps: install libraries, run the weekly deck. |
| `run_weekly.bat` | Runs `smt_weekly_deck.py` (the notebook calls this; you can also double-click it). |
| `smt_weekly_deck.py` | Builds the deck. |
| `looker_capture.py` | Opens Looker once and saves all chart screenshots. |
| `sharepoint_upload.py` | Uploads the finished deck to SharePoint. |
| `output\` | Screenshots, the spreadsheet images for slides 33/35/36, and the decks. |

## How a run works
1. **Screenshots first.** Edge opens once. Each Looker dashboard (main, Nike volumes, Recon) is opened once and every chart is taken from it. All screenshots are saved in `output\` (for example `output\slide_4_1.png`).
2. **Then the PowerPoint.** Last week's deck is copied, the screenshots are placed on their slides, and the new deck is saved in `output\SMT Weekly Deck\<year> - Weekly Exec Reporting\SMT Weekly Deck\`.

3. **Then SharePoint.** The finished deck is uploaded to Operations SharePoint › Weekly Exec Reporting › `<year> - Weekly Exec Reporting` › SMT Weekly Deck. A deck with the same name there is replaced. The first time, Edge shows a Microsoft sign-in: sign in with your work account and the upload carries on. Use `--no-upload` to skip this step, or `py smt_weekly_deck.py --upload-only` to upload the latest deck again without rebuilding it.

The deck file is named with the date of the run, e.g. `SMT Working Version - Weekly Service Update - OPS Week Comm 20.11.2026.pptx`.

If a chart can't be found, the run lists the missing chart names and does not build the deck.

## First time on a laptop
1. Install Python from https://www.python.org/downloads/ (tick **"Add python.exe to PATH"**).
2. Run the first notebook cell, or in Command Prompt: `py -m pip install playwright python-pptx pillow`
3. Log in to Looker once: `py smt_weekly_deck.py --login` → sign in in the Edge window → close the window.

## Which chart goes on which slide
Edit `TARGET_TILE_REPLACEMENTS` near the top of `smt_weekly_deck.py`. Charts from a different dashboard are set in `TILE_REPORT_URLS`.
