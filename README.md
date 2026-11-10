# MasterSheet Finder

A Windows app for searching and filtering a big giveaway/account sheet (CSV or Excel).

## Features

- Search by profile name, Steam name, Steam username, email, SteamID64 or trade link.
- Filter by trade link, email, and banned/active status.
- Filter by skin site: DMarket, KeyDrop, ForceDrop, CSGOFast, Hellcase, CSGO.Net, SkinClub, Juice.gg, Casehug, DaddySkins.
- Copy buttons for username, password, email, trade link and site IDs/links.
- Export the filtered rows as CSV.

The repo contains no sheet data. Keep your own CSV/XLSX files out of it.

## Run

```bash
pip install -r requirements.txt
python app.pyw
```

Or use `RUN_APP.bat`. `BUILD_EXE.bat` builds an exe.
