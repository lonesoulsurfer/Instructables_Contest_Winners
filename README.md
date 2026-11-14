# Instructables Contest Winners

This project finds the Instructables members with the most contest wins. A Python script collects the data from the Instructables website. It then makes a ranked table of the top 50 winners, with the prize level of each win and the profile statistics of each member.

**[See the top 50 winners](https://lonesoulsurfer.github.io/Instructables_Contest_Winners/)**

## What the results show

For each member in the top 50, the table shows:

- Rank, username, total wins and a link to the profile
- Followers, member date, location, number of Instructables and total views
- First win, most recent win and favourite category
- Number of wins at each prize level: Grand Prize, First Prize, Second Prize, Third Prize, Runner Up, Judges' Prize and others
- The list of contests won

## Files in this repository

| File | Contents |
|------|----------|
| `index.html` | The results web page |
| `instructables_top_winners.py` | The script that collects the data |
| `Results/instructables_winners.xlsx` | Results as an Excel workbook |
| `Results/instructables_winners.csv` | Results as a CSV file |
| `Results/instructables_top50.html` | Results as HTML, ready to paste into an Instructable |

## Run the script yourself

The script uses a headless web browser (Playwright). It takes some time, because it waits between page loads so that it does not overload the Instructables website.

### What you need

- Python 3
- These packages: `playwright` and `openpyxl`

### Steps

1. Download this repository. Click **Code → Download ZIP**, then extract the ZIP file.
2. Open a Command Prompt or PowerShell window in the folder.
3. Install the packages:

   ```
   pip install playwright openpyxl
   ```

4. Install the browser for Playwright:

   ```
   playwright install chromium
   ```

5. Run the script:

   ```
   python instructables_top_winners.py
   ```

6. Wait until the script stops. It saves the results next to the script.

> **Note:** To show more or fewer members, change `TOP_N` near the top of the script. The default is 50.

> **Note:** The script reads public pages of the Instructables website. If Instructables changes its website, the script can stop working.

## Support

Do you like this project? You can support my work on [GitHub Sponsors](https://github.com/sponsors/lonesoulsurfer).

## Licence

This project uses the [MIT License](LICENSE). You can use, change and share it. Copyright (c) 2026 Marcus (lonesoulsurfer).

The results contain public profile data from Instructables. The MIT License covers the code only.
