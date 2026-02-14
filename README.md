# NUSched

Export your Nile University class schedule to an ICS file and import it into Google Calendar (or any calendar app).

## Features

- Fetches your registered schedule directly from NU Self-Service
- Displays all classes in a table where you can review and toggle each one
- Generates a standard `.ics` file with weekly recurring events
- Correct Cairo timezone, weekday mapping, instructor names, and room locations
- Works for any student -- just paste your own browser request
- Built-in tutorial video and Google Calendar import instructions

## Requirements

- Python 3.8+
- `requests` library

```
pip install -r requirements.txt
```

## Usage

### 1. Run the app

```
python nusched.py
```

### 2. Get your request from the browser

1. Log in to [NU Self-Service](https://register.nu.edu.eg/PowerCampusSelfService) and open your schedule page.
2. Press **F12** to open Chrome DevTools, go to the **Network** tab.
3. Reload the schedule page.
4. Find the **Student** request (POST to `Schedule/Student`).
5. Right-click it and choose **Copy > Copy as fetch**.

> You can also click the **Show Tutorial** button inside the app to watch a video walkthrough.

### 3. Paste and fetch

1. Click **Fetch Schedule** in the app.
2. Paste the copied fetch command into the dialog and click **OK**.
3. Your classes will appear in the table.

### 4. Generate the ICS file

1. Use the checkmarks to select/deselect classes you want to include.
2. Click **Generate ICS File**.
3. Confirm the schedule is correct.
4. The file `schedule_export.ics` is created in the app folder.

### 5. Import into Google Calendar

1. Open [Google Calendar](https://calendar.google.com).
2. Click the gear icon and go to **Settings**.
3. In the left sidebar, click **Import & export**.
4. Click **Select file from your computer** and choose `schedule_export.ics`.
5. Pick which calendar to add it to and click **Import**.

> **Tip:** Create a separate calendar (e.g. "NU Schedule") so you can easily delete and re-import if your schedule changes.

## Project Structure

```
nusched/
  nusched.py              # Main application (GUI + fetch + ICS generation)
  requirements.txt        # Python dependencies
  assets/
    scheduletutorial.mp4  # Tutorial video
  schedule_export.ics     # Generated output (after running)
```

## License

MIT
