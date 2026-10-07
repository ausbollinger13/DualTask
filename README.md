# DualTask

A Windows desktop app for running a dual-task experiment. The subject keeps a mouse cursor inside a randomly moving target while typing 4-digit codes with the other hand. Each trial's raw data and summary scores are saved to an Excel file.

Developed in the NASA Neuroscience Lab by Austin Bollinger.

## Download

Get the latest Windows build from the [Releases](../../releases) page. Unzip it and run `DualTask_v2.2.exe`. Keep the `_internal` folder next to the exe; the program needs it to run.

> The public build doesn't include a passcode, so task parameters stay locked. To enable unlocking, see [Parameter passcode](#parameter-passcode).

## Task modes

| Mode | What the subject does |
|---|---|
| **Tracking Only** | Keep the white crosshair inside a moving green bubble using their usual mouse hand. |
| **Codes Only** | Type each 4-digit code shown at the top of the screen with the other hand. Each code stays up for 5 s, and a code counts as correct only if all four digits are entered before it disappears. |
| **Dual Task** | Do both at the same time. |

"Run Again" moves through the standard order: Tracking Only → Codes Only → Dual Task.

## Session flow

1. **Login.** Enter a Subject ID (required) and a Session ID (optional).
2. **Subjective ratings.** Stanford Sleepiness Scale, Subjective Fatigue Scale (Samm-Perelli), and mouse handedness. These are collected once per subject.
3. **Task parameters.** The parameters are shown but locked. The operator can unlock them with the passcode.
4. **Task.** Full screen, with a 5-second countdown. Press `Esc` to end early.
5. **Notes.** The subject can type optional remarks about the trial.
6. **Results.** Tracking and code-entry scores, plus options to run again, change parameters, start a new subject, or close.

## Parameters

| Parameter | Default |
|---|---|
| Duration (s) | 110 |
| Number of codes | 20 |
| Code interval ± (s) | 5 |
| Max speed (px/tick) | 6.0 |
| Min speed (px/tick) | 2.0 |
| Mouse gain | 1.0 |

The task mode and an optional buzzer are also set on this screen.

## Output

Each trial is saved to the save folder shown on the Task Parameters screen (you can change it there) as:

```
{SubjectID}[_session{SessionID}]_{mode}_trial{N}_{MM_DD_HHMM}.xlsx
```

The session part is left out when no Session ID is entered. Each workbook has three tabs:

- **Data:** raw samples. Tracking columns are `elapsed_s, cursor_x, cursor_y, bubble_x, bubble_y, error_px`, and code columns are `elapsed_s, code, response, rt_s, correct`. Dual-task files show both side by side.
- **Summary:** completion time, IDs, scale ratings, mouse hand, tracking error and time on target, code accuracy and mean correct RT, and the subject's notes.
- **Parameters:** every task setting used for the trial.

The file is written as soon as the task ends, then saved again with the notes once the subject continues.

## Running from source

Requires Windows and Python 3.12 (other 3.x versions will probably work).

```powershell
pip install openpyxl
python dual_task.py
```

## Building the executable

```powershell
pip install pyinstaller openpyxl
pyinstaller --noconfirm DualTask_v2.2.spec
```

The build is written to `dist\DualTask_v2.2\`. The older `.spec` files are kept for reference. Use the newest one, because only the newest one bundles `passcode.txt`.

## Parameter passcode

Task parameters are locked behind a passcode so subjects can't change them by accident. The passcode isn't stored in this repository. To set one:

- **From source:** create `passcode.txt` next to `dual_task.py` with the passcode on a single line. Git ignores this file.
- **When building:** create the same file before running PyInstaller, and it will be bundled into the build.
- **In a downloaded release:** put `passcode.txt` inside the `_internal` folder.

If no passcode file is found, the unlock dialog says so and the parameters stay at their defaults.
