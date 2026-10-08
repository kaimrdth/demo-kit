# Capturing scenarios

`gifcap` does the recording and drives the apps. Every command prints JSON; `gifcap <command> --help` lists options.
`templates/scene.json` is a working scenario.

## Scenario steps

A scenario is `{"name": ..., "steps": [...]}`; each step is an object with one key.

| Step | Value |
| --- | --- |
| `focus` | window (title or process substring). Also makes it the typing target. |
| `record_start` | `{"window" \| "region": [x,y,w,h] \| "monitor": 1, "pad", "name"}` |
| `record_stop` | `{"after": seconds}` (default 0.6) |
| `click`, `double_click`, `right_click`, `move` | point |
| `drag` | `{"from": point, "to": point, "duration"}` |
| `type` | text, or `{"text", "cps", "window"}`; `\n` presses Enter |
| `keys` | `"ctrl+v"`, `"ctrl+a delete"`, `"win+shift+d"`, or `{"keys", "window"}` |
| `wait` | seconds |
| `wait_idle` | `{"quiet": 2, "timeout": 90}`: waits until the recorded area stops changing |
| `mark` | name: records the time within the clip |
| `screenshot` | name, or `{"name", "window" \| "region" \| "monitor"}`: saves a PNG |

Points: `{"window": "Notepad", "frac": [0.5, 0.3]}` (fraction of the window), `{"window": ..., "px": [40, 120]}`
(pixels from its top-left), or `{"x": 900, "y": 400}` (screen pixels).

Keystrokes only go to the typing target and stop if anything else takes focus; `type` and `keys` fail unless a
`focus` or a click into a window came first. That's deliberate: never work around it with `--anywhere`.

## Set the stage

1. Ask the team to open the tool with demo-safe sample data, close anything confidential, and turn on Do Not
   Disturb (Windows: Focus) so notifications don't appear mid-take. If you can open things yourself, do.
2. `gifcap windows` shows what's open, where, and on which monitor. Pick the window by a distinctive part of its
   title or by process name (e.g. `"WINWORD"`, `"Contract Coach"`).
3. Size matters: GIFs are downscaled to 1280 px wide, so record a window rather than a whole 4K monitor, and make
   the app's text large enough to read (zoom to 110-125% if it's small).
4. Do a dry run of the steps by hand (or with input steps but no recording) if the UI is unfamiliar.

## Write the scenario

- Start with `focus` on the app, then `record_start` on its window (`"pad": 8` keeps the edge visible).
- Pause ~0.5 s at the start so the viewer sees the "before" state.
- After clicking into a text box, wait ~0.5 s before typing; apps often drop the first keystroke otherwise.
- Type at a natural speed (default 12 characters per second); type long prompts faster (`"cps": 25`) or paste them
  with `ctrl+v` after putting text on the clipboard yourself.
- For an AI reply: `mark` before sending, `keys: enter`, then `wait_idle` (quiet 2-3 s, a timeout well above the
  slowest reply you expect), then `mark` again. Speed up the gap later.
- End on the payoff and hold it: `record_stop` with `"after": 1.0`.
- Use `screenshot` steps for stills you'll want in the one-pager and deck.

## Run and review: the retake loop

```
gifcap run demo\scenes\01-flags-risks.json
gifcap sheet demo\takes\01-flags-risks.mkv demo\takes\01-flags-risks.sheet.png
```

Open the contact sheet and check:

- The whole action is in frame: nothing cut off at the edges, and no text truncated (e.g. the first word of a prompt).
- The text that was typed is exactly what you meant (compare with the scenario).
- No pop-ups, notifications, tooltips or unrelated windows got in.
- The ending shows the result clearly.
- Nothing from the never-on-screen list is visible (see confidentiality.md).

If anything's off, fix the scenario (or the stage) and rerun. If `gifcap run` stops with "took focus", something
else grabbed the screen mid-take: find out what (often a notification or the team clicking), then rerun.

## Turn takes into GIFs

Use the timeline's `marks` (in `demo/takes/<name>.timeline.json`) to trim and speed up:

```
gifcap gif demo\takes\01-flags-risks.mkv demo\gifs\01-flags-risks.gif --trim 0.4- --speed 6.1-28.4=4
```

Then `gifcap sheet` the GIF and look at it once more. Aim for under 5 MB per GIF so it pastes into Teams and Slack;
if it's bigger, trim harder, use `--fps 12` or `--max-width 960`.
