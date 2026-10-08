# Windows notes for agents

Things that cost time the first time round.

## Permissions and sandboxes

- Screen recording (`gifcap record`, `run`, `sheet` on a live screen) and simulated input need to run outside an
  agent sandbox. In Claude Code, approve running those commands without the sandbox; in Codex, the session needs
  permission to run commands outside its sandbox (approve when asked, or use a mode that allows it). Tell the team
  this before the first take so they aren't surprised by prompts.
- Sandboxed shells may also block the clipboard (`gifcap copy` fails with "Clipboard is busy") and PowerPoint
  automation. Run those outside the sandbox too.

## Paths

- `gifcap` is added to the user's PATH by the GIF Capture installer; terminals opened before the install don't see
  it. Fallback: `%LOCALAPPDATA%\Programs\GifCapture\cli\gifcap.exe`.
- Some agent hosts (e.g. desktop apps packaged as Windows apps) redirect writes under `%LOCALAPPDATA%` and
  `%APPDATA%` into a private copy. Keep the demo pack in the project folder or Documents, not in AppData.
- Paths with spaces (`OneDrive - Company`) are fine; quote them.

## PowerShell 5.1

- Scripts here are for Windows PowerShell 5.1, which ships with Windows. Run them with
  `powershell -ExecutionPolicy Bypass -File ...`.
- `&&` and `||` don't exist in 5.1; use `;` and `if ($?) { ... }`.
- Redirecting with `>` writes UTF-16 or UTF-8 with a byte-order mark. If you save `gifcap` JSON with `>`, read it
  back with `ConvertFrom-Json` in PowerShell (or `utf-8-sig` in Python). Write spec files with your editor tool,
  not `Out-File`, so they're plain UTF-8.

## PowerPoint

- The scripts start a windowless PowerPoint, and close it afterwards only if it wasn't already open with the
  user's own work. It's fine for the team to have PowerPoint open.
- Each script takes 5-20 seconds. If one fails with a COM error, rerun it once; if it fails again, report the error
  text.
- Without PowerPoint there's no deck, one-pager or diagram. Fall back to GIFs, stills and a reel, and say so.

## Screen

- Scaling: `gifcap` works in physical pixels, and so do window rects from `gifcap windows`. Don't convert.
- Multiple monitors: `gifcap monitors` lists them; monitor 1 is the primary. Window rects can be on any monitor,
  including negative coordinates.
- Hidden ("cloaked") windows don't appear in `gifcap windows`, even if their title shows in other tools.
