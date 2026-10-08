# Outputs

Every builder takes a JSON spec (paths inside are relative to the spec file) and prints JSON with the files it made.
Templates for each spec are in `templates/`. After building, look at the result: open the PNG, or read the PDF.

## Scenario GIFs and stills

- One GIF per scenario: `gifs/NN-short-name.gif`, numbered in plan order. 5-20 seconds, under ~5 MB.
- Stills: `screenshot` steps in a scenario, or `gifcap still take.mkv stills/name.png --at SECONDS`. Pick frames
  where the result is fully visible.
- PDFs and printed pages show a GIF's **first frame**, so in the one-pager use stills, not GIFs.

## How-it-works diagram

`powershell -ExecutionPolicy Bypass -File scripts\render-diagram.ps1 -Spec demo\diagrams\how-it-works.json`
makes `how-it-works.png` (and an editable `.pptx`). Decks and one-pagers can use the same JSON directly, which keeps
the diagram editable there.

```json
{"direction": "LR",
 "nodes": [{"id": "in", "label": "Attorney uploads an NDA", "kind": "person"},
           {"id": "ai", "label": "Contract Coach compares it to the playbook", "kind": "agent"},
           {"id": "pb", "label": "Clause playbook", "kind": "data"},
           {"id": "out", "label": "Redlines and a risk summary in Word", "kind": "end"}],
 "edges": [{"from": "in", "to": "ai"}, {"from": "pb", "to": "ai"}, {"from": "ai", "to": "out", "label": "~1 min"}]}
```

`kind`: `step` (default), `agent` (the AI part, highlighted), `decision`, `data`, `person`, `start`, `end`.
`direction`: `LR` (left to right, best for 3-7 steps) or `TB` (top to bottom). Keep labels under ~8 words; 4-7 nodes
read best. Show how the *user* experiences it, not the code architecture, unless the audience is technical.

## One-pager

`scripts\build-onepager.ps1 -Spec demo\one-pager.json` makes `one-pager.pdf`, `.png` and an editable `.pptx`.
A portrait Letter page: title, tagline, problem, what it does, why it matters, 1-3 screenshots, how it works, team.

```json
{"title": "Contract Coach", "tagline": "One sentence: what it does for whom.",
 "problem": "2-3 sentences, in the team's words.",
 "what": ["3-5 bullets, each a capability"], "impact": ["2-4 bullets: time saved, risk reduced, who benefits"],
 "screenshots": [{"image": "stills/summary.png", "caption": "Risk summary"}],
 "how": "diagrams/how-it-works.json", "team": "Team 4: names", "contact": "Teams channel or email"}
```

Keep it scannable: if the text gets small in the PNG, cut words rather than adding space. `accent` sets the colour.

## Deck

`scripts\build-deck.ps1 -Spec demo\deck.json [-Template company.potx]` makes `deck.pptx` (GIFs play in it) and
`deck.pdf`. With a company template, slides use its masters and branding.

Slide types:

| `type` | Fields |
| --- | --- |
| `title` | `title`, `subtitle`, `team` |
| `statement` | `text`, `attribution` (one big line: the problem, a quote, a number) |
| `bullets` | `title`, `bullets` (3-5, short) |
| `media` | `title`, `media` (GIF, PNG, JPG or MP4), `caption` |
| `media-bullets` | `title`, `media`, `bullets` (2-4) |
| `diagram` | `title`, `diagram` (path to a diagram spec, or the spec inline), `caption` |

Every slide can have `notes` (speaker notes). Top level: `title`, `team` (footer), `accent`, `template`.

A good 6-8 slide arc for a hackathon demo: title, the problem (statement), the 30-second moment (media), 1-2 more
scenarios (media or media-bullets), how it works (diagram), impact or what's next (bullets). Write speaker notes the
presenter can read cold.

## Sizzle reel

`gifcap reel demo\reel.json demo\reel.mp4`: a 45-90 second 1080p MP4 of title cards, clips and stills with captions
and crossfades.

```json
{"items": [
  {"title": "Contract Coach", "subtitle": "Commercial Legal - Team 4", "seconds": 3},
  {"clip": "takes/01-flags-risks.mkv", "trim": "0.4-30", "speed": ["6.1-28.4=4"],
   "caption": "Flags off-standard clauses in seconds", "redact": ["1830,2,190,28"]},
  {"image": "stills/redline.png", "seconds": 3, "caption": "Redlines land in Word"},
  {"title": "Thank you", "subtitle": "Find us in the Legal Ops channel", "seconds": 2}]}
```

Clips use the raw takes (best quality) with the same `trim`, `speed`, `redact`, `blur` and `crop` you used for the
GIFs. **Apply the same redactions as the GIFs.** Captions: 3-8 words, one idea each. Open with the payoff, not the
setup; end on a title card with how to find the team. Contact-sheet the result (`gifcap sheet reel.mp4 ...`) and check
every caption and redaction.
