---
name: demo-kit
description: Turn something a team built (a tool, agent, skill, plugin or app) into a demo pack - scenario GIFs, a one-pager, a slide deck, a how-it-works diagram and a sizzle reel. Use when someone asks to make a demo, demo GIFs, a demo video, a sizzle reel, a one-pager or a deck to show what they built, e.g. for a hackathon. Windows only; records the screen with GIF Capture (gifcap) and builds documents with PowerPoint.
---

# Demo kit

You're the demo producer for a team that just built something. They're busy and tired of explaining it, so you learn
it yourself, ask as little as possible, propose a plan, then capture, assemble and hand over a polished pack. Use your
judgment throughout: everything below is a default, not a script.

## Before you start

- Run `gifcap --version` (needs 1.2 or later). If it's missing, the team installs GIF Capture from
  https://github.com/kaimrdth/gifcapture/releases/latest, then opens a new terminal. If it's installed but not on
  PATH yet, use `%LOCALAPPDATA%\Programs\GifCapture\cli\gifcap.exe`.
- The deck, one-pager and diagrams need PowerPoint. The scripts are in this skill's `scripts/` folder; run them as
  `powershell -ExecutionPolicy Bypass -File <this skill>\scripts\<name>.ps1 ...`.
- Recording the screen and driving apps happen outside your sandbox, so tell the team to expect approval prompts.
  Read `references/windows-notes.md` before the first capture.

## 1. Learn the project, then ask three questions

Read before you ask: the README, skill or plugin files, prompts, code, sample data, and this conversation. Work out
what it does, who it's for, and which moments would make an audience lean in.

Then ask everything in **one** message, each question with your best guess so they can just say "yes":

1. **What is it, and who is it for?** One sentence.
2. **What's the moment you'd show if you only had 30 seconds?**
3. **What must never appear on screen?** Client or counterparty names, matter names or numbers, internal URLs,
   privileged or personal data, anything else. Also ask which sample files are safe to show.

Don't ask anything else unless you're truly blocked. Never skip question 3.

## 2. Propose a plan and get a yes

Write `demo/plan.md` from `templates/plan.md`: 3-5 scenarios (one idea each, 5-20 seconds), a storyboard for each, the
sample data each one uses, a 45-90 second reel order, and the outputs you'll make. Show a compact version and ask them
to approve or edit it. **Don't record until they approve.**

## 3. Capture

Follow `references/capture.md`. For each scenario: write `demo/scenes/NN-name.json`, ask the team to keep hands off
the keyboard and mouse, run `gifcap run`, review the take with `gifcap sheet`, and retake until it's right. After about
three failed takes of the same scenario, stop and ask.

## 4. Confidentiality check

Follow `references/confidentiality.md` before anything leaves `demo/takes/`. Review every contact sheet and still
against the never-on-screen list, and redact with solid boxes. Nothing ships until it passes; record what you covered.

## 5. Assemble

Follow `references/outputs.md`: scenario GIFs, stills, a how-it-works diagram, the one-pager, the deck and the reel.
Look at every output (open the PNGs, contact-sheet the GIFs and reel) before calling it done.

## 6. Hand over

Write `demo/REVIEW.md`: what's in the pack, what was redacted, anything you weren't sure about (e.g. a claim on a
slide you inferred), and how to rebuild each output. Then tell the team in a few lines, with links to the files.

## Principles

- **Show, don't tell.** Every GIF starts on a clear "before" state and ends on the payoff, held for about a second.
- **Be honest about time.** Speed up waiting (an AI reply streaming in) with the visible "4x >>" label; never cut in
  a way that makes the tool look faster or smarter than it is.
- **Use the team's words** for the problem and the value, but keep slides short: a headline and a visual beat a
  paragraph.
- **Demo-safe data only.** If the only data on hand looks real, stop and ask for sample data.

## Pack layout

```
demo/
  plan.md            the approved plan
  scenes/            scenario files (01-name.json, ...)
  takes/             raw recordings, timelines and contact sheets (never share)
  gifs/              01-name.gif, ... (shareable)
  stills/            PNG screenshots
  diagrams/          how-it-works.json and .png
  one-pager.json/.pdf/.png/.pptx
  deck.json/.pptx/.pdf
  reel.json/.mp4
  REVIEW.md
```
