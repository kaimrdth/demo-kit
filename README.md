# Demo kit

A skill for Claude Code and Codex that turns something your team built into a demo pack:

- **Scenario GIFs**: one per thing worth showing, named by scenario
- **A one-pager** (PDF): problem, what it does, screenshots, how it works
- **A slide deck** (PowerPoint, GIFs play in it; optionally on your company template)
- **A how-it-works diagram**, editable in the deck
- **A sizzle reel** (MP4) with title cards and captions

You answer three questions and approve a plan. The agent learns the project, drives the app on screen, records and
reviews its own takes, redacts anything confidential, and assembles the pack.

Windows only. Uses [GIF Capture](https://github.com/kaimrdth/gifcapture) to record and PowerPoint to build documents.

## Install

1. Install GIF Capture: run `GifCapture-Setup.exe` from its
   [latest release](https://github.com/kaimrdth/gifcapture/releases/latest). It puts `gifcap` on your PATH.
2. Install the skill, either way:
   - **Claude Code plugin:** in Claude Code, run
     ```
     /plugin marketplace add kaimrdth/demo-kit
     /plugin install demo-kit@demo-kit
     ```
     (The repo is private, so Git on your machine needs GitHub access to it.)
   - **Claude Code and/or Codex, by copying:** clone this repo, then
     ```
     powershell -ExecutionPolicy Bypass -File install.ps1
     ```
     This copies the skill to `~\.claude\skills\demo-kit` and `~\.codex\skills\demo-kit`. Re-run after pulling
     updates.
3. Restart the agent.

## Use

Open your agent in your project folder and say something like *"make a demo of what we built"*. It will:

1. Read your project, then ask: what it is and who it's for; the moment you'd show in 30 seconds; and what must
   never appear on screen.
2. Propose 3-5 scenarios, a storyboard for each, and a reel order, for you to approve or edit.
3. Record each scenario. Keep your hands off the mouse and keyboard while it does; it asks before each take. Expect
   approval prompts: recording the screen happens outside the agent's sandbox.
4. Check every frame for confidential content and cover it.
5. Build the pack in `demo\`, with a `REVIEW.md` listing what to check.

Use sample data, and turn on Do Not Disturb so notifications don't end up in a take.

## What's inside

```
plugins/demo-kit/skills/demo-kit/
  SKILL.md            the workflow the agent follows
  references/         capture, confidentiality, outputs, Windows notes
  templates/          plan, scenario, diagram, one-pager, deck and reel examples
  scripts/            build-deck.ps1, build-onepager.ps1, render-diagram.ps1 (PowerPoint)
```

The scripts also work on their own, e.g.
`powershell -ExecutionPolicy Bypass -File plugins\demo-kit\skills\demo-kit\scripts\build-deck.ps1 -Spec deck.json`.
