# Pitch Deck Check

A small, open-source page that checks a PowerPoint deck against the published rules of the Public Service Data/AI Challenge pitch training (November 19, 2026).

- **Local only.** The file is read in your browser. Nothing is uploaded, stored or sent.
- **No AI.** Every check is a fixed, readable rule in `analyse.js`.
- **Situates, does not judge.** It says which rule a deck meets, partly meets or does not meet. The judges decide.

## Use it
Online: https://pitchcheck.prismadatorum.ca · Source: https://github.com/dravitch/pitch-deck-check (the file is still read in your browser only).

1. Download or clone this repository.
2. Open `index.html` in a browser. It works offline: the ZIP reader (JSZip 3.10.1, MIT) is included.
3. Choose your `.pptx` file.

## Run it yourself, adapt it yourself
Each team runs its own copy, and owns its own use and its own changes. The authors do not host the tool for others, collect files, or maintain a service. To adapt the rules, edit `analyse.js`: each check is one block with an identifier (A1, D1, T1…). Please keep the rule text and its source visible in the page.

## What it checks
| Id | Rule | How |
|---|---|---|
| A1 | Text 18 pt or larger | font size of every text run; sizes inherited from a layout are reported as "To check" |
| A2 | Images have alternative text | image description field |
| A3 | Never colour alone; high contrast | manual check |
| D1 | No acronyms, on the slides or in the speech | words in capitals within sentences; all-capital headings are ignored |
| S1 | No outline slide | slide titles such as "Agenda" or "Outline" |
| P1 | No placeholders left | `[…]`, `___`, TBD, TODO |
| L1 | Links work on the day, and are hosted by the team or the department | external links and personal or third-party hosts |
| T1 | Use the full 10 minutes | timing marks `m:ss – m:ss` in the speaker notes |
| T2 | The first 60 seconds say what the problem is and why it matters | shows what is said; a person checks |
| S2 | Simple, uncluttered slides | more than 70 words per slide |
| S3 | Structure, including risks | mentions of risks, users, consultation, next steps, cost |

## Limits
- **Not validated on a corpus.** Tested on two real decks and the organizers' training deck only, with no baseline of known-good and known-bad decks. There is no measured rate of false alarms or missed problems. See the Methodology tab for the full list of limits.
- Automatic checks are signals, not verdicts.
- Contrast, colour and the quality of the story need a person.
- The acronym list will contain false positives: a word in capitals is not always an acronym.

## Source of the rules
Pitch Training for the Public Service Data/AI Challenge, November 19, 2026. This is an independent aid, not an official tool of the challenge.

## Licence
MIT (see `LICENSE`).
