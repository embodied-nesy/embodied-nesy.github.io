# ReS AI 2026 opening deck

Deliverable: `../iros_res_ai_2026_opening.pptx`.

The deck follows the workshop website's morning program. Its introduction,
organizer acknowledgements, research framing, program overview, paper track,
debate, award, and closing roles adapt the supplied SemRob deck. Session cards
follow the ReS AI schedule in order. The two poster-directory slides can be
skipped during the five-minute introduction.

## Sources

- `index.html`: workshop title, date, program, speaker affiliations and photos,
  paper titles, organizer names, links, and topic descriptions.
- `IROS_ReS_AI_Abstracts.docx`: speaker abstracts, cross-checked with the website.
- `ReS AI 2026 Submission Status.csv`: acceptance decisions cross-checked with
  the 18 published papers (4 oral, 14 poster-designated). Review scores and
  rejected submissions are not included in the slides.
- `ReS AI 2026 Reviewers Status.csv`: all 43 distinct reviewer display names,
  preserved verbatim and alphabetized across two program committee slides.
- `Copy of rss_semrob_2024-Opening.pptx`: structural reference, not current facts.
- https://2026.ieee-iros.org/: conference identity and official logo.
- https://2026.ieee-iros.org/program/workshops-tutorials/: Room 411 and the
  08:30–12:30 workshop block. The workshop website's 08:25 opening remarks are
  retained, followed by the first keynote at 08:30.
- https://www.ieee-ras.org/event/2026-ieee-rsj-international-conference-on-intelligent-robots-and-systems-iros-61738/:
  David L. Lawrence Convention Center, Pittsburgh.

Sources checked September 26, 2026. Each slide includes source/presenter notes.
The award slide intentionally contains no winner: none is supplied by the repo.
The full program committee acknowledgement uses the separately supplied
reviewer-status CSV; names are included regardless of assignment or completion
counts. Reviewer performance and individual assignments are not displayed.

## Design and rebuild

16:9 editable PowerPoint, with the repository's banner and portraits, the
conference's official logo, original editable PowerPoint icons, and a workshop
QR code. No new photo or AI-generated illustration was necessary.

`build_deck.py` reads `content.json` (extracted from `index.html`) and uses assets
in this directory plus the repository images. Install `python-pptx`, `Pillow`,
and `qrcode`, then run `python presentation/build_deck.py` from the repo root.

## Panel update

User-confirmed panel: Sebastian Scherer, Jiayuan Mao, Yezhou (YZ) Yang, and
Jean Oh. Slide 9 includes Jean as an invited panelist, and slide 23 displays
only these four panelists. The agenda uses “Panel debate” to reflect this update.
