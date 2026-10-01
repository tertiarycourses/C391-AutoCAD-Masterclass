#!/bin/sh
set -e
S=~/.claude/skills/wsq-to-non-wsq/scripts
python3 $S/convert_deck.py \
  --src "source-wsq/courseware/WSQ Technical Drawing with AutoCAD-v13.pptx" \
  --out "courseware/AutoCAD Masterclass (C391)-v1.0.pptx" \
  --old-title "WSQ Technical Drawing with AutoCAD" --new-title "AutoCAD Masterclass" \
  --old-code TGS-2021005538 --new-code C391 \
  --sub "WSQ Course Code: TGS-2021005538=>Course Code: C391" \
  --sub "Version v13  ·  21 August 2026=>Version v1.0  ·  1 October 2026" \
  --sub "COURSE SLIDES  ·  WSQ=>COURSE SLIDES  ·  NON-WSQ COURSE" \
  --sub "75% attendance is required.=>Attend both days in full." \
  --sub "WSQ · Tertiary Infotech Academy=>Tertiary Infotech Academy" \
  --wsq-slides 2,9,10,11,13,14,15,112,113,114 --feedback-slide 111
python3 build/fix_deck.py
