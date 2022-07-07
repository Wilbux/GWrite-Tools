#!/bin/bash
# Author:  marcia 'aicra' wilbur
# Purpose: validate and repair functionality for Linux if using M$ vendor/writer
# Usage:   grep character path/to/*.md
# Example: grep – filename.md
# for file selection on front end use yad "--file-selection"
# GUI version
# Must have the yad package.
command -v yad >/dev/null 2>&1 || { echo >&2 "yad package required but it is not installed.  Aborting."; exit 99; }

# Must have the zenity package.
command -v zenity >/dev/null 2>&1 || { echo >&2 "yad package required but it is not installed.  Aborting."; exit 99; }

# yad --title "MS Destroyer" --width=200
# zenity --info --text "THUG life" 2>/dev/null

yad --title destroyMS --width=500 --height=400 --geometry=400x400 --text="Destroy MS characters!"
result=$(grep – filename.md)
yad --title offending item --geometry=400x400 --text="$result"
sed -i 's/–/-/g' ./*.md
yad --title destroyed --geometry=400x400 --text="below are the instances of non functional dash (–)" 
yad --title leftovers --geometry=400x400 --text="$result"
goodchar=$(grep - filename.md)
yad --title "good char - fixed" --geometry=400x400 --width=500 --height=400 --text="$goodchar"

# peace out
