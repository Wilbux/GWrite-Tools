#!/bin/bash
# Author:  marcia 'aicra' wilbur
# Purpose: validate and repair functionality for Linux if using M$ writers
# Usage:   grep character path/to/*.md
# Example: grep – filename.md

echo "replacing MS weirdness"
result=$(grep – filename.md)
echo $result
sed -i 's/–/-/g' ./*.md
echo "below are the instances of non functional dash (–)" 
cat filename.md | grep '–'
