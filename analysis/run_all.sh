#!/bin/sh
# Rebuilds every dashboard data file and the workbook from the files in raw/.
set -e
for s in 01_deck_figures 02_mca_tiles 03_grade_totals 04_dashboard_data 06_mca_lexile_data 07_build_workbook; do python3 analysis/$s.py; done
