#!/usr/bin/env bash

readonly SCRIPT_DIR=$(dirname "$0")

python3 ${SCRIPT_DIR}/outlook_agenda_to_ics_file.py ${HOME}/Downloads/Outlook.CSV ${HOME}/Downloads/Outlook.ics
