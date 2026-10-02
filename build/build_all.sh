#!/usr/bin/env bash
set -euo pipefail
COURSE_REPO="$(cd "$(dirname "$0")/.." && pwd)"
export COURSE_REPO
cd "$COURSE_REPO"
python3 build/build_activity_packs.py
bash .claude/skills/non-wsq-courseware-build/build/build_courseware.sh
