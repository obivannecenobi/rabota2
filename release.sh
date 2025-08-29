#!/bin/sh
# Bump project version, run tests, build distribution and push tags
# Usage: ./release.sh [major|minor|patch]
set -e

part=${1:-patch}

# run tests if present
if [ -d tests ]; then
    python -m pytest
fi

# bump version in VERSION and config.json
bump2version "$part"

# build distribution
if command -v pyinstaller >/dev/null 2>&1; then
    pyinstaller --onefile app/main.py -n webnovel-planner
else
    mkdir -p dist
    zip -r dist/app.zip app data assets requirements.txt VERSION
fi

# push commit and tags
git push
git push --tags
