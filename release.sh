#!/bin/sh
# Bump project version and create git tag
# Usage: ./release.sh [major|minor|patch]
set -e
part=${1:-patch}
bump2version "$part"
