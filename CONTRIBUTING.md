# Contributing

## Changelog and versioning
- Each merge to the main branch must append a new section to `CHANGELOG.md2` with the new version and a summary of changes.
- After editing `CHANGELOG.md2`, run `bump2version` to update `VERSION` and `data/config.json` before committing.

## Release
Use `release.sh` to create a release. The script will fail if `CHANGELOG.md2` was not updated and will automatically bump the version using `bump2version`.
