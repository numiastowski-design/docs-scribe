# Changelog

All notable changes to `docs-scribe` are written here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the package follows
[Semantic Versioning](https://semver.org/spec/v2.0.0.html): while the version is `0.x`, a minor
version may break the published surface, and a patch never does. Dates are the publication dates on
the npm registry (UTC).

## [Unreleased]

### Changed

- Licence: MIT (was AGPL-3.0-or-later up to 0.5.0). The versions already published keep their
  licence. A new `THIRD-PARTY.md` carries the third-party notices the bundle owes: SheetJS SSF,
  jpeg-js, the Adobe glyph list, pdf.js, and the text of the Apache License 2.0.
- The type declarations of the package's own modules moved inside the archive, from
  `dist/types/packages/doc-read/src/…` to `dist/types/packages/docs-scribe/src/…`. The `exports` map
  follows, so the four entries resolve as before; only a deep import of a declaration file sees it.
- The declarations of the slide layout, of the PNG decoder and of the zip gate that only this reader
  uses moved too, from `dist/types/packages/doc-ir/src/…` and `dist/types/packages/doc-gate/src/…`
  to `dist/types/packages/docs-scribe/src/…`. What stays under `dist/types/packages/doc-ir/` is what
  both readers share. The four entries resolve as before.

## [0.5.0] - 2026-11-17

### Added

- `require("docs-scribe")` works from CommonJS on the announced Node.js versions: every entry is
  served under the `default` condition of `exports`. It is the same ES module; no CommonJS build is
  added.
- `"sideEffects": false` in `package.json`: importing any entry only declares, so a bundler may drop
  what you do not use.
- This changelog ships in the package.

### Changed

- The package announces Node.js 24 or later (`engines`), instead of 20: these are the versions its
  checks run on. npm warns on an older Node and still installs.

### Fixed

- The type declarations name their files with an extension (`./docx.js`). A project compiled with
  `moduleResolution` set to `node16` or `nodenext` resolved none of the package's internal type
  imports, so every type of the package became `any`.

## [0.4.2] - 2026-11-17

### Changed

- No change in behaviour. The code was made easier to read (PowerPoint reader restructured,
  formatting, lint fixes); the output is identical before and after on 205 Office documents.

## [0.4.1] - 2026-11-17

### Changed

- The package ships readable code: `dist/index.js` is the default entry, with its source map, and
  the same code minified is available as `docs-scribe/min` and `docs-scribe/light/min`.

### Added

- A chart whose category axis runs last to first carries `categoriesReversed`; its table keeps the
  order of the sheet.

### Fixed

- A chart that carries no copy of its values: a range shared by several series is counted once,
  and a chart set aside by its cell ceiling is reported by a `chart-withheld` notice instead of
  being published as a title and an empty table.
- An Office file whose compound-file sector table holds a value that is not a sector number reads
  that entry as empty instead of failing.

## [0.4.0] - 2026-11-16

### Added

- A chart that carries no copy of its values is read from its sheet, within a ceiling on the cells
  it may read.

### Fixed

- The published type declarations no longer import a package that is not published, so they
  resolve.

## [0.3.1] - 2026-11-14

### Changed

- `repository` points to the public source repository, and the licence section says the
  Corresponding Source is published there. No code change.

## [0.3.0] - 2026-11-13

### Changed

- No nameable change in behaviour: the bundle differs from 0.2.0 by two bytes (one renamed
  minified identifier). The version was raised to keep both packages in step.

## [0.2.0] - 2026-11-10

### Added

- Direct colours of Word text: `color` and `background`.

## [0.1.1] - 2026-11-09

### Changed

- Nothing but the number: the first version published from continuous integration, to prove that
  path.

## [0.1.0] - 2026-11-09

### Added

- First release: the readers of Word, PowerPoint and Excel files and of Google Docs and Sheets
  payloads, split from the PDF package (then `skryba`, now `pdf-skryba`), which carried them until
  this day.

The registry also lists `0.0.0-stage`, a placeholder npm creates for staged publishing. It is not a
release.

[Unreleased]: https://github.com/numiastowski-design/docs-scribe
[0.5.0]: https://www.npmjs.com/package/docs-scribe/v/0.5.0
[0.4.2]: https://www.npmjs.com/package/docs-scribe/v/0.4.2
[0.4.1]: https://www.npmjs.com/package/docs-scribe/v/0.4.1
[0.4.0]: https://www.npmjs.com/package/docs-scribe/v/0.4.0
[0.3.1]: https://www.npmjs.com/package/docs-scribe/v/0.3.1
[0.3.0]: https://www.npmjs.com/package/docs-scribe/v/0.3.0
[0.2.0]: https://www.npmjs.com/package/docs-scribe/v/0.2.0
[0.1.1]: https://www.npmjs.com/package/docs-scribe/v/0.1.1
[0.1.0]: https://www.npmjs.com/package/docs-scribe/v/0.1.0
