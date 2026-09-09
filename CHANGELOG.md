# Changelog

All notable changes to this fork are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the versioning
[SemVer](https://semver.org/).

This is a fork of [Beyond-Finance/docx_replace](https://github.com/Beyond-Finance/docx_replace),
picked up at its unreleased `master` (December 2022) because header and footer
support never made it into a release. It is not published to RubyGems;
consumers pin it by git tag. Versions continue the inherited 1.2.1.

## [1.3.0] - 2026-09-09

### Added

- `Doc.new` accepts any IO that responds to `#read`, so a template can be read
  straight out of object storage without a local copy first.
- `#to_io` returns the result as a rewound `StringIO`, the counterpart to
  `#commit` for consumers that never want the document on disk.
- Support for rubyzip 2 and 3; the dependency is now `>= 1.2.1, < 4`.
- CI across rubyzip 1, 2 and 3 (see `gemfiles/`), plus standardrb and
  dependabot.
- Gemspec metadata: this fork as homepage and source, a changelog link, and an
  `allowed_push_host` that makes an accidental `gem push` fail rather than
  publish under the original gem's name.

### Changed

- `#replace` returns the number of occurrences it replaced. It used to return
  the internal contents hash, which gave callers no way to tell a replacement
  from a silent miss - the common failure mode, since Word splits text across
  `<w:r>` runs.
- `#unique_matches` (and its alias `#uniq_matches`) now deduplicates. It only
  ever delegated to `#matches`.
- `#commit` without a path raises `ArgumentError` when the template came from
  an IO, because there is no file to write back to.
- Declares `required_ruby_version >= 3.1`; CI exercises Ruby 3.4.

### Fixed

- Reopening the freshly written document raised `ArgumentError` on rubyzip 3,
  which removed the positional `create` argument from `Zip::File#initialize`.
  The file always exists at that point, so the flag was a no-op.
- `fileutils` is required explicitly instead of relying on `tempfile` to pull
  it in.
- Committing in place moves the result over the template with `force` instead
  of removing the template first, which left a window in which neither file
  existed. The intermediate tempfile is closed once it has been written.

## [1.2.1] - 2020-02-14

Last release of the upstream gem; see its history for earlier changes.
