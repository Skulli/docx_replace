# Docx Replace

Replace placeholders inside a `.docx` template — in the document body as well as
in headers and footers. Purposefully simple and feature-light: it does not parse
WordprocessingML, it substitutes strings in the XML parts and repackages the
document.

## About this fork

This is a fork of [Beyond-Finance/docx_replace](https://github.com/Beyond-Finance/docx_replace),
taken from its unreleased `master`, because header and footer support never made
it into a published release (the last one, 1.2.1, is from 2020 and only touches
`word/document.xml`).

It is **not published to RubyGems** — that name belongs to the original gem.
Pin it by git tag:

```ruby
gem "docx_replace", git: "https://github.com/Skulli/docx_replace.git", tag: "v1.3.0"
```

Works with rubyzip 1, 2 and 3; every release is tested against all three.

## Usage

```ruby
doc = DocxReplace::Doc.new("template.docx", Rails.root.join("tmp"))

# Returns the number of occurrences replaced - see the caveat below, a zero
# is worth acting on.
doc.replace("FIRSTNAME", user.first_name)
doc.replace("BIRTHDATE", user.birth_date, true)   # all occurrences

doc.commit("result.docx")
```

Headers and footers are handled the same way: `#replace` covers
`word/document.xml`, `word/header*.xml` and `word/footer*.xml`.

### Working without the disk

`Doc.new` also accepts any IO that responds to `#read`, and `#to_io` hands the
result back as a rewound `StringIO`. Useful when the template lives in object
storage and the result goes straight back there:

```ruby
doc = DocxReplace::Doc.new(StringIO.new(blob.download))
doc.replace("FIRSTNAME", user.first_name)

record.document.attach(io: doc.to_io, filename: "report.docx")
```

`#commit` without a path overwrites the template in place. That needs a path to
overwrite, so it raises `ArgumentError` when the template came from an IO.

### Finding placeholders

```ruby
doc.matches(/\$([A-Z_]+)\$/)         # every match, in document order
doc.unique_matches(/\$([A-Z_]+)\$/)  # deduplicated
```

The pattern needs a capture group; the first group of each match is returned.

## Caveat: Word splits text across runs

Word may distribute a single word over several `<w:r>` runs, typically after
editing. A placeholder that looks contiguous in Word then does not exist as a
contiguous string in the XML, and the replacement silently does nothing. That is
why `#replace` reports how many occurrences it replaced — check for zero rather
than trusting the template. Placeholders without special characters survive
better.

## Development

```bash
bundle exec rspec          # 19 examples, 100 % line and branch coverage (enforced)
bundle exec standardrb     # lint
```

New fixtures are built at runtime by `spec/support/docx_builder.rb` rather than
committed as binaries, so a spec states the document it relies on. The three
examples inherited from upstream still use `spec/fixtures/*.docx`.

To run the suite against one specific rubyzip generation:

```bash
BUNDLE_GEMFILE=gemfiles/rubyzip3.gemfile bundle exec rspec
```

CI does exactly that for rubyzip 1, 2 and 3 on Ruby 3.4.

## Credits

Originally written by Adam Albrecht, later maintained by Beyond Finance, and
based in turn on an older gem called [docxedit](https://github.com/oliamb/docxedit).
MIT licensed — see LICENSE.txt.
