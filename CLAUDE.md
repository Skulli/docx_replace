# docx_replace (Fork)

Fork von [Beyond-Finance/docx_replace](https://github.com/Beyond-Finance/docx_replace):
Platzhalter-Ersetzung in `.docx`-Vorlagen. Einziger Konsument ist symdok
(Git-Gem, per Tag gepinnt). Nicht auf RubyGems veröffentlicht — der Name gehört
dem fremden Original.

**Sprachregel in diesem Repo:** nach außen englisch (README, CHANGELOG,
Commit-Messages, Code-Kommentare — öffentliches Repo, englischsprachiges Erbe),
diese Datei deutsch.

## Warum der Fork

Das letzte RubyGems-Release 1.2.1 (2020) ersetzt nur in `word/document.xml`.
Kopf- und Fußzeilen (`word/header*.xml`, `word/footer*.xml`) kann erst der
unveröffentlichte Upstream-`master` von Dezember 2022 — genau deswegen hängt
symdok an der Git-Quelle und nicht am Release. Der Upstream gilt als tot
(letzter Push 2022); Rückportierungen dorthin sind nicht zu erwarten.

## Stolperfallen

- **Word zerlegt Platzhalter über `<w:r>`-Runs.** Ein Marker, der im Dokument
  sichtbar ist, existiert dann nicht als zusammenhängender String und wird
  stillschweigend nicht ersetzt. Deshalb gibt `#replace` die Trefferzahl
  zurück — eine 0 ist ein Befund, kein Nichts. Marker ohne Sonderzeichen
  halten besser.
- **`Zip::File#read` liefert `ASCII-8BIT`.** Vergleiche gegen UTF-8-Strings
  scheitern mit `Encoding::CompatibilityError`; der Spec-Helper `docx_part`
  setzt die Kodierung deshalb zurück.
- **`#commit` ohne Pfad ist destruktiv:** es löscht die Quelldatei und schiebt
  das Ergebnis an deren Stelle. Bei IO-Quellen gibt es keinen Pfad, dort ist es
  ein `ArgumentError`.
- **Actions laufen in diesem Fork** (in vielen Forks sind sie es nicht — hier
  waren sie über den alten Upstream-Workflow schon freigegeben). Nachgewiesen
  mit Lauf 34330670522: lint plus rspec über rubyzip 1/2/3, alle grün.

## Tests

- `bundle exec rspec` — 19 Beispiele, Abdeckung ist per SimpleCov auf 100 %
  Zeilen und Branch festgenagelt (`minimum_coverage`). Neuer Code ohne Spec
  lässt die Suite durchfallen; das ist Absicht.
- Fixtures entstehen zur Laufzeit über `spec/support/docx_builder.rb` statt als
  Binärdateien im Repo — was ein Spec voraussetzt, steht damit im Spec. Keine
  echten Kundenvorlagen hier ablegen, das Repo ist öffentlich.
- Matrix lokal: `BUNDLE_GEMFILE=gemfiles/rubyzip1.gemfile bundle exec rspec`
  (analog `rubyzip2`, `rubyzip3`).

## Release

1. Version in `lib/docx_replace/version.rb` erhöhen
2. `CHANGELOG.md` ergänzen
3. Annotiertes Tag `vX.Y.Z` setzen, `git push origin master vX.Y.Z`
4. In symdok den Tag-Pin im Gemfile nachziehen (`tag: "vX.Y.Z"`) + `bundle install`

Kein `gem push` — siehe oben.

## Konventionen

- `standardrb` muss grün sein (läuft in der CI als eigener Job)
- Commit-Messages englisch, im Stil des Repos (kein `chg:dev:`-Präfix wie in
  symdok — hier gilt die Upstream-Historie)
