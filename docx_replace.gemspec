lib = File.expand_path("../lib", __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require "docx_replace/version"

Gem::Specification.new do |gem|
  gem.name = "docx_replace"
  gem.version = DocxReplace::VERSION
  # Ursprung zuerst, dann die Pflege dieses Forks. Adam Albrecht hat das Gem
  # geschrieben, Beyond Finance es danach betreut; der Copyright-Vermerk in
  # LICENSE.txt bleibt unveraendert.
  gem.authors = ["Adam Albrecht", "Sascha Skulima"]
  gem.email = ["adam.albrecht@gmail.com", "sskulima@symdok.de"]
  gem.summary = "Find and replace variables inside a Microsoft Word (.docx) template"
  gem.description = "Replaces placeholders in the body, headers and footers of a .docx template. Fork maintained for symdok; reads from and writes to paths or IO."
  gem.homepage = "https://github.com/Skulli/docx_replace"
  gem.license = "MIT"

  gem.metadata = {
    # Kein homepage_uri: gem.homepage steht schon oben, und zwei gleiche Links
    # nebeneinander quittiert RubyGems beim Bauen mit einer Warnung.
    "source_code_uri" => "https://github.com/Skulli/docx_replace",
    "changelog_uri" => "https://github.com/Skulli/docx_replace/blob/main/CHANGELOG.md",
    # Der Name "docx_replace" gehoert dem Original auf RubyGems. Konsumenten
    # pinnen diesen Fork per git-Tag; ein "gem push" waere ein Fehlgriff und
    # scheitert an diesem Host, statt unter fremdem Namen zu veroeffentlichen.
    "allowed_push_host" => "https://rubygems.invalid"
  }

  gem.files = `git ls-files`.split($/)
  gem.executables = gem.files.grep(%r{^bin/}).map { |f| File.basename(f) }
  gem.require_paths = ["lib"]
  gem.required_ruby_version = ">= 3.1"
  gem.add_runtime_dependency "rubyzip", ">= 1.2.1", "< 4"
  gem.add_development_dependency "rake"
  gem.add_development_dependency "rspec"
  gem.add_development_dependency "simplecov"
  gem.add_development_dependency "standard"
end
