lib = File.expand_path("../lib", __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require "docx_replace/version"

Gem::Specification.new do |gem|
  gem.name = "docx_replace"
  gem.version = DocxReplace::VERSION
  gem.authors = ["Adam Albrecht"]
  gem.email = ["adam.albrecht@gmail.com"]
  gem.description = "Find and replace variables inside a Micorsoft Word (.docx) template"
  gem.summary = "Find and replace variables inside a Micorsoft Word (.docx) template"
  gem.homepage = "https://github.com/adamalbrecht/docx_replace"
  gem.license = "MIT"

  gem.files = `git ls-files`.split($/)
  gem.executables = gem.files.grep(%r{^bin/}).map { |f| File.basename(f) }
  gem.require_paths = ["lib"]
  gem.add_runtime_dependency "rubyzip", ">= 1.2.1", "< 4"
  gem.add_development_dependency "rake"
  gem.add_development_dependency "rspec"
  gem.add_development_dependency "simplecov"
  gem.add_development_dependency "standard"
end
