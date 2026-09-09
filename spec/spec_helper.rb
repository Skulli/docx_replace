require "simplecov"

# Coverage is only complete when the whole suite runs, so the gate has to skip
# partial runs - otherwise `rspec <file>:<line>` exits 2 despite "0 failures"
# and every attempt at debugging one example looks like a failure.
partial_run = ARGV.any? { |arg| arg.include?("_spec.rb") }

SimpleCov.start do
  add_filter "/spec/"
  enable_coverage :branch
  minimum_coverage line: 100, branch: 100 unless partial_run
end

$LOAD_PATH.unshift(File.expand_path("../lib", __dir__))

require "docx_replace"
require "zip"
require "tmpdir"
require "pathname"

Dir[File.expand_path("support/**/*.rb", __dir__)].sort.each { |f| require f }

RSpec.configure do |config|
  config.include DocxBuilder

  # Run specs in random order to surface order dependencies. If you find an
  # order dependency and want to debug it, you can fix the order by providing
  # the seed, which is printed after each run.
  #     --seed 1234
  config.order = "random"
end
