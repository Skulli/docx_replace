require "docx_replace/version"
require "zip"
require "tempfile"
require "fileutils"
require "stringio"

module DocxReplace
  class Doc
    attr_reader :document_contents

    # +source+ is either a path or an IO, which lets a template be read straight
    # out of object storage without a local copy.
    def initialize(source, temp_dir = nil)
      if path_like?(source)
        @zip_file = Zip::File.new(source)
        @source_path = @zip_file.name
      else
        @zip_file = Zip::File.open_buffer(source)
      end

      @document_file_paths = find_query_file_paths
      @temp_dir = temp_dir
      @visible_text = {}
      read_docx_files
    end

    # Tags Word may put between the characters of one visible word: run
    # boundaries, proofing marks, bookmarks. Everything that shows on the page
    # (paragraphs, breaks, tabs, symbols, note references, drawings) or holds
    # text that does not (deletions, field codes) ends a match.
    TAG_WITHIN_TEXT = %r{<(?!/?w:(?:p|br|cr|tab|ptab|noBreakHyphen|softHyphen|sym|
      footnoteReference|endnoteReference|drawing|pict|object|fldChar|delText|instrText)[\s/>])[^>]*+>}x
    TAG = /<[^>]*>/

    # Returns how many occurrences were replaced across all parts.
    #
    # Word splits text across <w:r> runs, so a String pattern also matches with
    # tags between its characters. The replacement takes the place of the first
    # character and with it that run's formatting; tags inside a match are kept,
    # so the XML stays balanced. A Regexp, or a String containing markup, is
    # matched against the raw XML as given.
    def replace(pattern, replacement, multiple_occurrences = false)
      replace = replacement.to_s.encode(xml: :text)
      text = pattern.to_s unless pattern.is_a?(Regexp)
      raise ArgumentError, "pattern must not be empty" if text == ""

      tolerant = text && !text.include?("<")
      pattern = run_tolerant(text) if tolerant

      @document_contents.sum do |path, document|
        document.force_encoding("UTF-8")
        # Most placeholders of a template set are absent from a given template;
        # a plain substring check spares them the much slower tolerant match.
        next 0 if tolerant && !document.include?(text) && !visible_text(path).include?(text)

        # Block form on purpose: with a string replacement, sub!/gsub! read \1,
        # \0 and \\ as backreferences, so a value like 'C:\10\Neu' would be
        # mangled into 'C:0\Neu' - and reported as a successful replacement.
        count = 0
        document.public_send(multiple_occurrences ? :gsub! : :sub!, pattern) do |match|
          count += 1
          tolerant ? replace + match.scan(TAG).join : replace
        end
        count
      end
    end

    def matches(pattern)
      @document_contents.values.join.scan(pattern).map { |match| match.first }
    end

    def unique_matches(pattern)
      matches(pattern).uniq
    end

    alias_method :uniq_matches, :unique_matches

    def commit(new_path = nil)
      write_back_to_file(new_path)
    end

    # Returns the result as a rewound StringIO rather than writing it to disk -
    # the counterpart to #commit for object storage.
    def to_io
      buffer = Zip::OutputStream.write_buffer(StringIO.new("")) do |zos|
        write_entries(zos)
      end

      buffer.rewind
      buffer
    end

    private

    def run_tolerant(text)
      Regexp.new(text.chars.map { |char| Regexp.escape(char) }.join("(?:#{TAG_WITHIN_TEXT})*+"))
    end

    # The part with run boundaries removed. Cached against a copy of the part,
    # because document_contents is public and may be changed from outside.
    def visible_text(path)
      document = @document_contents[path]
      snapshot, visible = @visible_text[path]
      return visible if snapshot == document

      visible = document.gsub(TAG_WITHIN_TEXT, "")
      @visible_text[path] = [document.dup, visible]
      visible
    end

    # Anything carrying a path of its own is opened as a file - Pathname and
    # File answer to #read too, so #read alone does not identify an IO.
    def path_like?(source)
      source.is_a?(String) || source.respond_to?(:to_path)
    end

    def find_query_file_paths
      @zip_file.entries.map(&:name).select do |entry|
        !(/^word\/(document|footer[0-9]+|header[0-9]+).xml$/ =~ entry).nil?
      end
    end

    def read_docx_files
      @document_contents = {}
      @document_file_paths.each do |path|
        @document_contents[path] = @zip_file.read(path)
      end
    end

    def write_back_to_file(new_path = nil)
      temp_file = if @temp_dir.nil?
        Tempfile.new("docxedit-")
      else
        Tempfile.new("docxedit-", @temp_dir)
      end

      Zip::OutputStream.open(temp_file.path) { |zos| write_entries(zos) }
      # The path is all that is needed from here on; leaving the handle open
      # until GC gets around to it serves nothing.
      temp_file.close

      if new_path.nil?
        raise ArgumentError, "#commit needs a path when the template was read from an IO" if @source_path.nil?

        path = @source_path
      else
        path = new_path
      end
      # No rm before the mv: that left a window in which neither the template nor
      # the result existed. mv overwrites on its own. And no force: either -
      # FileUtils.mv swallows every SystemCallError when forced, so #commit
      # would report success while the destination kept its old content.
      FileUtils.mv(temp_file.path, path)
      # The file exists at this point, so rubyzip's `create` flag was always a
      # no-op here. It was removed as a positional argument in rubyzip 3.
      @zip_file = Zip::File.new(path)
      @source_path = @zip_file.name
    end

    # Copies every untouched entry over verbatim and writes the replaced parts
    # in their place.
    def write_entries(zos)
      @zip_file.entries.each do |entry|
        next if @document_file_paths.include?(entry.name)

        zos.put_next_entry(entry.name)
        zos.print entry.get_input_stream.read
      end

      @document_contents.each do |path, document|
        zos.put_next_entry(path)
        zos.print document
      end
    end
  end
end
