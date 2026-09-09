require "docx_replace/version"
require "zip"
require "tempfile"
require "fileutils"
require "stringio"

module DocxReplace
  class Doc
    attr_reader :document_contents

    # +source+ is either a path or any IO that responds to #read, which lets a
    # template be read straight out of object storage without a local copy.
    def initialize(source, temp_dir = nil)
      @zip_file = source.respond_to?(:read) ? Zip::File.open_buffer(source) : Zip::File.new(source)
      @source_path = @zip_file.name unless source.respond_to?(:read)
      @document_file_paths = find_query_file_paths
      @temp_dir = temp_dir
      read_docx_files
    end

    # Returns how many occurrences were replaced across all parts. Zero is worth
    # acting on: Word splits text across <w:r> runs, so a pattern that is
    # visually present in the document may not exist as a contiguous string.
    def replace(pattern, replacement, multiple_occurrences = false)
      replace = replacement.to_s.encode(xml: :text)

      @document_contents.sum do |_path, document|
        document.force_encoding("UTF-8")
        occurrences = document.scan(pattern).size
        next 0 if occurrences.zero?

        if multiple_occurrences
          document.gsub!(pattern, replace)
          occurrences
        else
          document.sub!(pattern, replace)
          1
        end
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
      # force, rather than removing the destination first: an rm followed by an
      # mv leaves a window in which neither the template nor the result exists.
      FileUtils.mv(temp_file.path, path, force: true)
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
