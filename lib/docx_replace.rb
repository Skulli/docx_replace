require "docx_replace/version"
require 'zip'
require 'tempfile'

module DocxReplace
  class Doc
    def initialize(path, temp_dir=nil)
      @zip_file = Zip::File.new(path)
      @temp_dir = temp_dir
      read_docx_file
    end

    def replace(pattern, replacement, multiple_occurrences=false)
      if multiple_occurrences
        @document_content.gsub!(pattern, replacement)
      else
        @document_content.sub!(pattern, replacement)
      end
    end

    def matches(pattern)
      @document_content.scan(pattern).map{|match| match.first}
    end

    def unique_matches(pattern)
      matches(pattern)
    end

    alias_method :uniq_matches, :unique_matches


    def commit(new_path=nil)
      write_back_to_file(new_path)
    end

    private
    DOCUMENT_FILE_PATH = 'word/document.xml'

    def new_temp_file(key = 'docxedit-')
      if @temp_dir.nil?
        temp_file = Tempfile.new(key)
      else
        temp_file = Tempfile.new(key, @temp_dir)
      end
    end

    def read_docx_file
      Dir.mktmpdir("docx_replace") do |dir|
        file = File.join(dir, "document.xml")
        @zip_file.extract(DOCUMENT_FILE_PATH, file)
        @document_content = file_content(file)
      end

      @document_content
    end

    def file_content(path, delete = false)
      content = nil
      File.open(path) do |f|
        content = f.read
        f.close
      end

      FileUtils.rm_rf(path) if delete

      return content
    end

    def write_back_to_file(new_path=nil)
      temp_file = new_temp_file("docxedit")

      Zip::OutputStream.open(temp_file.path) do |zos|
        @zip_file.entries.each do |e|
          unless e.name == DOCUMENT_FILE_PATH
            zos.put_next_entry(e.name)
            zos.print e.get_input_stream.read
          end
        end

        zos.put_next_entry(DOCUMENT_FILE_PATH)
        zos.print @document_content
      end

      if new_path.nil?
        path = @zip_file.name
        FileUtils.rm(path)
      else
        path = new_path
      end
      FileUtils.mv(temp_file.path, path)
      @zip_file = Zip::File.new(path)
    end
  end
end
