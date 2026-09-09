# Builds minimal .docx files at runtime instead of shipping binary fixtures, so
# that what a spec relies on is readable in the spec itself. Word only needs the
# content types, the package relationships and the parts we address.
module DocxBuilder
  CONTENT_TYPES = <<~XML
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
      <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
    </Types>
  XML

  PACKAGE_RELS = <<~XML
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>
  XML

  DOCUMENT_RELS = <<~XML
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
    </Relationships>
  XML

  # Writes a .docx to +path+. Every keyword takes the visible text of that part;
  # pass nil for header/footer to leave the part out of the package entirely.
  def build_docx(path, body:, header: nil, footer: nil)
    Zip::OutputStream.open(path) do |zos|
      zos.put_next_entry("[Content_Types].xml")
      zos.print CONTENT_TYPES
      zos.put_next_entry("_rels/.rels")
      zos.print PACKAGE_RELS
      zos.put_next_entry("word/_rels/document.xml.rels")
      zos.print DOCUMENT_RELS
      zos.put_next_entry("word/document.xml")
      zos.print wordprocessing_part("w:document", "w:body", body)

      if header
        zos.put_next_entry("word/header1.xml")
        zos.print wordprocessing_part("w:hdr", nil, header)
      end

      if footer
        zos.put_next_entry("word/footer1.xml")
        zos.print wordprocessing_part("w:ftr", nil, footer)
      end
    end

    path
  end

  # Reads one part out of a .docx, defaulting to the main document body.
  # rubyzip hands back ASCII-8BIT, so the encoding has to be restored before
  # the content can be compared against UTF-8 expectations.
  def docx_part(path, part = "word/document.xml")
    Zip::File.open(path) { |zip| zip.read(part) }.force_encoding("UTF-8")
  end

  def docx_entries(path)
    Zip::File.open(path) { |zip| zip.entries.map(&:name).sort }
  end

  private

  def wordprocessing_part(root, wrapper, text)
    paragraphs = Array(text).map { |line| "<w:p><w:r><w:t>#{line}</w:t></w:r></w:p>" }.join
    inner = wrapper ? "<#{wrapper}>#{paragraphs}</#{wrapper}>" : paragraphs

    <<~XML
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <#{root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">#{inner}</#{root}>
    XML
  end
end
