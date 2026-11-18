describe DocxReplace::Doc do
  let(:output_file) { Tempfile.new("docx_replace_tmp_output") }
  let(:output_text) { docx_content(output_file) }

  it "replaces a single variable in a very basic document" do
    doc = described_class.new(get_fixture("basic.docx"))
    doc.replace("FOOBAR", "hello world")
    doc.commit(output_file)

    expect(output_text).to match(/hello world/)
    expect(output_text).not_to match(/FOOBAR/)
  end

  it "can replace multiple occurrences of the same variable" do
    doc = described_class.new(get_fixture("multiple.docx"))
    doc.replace("FOOBAR", "hello world", true)
    doc.commit(output_file)

    expect(output_text).not_to match(/FOOBAR/)
    expect(output_text.scan("hello world").size).to eq(2)
  end

  it "does not replace multiple occurrences unless instructed to do so" do
    doc = described_class.new(get_fixture("multiple.docx"))
    doc.replace("FOOBAR", "hello world", false)
    doc.commit(output_file)

    expect(output_text.scan("FOOBAR").size).to eq(1)
    expect(output_text.scan("hello world").size).to eq(1)
  end

  describe "headers and footers" do
    it "replaces variables in headers and footers, not just the body" do
      Dir.mktmpdir do |dir|
        template = build_docx(File.join(dir, "template.docx"),
          body: "Body: FOOBAR", header: "Header: FOOBAR", footer: "Footer: FOOBAR")

        doc = described_class.new(template)
        doc.replace("FOOBAR", "hello world", true)
        output = File.join(dir, "output.docx")
        doc.commit(output)

        expect(docx_part(output)).to match(/Body: hello world/)
        expect(docx_part(output, "word/header1.xml")).to match(/Header: hello world/)
        expect(docx_part(output, "word/footer1.xml")).to match(/Footer: hello world/)
      end
    end

    it "keeps every entry of the package intact" do
      Dir.mktmpdir do |dir|
        template = build_docx(File.join(dir, "template.docx"),
          body: "FOOBAR", header: "FOOBAR", footer: "FOOBAR")
        before = docx_entries(template)

        doc = described_class.new(template)
        doc.replace("FOOBAR", "hello world", true)
        output = File.join(dir, "output.docx")
        doc.commit(output)

        expect(docx_entries(output)).to eq(before)
      end
    end

    it "works on documents without headers or footers" do
      Dir.mktmpdir do |dir|
        template = build_docx(File.join(dir, "template.docx"), body: "FOOBAR")

        doc = described_class.new(template)
        doc.replace("FOOBAR", "hello world")
        output = File.join(dir, "output.docx")
        doc.commit(output)

        expect(docx_part(output)).to match(/hello world/)
        expect(docx_entries(output)).not_to include("word/header1.xml")
      end
    end
  end

  describe "#commit" do
    it "replaces the template in place when called without a path" do
      Dir.mktmpdir do |dir|
        template = build_docx(File.join(dir, "template.docx"), body: "FOOBAR")

        doc = described_class.new(template)
        doc.replace("FOOBAR", "hello world")
        doc.commit

        expect(docx_part(template)).to match(/hello world/)
        expect(Dir.children(dir)).to eq(["template.docx"])
      end
    end

    it "raises instead of reporting success when the destination is unwritable" do
      Dir.mktmpdir do |dir|
        template = build_docx(File.join(dir, "template.docx"), body: "FOOBAR")

        doc = described_class.new(template)
        doc.replace("FOOBAR", "hello world")

        expect { doc.commit(File.join(dir, "nicht-da", "output.docx")) }
          .to raise_error(SystemCallError)
      end
    end

    it "writes its intermediate file to the given temp dir" do
      Dir.mktmpdir do |dir|
        Dir.mktmpdir do |temp_dir|
          template = build_docx(File.join(dir, "template.docx"), body: "FOOBAR")

          doc = described_class.new(template, temp_dir)
          doc.replace("FOOBAR", "hello world")
          output = File.join(dir, "output.docx")
          doc.commit(output)

          expect(docx_part(output)).to match(/hello world/)
        end
      end
    end
  end

  describe "#replace" do
    it "escapes XML special characters in the replacement" do
      Dir.mktmpdir do |dir|
        template = build_docx(File.join(dir, "template.docx"), body: "FOOBAR")

        doc = described_class.new(template)
        doc.replace("FOOBAR", "Müller & Söhne <GmbH>")
        output = File.join(dir, "output.docx")
        doc.commit(output)

        expect(docx_part(output)).to include("Müller &amp; Söhne &lt;GmbH&gt;")
      end
    end

    it "returns the number of replaced occurrences" do
      Dir.mktmpdir do |dir|
        template = build_docx(File.join(dir, "template.docx"),
          body: "FOOBAR and FOOBAR", header: "FOOBAR")

        doc = described_class.new(template)

        expect(doc.replace("FOOBAR", "hello world", true)).to eq(3)
      end
    end

    it "returns zero when the pattern is not present" do
      Dir.mktmpdir do |dir|
        template = build_docx(File.join(dir, "template.docx"), body: "nothing to see")

        doc = described_class.new(template)

        expect(doc.replace("FOOBAR", "hello world", true)).to eq(0)
      end
    end

    it "counts one replacement per part when not replacing multiples" do
      Dir.mktmpdir do |dir|
        template = build_docx(File.join(dir, "template.docx"),
          body: "FOOBAR and FOOBAR", header: "FOOBAR")

        doc = described_class.new(template)

        expect(doc.replace("FOOBAR", "hello world")).to eq(2)
      end
    end

    it "treats the replacement as literal text, not as a pattern" do
      Dir.mktmpdir do |dir|
        template = build_docx(File.join(dir, "template.docx"), body: "FOOBAR")

        doc = described_class.new(template)
        doc.replace("FOOBAR", 'C:\10\Neu')
        output = File.join(dir, "output.docx")
        doc.commit(output)

        # With a string replacement, gsub! would read \1 as a backreference and
        # write "C:0\Neu" instead.
        expect(docx_part(output)).to include('C:\10\Neu')
      end
    end

    describe "placeholders Word split across runs" do
      # What Word writes for a placeholder after editing or proofing, taken
      # from a real template: one visible word, three runs.
      let(:split) do
        'AKT_DAT_</w:t></w:r><w:proofErr w:type="gramStart"/>' \
          '<w:r w:rsidRPr="009F09F4"><w:rPr><w:b/></w:rPr><w:t>FORMAT(</w:t></w:r>' \
          '<w:proofErr w:type="gramEnd"/><w:r><w:t>TT.MM.JJ)'
      end

      it "replaces them and keeps every tag, so the XML stays balanced" do
        Dir.mktmpdir do |dir|
          template = build_docx(File.join(dir, "template.docx"), body: "Datum: #{split}", footer: split)

          doc = described_class.new(template)
          expect(doc.replace("AKT_DAT_FORMAT(TT.MM.JJ)", "17.11.26", true)).to eq(2)
          output = File.join(dir, "output.docx")
          doc.commit(output)

          expect(docx_part(output)).to include(
            '<w:t>Datum: 17.11.26</w:t></w:r><w:proofErr w:type="gramStart"/>' \
            '<w:r w:rsidRPr="009F09F4"><w:rPr><w:b/></w:rPr><w:t></w:t></w:r>' \
            '<w:proofErr w:type="gramEnd"/><w:r><w:t></w:t></w:r>'
          )
          expect(docx_part(output, "word/footer1.xml")).to include("<w:t>17.11.26</w:t>")
        end
      end

      it "does not match across anything that shows on the page or holds hidden text" do
        separators = %w[<w:br/> <w:cr/> <w:tab/> <w:ptab/> <w:noBreakHyphen/> <w:softHyphen/>]
        separators += ['<w:sym w:char="F0E0"/>', '<w:footnoteReference w:id="1"/>', "<w:drawing></w:drawing>"]
        lines = separators.map { |tag| "FOO</w:t>#{tag}<w:t>BAR" }
        lines << "x</w:t></w:r><w:del><w:r><w:delText>FOO</w:delText></w:r></w:del><w:r><w:t>BAR"
        lines << "x</w:t></w:r><w:r><w:instrText>FOO</w:instrText></w:r><w:r><w:t>BAR"

        Dir.mktmpdir do |dir|
          template = build_docx(File.join(dir, "template.docx"), body: ["FOO", "BAR", *lines])

          doc = described_class.new(template)

          # The second call is answered from the cached visible text.
          2.times { expect(doc.replace("FOOBAR", "hello world", true)).to eq(0) }
        end
      end

      it "notices when a part was changed from outside" do
        Dir.mktmpdir do |dir|
          template = build_docx(File.join(dir, "template.docx"), body: "nothing")

          doc = described_class.new(template)
          expect(doc.replace("FOOBAR", "hello world")).to eq(0)

          doc.document_contents["word/document.xml"].sub!("nothing", "FOO</w:t></w:r><w:r><w:t>BAR")

          expect(doc.replace("FOOBAR", "hello world")).to eq(1)
        end
      end

      it "uses a Regexp pattern as given" do
        Dir.mktmpdir do |dir|
          template = build_docx(File.join(dir, "template.docx"), body: ["FOO1", "FOO</w:t></w:r><w:r><w:t>2"], header: "BAR")

          doc = described_class.new(template)

          expect(doc.replace(/FOO\d/, "hello world", true)).to eq(1)
        end
      end

      it "matches a String containing markup against the raw XML, as before" do
        Dir.mktmpdir do |dir|
          template = build_docx(File.join(dir, "template.docx"), body: "FOO")

          doc = described_class.new(template)
          expect(doc.replace("<w:t>FOO</w:t>", "hello world")).to eq(1)

          expect(doc.document_contents["word/document.xml"]).to include("<w:r>hello world</w:r>")
        end
      end

      it "refuses an empty pattern instead of matching between every character" do
        Dir.mktmpdir do |dir|
          template = build_docx(File.join(dir, "template.docx"), body: "FOO")

          doc = described_class.new(template)

          expect { doc.replace(nil, "X", true) }.to raise_error(ArgumentError, /empty/)
          expect { doc.replace("", "X", true) }.to raise_error(ArgumentError, /empty/)
        end
      end
    end

    describe "line breaks in the replacement" do
      it "writes them as Word line breaks, also after a split placeholder" do
        Dir.mktmpdir do |dir|
          template = build_docx(File.join(dir, "template.docx"),
            body: ["FOOBAR", "FOO</w:t></w:r><w:r><w:t>BAR"])

          doc = described_class.new(template)
          expect(doc.replace("FOOBAR", "eins\nzwei\r\ndrei", true)).to eq(2)

          breaks = '</w:t><w:br/><w:t xml:space="preserve">'
          expect(doc.document_contents["word/document.xml"])
            .to include("<w:t>eins#{breaks}zwei#{breaks}drei</w:t></w:r>")
            .and include("<w:t>eins#{breaks}zwei#{breaks}drei</w:t></w:r><w:r><w:t></w:t></w:r>")
        end
      end

      it "keeps the newline outside a text run, where no element may go" do
        Dir.mktmpdir do |dir|
          template = build_docx(File.join(dir, "template.docx"),
            body: "x</w:t></w:r><w:r><w:instrText>FOOBAR</w:instrText></w:r><w:r><w:t>y")

          doc = described_class.new(template)
          doc.replace("FOOBAR", "eins\nzwei")

          expect(doc.document_contents["word/document.xml"]).to include("<w:instrText>eins\nzwei</w:instrText>")
        end
      end
    end

    it "accepts replacements that are not strings" do
      Dir.mktmpdir do |dir|
        template = build_docx(File.join(dir, "template.docx"), body: "FOOBAR")

        doc = described_class.new(template)
        doc.replace("FOOBAR", 42)
        output = File.join(dir, "output.docx")
        doc.commit(output)

        expect(docx_part(output)).to match(/42/)
      end
    end
  end

  describe "reading from and writing to IO" do
    it "opens a Pathname as a file, not as an IO" do
      Dir.mktmpdir do |dir|
        # Pathname answers to #read, which is exactly why it must not be
        # mistaken for an IO.
        template = Pathname.new(build_docx(File.join(dir, "template.docx"), body: "FOOBAR"))

        doc = described_class.new(template)
        expect(doc.replace("FOOBAR", "hello world")).to eq(1)

        output = File.join(dir, "output.docx")
        doc.commit(output)

        expect(docx_part(output)).to match(/hello world/)
      end
    end

    it "opens a File as a file, too" do
      Dir.mktmpdir do |dir|
        path = build_docx(File.join(dir, "template.docx"), body: "FOOBAR")

        File.open(path) do |file|
          doc = described_class.new(file)
          expect(doc.replace("FOOBAR", "hello world")).to eq(1)

          output = File.join(dir, "output.docx")
          doc.commit(output)

          expect(docx_part(output)).to match(/hello world/)
        end
      end
    end

    it "reads a template straight out of an IO" do
      Dir.mktmpdir do |dir|
        template = build_docx(File.join(dir, "template.docx"), body: "FOOBAR", header: "FOOBAR")

        doc = described_class.new(StringIO.new(File.binread(template)))
        expect(doc.replace("FOOBAR", "hello world", true)).to eq(2)

        output = File.join(dir, "output.docx")
        doc.commit(output)

        expect(docx_part(output)).to match(/hello world/)
        expect(docx_part(output, "word/header1.xml")).to match(/hello world/)
      end
    end

    it "returns the result as a rewound IO" do
      Dir.mktmpdir do |dir|
        template = build_docx(File.join(dir, "template.docx"), body: "FOOBAR", header: "FOOBAR")

        doc = described_class.new(template)
        doc.replace("FOOBAR", "hello world", true)
        io = doc.to_io

        expect(io).to be_a(StringIO)
        expect(io.pos).to eq(0)

        # The buffer has to be a readable .docx again, not just any bytes.
        roundtrip = File.join(dir, "roundtrip.docx")
        File.binwrite(roundtrip, io.read)
        expect(docx_part(roundtrip)).to match(/hello world/)
        expect(docx_part(roundtrip, "word/header1.xml")).to match(/hello world/)
        expect(docx_entries(roundtrip)).to eq(docx_entries(template))
      end
    end

    it "goes from IO to IO without touching the disk" do
      Dir.mktmpdir do |dir|
        template = build_docx(File.join(dir, "template.docx"), body: "FOOBAR")
        source = StringIO.new(File.binread(template))

        doc = described_class.new(source)
        doc.replace("FOOBAR", "hello world")

        again = described_class.new(doc.to_io)

        expect(again.matches(/(hello world)/)).to eq(["hello world"])
      end
    end

    it "refuses to commit in place when there is no path to commit to" do
      Dir.mktmpdir do |dir|
        template = build_docx(File.join(dir, "template.docx"), body: "FOOBAR")

        doc = described_class.new(StringIO.new(File.binread(template)))
        doc.replace("FOOBAR", "hello world")

        expect { doc.commit }.to raise_error(ArgumentError, /needs a path/)
      end
    end
  end

  describe "#matches" do
    it "returns the captured group of every match across all parts" do
      Dir.mktmpdir do |dir|
        template = build_docx(File.join(dir, "template.docx"),
          body: "$FIRST_NAME$ $LAST_NAME$", header: "$COMPANY$")

        doc = described_class.new(template)

        expect(doc.matches(/\$([A-Z_]+)\$/)).to contain_exactly("FIRST_NAME", "LAST_NAME", "COMPANY")
      end
    end

    it "deduplicates in #unique_matches and its alias" do
      Dir.mktmpdir do |dir|
        template = build_docx(File.join(dir, "template.docx"), body: "$NAME$ and $NAME$ again")

        doc = described_class.new(template)

        expect(doc.matches(/\$([A-Z_]+)\$/)).to eq(%w[NAME NAME])
        expect(doc.unique_matches(/\$([A-Z_]+)\$/)).to eq(%w[NAME])
        expect(doc.uniq_matches(/\$([A-Z_]+)\$/)).to eq(%w[NAME])
      end
    end
  end

  private

  def get_fixture(name)
    File.join(project_root, "spec", "fixtures", name)
  end

  def project_root
    File.expand_path(File.dirname(__FILE__, 2))
  end

  def docx_content(path)
    zip_file = Zip::File.new(path)
    zip_file.read("word/document.xml")
  end
end
