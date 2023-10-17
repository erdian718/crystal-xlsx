# Represents the XLSX workbook.
class XLSX::Book
  private REL_OFFICE_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"

  @dir = Path.posix
  @rels = Relationships.new
  @sheets = Sheets.new
  @entries = {} of String => Bytes

  private def initialize(file : Compress::Zip::File)
    file.entries.each do |entry|
      if entry.file?
        entry.open do |io|
          @entries[entry.filename] = io.getb_to_end
        end
      end
    end

    rels = Relationships.new(fetch("_rels/.rels"))
    target = Path.posix(rels.fetch(REL_OFFICE_DOCUMENT) { "xl/workbook.xml" })
    @dir = target.parent
    @rels = Relationships.new(fetch(@dir.join("_rels", "#{target.basename}.rels")))

    workbook = XML.parse(IO::Memory.new(fetch(target)), XML_PARSER_OPTIONS).first_element_child
    raise Error.new("Invalid workbook entry") unless !workbook.nil? && workbook.name == "workbook"

    workbook.children.each do |node|
      case node.name
      when "sheets"
        @sheets = Sheets.new(node)
      when "calcPr"
        # TODO
      else
        # TODO
      end
    end
  end

  def self.new(io : IO, sync_close : Bool = false) : self
    Compress::Zip::File.open(io, sync_close) do |file|
      new(file)
    end
  end

  def self.new(filename : Path | String) : self
    Compress::Zip::File.open(filename) do |file|
      new(file)
    end
  end

  private macro fetch(path)
    %path = {{path}}.to_s
    entry = @entries[%path]?
    raise Error.new("Missing entry: #{%path}") if entry.nil?
    entry
  end

  def sheet_size : Int
    @sheets.size
  end

  def sheet_names : Array(String)
    @sheets.names
  end

  def [](index : Int) : Sheet
    sheet(@sheets.sheet(index))
  end

  def [](name : String) : Sheet
    sheet(@sheets.sheet(name))
  end

  private def sheet(sheet)
    entry = fetch(@dir.join(@rels[sheet[:id]]))
    if entry.is_a?(Sheet)
      entry
    else
      Sheet.new(sheet[:name], entry)
    end
  end

  private class Sheets
    @sheets = [] of {id: String, name: String, sheet_id: String}

    def initialize(node = nil)
      return if node.nil?
      node.children.each do |sheet|
        @sheets << {id: sheet["id"], name: sheet["name"], sheet_id: sheet["sheetId"]}
      end
    end

    def size : Int
      @sheets.size
    end

    def names : Array(String)
      @sheets.map { |sheet| sheet[:name] }
    end

    def sheet(index : Int)
      @sheets[index]
    end

    def sheet(name : String)
      @sheets.each do |sheet|
        return sheet if sheet[:name] == name
      end
      raise Error.new("Sheet not found: #{name}")
    end
  end

  private class CalcPr
    def initialize(node = nil)
      # Clears `calcId` to achieve automatic formula update.
    end
  end
end
