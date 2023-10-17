# Represents the XLSX worksheet.
class XLSX::Sheet
  getter name : String

  @data = Data.new

  def initialize(@name, bytes : Bytes)
    worksheet = XML.parse(IO::Memory.new(bytes), XML_PARSER_OPTIONS).first_element_child
    raise Error.new("Invalid worksheet entry") unless !worksheet.nil? && worksheet.name == "worksheet"

    worksheet.children.each do |node|
      case node.name
      when "sheetData"
        @data = Data.new(node)
      else
        # TODO
      end
    end
  end

  private class Data
    def initialize(node = nil)
      return if node.nil?

      # TODO
    end

    def [](row_index : Int, col_index : Int) : Cell
      # TODO
      Cell.new
    end

    def [](reference : String) : Cell
      # TODO
      Cell.new
    end
  end
end
