# :nodoc:
class XLSX::Relationships
  @entries = {} of String => {type: String, target: String}

  def [](id : String) : String
    @entries[id].target
  end

  def fetch(type : String) : String
    fetch(type) do |type|
      raise Error.new("Relationship type not found: #{type}")
    end
  end

  def fetch(type : String, & : String -> String) : String
    @entries.each_value do |rel|
      return rel.target if rel.type == type
    end

    target = yield type
    idx = @entries.size
    loop do
      id = "rId#{idx += 1}"
      if !@entries.has_key?(id)
        @entries[id] = {type, target}
        return target
      end
    end
  end
end
