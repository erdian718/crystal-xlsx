# :nodoc:
#
# Represents a collection of relationships.
class XLSX::Relationships
  @entries = {} of String => {type: String, target: String}

  def initialize(bytes : Bytes? = nil)
    return if bytes.nil?

    rels = XML.parse(IO::Memory.new(bytes), XML_PARSER_OPTIONS).first_element_child
    raise Error.new("Invalid relationships entry") unless !rels.nil? && rels.name == "Relationships"

    rels.children.each do |rel|
      if rel.name == "Relationship"
        @entries[rel["Id"]] = {type: rel["Type"], target: rel["Target"]}
      end
    end
  end

  # Returns the target by *id*.
  # If not found, raises an error.
  def [](id : String) : String
    entry = @entries[id]?
    raise Error.new("Relationship id not found: #{id}") if entry.nil?
    entry[:target]
  end

  # Returns the target by *type*.
  # If not found, raises an error.
  def fetch(type : String) : String
    fetch(type) do |type|
      raise Error.new("Relationship type not found: #{type}")
    end
  end

  # Returns the target by *type*.
  # If not found, calls the given block with *type*.
  def fetch(type : String, & : String -> String) : String
    @entries.each_value do |rel|
      return rel[:target] if rel[:type] == type
    end

    target = yield type
    idx = @entries.size
    loop do
      id = "rId#{idx += 1}"
      if !@entries.has_key?(id)
        @entries[id] = {type: type, target: target}
        return target
      end
    end
  end
end
