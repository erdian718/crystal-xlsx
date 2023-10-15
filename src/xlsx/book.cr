require "compress/zip"

class XLSX::Book
  @entries = {} of String => XLSX::Entry | Bytes

  def initialize(file : Compress::Zip::File)
    file.entries.each do |entry|
      entry.open do |io|
        @entries[entry.filename] = io.getb_to_end
      end
    end
  end

  def self.new(bytes : Bytes) : self
    new(IO::Memory.new(bytes))
  end

  def self.new(io : IO, sync_close : Bool = false) : self
    Compress::Zip::File.open(io, sync_close) do |file|
      self.new(file)
    end
  end

  def self.new(filename : Path | String) : self
    Compress::Zip::File.open(filename) do |file|
      self.new(file)
    end
  end
end
