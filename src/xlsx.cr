require "compress/zip"
require "xml"

# TODO: Write documentation for `XLSX`
module XLSX
  VERSION = "0.1.0"

  private XML_PARSER_OPTIONS = XML::ParserOptions.default | XML::ParserOptions::NOBLANKS

  def self.ref(ridx : Int, cidx : Int) : String
    "#{cref(cidx)}#{rref(ridx)}"
  end

  def self.rref(ridx : Int) : String
    raise Error.new("Invalid row index: #{ridx}") if ridx < 0
    (ridx + 1).to_s
  end

  def self.cref(cidx : Int) : String
    raise Error.new("Invalid column index: #{cidx}") if cidx < 0
    return ('A' + cidx).to_s if cidx < 26
    "#{cref(cidx//26 - 1)}#{cref(cidx % 26)}"
  end

  def self.idx(ref : String)
    ords = ref.to_slice
    ords.each_with_index do |ord, i|
      return {ridx(String.new(ords[i..])), cidx(ords[...i])} if '0'.ord <= ord <= '9'.ord
    end
    raise Error.new("Invalid reference: #{ref}")
  end

  def self.ridx(rref : String) : Int
    ridx = rref.to_i - 1
    raise Error.new("Invalid row reference: #{rref}") if ridx < 0
    ridx
  end

  def self.cidx(cref : String) : Int
    cidx(cref.to_slice)
  end

  private def self.cidx(cref : Slice(UInt8)) : Int
    raise Error.new("Invalid column reference: #{String.new(cref)}") if cref.size < 1

    a = cidx(cref[-1])
    return a if cref.size == 1

    b = cidx(cref[...-1])
    26 + 26*b + a
  end

  private def self.cidx(cref : UInt8) : Int
    ord = cref.to_i
    return ord - 'A'.ord if 'A'.ord <= ord <= 'Z'.ord
    return ord - 'a'.ord if 'a'.ord <= ord <= 'z'.ord
    raise Error.new("Invalid column reference char: #{cref.chr}")
  end

  class Error < Exception
  end
end

require "./xlsx/*"
