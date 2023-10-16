require "compress/zip"
require "xml"

# TODO: Write documentation for `XLSX`
module XLSX
  VERSION = "0.1.0"

  private XML_PARSER_OPTIONS = XML::ParserOptions.default | XML::ParserOptions::NOBLANKS

  class Error < Exception
  end
end

require "./xlsx/*"
