# :nodoc:
abstract class XLSX::Entry
  abstract def new(bytes : Bytes) : self
end
