///|
/// A struct to handle any error and a message
pub(all) enum CalamineError {
  /// IO error
  Io(String)
  /// Ods specific error
  Ods(OdsError)
  /// xls specific error
  Xls(XlsError)
  /// xlsb specific error
  Xlsb(XlsbError)
  /// xlsx specific error
  Xlsx(XlsxError)
  /// vba specific error
  Vba(VbaError)
  /// cfb specific error
  De(DeError)
  /// General error message
  Msg(String)
} derive(Show, Eq)

///|
/// An enum for ODS specific errors.
pub(all) enum OdsError {
  /// A wrapper for IO errors
  Io(String)
  /// A wrapper for XML parsing errors
  Xml(String)
  /// A wrapper for XML attribute errors
  XmlAttr(String)
  /// A wrapper for ZIP errors
  Zip(String)
  /// Unexpected end of XML file
  XmlEof(String)
  /// File not found in ODS container
  FileNotFound(String)
  /// Password protected workbook
  Password
  /// Worksheet not found
  WorksheetNotFound(String)
  /// Cell error
  CellError(String)
  /// Unexpected node in XML
  UnexpectedNode(String)
} derive(Show, Eq)

///|
/// An enum for XLS specific errors.
pub(all) enum XlsError {
  /// A wrapper for IO errors
  Io(String)
  /// Invalid BOM
  InvalidBOM
  /// Biff version error
  BiffVersion(String)
  /// Parse error
  Parse(String)
  /// Workbook is password protected
  Password
  /// Worksheet not found
  WorksheetNotFound(String)
  /// Dimensions error
  Dimensions(String)
  /// General XLS error
  General(String)
} derive(Show, Eq)

///|
/// An enum for XLSB specific errors.
pub(all) enum XlsbError {
  /// A wrapper for IO errors
  Io(String)
  /// A wrapper for ZIP errors
  Zip(String)
  /// A wrapper for VBA errors
  Vba(VbaError)
  /// File not found in XLSB container
  FileNotFound(String)
  /// Relationship not found
  RelationshipNotFound
  /// Password protected workbook
  Password
  /// Worksheet not found
  WorksheetNotFound(String)
  /// Cell error
  CellError(String)
  /// Parse error
  Parse(String)
  /// Unexpected error
  Unexpected(String)
} derive(Show, Eq)

///|
/// An enum for XLSX specific errors.
pub(all) enum XlsxError {
  /// A wrapper for IO errors
  Io(String)
  /// A wrapper for ZIP errors
  Zip(String)
  /// A wrapper for VBA errors
  Vba(VbaError)
  /// A wrapper for XML parsing errors
  Xml(String)
  /// A wrapper for XML attribute errors
  XmlAttr(String)
  /// Parse string error
  Parse(String)
  /// Parse integer error
  ParseInt(String)
  /// Parse float error
  ParseFloat(String)
  /// Unexpected end of XML file
  XmlEof(String)
  /// Unexpected node in XML
  UnexpectedNode(String)
  /// File not found in XLSX container
  FileNotFound(String)
  /// Relationship not found
  RelationshipNotFound
  /// Non alphanumeric character found when parsing A1 style range
  Alphanumeric(UInt)
  /// Error when parsing column name in A1 style range
  NumericColumn(UInt)
  /// Missing column name when parsing A1 style range
  RangeWithoutColumnComponent
  /// Missing row number when parsing A1 style range
  RangeWithoutRowComponent
  /// Error when parsing dimensions of worksheet
  DimensionCount(UInt)
  /// Unknown cell type attribute
  CellTAttribute(String)
  /// Unexpected XML element or attribute error
  Unexpected(String)
  /// Unrecognized worksheet type or state
  Unrecognized(String)
  /// Unrecognized cell error type
  CellError(String)
  /// Password protected workbook
  Password
  /// Worksheet not found
  WorksheetNotFound(String)
  /// Table not found
  TableNotFound(String)
  /// The specified sheet is not a worksheet
  NotAWorksheet(String)
  /// Encoding error
  Encoding(String)
} derive(Show, Eq)

///|
/// An enum for VBA specific errors.
pub(all) enum VbaError {
  /// A wrapper for IO errors
  Io(String)
  /// Invalid directory entry
  InvalidDirectory
  /// Parse error
  Parse(String)
  /// General VBA error
  General(String)
} derive(Show, Eq)

///|
/// An enum for deserialization errors.
pub(all) enum DeError {
  /// Missing header
  MissingHeader(String)
  /// Unknown header
  UnknownHeader(String)
  /// Unexpected data type
  UnexpectedDataType(String)
  /// Index out of bounds
  IndexOutOfBounds
  /// General deserialization error
  General(String)
} derive(Show, Eq)

///|
/// Result type for calamine operations
pub typealias Result[T, CalamineError] as CalamineResult[T]

///|
/// Result type for XLSX operations
pub typealias Result[T, XlsxError] as XlsxResult[T]

///|
/// Result type for XLS operations
pub typealias Result[T, XlsError] as XlsResult[T]

///|
/// Result type for XLSB operations
pub typealias Result[T, XlsbError] as XlsbResult[T]

///|
/// Result type for ODS operations
pub typealias Result[T, OdsError] as OdsResult[T]

///|
/// Convert XLSX error to general error
pub fn XlsxError::to_error(self : XlsxError) -> CalamineError {
  CalamineError::Xlsx(self)
}

///|
/// Convert XLS error to general error
pub fn XlsError::to_error(self : XlsError) -> CalamineError {
  CalamineError::Xls(self)
}

///|
/// Convert XLSB error to general error
pub fn XlsbError::to_error(self : XlsbError) -> CalamineError {
  CalamineError::Xlsb(self)
}

///|
/// Convert ODS error to general error
pub fn OdsError::to_error(self : OdsError) -> CalamineError {
  CalamineError::Ods(self)
}

///|
/// Convert VBA error to general error
pub fn VbaError::to_error(self : VbaError) -> CalamineError {
  CalamineError::Vba(self)
}

///|
/// Convert deserialization error to general error
pub fn DeError::to_error(self : DeError) -> CalamineError {
  CalamineError::De(self)
}
