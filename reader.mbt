///|
/// A trait to share spreadsheet reader functions across different FileTypes
pub trait Reader {
  /// Creates a new instance from bytes
  new(Bytes) -> Self raise CalamineError

  /// Set header row (i.e. first row to be read)
  with_header_row(Self, HeaderRow) -> Self

  /// Gets VBA Project  
  vba_project(Self) -> VbaProject? raise CalamineError

  /// Get metadata
  metadata(Self) -> Metadata

  /// Read worksheet data in corresponding worksheet name
  worksheet_range(Self, String) -> Range[Data] raise CalamineError

  /// Fetch all worksheet data & names
  worksheets(Self) -> Array[(String, Range[Data])]

  /// Read worksheet formula in corresponding worksheet name
  worksheet_formula(Self, String) -> Range[String] raise CalamineError

  /// Get all sheet names of this workbook, in workbook order
  sheet_names(Self) -> Array[String]

  /// Fetch all sheets metadata
  sheets_metadata(Self) -> Array[Sheet]

  /// Get all defined names (Ranges names etc)
  defined_names(Self) -> Array[(String, String)]

  /// Get the nth worksheet. Shortcut for getting the nth
  /// worksheet name, then the corresponding worksheet.
  worksheet_range_at(Self, Int) -> Range[Data]? raise CalamineError
}

///|
/// A trait for readers that support reference-based data access
pub trait ReaderRef: Reader {
  /// Get worksheet range where shared string values are only borrowed.
  worksheet_range_ref(Self, String) -> Range[DataRef] raise CalamineError

  /// Get the nth worksheet range where shared string values are only borrowed.
  worksheet_range_at_ref(Self, Int) -> Range[DataRef]? raise CalamineError
}

///|
/// VBA Project placeholder
pub(all) struct VbaProject {
  data : Bytes
} derive(Show, Eq)

///|
/// Create a new VBA project
pub fn VbaProject::new(data : Bytes) -> VbaProject {
  { data, }
}

///|
/// Get VBA project data
pub fn VbaProject::data(self : VbaProject) -> Bytes {
  self.data
}

///|
/// Convenient function to open a file from bytes
pub fn[R : Reader] open_workbook(data : Bytes) -> R raise CalamineError {
  R::new(data)
}

///|
/// Workbook detection and automatic reader selection
pub enum WorkbookType {
  /// Excel 2007+ format (XLSX)
  Xlsx
  /// Excel 2003 format (XLS)
  Xls
  /// Excel Binary format (XLSB)
  Xlsb
  /// OpenDocument Spreadsheet (ODS)
  Ods
  /// Unknown format
  Unknown
} derive(Show, Eq)

///|
/// Detect workbook type from file signature
pub fn detect_workbook_type(data : Bytes) -> WorkbookType {
  if data.length() < 8 {
    return Unknown
  }

  // Check for ZIP signature (XLSX, XLSB, ODS)
  if data[0] == 80 && data[1] == 75 && data[2] == 3 && data[3] == 4 {
    // It's a ZIP file, need to check content
    // For now, assume XLSX (we can improve this later)
    return Xlsx
  }

  // Check for OLE2/CFB signature (XLS)
  if data[0] == 208 &&
    data[1] == 207 &&
    data[2] == 17 &&
    data[3] == 224 &&
    data[4] == 161 &&
    data[5] == 177 &&
    data[6] == 26 &&
    data[7] == 225 {
    return Xls
  }
  Unknown
}

///|
/// Auto-detect and open any supported workbook format
pub fn open_workbook_auto(data : Bytes) -> AutoWorkbook raise CalamineError {
  match detect_workbook_type(data) {
    Xlsx => AutoWorkbook::Xlsx(XlsxReader::new(data))
    Xls => AutoWorkbook::Xls(XlsReader::new(data))
    Xlsb => AutoWorkbook::Xlsb(XlsbReader::new(data))
    Ods => AutoWorkbook::Ods(OdsReader::new(data))
    Unknown => raise CalamineError::Msg("Unknown workbook format")
  }
}

///|
/// Enum that can hold any type of workbook reader
pub(all) enum AutoWorkbook {
  Xlsx(XlsxReader)
  Xls(XlsReader)
  Xlsb(XlsbReader)
  Ods(OdsReader)
} derive(Show, Eq)

///|
/// Forward Reader methods to the appropriate concrete reader
impl Reader for AutoWorkbook with new(data) {
  match detect_workbook_type(data) {
    Xlsx => Xlsx(XlsxReader::new(data))
    Xls => Xls(XlsReader::new(data))
    Xlsb => Xlsb(XlsbReader::new(data))
    Ods => Ods(OdsReader::new(data))
    Unknown => raise CalamineError::Msg("Unknown workbook format")
  }
}

///|
impl Reader for AutoWorkbook with with_header_row(self, header_row) {
  match self {
    Xlsx(reader) => Xlsx(reader.with_header_row(header_row))
    Xls(reader) => Xls(reader.with_header_row(header_row))
    Xlsb(reader) => Xlsb(reader.with_header_row(header_row))
    Ods(reader) => Ods(reader.with_header_row(header_row))
  }
}

///|
impl Reader for AutoWorkbook with vba_project(self) {
  match self {
    Xlsx(reader) => reader.vba_project()
    Xls(reader) => reader.vba_project()
    Xlsb(reader) => reader.vba_project()
    Ods(reader) => reader.vba_project()
  }
}

///|
impl Reader for AutoWorkbook with metadata(self) {
  match self {
    Xlsx(reader) => reader.metadata()
    Xls(reader) => reader.metadata()
    Xlsb(reader) => reader.metadata()
    Ods(reader) => reader.metadata()
  }
}

///|
impl Reader for AutoWorkbook with worksheet_range(self, name) {
  match self {
    Xlsx(reader) => reader.worksheet_range(name)
    Xls(reader) => reader.worksheet_range(name)
    Xlsb(reader) => reader.worksheet_range(name)
    Ods(reader) => reader.worksheet_range(name)
  }
}

///|
impl Reader for AutoWorkbook with worksheets(self) {
  match self {
    Xlsx(reader) => reader.worksheets()
    Xls(reader) => reader.worksheets()
    Xlsb(reader) => reader.worksheets()
    Ods(reader) => reader.worksheets()
  }
}

///|
impl Reader for AutoWorkbook with worksheet_formula(self, name) {
  match self {
    Xlsx(reader) => reader.worksheet_formula(name)
    Xls(reader) => reader.worksheet_formula(name)
    Xlsb(reader) => reader.worksheet_formula(name)
    Ods(reader) => reader.worksheet_formula(name)
  }
}

///|
impl Reader for AutoWorkbook with sheet_names(self) {
  match self {
    Xlsx(reader) => reader.sheet_names()
    Xls(reader) => reader.sheet_names()
    Xlsb(reader) => reader.sheet_names()
    Ods(reader) => reader.sheet_names()
  }
}

///|
impl Reader for AutoWorkbook with sheets_metadata(self) {
  match self {
    Xlsx(reader) => reader.sheets_metadata()
    Xls(reader) => reader.sheets_metadata()
    Xlsb(reader) => reader.sheets_metadata()
    Ods(reader) => reader.sheets_metadata()
  }
}

///|
impl Reader for AutoWorkbook with defined_names(self) {
  match self {
    Xlsx(reader) => reader.defined_names()
    Xls(reader) => reader.defined_names()
    Xlsb(reader) => reader.defined_names()
    Ods(reader) => reader.defined_names()
  }
}

///|
impl Reader for AutoWorkbook with worksheet_range_at(self, n) {
  match self {
    Xlsx(reader) => reader.worksheet_range_at(n)
    Xls(reader) => reader.worksheet_range_at(n)
    Xlsb(reader) => reader.worksheet_range_at(n)
    Ods(reader) => reader.worksheet_range_at(n)
  }
}

///|
/// Placeholder reader types (to be implemented)
pub(all) struct XlsxReader {
  data : Bytes
  header_row : HeaderRow
  metadata : Metadata
} derive(Show, Eq)

///|
pub(all) struct XlsReader {
  data : Bytes
  header_row : HeaderRow
  metadata : Metadata
} derive(Show, Eq)

///|
pub(all) struct XlsbReader {
  data : Bytes
  header_row : HeaderRow
  metadata : Metadata
} derive(Show, Eq)

///|
pub(all) struct OdsReader {
  data : Bytes
  header_row : HeaderRow
  metadata : Metadata
} derive(Show, Eq)

///|
/// Basic XlsxReader implementation
impl Reader for XlsxReader with new(data) {
  { data, header_row: FirstNonEmptyRow, metadata: Metadata::default() }
}

///|
impl Reader for XlsxReader with with_header_row(self, header_row) {
  { ..self, header_row, }
}

///|
impl Reader for XlsxReader with vba_project(self) {
  None
}

///|
impl Reader for XlsxReader with metadata(self) {
  self.metadata
}

///|
impl Reader for XlsxReader with worksheet_range(self, name) {
  // Placeholder implementation
  raise CalamineError::Msg("Not implemented yet")
}

///|
impl Reader for XlsxReader with worksheets(self) {
  []
}

///|
impl Reader for XlsxReader with worksheet_formula(self, name) {
  Range::empty()
}

///|
impl Reader for XlsxReader with sheet_names(self) {
  self.metadata().sheet_names()
}

///|
impl Reader for XlsxReader with sheets_metadata(self) {
  self.metadata().sheets
}

///|
impl Reader for XlsxReader with defined_names(self) {
  self.metadata().names
}

///|
impl Reader for XlsxReader with worksheet_range_at(self, n) {
  let sheet_names = self.sheet_names()
  if n >= 0 && n < sheet_names.length() {
    let name = sheet_names[n]
    Some(self.worksheet_range(name))
  } else {
    None
  }
}

///|
/// Basic XlsReader implementation  
impl Reader for XlsReader with new(data) {
  { data, header_row: FirstNonEmptyRow, metadata: Metadata::default() }
}

///|
impl Reader for XlsReader with with_header_row(self, header_row) {
  { ..self, header_row, }
}

///|
impl Reader for XlsReader with vba_project(self) {
  None
}

///|
impl Reader for XlsReader with metadata(self) {
  self.metadata
}

///|
impl Reader for XlsReader with worksheet_range(self, name) {
  raise CalamineError::Msg("Not implemented yet")
}

///|
impl Reader for XlsReader with worksheets(self) {
  []
}

///|
impl Reader for XlsReader with worksheet_formula(self, name) {
  Range::empty()
}

///|
impl Reader for XlsReader with sheet_names(self) {
  self.metadata().sheet_names()
}

///|
impl Reader for XlsReader with sheets_metadata(self) {
  self.metadata().sheets
}

///|
impl Reader for XlsReader with defined_names(self) {
  self.metadata().names
}

///|
impl Reader for XlsReader with worksheet_range_at(self, n) {
  let sheet_names = self.sheet_names()
  if n >= 0 && n < sheet_names.length() {
    let name = sheet_names[n]
    Some(self.worksheet_range(name))
  } else {
    None
  }
}

///|
/// Basic XlsbReader implementation
impl Reader for XlsbReader with new(data) {
  { data, header_row: FirstNonEmptyRow, metadata: Metadata::default() }
}

///|
impl Reader for XlsbReader with with_header_row(self, header_row) {
  { ..self, header_row, }
}

///|
impl Reader for XlsbReader with vba_project(self) {
  None
}

///|
impl Reader for XlsbReader with metadata(self) {
  self.metadata
}

///|
impl Reader for XlsbReader with worksheet_range(self, name) {
  raise CalamineError::Msg("Not implemented yet")
}

///|
impl Reader for XlsbReader with worksheets(self) {
  []
}

///|
impl Reader for XlsbReader with worksheet_formula(self, name) {
  Range::empty()
}

///|
impl Reader for XlsbReader with sheet_names(self) {
  self.metadata().sheet_names()
}

///|
impl Reader for XlsbReader with sheets_metadata(self) {
  self.metadata().sheets
}

///|
impl Reader for XlsbReader with defined_names(self) {
  self.metadata().names
}

///|
impl Reader for XlsbReader with worksheet_range_at(self, n) {
  let sheet_names = self.sheet_names()
  if n >= 0 && n < sheet_names.length() {
    let name = sheet_names[n]
    Some(self.worksheet_range(name))
  } else {
    None
  }
}

///|
/// Basic OdsReader implementation
impl Reader for OdsReader with new(data) {
  { data, header_row: FirstNonEmptyRow, metadata: Metadata::default() }
}

///|
impl Reader for OdsReader with with_header_row(self, header_row) {
  { ..self, header_row, }
}

///|
impl Reader for OdsReader with vba_project(self) {
  None
}

///|
impl Reader for OdsReader with metadata(self) {
  self.metadata
}

///|
impl Reader for OdsReader with worksheet_range(self, name) {
  raise CalamineError::Msg("Not implemented yet")
}

///|
impl Reader for OdsReader with worksheets(self) {
  []
}

///|
impl Reader for OdsReader with worksheet_formula(self, name) {
  Range::empty()
}

///|
impl Reader for OdsReader with sheet_names(self) {
  self.metadata().sheet_names()
}

///|
impl Reader for OdsReader with sheets_metadata(self) {
  self.metadata().sheets
}

///|
impl Reader for OdsReader with defined_names(self) {
  self.metadata().names
}

///|
impl Reader for OdsReader with worksheet_range_at(self, n) {
  let sheet_names = self.sheet_names()
  if n >= 0 && n < sheet_names.length() {
    let name = sheet_names[n]
    Some(self.worksheet_range(name))
  } else {
    None
  }
}
