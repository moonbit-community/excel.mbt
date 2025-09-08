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

  /// Load tables in the workbook (for formats that support it)
  load_tables(Self) -> Unit raise CalamineError

  /// Get table by name (for formats that support it)
  table_by_name(Self, String) -> Table[Data]? raise CalamineError

  /// Get all table names in the workbook
  table_names(Self) -> Array[String]

  /// Get table names in a specific sheet
  table_names_in_sheet(Self, String) -> Array[String]
}

///|
/// A trait for readers that support reference-based data access
pub trait ReaderRef: Reader {
  /// Get worksheet range where shared string values are only borrowed.
  worksheet_range_ref(Self, String) -> Range[DataRef] raise CalamineError

  /// Get the nth worksheet range where shared string values are only borrowed.
  worksheet_range_at_ref(Self, Int) -> Range[DataRef]? raise CalamineError

  /// Get table by name where shared string values are only borrowed.
  table_by_name_ref(Self, String) -> Table[DataRef]? raise CalamineError
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
    // It's a ZIP file, need to check content type
    detect_zip_based_format(data)

    // Check for OLE2/CFB signature (XLS)  
  } else if data[0] == 208 &&
    data[1] == 207 &&
    data[2] == 17 &&
    data[3] == 224 &&
    data[4] == 161 &&
    data[5] == 177 &&
    data[6] == 26 &&
    data[7] == 225 {
    Xls
  } else {
    Unknown
  }
}

///|
/// Detect ZIP-based format by examining content
fn detect_zip_based_format(data : Bytes) -> WorkbookType {
  // Try to detect the format by looking for specific patterns or filenames
  // This is a simplified implementation - a full implementation would parse the ZIP

  // Look for typical file patterns in the data
  let data_str = data.to_string()

  // Check for ODS MIME type
  if data_str.contains("application/vnd.oasis.opendocument.spreadsheet") {
    return Ods
  }

  // Check for XLSX patterns
  if data_str.contains("xl/") ||
    data_str.contains("xl/workbook.xml") ||
    data_str.contains("xl/worksheets/") {
    return Xlsx
  }

  // Check for XLSB patterns
  if data_str.contains("xl/workbook.bin") {
    return Xlsb
  }

  // Check for ODS patterns
  if data_str.contains("content.xml") &&
    data_str.contains("META-INF/manifest.xml") {
    return Ods
  }

  // Default to XLSX for unknown ZIP files (most common)
  Xlsx
}

///|
/// Enhanced detection with file extension hints
pub fn detect_workbook_type_with_extension(
  data : Bytes,
  filename : String?,
) -> WorkbookType {
  // First try detection by content
  let detected = detect_workbook_type(data)

  // If we couldn't determine from content, use filename
  if detected == Unknown {
    match filename {
      Some(name) => {
        let name_lower = name // MoonBit doesn't have to_lowercase, use as-is for now
        if name_lower.ends_with(".xlsx") is Some(_) ||
          name_lower.ends_with(".xlsm") is Some(_) {
          return Xlsx
        } else if name_lower.ends_with(".xls") is Some(_) ||
          name_lower.ends_with(".xla") is Some(_) {
          return Xls
        } else if name_lower.ends_with(".xlsb") is Some(_) {
          return Xlsb
        } else if name_lower.ends_with(".ods") is Some(_) {
          return Ods
        }
        Unknown
      }
      None => Unknown
    }
  } else {
    detected
  }
}

///|
/// Auto-detect and open any supported workbook format
pub fn open_workbook_auto(data : Bytes) -> AutoWorkbook raise CalamineError {
  match detect_workbook_type(data) {
    Xlsx => AutoWorkbook::Xlsx(XlsxReader::new(data))
    Xls => AutoWorkbook::Xls(XlsReader::new(data))
    Xlsb => AutoWorkbook::Xlsb(XlsbReader::new(data))
    Ods => AutoWorkbook::Ods(OdsReader::new(data))
    Unknown =>
      raise CalamineError::Msg(
        "Unknown workbook format - unable to detect file type",
      )
  }
}

///|
/// Auto-detect and open any supported workbook format with filename hint
pub fn open_workbook_auto_with_filename(
  data : Bytes,
  filename : String?,
) -> AutoWorkbook raise CalamineError {
  match detect_workbook_type_with_extension(data, filename) {
    Xlsx => AutoWorkbook::Xlsx(XlsxReader::new(data))
    Xls => AutoWorkbook::Xls(XlsReader::new(data))
    Xlsb => AutoWorkbook::Xlsb(XlsbReader::new(data))
    Ods => AutoWorkbook::Ods(OdsReader::new(data))
    Unknown =>
      raise CalamineError::Msg(
        "Unknown workbook format - unable to detect file type even with filename hint",
      )
  }
}

///|
impl Reader for OdsReader with load_tables(self) {
  // ODS format doesn't support Excel-style tables
  return
}

///|
impl Reader for OdsReader with table_by_name(self, name) {
  // ODS format doesn't support Excel-style tables
  None
}

///|
impl Reader for OdsReader with table_names(self) {
  // ODS format doesn't support Excel-style tables
  []
}

///|
impl Reader for OdsReader with table_names_in_sheet(self, sheet_name) {
  // ODS format doesn't support Excel-style tables
  []
}

///|
impl Reader for XlsbReader with load_tables(self) {
  // For now, no-op - real implementation would parse tables from XLSB
  return
}

///|
impl Reader for XlsbReader with table_by_name(self, name) {
  // For now, return None - real implementation would lookup table by name
  None
}

///|
impl Reader for XlsbReader with table_names(self) {
  // For now, return empty array - real implementation would return all table names
  []
}

///|
impl Reader for XlsbReader with table_names_in_sheet(self, sheet_name) {
  // For now, return empty array - real implementation would return table names in sheet
  []
}

///|
/// XlsbReader implements ReaderRef for efficient reference-based data access  
impl ReaderRef for XlsbReader with worksheet_range_ref(self, name) {
  // For now, convert Data to DataRef - real implementation would avoid copying strings
  let data_range = self.worksheet_range(name)
  let cells = data_range
    .cells()
    .map(fn(cell) {
      let (row, col, data) = cell
      let data_ref = match data {
        String(s) => SharedString(s)
        Int(i) => Int(i)
        Float(f) => Float(f)
        Bool(b) => Bool(b)
        DateTime(dt) => DateTime(dt)
        DateTimeIso(s) => DateTimeIso(s)
        DurationIso(s) => DurationIso(s)
        Error(e) => Error(e)
        Empty => Empty
      }
      Cell::new((row, col), data_ref)
    })
  Range::from_sparse(cells)
}

///|
impl Reader for AutoWorkbook with new(data) {
  match detect_workbook_type(data) {
    Xlsx => Xlsx(XlsxReader::new(data))
    Xls => Xls(XlsReader::new(data))
    Xlsb => Xlsb(XlsbReader::new(data))
    Ods => Ods(OdsReader::new(data))
    Unknown =>
      raise CalamineError::Msg(
        "Unknown workbook format - unable to detect file type from content",
      )
  }
}

///|
impl ReaderRef for XlsbReader with worksheet_range_at_ref(self, n) {
  let sheet_names = self.sheet_names()
  if n >= 0 && n < sheet_names.length() {
    let name = sheet_names[n]
    Some(self.worksheet_range_ref(name))
  } else {
    None
  }
}

///|
impl ReaderRef for XlsbReader with table_by_name_ref(self, name) {
  // For now, return None - real implementation would return Table[DataRef]
  None
}

///|
impl Reader for XlsReader with load_tables(self) {
  // XLS format doesn't support tables
  return
}

///|
impl Reader for XlsReader with table_by_name(self, name) {
  // XLS format doesn't support tables
  None
}

///|
impl Reader for XlsReader with table_names(self) {
  // XLS format doesn't support tables
  []
}

///|
impl Reader for XlsReader with table_names_in_sheet(self, sheet_name) {
  // XLS format doesn't support tables
  []
}

///|
impl Reader for XlsxReader with load_tables(self) {
  // For now, no-op - real implementation would parse tables from XLSX
  return
}

///|
impl Reader for XlsxReader with table_by_name(self, name) {
  // For now, return None - real implementation would lookup table by name
  None
}

///|
impl Reader for XlsxReader with table_names(self) {
  // For now, return empty array - real implementation would return all table names
  []
}

///|
impl Reader for XlsxReader with table_names_in_sheet(self, sheet_name) {
  // For now, return empty array - real implementation would return table names in sheet
  []
}

///|
/// XlsxReader implements ReaderRef for efficient reference-based data access
impl ReaderRef for XlsxReader with worksheet_range_ref(self, name) {
  // For now, convert Data to DataRef - real implementation would avoid copying strings
  let data_range = self.worksheet_range(name)
  let cells = data_range
    .cells()
    .map(fn(cell) {
      let (row, col, data) = cell
      let data_ref = match data {
        String(s) => SharedString(s)
        Int(i) => Int(i)
        Float(f) => Float(f)
        Bool(b) => Bool(b)
        DateTime(dt) => DateTime(dt)
        DateTimeIso(s) => DateTimeIso(s)
        DurationIso(s) => DurationIso(s)
        Error(e) => Error(e)
        Empty => Empty
      }
      Cell::new((row, col), data_ref)
    })
  Range::from_sparse(cells)
}

///|
impl ReaderRef for XlsxReader with worksheet_range_at_ref(self, n) {
  let sheet_names = self.sheet_names()
  if n >= 0 && n < sheet_names.length() {
    let name = sheet_names[n]
    Some(self.worksheet_range_ref(name))
  } else {
    None
  }
}

///|
impl ReaderRef for XlsxReader with table_by_name_ref(self, name) {
  // For now, return None - real implementation would return Table[DataRef]
  None
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
impl Reader for AutoWorkbook with load_tables(self) {
  match self {
    Xlsx(reader) => reader.load_tables()
    Xls(reader) => reader.load_tables()
    Xlsb(reader) => reader.load_tables()
    Ods(reader) => reader.load_tables()
  }
}

///|
impl Reader for AutoWorkbook with table_by_name(self, name) {
  match self {
    Xlsx(reader) => reader.table_by_name(name)
    Xls(reader) => reader.table_by_name(name)
    Xlsb(reader) => reader.table_by_name(name)
    Ods(reader) => reader.table_by_name(name)
  }
}

///|
impl Reader for AutoWorkbook with table_names(self) {
  match self {
    Xlsx(reader) => reader.table_names()
    Xls(reader) => reader.table_names()
    Xlsb(reader) => reader.table_names()
    Ods(reader) => reader.table_names()
  }
}

///|
impl Reader for AutoWorkbook with table_names_in_sheet(self, sheet_name) {
  match self {
    Xlsx(reader) => reader.table_names_in_sheet(sheet_name)
    Xls(reader) => reader.table_names_in_sheet(sheet_name)
    Xlsb(reader) => reader.table_names_in_sheet(sheet_name)
    Ods(reader) => reader.table_names_in_sheet(sheet_name)
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
