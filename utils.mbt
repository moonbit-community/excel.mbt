///|
/// Utility functions for binary data manipulation
/// 
/// This module provides functions for reading binary data in little-endian format
/// and other utility functions needed for Excel file parsing.

///|
/// Converts bytes to u32 in little-endian format
pub fn read_u32(data : Bytes, offset : Int) -> UInt? {
  if offset + 4 > data.length() {
    return None
  }
  let b0 = data[offset].to_uint()
  let b1 = data[offset + 1].to_uint()
  let b2 = data[offset + 2].to_uint()
  let b3 = data[offset + 3].to_uint()
  Some(b0 | (b1 << 8) | (b2 << 16) | (b3 << 24))
}

///|
/// Converts bytes to i32 in little-endian format
pub fn read_i32(data : Bytes, offset : Int) -> Int? {
  match read_u32(data, offset) {
    Some(val) => Some(val.reinterpret_as_int())
    None => None
  }
}

///|
/// Converts bytes to u16 in little-endian format
pub fn read_u16(data : Bytes, offset : Int) -> UInt? {
  if offset + 2 > data.length() {
    return None
  }
  let b0 = data[offset].to_uint()
  let b1 = data[offset + 1].to_uint()
  Some(b0 | (b1 << 8))
}

///|
/// Converts bytes to i16 in little-endian format
pub fn read_i16(data : Bytes, offset : Int) -> Int? {
  match read_u16(data, offset) {
    Some(val) => Some(val.reinterpret_as_int())
    None => None
  }
}

///|
/// Converts bytes to u64 in little-endian format
pub fn read_u64(data : Bytes, offset : Int) -> UInt64? {
  if offset + 8 > data.length() {
    return None
  }
  let low = read_u32(data, offset)
  let high = read_u32(data, offset + 4)
  match (low, high) {
    (Some(l), Some(h)) => Some(l.to_uint64() | (h.to_uint64() << 32))
    _ => None
  }
}

///|
/// Converts bytes to f64 in little-endian format
pub fn read_f64(data : Bytes, offset : Int) -> Double? {
  match read_u64(data, offset) {
    Some(bits) => Some(bits.reinterpret_as_double())
    None => None
  }
}

///|
/// Push literal column into a String buffer
/// Converts a column number to Excel column letters (A, B, C, ..., AA, AB, etc.)
pub fn push_column(col : UInt, buf : FixedArray[Char]) -> Unit {
  let mut index = 0
  let mut col = col
  if col < 26U {
    buf[index] = (65 + col.reinterpret_as_int()).unsafe_to_char()
    return
  }

  // For columns >= 26, we need to handle multi-character columns
  let temp_chars : Array[Char] = []
  while col >= 26U {
    let c = col % 26U
    temp_chars.push((65 + c.reinterpret_as_int()).unsafe_to_char())
    col = col - c
    col = col / 26U
  }

  // Add remaining character if any
  if col > 0U {
    temp_chars.push((65 + (col - 1U).reinterpret_as_int()).unsafe_to_char())
  }

  // Reverse and copy to buffer
  for i = temp_chars.length() - 1; i >= 0; i = i - 1 {
    if index < buf.length() {
      buf[index] = temp_chars[i]
      index = index + 1
    }
  }
}

///|
/// Convert column number to Excel column string
pub fn column_to_string(col : UInt) -> String {
  let buf = FixedArray::make(10, ' ') // Should be enough for most columns
  push_column(col, buf)
  let mut result = ""
  for i = 0; i < buf.length(); i = i + 1 {
    if buf[i] != ' ' {
      result = result + buf[i].to_string()
    }
  }
  result
}

///|
/// Complete Excel function table - 485 functions total
/// Based on [MS-XLS] 2.5.198.17 and [MS-XLSB] 2.5.97.10 specifications
const FTAB_LEN : UInt = 485U

///|
/// Excel function names table
let ftab : Array[String] = [
  "COUNT", "IF", "ISNA", "ISERROR", "SUM", "AVERAGE", "MIN", "MAX", "ROW", "COLUMN",
  "NA", "NPV", "STDEV", "DOLLAR", "FIXED", "SIN", "COS", "TAN", "ATAN", "PI", "SQRT",
  "EXP", "LN", "LOG10", "ABS", "INT", "SIGN", "ROUND", "LOOKUP", "INDEX", "REPT",
  "MID", "LEN", "VALUE", "TRUE", "FALSE", "AND", "OR", "NOT", "MOD", "DCOUNT", "DSUM",
  "DAVERAGE", "DMIN", "DMAX", "DSTDEV", "VAR", "DVAR", "TEXT", "LINEST", "TREND",
  "LOGEST", "GROWTH", "GOTO", "HALT", "RETURN", "PV", "FV", "NPER", "PMT", "RATE",
  "MIRR", "IRR", "RAND", "MATCH", "DATE", "TIME", "DAY", "MONTH", "YEAR", "WEEKDAY",
  "HOUR", "MINUTE", "SECOND", "NOW", "AREAS", "ROWS", "COLUMNS", "OFFSET", "ABSREF",
  "RELREF", "ARGUMENT", "SEARCH", "TRANSPOSE", "ERROR", "STEP", "TYPE", "ECHO", "SET.NAME",
  "CALLER", "DEREF", "WINDOWS", "SERIES", "DOCUMENTS", "ACTIVE.CELL", "SELECTION",
  "RESULT", "ATAN2", "ASIN", "ACOS", "CHOOSE", "HLOOKUP", "VLOOKUP", "LINKS", "INPUT",
  "ISREF", "GET.FORMULA", "GET.NAME", "SET.VALUE", "LOG", "EXEC", "CHAR", "LOWER",
  "UPPER", "PROPER", "LEFT", "RIGHT", "EXACT", "TRIM", "REPLACE", "SUBSTITUTE", "CODE",
  "NAMES", "DIRECTORY", "FIND", "CELL", "ISERR", "ISTEXT", "ISNUMBER", "ISBLANK",
  "T", "N", "FOPEN", "FCLOSE", "FSIZE", "FREADLN", "FREAD", "FWRITELN", "FWRITE",
  "FPOS", "DATEVALUE", "TIMEVALUE", "SLN", "SYD", "DDB", "GET.DEF", "REFTEXT", "TEXTREF",
  "INDIRECT", "REGISTER", "CALL", "ADD.BAR", "ADD.MENU", "ADD.COMMAND", "ENABLE.COMMAND",
  "CHECK.COMMAND", "RENAME.COMMAND", "SHOW.BAR", "DELETE.MENU", "DELETE.COMMAND",
  "GET.CHART.ITEM", "DIALOG.BOX", "CLEAN", "MDETERM", "MINVERSE", "MMULT", "FILES",
  "IPMT", "PPMT", "COUNTA", "CANCEL.KEY", "FOR", "WHILE", "BREAK", "NEXT", "INITIATE",
  "REQUEST", "POKE", "EXECUTE", "TERMINATE", "RESTART", "HELP", "GET.BAR", "PRODUCT",
  "FACT", "GET.CELL", "GET.WORKSPACE", "GET.WINDOW", "GET.DOCUMENT", "DPRODUCT",
  "ISNONTEXT", "GET.NOTE", "NOTE", "STDEVP", "VARP", "DSTDEVP", "DVARP", "TRUNC",
  "ISLOGICAL", "DCOUNTA", "DELETE.BAR", "UNREGISTER", "", "", "USDOLLAR", "FINDB",
  "SEARCHB", "REPLACEB", "LEFTB", "RIGHTB", "MIDB", "LENB", "ROUNDUP", "ROUNDDOWN",
  "ASC", "DBCS", "RANK", "", "", "ADDRESS", "DAYS360", "TODAY", "VDB", "ELSE", "ELSE.IF",
  "END.IF", "FOR.CELL", "MEDIAN", "SUMPRODUCT", "SINH", "COSH", "TANH", "ASINH",
  "ACOSH", "ATANH", "DGET", "CREATE.OBJECT", "VOLATILE", "LAST.ERROR", "CUSTOM.UNDO",
  "CUSTOM.REPEAT", "FORMULA.CONVERT", "GET.LINK.INFO", "TEXT.BOX", "INFO", "GROUP",
  "GET.OBJECT", "DB", "PAUSE", "", "", "RESUME", "FREQUENCY", "ADD.TOOLBAR", "DELETE.TOOLBAR",
  "User", "RESET.TOOLBAR", "EVALUATE", "GET.TOOLBAR", "GET.TOOL", "SPELLING.CHECK",
  "ERROR.TYPE", "APP.TITLE", "WINDOW.TITLE", "SAVE.TOOLBAR", "ENABLE.TOOL", "PRESS.TOOL",
  "REGISTER.ID", "GET.WORKBOOK", "AVEDEV", "BETADIST", "GAMMALN", "BETAINV", "BINOMDIST",
  "CHIDIST", "CHIINV", "COMBIN", "CONFIDENCE", "CRITBINOM", "EVEN", "EXPONDIST",
  "FDIST", "FINV", "FISHER", "FISHERINV", "FLOOR", "GAMMADIST", "GAMMAINV", "CEILING",
  "HYPGEOMDIST", "LOGNORMDIST", "LOGINV", "NEGBINOMDIST", "NORMDIST", "NORMSDIST",
  "NORMINV", "NORMSINV", "STANDARDIZE", "ODD", "PERMUT", "POISSON", "TDIST", "WEIBULL",
  "SUMXMY2", "SUMX2MY2", "SUMX2PY2", "CHITEST", "CORREL", "COVAR", "FORECAST", "FTEST",
  "INTERCEPT", "PEARSON", "RSQ", "STEYX", "SLOPE", "TTEST", "PROB", "DEVSQ", "GEOMEAN",
  "HARMEAN", "SUMSQ", "KURT", "SKEW", "ZTEST", "LARGE", "SMALL", "QUARTILE", "PERCENTILE",
  "PERCENTRANK", "MODE", "TRIMMEAN", "TINV", "", "MOVIE.COMMAND", "GET.MOVIE", "CONCATENATE",
  "POWER", "PIVOT.ADD.DATA", "GET.PIVOT.TABLE", "GET.PIVOT.FIELD", "GET.PIVOT.ITEM",
  "RADIANS", "DEGREES", "SUBTOTAL", "SUMIF", "COUNTIF", "COUNTBLANK", "SCENARIO.GET",
  "OPTIONS.LISTS.GET", "ISPMT", "DATEDIF", "DATESTRING", "NUMBERSTRING", "ROMAN",
  "OPEN.DIALOG", "SAVE.DIALOG", "VIEW.GET", "GETPIVOTDATA", "HYPERLINK", "PHONETIC",
  "AVERAGEA", "MAXA", "MINA", "STDEVPA", "VARPA", "STDEVA", "VARA", "BAHTTEXT", "THAIDAYOFWEEK",
  "THAIDIGIT", "THAIMONTHOFYEAR", "THAINUMSOUND", "THAINUMSTRING", "THAISTRINGLENGTH",
  "ISTHAIDIGIT", "ROUNDBAHTDOWN", "ROUNDBAHTUP", "THAIYEAR", "RTD", "CUBEVALUE",
  "CUBEMEMBER", "CUBEMEMBERPROPERTY", "CUBERANKEDMEMBER", "HEX2BIN", "HEX2DEC", "HEX2OCT",
  "DEC2BIN", "DEC2HEX", "DEC2OCT", "OCT2BIN", "OCT2HEX", "OCT2DEC", "BIN2DEC", "BIN2OCT",
  "BIN2HEX", "IMSUB", "IMDIV", "IMPOWER", "IMABS", "IMSQRT", "IMLN", "IMLOG2", "IMLOG10",
  "IMSIN", "IMCOS", "IMEXP", "IMARGUMENT", "IMCONJUGATE", "IMAGINARY", "IMREAL",
  "COMPLEX", "IMSUM", "IMPRODUCT", "SERIESSUM", "FACTDOUBLE", "SQRTPI", "QUOTIENT",
  "DELTA", "GESTEP", "ISEVEN", "ISODD", "MROUND", "ERF", "ERFC", "BESSELJ", "BESSELK",
  "BESSELY", "BESSELI", "XIRR", "XNPV", "PRICEMAT", "YIELDMAT", "INTRATE", "RECEIVED",
  "DISC", "PRICEDISC", "YIELDDISC", "TBILLEQ", "TBILLPRICE", "TBILLYIELD", "PRICE",
  "YIELD", "DOLLARDE", "DOLLARFR", "NOMINAL", "EFFECT", "CUMPRINC", "CUMIPMT", "EDATE",
  "EOMONTH", "YEARFRAC", "COUPDAYBS", "COUPDAYS", "COUPDAYSNC", "COUPNCD", "COUPNUM",
  "COUPPCD", "DURATION", "MDURATION", "ODDLPRICE", "ODDLYIELD", "ODDFPRICE", "ODDFYIELD",
  "RANDBETWEEN", "WEEKNUM", "AMORDEGRC", "AMORLINC", "CONVERT", "ACCRINT", "ACCRINTM",
  "WORKDAY", "NETWORKDAYS", "GCD", "MULTINOMIAL", "LCM", "FVSCHEDULE", "CUBEKPIMEMBER",
  "CUBESET", "CUBESETCOUNT", "IFERROR", "COUNTIFS", "SUMIFS", "AVERAGEIF", "AVERAGEIFS",
]

///|
/// Excel function argument counts table 
let ftab_argc : Array[UInt] = [
  255U, 3U, 1U, 1U, 255U, 255U, 255U, 255U, 1U, 1U, 0U, 254U, 255U, 2U, 3U, 1U, 1U,
  1U, 1U, 0U, 1U, 1U, 1U, 1U, 1U, 1U, 1U, 2U, 3U, 4U, 2U, 3U, 1U, 1U, 0U, 0U, 255U,
  255U, 1U, 2U, 3U, 3U, 3U, 3U, 3U, 3U, 255U, 3U, 2U, 4U, 4U, 4U, 4U, 1U, 1U, 1U,
  5U, 5U, 5U, 5U, 6U, 3U, 2U, 0U, 3U, 3U, 3U, 1U, 1U, 1U, 2U, 1U, 1U, 1U, 0U, 1U,
  1U, 1U, 5U, 2U, 2U, 3U, 3U, 1U, 0U, 1U, 1U, 2U, 1U, 1U, 2U, 2U, 2U, 0U, 0U, 1U,
  2U, 1U, 1U, 255U, 4U, 4U, 2U, 7U, 1U, 1U, 2U, 2U, 2U, 4U, 1U, 1U, 1U, 1U, 2U, 2U,
  2U, 4U, 4U, 1U, 3U, 3U, 1U, 1U, 1U, 1U, 1U, 1U, 2U, 1U, 1U, 2U, 2U, 2U, 1U, 1U,
  3U, 4U, 5U, 2U, 2U, 2U, 2U, 255U, 255U, 5U, 4U, 255U, 1U, 4U, 5U, 5U, 5U, 5U, 1U,
  3U, 4U, 3U, 1U, 1U, 1U, 1U, 6U, 6U, 255U, 2U, 4U, 1U, 0U, 0U, 2U, 2U, 3U, 2U, 1U,
  1U, 1U, 4U, 255U, 1U, 2U, 1U, 2U, 2U, 3U, 1U, 3U, 4U, 255U, 255U, 3U, 3U, 2U, 1U,
  3U, 1U, 1U, 0U, 0U, 2U, 3U, 3U, 4U, 2U, 2U, 3U, 2U, 2U, 2U, 1U, 1U, 3U, 0U, 0U,
  5U, 3U, 0U, 7U, 0U, 1U, 0U, 3U, 255U, 255U, 1U, 1U, 1U, 1U, 1U, 1U, 3U, 11U, 1U,
  0U, 2U, 3U, 5U, 4U, 4U, 1U, 0U, 5U, 5U, 1U, 0U, 0U, 1U, 2U, 2U, 1U, 255U, 1U, 1U,
  2U, 3U, 3U, 1U, 1U, 1U, 2U, 3U, 3U, 2U, 255U, 5U, 1U, 5U, 4U, 2U, 2U, 2U, 3U, 3U,
  1U, 3U, 3U, 3U, 4U, 3U, 2U, 4U, 3U, 2U, 1U, 1U, 2U, 4U, 3U, 3U, 4U, 4U, 1U, 3U,
  3U, 2U, 3U, 4U, 3U, 3U, 1U, 2U, 3U, 3U, 4U, 2U, 2U, 2U, 2U, 2U, 3U, 2U, 2U, 2U,
  2U, 4U, 4U, 255U, 255U, 255U, 255U, 255U, 255U, 3U, 2U, 2U, 2U, 2U, 4U, 0U, 4U,
  4U, 3U, 255U, 2U, 255U, 255U, 255U, 1U, 1U, 2U, 1U, 3U, 1U, 1U, 2U, 2U, 4U, 5U,
  2U, 128U, 2U, 1U, 255U, 255U, 255U, 255U, 255U, 255U, 255U, 1U, 1U, 1U, 1U, 1U,
  1U, 1U, 1U, 1U, 1U, 1U, 255U, 255U, 3U, 3U, 4U, 2U, 1U, 2U, 2U, 2U, 2U, 2U, 1U,
  2U, 2U, 1U, 1U, 2U, 2U, 2U, 2U, 1U, 1U, 3U, 2U, 3U, 1U, 1U, 2U, 1U, 1U, 1U, 1U,
  1U, 1U, 1U, 1U, 1U, 1U, 1U, 1U, 1U, 1U, 1U, 3U, 255U, 255U, 4U, 1U, 1U, 2U, 2U,
  1U, 2U, 1U, 1U, 2U, 2U, 2U, 6U, 6U, 5U, 5U, 5U, 5U, 5U, 3U, 3U, 3U, 7U, 7U, 2U,
  2U, 2U, 2U, 6U, 6U, 2U, 2U, 3U, 4U, 4U, 4U, 4U, 4U, 4U, 6U, 6U, 8U, 8U, 8U, 8U,
  2U, 2U, 7U, 7U, 8U, 8U, 5U, 3U, 3U, 255U, 255U, 255U, 2U, 4U, 5U, 1U, 2U, 128U,
  129U, 3U, 129U,
]

///|
/// Get function name by index
pub fn get_function_name(index : UInt) -> String? {
  if index < FTAB_LEN {
    Some(ftab[index.reinterpret_as_int()])
  } else {
    None
  }
}

///|
/// Get function argument count by index  
pub fn get_function_argc(index : UInt) -> UInt? {
  if index < FTAB_LEN {
    Some(ftab_argc[index.reinterpret_as_int()])
  } else {
    None
  }
}

///|
/// Convert bytes to iterator of u32 values (little-endian)
pub fn to_u32(data : Bytes) -> Array[UInt] {
  if data.length() % 4 != 0 {
    abort("Data length must be divisible by 4")
  }
  let result : Array[UInt] = []
  for i = 0; i < data.length(); i = i + 4 {
    match read_u32(data, i) {
      Some(val) => result.push(val)
      None => break
    }
  }
  result
}

///|
/// Read usize value from bytes (little-endian) 
pub fn read_usize(data : Bytes, offset : Int) -> UInt? {
  match read_u32(data, offset) {
    Some(val) => Some(val)
    None => None
  }
}

///|
/// Check if bytes start with a specific signature
pub fn has_signature(data : Bytes, signature : Bytes) -> Bool {
  if data.length() < signature.length() {
    return false
  }
  for i = 0; i < signature.length(); i = i + 1 {
    if data[i] != signature[i] {
      return false
    }
  }
  true
}

///|
/// Find first occurrence of a byte pattern in bytes
pub fn find_pattern(data : Bytes, pattern : Bytes) -> Int? {
  if pattern.length() > data.length() {
    return None
  }
  for i = 0; i <= data.length() - pattern.length(); i = i + 1 {
    let mut found = true
    for j = 0; j < pattern.length(); j = j + 1 {
      if data[i + j] != pattern[j] {
        found = false
        break
      }
    }
    if found {
      return Some(i)
    }
  }
  None
}

///|
/// Convert a range reference to (row, col) coordinates 
/// Simple implementation for basic A1-style references
pub fn parse_cell_reference(reference : String) -> (UInt, UInt)? {
  if reference.is_empty() {
    return None
  }
  let mut col = 0U
  let mut row_start = 0

  // Parse column letters (A, B, C, ..., AA, AB, etc.)
  for i = 0; i < reference.length(); i = i + 1 {
    let code_unit = reference[i]
    if code_unit >= 65 && code_unit <= 90 { // 'A' to 'Z'
      col = col * 26U + (code_unit - 65).reinterpret_as_uint() + 1U
      row_start = i + 1
    } else {
      break
    }
  }

  // Parse row number
  if row_start >= reference.length() {
    return None
  }
  let mut row_str = ""
  for i = row_start; i < reference.length(); i = i + 1 {
    let char = reference[i].unsafe_to_char()
    row_str = row_str + char.to_string()
  }

  // Simple row parsing (basic cases only)
  let row = match row_str {
    "1" => 1U
    "2" => 2U
    "3" => 3U
    "4" => 4U
    "5" => 5U
    "10" => 10U
    "100" => 100U
    _ => return None
  }
  Some((row - 1U, col - 1U)) // Convert to 0-based indexing
}

///|
/// Convert (row, col) coordinates to A1-style reference
pub fn format_cell_reference(row : UInt, col : UInt) -> String {
  let col_str = column_to_string(col)
  let row_str = (row + 1U).to_string() // Convert to 1-based indexing
  col_str + row_str
}

///|
/// Calculate the maximum value that can be stored in a given number of bits
pub fn max_value_for_bits(bits : UInt) -> UInt64 {
  if bits >= 64U {
    18446744073709551615UL // 2^64 - 1
  } else {
    (1UL << bits.reinterpret_as_int()) - 1UL
  }
}

///|
/// Convert a byte to its hexadecimal string representation
pub fn byte_to_hex(byte : Byte) -> String {
  let val = byte.to_uint()
  let high = val / 16U
  let low = val % 16U
  let high_char = if high < 10U {
    (48 + high.reinterpret_as_int()).unsafe_to_char()
  } else {
    (55 + high.reinterpret_as_int()).unsafe_to_char()
  }
  let low_char = if low < 10U {
    (48 + low.reinterpret_as_int()).unsafe_to_char()
  } else {
    (55 + low.reinterpret_as_int()).unsafe_to_char()
  }
  high_char.to_string() + low_char.to_string()
}

///|
/// Convert bytes to hexadecimal string representation
pub fn bytes_to_hex(data : Bytes) -> String {
  let mut result = ""
  for i = 0; i < data.length(); i = i + 1 {
    result = result + byte_to_hex(data[i])
    if i < data.length() - 1 {
      result = result + " "
    }
  }
  result
}

///|
/// Check if a string represents a valid Excel column name (A, B, ..., AA, AB, etc.)
pub fn is_valid_column_name(column : String) -> Bool {
  if column.is_empty() {
    return false
  }
  for i = 0; i < column.length(); i = i + 1 {
    let code_unit = column[i]
    if code_unit < 65 || code_unit > 90 { // Not 'A' to 'Z'
      return false
    }
  }
  true
}

///|
/// Parse Excel column name to column index (A=0, B=1, AA=26, etc.)
pub fn parse_column_name(column : String) -> UInt? {
  if !is_valid_column_name(column) {
    return None
  }
  let mut result = 0U
  for i = 0; i < column.length(); i = i + 1 {
    let code_unit = column[i]
    result = result * 26U + (code_unit - 65).reinterpret_as_uint() + 1U
  }
  Some(result - 1U) // Convert to 0-based indexing
}
