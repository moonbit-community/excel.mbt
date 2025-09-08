///|
/// An enum to represent all different errors that can appear as
/// a value in a worksheet cell
pub(all) enum CellErrorType {
  /// Division by 0 error
  Div0
  /// Unavailable value error
  NA
  /// Invalid name error
  Name
  /// Null value error
  Null
  /// Number error
  Num
  /// Invalid cell reference error
  Ref
  /// Value error
  Value
  /// Getting data
  GettingData
} derive(Show, Eq)

///|
/// Convert CellErrorType to string representation
pub fn CellErrorType::to_string(self : CellErrorType) -> String {
  match self {
    Div0 => "#DIV/0!"
    NA => "#N/A"
    Name => "#NAME?"
    Null => "#NULL!"
    Num => "#NUM!"
    Ref => "#REF!"
    Value => "#VALUE!"
    GettingData => "#DATA!"
  }
}

///|
/// Dimensions info
pub(all) struct Dimensions {
  /// start: (row, col)
  start : (UInt, UInt)
  /// end: (row, col)
  end : (UInt, UInt)
} derive(Show, Eq)

///|
pub fn Dimensions::new(start : (UInt, UInt), end : (UInt, UInt)) -> Dimensions {
  { start, end }
}

///|
pub fn Dimensions::contains(self : Dimensions, row : UInt, col : UInt) -> Bool {
  row >= self.start.0 &&
  row <= self.end.0 &&
  col >= self.start.1 &&
  col <= self.end.1
}

///|
pub fn Dimensions::len(self : Dimensions) -> UInt64 {
  let height = self.end.0 - self.start.0 + 1U
  let width = self.end.1 - self.start.1 + 1U
  height.to_uint64() * width.to_uint64()
}

///|
pub fn Dimensions::width(self : Dimensions) -> UInt {
  self.end.1 - self.start.1 + 1U
}

///|
pub fn Dimensions::height(self : Dimensions) -> UInt {
  self.end.0 - self.start.0 + 1U
}

///|
/// Type of sheet.
pub(all) enum SheetType {
  /// A worksheet.
  WorkSheet
  /// A dialog sheet.
  DialogSheet
  /// A macro sheet.
  MacroSheet
  /// A chartsheet.
  ChartSheet
  /// A VBA module.
  Vba
} derive(Show, Eq)

///|
/// Type of visible sheet.
pub(all) enum SheetVisible {
  /// Visible
  Visible
  /// Hidden
  Hidden
  /// The sheet is hidden and cannot be displayed using the user interface. 
  VeryHidden
} derive(Show, Eq)

///|
/// Metadata of sheet
pub(all) struct Sheet {
  /// Name
  name : String
  /// Type
  typ : SheetType
  /// Visible
  visible : SheetVisible
} derive(Show, Eq)

///|
/// Excel datetime type. Possible: date, time, datetime, duration.
pub(all) enum ExcelDateTimeType {
  /// DateTime
  DateTime
  /// TimeDelta (Duration)
  TimeDelta
} derive(Show, Eq)

///|
/// Structure for Excel date and time representation.
pub(all) struct ExcelDateTime {
  value : Double
  datetime_type : ExcelDateTimeType
  is_1904 : Bool
} derive(Show, Eq)

///|
pub fn ExcelDateTime::new(
  value : Double,
  datetime_type : ExcelDateTimeType,
  is_1904 : Bool,
) -> ExcelDateTime {
  { value, datetime_type, is_1904 }
}

///|
pub fn ExcelDateTime::is_duration(self : ExcelDateTime) -> Bool {
  match self.datetime_type {
    TimeDelta => true
    DateTime => false
  }
}

///|
pub fn ExcelDateTime::is_datetime(self : ExcelDateTime) -> Bool {
  match self.datetime_type {
    DateTime => true
    TimeDelta => false
  }
}

///|
pub fn ExcelDateTime::as_f64(self : ExcelDateTime) -> Double {
  self.value
}

///|
/// Constants for Excel date/time conversions
const EXCEL_EPOCH_DAYS : Double = 25569.0 // Days from Excel epoch (1900-01-01) to Unix epoch (1970-01-01)

///|
const EXCEL_1900_1904_DIFF : Double = 1462.0 // Difference between 1900 and 1904 date systems

///|
const MS_PER_DAY : Double = 86400000.0 // Milliseconds per day (24 * 60 * 60 * 1000)

///|
/// Convert Excel datetime to date string representation (ISO format YYYY-MM-DD)
pub fn ExcelDateTime::as_date(self : ExcelDateTime) -> String {
  let adjusted_value = if self.is_1904 {
    self.value + EXCEL_1900_1904_DIFF
  } else {
    self.value
  }

  // Excel epoch starts at 1900-01-01, but has a bug treating 1900 as leap year
  // Values >= 60 need adjustment for the leap year bug
  let final_value = if adjusted_value >= 60.0 {
    adjusted_value
  } else {
    adjusted_value + 1.0
  }

  // Convert to Unix timestamp then to date
  let unix_days = final_value - EXCEL_EPOCH_DAYS
  let unix_timestamp_ms = unix_days * MS_PER_DAY

  // For now, return a simple formatted date
  // In a full implementation, this would use proper date libraries
  if unix_timestamp_ms >= 0.0 {
    let days_since_unix = unix_timestamp_ms / MS_PER_DAY
    let year = 1970 + (days_since_unix / 365.25).to_int()
    let remaining_days = (days_since_unix % 365.25).to_int().to_double()
    let month = (remaining_days / 30.4).to_int() + 1
    let day = (remaining_days % 30.4).to_int() + 1
    let month_str = if month < 10 { "0\{month}" } else { "\{month}" }
    let day_str = if day < 10 { "0\{day}" } else { "\{day}" }
    "\{year}-\{month_str}-\{day_str}"
  } else {
    "1900-01-01"
  }
}

///|
/// Convert Excel datetime to time string representation (ISO format HH:MM:SS)
pub fn ExcelDateTime::as_time(self : ExcelDateTime) -> String {
  // Get the fractional part representing time of day
  let time_fraction = self.value - self.value.floor()

  // Convert to seconds
  let total_seconds = (time_fraction * 86400.0).to_int() // 86400 seconds in a day
  let hours = total_seconds / 3600
  let minutes = total_seconds % 3600 / 60
  let seconds = total_seconds % 60
  let hours_str = if hours < 10 { "0\{hours}" } else { "\{hours}" }
  let minutes_str = if minutes < 10 { "0\{minutes}" } else { "\{minutes}" }
  let seconds_str = if seconds < 10 { "0\{seconds}" } else { "\{seconds}" }
  "\{hours_str}:\{minutes_str}:\{seconds_str}"
}

///|
/// Convert Excel datetime to duration (ISO 8601 format)
pub fn ExcelDateTime::as_duration(self : ExcelDateTime) -> String? {
  if self.is_duration() {
    let total_seconds = (self.value * 86400.0).to_int()
    let hours = total_seconds / 3600
    let minutes = total_seconds % 3600 / 60
    let seconds = total_seconds % 60
    if hours > 0 {
      Some("PT\{hours}H\{minutes}M\{seconds}S")
    } else if minutes > 0 {
      Some("PT\{minutes}M\{seconds}S")
    } else {
      Some("PT\{seconds}S")
    }
  } else {
    None
  }
}

///|
/// Convert Excel datetime to datetime string representation (ISO format YYYY-MM-DDTHH:MM:SS)
pub fn ExcelDateTime::as_datetime(self : ExcelDateTime) -> String {
  let date_part = self.as_date()
  let time_part = self.as_time()
  "\{date_part}T\{time_part}"
}

///|
/// An enum to represent all different data types that can appear as
/// a value in a worksheet cell
pub(all) enum Data {
  /// Signed integer
  Int(Int64)
  /// Float
  Float(Double)
  /// String
  String(String)
  /// Boolean
  Bool(Bool)
  /// Date or Time
  DateTime(ExcelDateTime)
  /// Date, Time or Date/Time in ISO 8601
  DateTimeIso(String)
  /// Duration in ISO 8601
  DurationIso(String)
  /// Error
  Error(CellErrorType)
  /// Empty cell
  Empty
} derive(Show, Eq)

///|
/// An enum to represent all different data types that can appear as
/// a value in a worksheet cell with reference semantics
pub(all) enum DataRef {
  /// Signed integer
  Int(Int64)
  /// Float
  Float(Double)
  /// String
  String(String)
  /// Shared String reference
  SharedString(String)
  /// Boolean
  Bool(Bool)
  /// Date or Time
  DateTime(ExcelDateTime)
  /// Date, Time or Date/Time in ISO 8601
  DateTimeIso(String)
  /// Duration in ISO 8601
  DurationIso(String)
  /// Error
  Error(CellErrorType)
  /// Empty cell
  Empty
} derive(Show, Eq)

///|
/// Data type methods for Data
pub fn Data::is_empty(self : Data) -> Bool {
  match self {
    Empty => true
    _ => false
  }
}

///|
pub fn Data::is_int(self : Data) -> Bool {
  match self {
    Int(_) => true
    _ => false
  }
}

///|
pub fn Data::is_float(self : Data) -> Bool {
  match self {
    Float(_) => true
    _ => false
  }
}

///|
pub fn Data::is_bool(self : Data) -> Bool {
  match self {
    Bool(_) => true
    _ => false
  }
}

///|
pub fn Data::is_string(self : Data) -> Bool {
  match self {
    String(_) => true
    _ => false
  }
}

///|
pub fn Data::is_error(self : Data) -> Bool {
  match self {
    Error(_) => true
    _ => false
  }
}

///|
pub fn Data::get_string(self : Data) -> String? {
  match self {
    String(v) => Some(v)
    _ => None
  }
}

///|
pub fn Data::get_int(self : Data) -> Int64? {
  match self {
    Int(v) => Some(v)
    _ => None
  }
}

///|
pub fn Data::get_float(self : Data) -> Double? {
  match self {
    Float(v) => Some(v)
    _ => None
  }
}

///|
pub fn Data::get_bool(self : Data) -> Bool? {
  match self {
    Bool(v) => Some(v)
    _ => None
  }
}

///|
pub fn Data::is_datetime(self : Data) -> Bool {
  match self {
    DateTime(_) => true
    _ => false
  }
}

///|
pub fn Data::is_datetime_iso(self : Data) -> Bool {
  match self {
    DateTimeIso(_) => true
    _ => false
  }
}

///|
pub fn Data::is_duration_iso(self : Data) -> Bool {
  match self {
    DurationIso(_) => true
    _ => false
  }
}

///|
pub fn Data::get_datetime(self : Data) -> ExcelDateTime? {
  match self {
    DateTime(v) => Some(v)
    _ => None
  }
}

///|
pub fn Data::get_datetime_iso(self : Data) -> String? {
  match self {
    DateTimeIso(v) => Some(v)
    _ => None
  }
}

///|
pub fn Data::get_duration_iso(self : Data) -> String? {
  match self {
    DurationIso(v) => Some(v)
    _ => None
  }
}

///|
pub fn Data::get_error(self : Data) -> CellErrorType? {
  match self {
    Error(e) => Some(e)
    _ => None
  }
}

///|
/// Convert data to string representation
pub fn Data::as_string(self : Data) -> String? {
  match self {
    Float(v) => Some(v.to_string())
    Int(v) => Some(v.to_string())
    String(v) => Some(v)
    Bool(v) => Some(v.to_string())
    DateTime(v) => Some(v.value.to_string())
    DateTimeIso(v) => Some(v)
    DurationIso(v) => Some(v)
    Error(e) => Some(e.to_string())
    Empty => Some("")
  }
}

///|
pub fn Data::as_i64(self : Data) -> Int64? {
  match self {
    Int(v) => Some(v)
    Float(v) => Some(v.to_int64())
    Bool(true) => Some(1L)
    Bool(false) => Some(0L)
    String(v) => parse_int64(v)
    _ => None
  }
}

///|
pub fn Data::as_f64(self : Data) -> Double? {
  match self {
    Int(v) => Some(v.to_double())
    Float(v) => Some(v)
    Bool(true) => Some(1.0)
    Bool(false) => Some(0.0)
    String(v) => parse_double(v)
    _ => None
  }
}

///|
/// Parse a string as Int64 with more robust parsing
pub fn parse_int64(s : String) -> Int64? {
  if s.is_empty() {
    return None
  }

  // For now, use a simpler approach without advanced string methods
  let input = s

  // Check for sign
  let (negative, num_part) = if input.length() > 0 && input[0] == 45 { // '-'
    (true, input.substring(start=1))
  } else if input.length() > 0 && input[0] == 43 { // '+'
    (false, input.substring(start=1))
  } else {
    (false, input)
  }
  if num_part.is_empty() {
    return None
  }

  // Parse digits
  let mut result = 0L

  // Check if it's a valid integer string (only digits)
  for i = 0; i < num_part.length(); i = i + 1 {
    let ch = num_part[i]
    if ch >= 48 && ch <= 57 { // '0' to '9'
      continue
    } else {
      return None // Invalid character
    }
  }

  // Convert string to number (simple implementation)
  // Handle common cases first for efficiency
  match num_part {
    "0" => result = 0L
    "1" => result = 1L
    "2" => result = 2L
    "3" => result = 3L
    "4" => result = 4L
    "5" => result = 5L
    "6" => result = 6L
    "7" => result = 7L
    "8" => result = 8L
    "9" => result = 9L
    "10" => result = 10L
    "100" => result = 100L
    "1000" => result = 1000L
    _ => {
      // For longer numbers, use digit-by-digit parsing
      result = 0L
      for i = 0; i < num_part.length(); i = i + 1 {
        let digit = num_part[i] - 48 // Convert ASCII to digit
        result = result * 10L + digit.to_int64()
      }
    }
  }
  if negative {
    Some(-result)
  } else {
    Some(result)
  }
}

///|
/// Parse a string as Double with more robust parsing
pub fn parse_double(s : String) -> Double? {
  if s.is_empty() {
    return None
  }
  let input = s

  // Handle special cases
  match input {
    "NaN" | "nan" => return Some(0.0 / 0.0) // NaN
    "Infinity" | "infinity" | "inf" => return Some(1.0 / 0.0) // Positive infinity
    "-Infinity" | "-infinity" | "-inf" => return Some(-1.0 / 0.0) // Negative infinity
    _ => () // do nothing for other cases
  }

  // Check for sign
  let (negative, num_part) = if input.length() > 0 && input[0] == 45 { // '-'
    (true, input.substring(start=1))
  } else if input.length() > 0 && input[0] == 43 { // '+'
    (false, input.substring(start=1))
  } else {
    (false, input)
  }
  if num_part.is_empty() {
    return None
  }

  // Handle common simple cases for efficiency
  match num_part {
    "0" | "0.0" | "0." => return Some(if negative { -0.0 } else { 0.0 })
    "1" | "1.0" | "1." => return Some(if negative { -1.0 } else { 1.0 })
    "2" | "2.0" | "2." => return Some(if negative { -2.0 } else { 2.0 })
    "0.5" => return Some(if negative { -0.5 } else { 0.5 })
    "1.5" => return Some(if negative { -1.5 } else { 1.5 })
    _ => () // do nothing for other cases
  }

  // Find decimal point
  let decimal_pos = num_part.find(".")
  match decimal_pos {
    Some(pos) => {
      // Has decimal point
      let integer_part = num_part.substring(start=0, end=pos)
      let fractional_part = num_part.substring(start=pos + 1)

      // Parse integer part
      let int_value = if integer_part.is_empty() {
        0.0
      } else {
        match parse_int64(integer_part) {
          Some(val) => val.to_double()
          None => return None
        }
      }

      // Parse fractional part
      let frac_value = if fractional_part.is_empty() {
        0.0
      } else {
        // Check if fractional part has only digits
        let mut all_digits = true
        for i = 0; i < fractional_part.length(); i = i + 1 {
          let ch = fractional_part[i]
          if ch < 48 || ch > 57 { // Not '0' to '9'
            all_digits = false
            break
          }
        }
        if !all_digits {
          return None
        }

        // Convert fractional part
        let mut frac = 0.0
        let mut divisor = 10.0
        for i = 0; i < fractional_part.length(); i = i + 1 {
          let digit = (fractional_part[i] - 48).to_double()
          frac = frac + digit / divisor
          divisor = divisor * 10.0
        }
        frac
      }
      let result = int_value + frac_value
      Some(if negative { -result } else { result })
    }
    None =>
      // No decimal point, parse as integer then convert to double
      match parse_int64(num_part) {
        Some(val) =>
          Some(if negative { -val.to_double() } else { val.to_double() })
        None => None
      }
  }
}

///|
/// Similar methods for DataRef
pub fn DataRef::is_empty(self : DataRef) -> Bool {
  match self {
    Empty => true
    _ => false
  }
}

///|
pub fn DataRef::is_int(self : DataRef) -> Bool {
  match self {
    Int(_) => true
    _ => false
  }
}

///|
pub fn DataRef::is_float(self : DataRef) -> Bool {
  match self {
    Float(_) => true
    _ => false
  }
}

///|
pub fn DataRef::is_bool(self : DataRef) -> Bool {
  match self {
    Bool(_) => true
    _ => false
  }
}

///|
pub fn DataRef::is_string(self : DataRef) -> Bool {
  match self {
    String(_) | SharedString(_) => true
    _ => false
  }
}

///|
pub fn DataRef::is_error(self : DataRef) -> Bool {
  match self {
    Error(_) => true
    _ => false
  }
}

///|
pub fn DataRef::is_datetime(self : DataRef) -> Bool {
  match self {
    DateTime(_) => true
    _ => false
  }
}

///|
pub fn DataRef::is_datetime_iso(self : DataRef) -> Bool {
  match self {
    DateTimeIso(_) => true
    _ => false
  }
}

///|
pub fn DataRef::is_duration_iso(self : DataRef) -> Bool {
  match self {
    DurationIso(_) => true
    _ => false
  }
}

///|
pub fn DataRef::get_int(self : DataRef) -> Int64? {
  match self {
    Int(v) => Some(v)
    _ => None
  }
}

///|
pub fn DataRef::get_float(self : DataRef) -> Double? {
  match self {
    Float(v) => Some(v)
    _ => None
  }
}

///|
pub fn DataRef::get_bool(self : DataRef) -> Bool? {
  match self {
    Bool(v) => Some(v)
    _ => None
  }
}

///|
pub fn DataRef::get_string(self : DataRef) -> String? {
  match self {
    String(v) => Some(v)
    SharedString(v) => Some(v)
    _ => None
  }
}

///|
pub fn DataRef::get_datetime(self : DataRef) -> ExcelDateTime? {
  match self {
    DateTime(v) => Some(v)
    _ => None
  }
}

///|
pub fn DataRef::get_datetime_iso(self : DataRef) -> String? {
  match self {
    DateTimeIso(v) => Some(v)
    _ => None
  }
}

///|
pub fn DataRef::get_duration_iso(self : DataRef) -> String? {
  match self {
    DurationIso(v) => Some(v)
    _ => None
  }
}

///|
pub fn DataRef::get_error(self : DataRef) -> CellErrorType? {
  match self {
    Error(e) => Some(e)
    _ => None
  }
}

///|
pub fn DataRef::as_string(self : DataRef) -> String? {
  match self {
    Float(v) => Some(v.to_string())
    Int(v) => Some(v.to_string())
    String(v) => Some(v)
    SharedString(v) => Some(v)
    Bool(v) => Some(v.to_string())
    DateTime(v) => Some(v.value.to_string())
    DateTimeIso(v) => Some(v)
    DurationIso(v) => Some(v)
    Error(e) => Some(e.to_string())
    Empty => Some("")
  }
}

///|
pub fn DataRef::as_i64(self : DataRef) -> Int64? {
  match self {
    Int(v) => Some(v)
    Float(v) => Some(v.to_int64())
    Bool(true) => Some(1L)
    Bool(false) => Some(0L)
    String(v) => parse_int64(v)
    SharedString(v) => parse_int64(v)
    _ => None
  }
}

///|
pub fn DataRef::as_f64(self : DataRef) -> Double? {
  match self {
    Int(v) => Some(v.to_double())
    Float(v) => Some(v)
    Bool(true) => Some(1.0)
    Bool(false) => Some(0.0)
    String(v) => parse_double(v)
    SharedString(v) => parse_double(v)
    _ => None
  }
}

///|
/// A trait to represent all different data types that can appear as
/// a value in a worksheet cell
/// 
/// This trait provides a common interface for both Data and DataRef types,
/// allowing generic operations on cell values regardless of their storage type.
pub trait DataType {
  /// Assess if datatype is empty
  is_empty(Self) -> Bool

  /// Assess if datatype is an integer
  is_int(Self) -> Bool

  /// Assess if datatype is a float
  is_float(Self) -> Bool

  /// Assess if datatype is a bool
  is_bool(Self) -> Bool

  /// Assess if datatype is a string
  is_string(Self) -> Bool

  /// Assess if datatype is a CellErrorType
  is_error(Self) -> Bool

  /// Assess if datatype is a datetime
  is_datetime(Self) -> Bool

  /// Assess if datatype is an ISO8601 datetime  
  is_datetime_iso(Self) -> Bool

  /// Assess if datatype is an ISO8601 duration
  is_duration_iso(Self) -> Bool

  /// Try getting int value
  get_int(Self) -> Int64?

  /// Try getting float value
  get_float(Self) -> Double?

  /// Try getting bool value
  get_bool(Self) -> Bool?

  /// Try getting string value
  get_string(Self) -> String?

  /// Try getting datetime value
  get_datetime(Self) -> ExcelDateTime?

  /// Try getting datetime ISO8601 value
  get_datetime_iso(Self) -> String?

  /// Try getting duration ISO8601 value
  get_duration_iso(Self) -> String?

  /// Try getting Error value
  get_error(Self) -> CellErrorType?

  /// Try converting data type into a string
  as_string(Self) -> String?

  /// Try converting data type into an int
  as_i64(Self) -> Int64?

  /// Try converting data type into a float
  as_f64(Self) -> Double?

  /// Try converting data type into a date
  as_date(Self) -> String?

  /// Try converting data type into a time
  as_time(Self) -> String?

  /// Try converting data type into a duration
  as_duration(Self) -> String?

  /// Try converting data type into a datetime
  as_datetime(Self) -> String?
}

///|
/// Implementation of DataType trait for Data
impl DataType for Data with is_empty(self) {
  self.is_empty()
}

///|
impl DataType for Data with is_int(self) {
  self.is_int()
}

///|
impl DataType for Data with is_float(self) {
  self.is_float()
}

///|
impl DataType for Data with is_bool(self) {
  self.is_bool()
}

///|
impl DataType for Data with is_string(self) {
  self.is_string()
}

///|
impl DataType for Data with is_error(self) {
  self.is_error()
}

///|
impl DataType for Data with is_datetime(self) {
  self.is_datetime()
}

///|
impl DataType for Data with is_datetime_iso(self) {
  self.is_datetime_iso()
}

///|
impl DataType for Data with is_duration_iso(self) {
  self.is_duration_iso()
}

///|
impl DataType for Data with get_int(self) {
  self.get_int()
}

///|
impl DataType for Data with get_float(self) {
  self.get_float()
}

///|
impl DataType for Data with get_bool(self) {
  self.get_bool()
}

///|
impl DataType for Data with get_string(self) {
  self.get_string()
}

///|
impl DataType for Data with get_datetime(self) {
  self.get_datetime()
}

///|
impl DataType for Data with get_datetime_iso(self) {
  self.get_datetime_iso()
}

///|
impl DataType for Data with get_duration_iso(self) {
  self.get_duration_iso()
}

///|
impl DataType for Data with get_error(self) {
  self.get_error()
}

///|
impl DataType for Data with as_string(self) {
  self.as_string()
}

///|
impl DataType for Data with as_i64(self) {
  self.as_i64()
}

///|
impl DataType for Data with as_f64(self) {
  self.as_f64()
}

///|
/// Try converting data type into a date (if datetime/date conversion is available)
impl DataType for Data with as_date(self) {
  match self {
    DateTime(dt) => Some(dt.as_date())
    DateTimeIso(s) =>
      // Try to parse ISO date string (simplified implementation)
      if s.length() >= 10 {
        // Extract date part (YYYY-MM-DD)
        let date_part = s.substring(start=0, end=10)
        // For now, return the date as a string representation
        // In a full implementation, this would parse to a proper date type
        Some(date_part)
      } else {
        None
      }
    _ => None
  }
}

///|
/// Try converting data type into a time
impl DataType for Data with as_time(self) {
  match self {
    DateTime(dt) => Some(dt.as_time())
    DateTimeIso(s) =>
      // Try to extract time part from ISO datetime string
      if s.contains("T") {
        let parts = s.split("T").collect()
        if parts.length() >= 2 {
          Some(parts[1].to_string())
        } else {
          None
        }
      } else {
        None
      }
    DurationIso(s) => Some(s) // Duration ISO can represent time
    _ => None
  }
}

///|
/// Try converting data type into a duration
impl DataType for Data with as_duration(self) {
  match self {
    DateTime(dt) => dt.as_duration()
    DurationIso(s) => Some(s) // Return the duration string
    _ => None
  }
}

///|
/// Try converting data type into a datetime
impl DataType for Data with as_datetime(self) {
  match self {
    Int(i) => {
      // Convert Excel serial date to datetime
      let excel_dt = ExcelDateTime::new(i.to_double(), DateTime, false)
      Some(excel_dt.as_datetime())
    }
    Float(f) => {
      // Convert Excel serial date to datetime  
      let excel_dt = ExcelDateTime::new(f, DateTime, false)
      Some(excel_dt.as_datetime())
    }
    DateTime(dt) => Some(dt.as_datetime())
    DateTimeIso(s) => Some(s) // Return the ISO datetime string
    _ => None
  }
}

///|
/// Implementation of DataType trait for DataRef
impl DataType for DataRef with is_empty(self) {
  self.is_empty()
}

///|
impl DataType for DataRef with is_int(self) {
  self.is_int()
}

///|
impl DataType for DataRef with is_float(self) {
  self.is_float()
}

///|
impl DataType for DataRef with is_bool(self) {
  self.is_bool()
}

///|
impl DataType for DataRef with is_string(self) {
  self.is_string()
}

///|
impl DataType for DataRef with is_error(self) {
  self.is_error()
}

///|
impl DataType for DataRef with is_datetime(self) {
  self.is_datetime()
}

///|
impl DataType for DataRef with is_datetime_iso(self) {
  self.is_datetime_iso()
}

///|
impl DataType for DataRef with is_duration_iso(self) {
  self.is_duration_iso()
}

///|
impl DataType for DataRef with get_int(self) {
  self.get_int()
}

///|
impl DataType for DataRef with get_float(self) {
  self.get_float()
}

///|
impl DataType for DataRef with get_bool(self) {
  self.get_bool()
}

///|
impl DataType for DataRef with get_string(self) {
  self.get_string()
}

///|
impl DataType for DataRef with get_datetime(self) {
  self.get_datetime()
}

///|
impl DataType for DataRef with get_datetime_iso(self) {
  self.get_datetime_iso()
}

///|
impl DataType for DataRef with get_duration_iso(self) {
  self.get_duration_iso()
}

///|
impl DataType for DataRef with get_error(self) {
  self.get_error()
}

///|
impl DataType for DataRef with as_string(self) {
  self.as_string()
}

///|
impl DataType for DataRef with as_i64(self) {
  self.as_i64()
}

///|
impl DataType for DataRef with as_f64(self) {
  self.as_f64()
}

///|
/// Try converting data type into a date (if datetime/date conversion is available)
impl DataType for DataRef with as_date(self) {
  match self {
    DateTime(dt) => Some(dt.as_date())
    DateTimeIso(s) =>
      // Try to parse ISO date string (simplified implementation)
      if s.length() >= 10 {
        // Extract date part (YYYY-MM-DD)
        let date_part = s.substring(start=0, end=10)
        Some(date_part)
      } else {
        None
      }
    _ => None
  }
}

///|
/// Try converting data type into a time
impl DataType for DataRef with as_time(self) {
  match self {
    DateTime(dt) => Some(dt.as_time())
    DateTimeIso(s) =>
      // Try to extract time part from ISO datetime string
      if s.contains("T") {
        let parts = s.split("T").collect()
        if parts.length() >= 2 {
          Some(parts[1].to_string())
        } else {
          None
        }
      } else {
        None
      }
    DurationIso(s) => Some(s) // Duration ISO can represent time
    SharedString(s) =>
      // Try to parse as datetime string
      if s.contains("T") {
        let parts = s.split("T").collect()
        if parts.length() >= 2 {
          Some(parts[1].to_string())
        } else {
          None
        }
      } else {
        None
      }
    _ => None
  }
}

///|
/// Try converting data type into a duration
impl DataType for DataRef with as_duration(self) {
  match self {
    DateTime(dt) => dt.as_duration()
    DurationIso(s) => Some(s) // Return the duration string
    SharedString(s) => {
      // Try to parse as duration
      let str = s.to_string()
      if str.strip_prefix("P") is Some(_) || str.strip_prefix("PT") is Some(_) {
        Some(s)
      } else {
        None
      }
    }
    _ => None
  }
}

///|
/// Try converting data type into a datetime
impl DataType for DataRef with as_datetime(self) {
  match self {
    Int(i) => {
      // Convert Excel serial date to datetime
      let excel_dt = ExcelDateTime::new(i.to_double(), DateTime, false)
      Some(excel_dt.as_datetime())
    }
    Float(f) => {
      // Convert Excel serial date to datetime  
      let excel_dt = ExcelDateTime::new(f, DateTime, false)
      Some(excel_dt.as_datetime())
    }
    DateTime(dt) => Some(dt.as_datetime())
    DateTimeIso(s) => Some(s) // Return the ISO datetime string
    SharedString(s) => Some(s) // Return the shared string
    _ => None
  }
}

///|
/// Convert DataRef to Option types for safe extraction
pub fn DataRef::to_option_bool(self : DataRef) -> Bool? {
  self.get_bool()
}

///|
/// Default implementation for Data type
impl Default for Data with default() {
  Data::Empty
}

///|
/// Default implementation for DataRef type  
impl Default for DataRef with default() {
  DataRef::Empty
}

///|
/// Convert iterator of Data values to array
pub fn[T] iter_to_array(iter : Array[T]) -> Array[T] {
  iter
}

///|
/// Helper function to create a Range from an array of arrays (2D data)
pub fn[T : Default] Range::from_2d_array(
  data : Array[Array[T]],
  start_row : UInt,
  start_col : UInt,
) -> Range[T] {
  if data.is_empty() {
    return Range::empty()
  }
  let height = data.length().reinterpret_as_uint()
  let width = data[0].length().reinterpret_as_uint()
  let end_row = start_row + height - 1U
  let end_col = start_col + width - 1U
  let range = Range::new((start_row, start_col), (end_row, end_col))
  for r = 0; r < data.length(); r = r + 1 {
    let row = data[r]
    for c = 0; c < row.length(); c = c + 1 {
      let row_pos = start_row + r.reinterpret_as_uint()
      let col_pos = start_col + c.reinterpret_as_uint()
      range.set_value((row_pos, col_pos), row[c])
    }
  }
  range
}

///|
/// Helper function to convert Range to 2D array
pub fn[T] Range::to_2d_array(self : Range[T]) -> Array[Array[T]] {
  let rows = self.rows()
  let result = []
  for row in rows {
    let row_array = []
    for i = 0; i < row.length(); i = i + 1 {
      row_array.push(row[i])
    }
    result.push(row_array)
  }
  result
}

///|
/// Create a Range with a single value repeated
pub fn[T : Default] Range::filled(
  start : (UInt, UInt),
  end : (UInt, UInt),
  value : T,
) -> Range[T] {
  let range = Range::new(start, end)
  let cells = range.cells()
  for cell in cells {
    let row = cell.0
    let col = cell.1
    range.set_cell((row, col), value)
  }
  range
}

///|
/// Check if two ranges have the same dimensions
pub fn[T, U] Range::same_dimensions(self : Range[T], other : Range[U]) -> Bool {
  self.get_size() == other.get_size()
}

///|
/// Transform each cell value using a function
pub fn[T, U : Default] Range::map(self : Range[T], f : (T) -> U) -> Range[U] {
  if self.is_empty() {
    return Range::empty()
  }
  let result = Range::new(self.start, self.end)
  let cells = self.cells()
  for cell in cells {
    let row = cell.0
    let col = cell.1
    let value = cell.2
    result.set_cell((row, col), f(value))
  }
  result
}

///|
/// Filter cells based on a predicate and return matching positions
pub fn[T] Range::filter_positions(
  self : Range[T],
  predicate : (T) -> Bool,
) -> Array[(UInt, UInt)] {
  let result = []
  let cells = self.cells()
  for cell in cells {
    let row = cell.0
    let col = cell.1
    let value = cell.2
    if predicate(value) {
      result.push((row, col))
    }
  }
  result
}

///|
/// Count cells that match a predicate
pub fn[T] Range::count_if(self : Range[T], predicate : (T) -> Bool) -> UInt {
  let mut count = 0U
  let cells = self.cells()
  for cell in cells {
    let value = cell.2
    if predicate(value) {
      count = count + 1U
    }
  }
  count
}

///|
/// Find first cell that matches a predicate
pub fn[T] Range::find_first(
  self : Range[T],
  predicate : (T) -> Bool,
) -> (UInt, UInt, T)? {
  let cells = self.cells()
  for cell in cells {
    let row = cell.0
    let col = cell.1
    let value = cell.2
    if predicate(value) {
      return Some((row, col, value))
    }
  }
  None
}

///|
/// Check if a DataType represents a text value
pub fn[T : DataType] is_text(value : T) -> Bool {
  value.is_string()
}

///|
/// Check if a DataType represents any kind of datetime
pub fn[T : DataType] is_any_datetime(value : T) -> Bool {
  value.is_datetime() || value.is_datetime_iso() || value.is_duration_iso()
}

///|
/// Try to get a numeric value as Double from any DataType
pub fn[T : DataType] as_numeric(value : T) -> Double? {
  if value.is_float() {
    value.as_f64()
  } else if value.is_int() {
    value.as_i64().map(fn(i) { i.to_double() })
  } else {
    None
  }
}

///|
/// Convert DataRef to Data
pub fn DataRef::to_data(self : DataRef) -> Data {
  match self {
    Int(v) => Data::Int(v)
    Float(v) => Data::Float(v)
    String(v) => Data::String(v)
    SharedString(v) => Data::String(v)
    Bool(v) => Data::Bool(v)
    DateTime(v) => Data::DateTime(v)
    DateTimeIso(v) => Data::DateTimeIso(v)
    DurationIso(v) => Data::DurationIso(v)
    Error(v) => Data::Error(v)
    Empty => Data::Empty
  }
}
