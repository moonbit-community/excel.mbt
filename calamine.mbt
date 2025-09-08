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
