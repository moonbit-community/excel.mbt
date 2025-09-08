///|
/// Cell format detection for Excel files
/// 
/// This module provides functionality to detect whether a cell format
/// represents a date/time, duration, or other format types.

///|
/// Cell format classification
pub(all) enum CellFormat {
  /// Any format that is not date/time related
  Other
  /// Date and time formats  
  DateTime
  /// Duration/time delta formats
  TimeDelta
} derive(Show, Eq)

///|
/// Check excel number format string to detect if it represents datetime
pub fn detect_custom_number_format(format : String) -> CellFormat {
  let mut escaped = false
  let mut is_quote = false
  let mut brackets = 0
  let mut prev = ' '
  let mut hms = false
  let mut ap = false
  for char in format {
    match (char, escaped, is_quote, ap, brackets) {
      // If escaped, ignore this character  
      (_, true, _, _, _) => escaped = false
      // Set escaped flag for next character
      ('_', false, _, _, _) | ('\\', false, _, _, _) => escaped = true
      // Handle quoted strings
      ('"', false, true, _, _) => is_quote = false
      (_, false, true, _, _) => () // Ignore characters inside quotes
      ('"', false, false, _, _) => is_quote = true
      // First format only - stop at semicolon
      (';', false, false, _, _) => return Other
      // Handle brackets for duration detection
      ('[', false, false, _, _) => brackets = brackets + 1
      (']', false, false, _, b) if b > 0 && hms => return TimeDelta
      (']', false, false, _, b) if b > 0 =>
        brackets = if brackets > 0 { brackets - 1 } else { 0 }
      // AM/PM detection
      ('a', false, false, false, 0) | ('A', false, false, false, 0) => ap = true
      ('p', false, false, true, 0)
      | ('P', false, false, true, 0)
      | ('m', false, false, true, 0)
      | ('M', false, false, true, 0)
      | ('/', false, false, true, 0) => return DateTime
      // Date/time format characters
      ('d', false, false, false, 0)
      | ('D', false, false, false, 0)
      | ('m', false, false, false, 0)
      | ('M', false, false, false, 0)
      | ('h', false, false, false, 0)
      | ('H', false, false, false, 0)
      | ('y', false, false, false, 0)
      | ('Y', false, false, false, 0)
      | ('s', false, false, false, 0)
      | ('S', false, false, false, 0) => return DateTime
      _ =>
        if hms && char == prev {
          // Continue hms detection - do nothing
        } else {
          hms = prev == '[' &&
            (
              char == 'm' ||
              char == 'M' ||
              char == 'h' ||
              char == 'H' ||
              char == 's' ||
              char == 'S'
            )
        }
    }
    prev = char
  }
  Other
}

///|
/// Check if builtin format ID represents a date format
pub fn builtin_format_by_id(id : String) -> CellFormat {
  match id {
    // mm-dd-yy
    "14"
    // d-mmm-yy  
    | "15"
    // d-mmm
    | "16"
    // mmm-yy
    | "17"
    // h:mm AM/PM
    | "18"
    // h:mm:ss AM/PM
    | "19"
    // h:mm
    | "20"
    // h:mm:ss
    | "21"
    // m/d/yy h:mm
    | "22"
    // mm:ss
    | "45"
    // mmss.0
    | "47" => DateTime
    // [h]:mm:ss
    "46" => TimeDelta
    _ => Other
  }
}

///|
/// Check if builtin format code represents a date format
pub fn builtin_format_by_code(code : UInt) -> CellFormat {
  match code {
    14U | 15U | 16U | 17U | 18U | 19U | 20U | 21U | 22U | 45U | 47U => DateTime
    46U => TimeDelta
    _ => Other
  }
}

///|
/// Format i64 value according to cell format
pub fn format_excel_i64(
  value : Int64,
  format : CellFormat?,
  is_1904 : Bool,
) -> Data {
  match format {
    Some(DateTime) =>
      Data::DateTime(
        ExcelDateTime::new(
          value.to_double(),
          ExcelDateTimeType::DateTime,
          is_1904,
        ),
      )
    Some(TimeDelta) =>
      Data::DateTime(
        ExcelDateTime::new(
          value.to_double(),
          ExcelDateTimeType::TimeDelta,
          is_1904,
        ),
      )
    _ => Data::Int(value)
  }
}

///|
/// Format f64 value according to cell format
pub fn format_excel_f64(
  value : Double,
  format : CellFormat?,
  is_1904 : Bool,
) -> Data {
  match format {
    Some(DateTime) =>
      Data::DateTime(
        ExcelDateTime::new(value, ExcelDateTimeType::DateTime, is_1904),
      )
    Some(TimeDelta) =>
      Data::DateTime(
        ExcelDateTime::new(value, ExcelDateTimeType::TimeDelta, is_1904),
      )
    _ => Data::Float(value)
  }
}

///|
/// Format f64 value according to cell format for DataRef
pub fn format_excel_f64_ref(
  value : Double,
  format : CellFormat?,
  is_1904 : Bool,
) -> DataRef {
  match format {
    Some(DateTime) =>
      DataRef::DateTime(
        ExcelDateTime::new(value, ExcelDateTimeType::DateTime, is_1904),
      )
    Some(TimeDelta) =>
      DataRef::DateTime(
        ExcelDateTime::new(value, ExcelDateTimeType::TimeDelta, is_1904),
      )
    _ => DataRef::Float(value)
  }
}
