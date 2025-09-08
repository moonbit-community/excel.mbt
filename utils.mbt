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
    buf[index] = Char::from_int((65 + col.reinterpret_as_int()))
    return
  }
  
  // For columns >= 26, we need to handle multi-character columns
  let temp_chars : Array[Char] = []
  while col >= 26U {
    let c = col % 26U
    temp_chars.push(Char::from_int((65 + c.reinterpret_as_int())))
    col = col - c
    col = col / 26U
  }
  
  // Add remaining character if any
  if col > 0U {
    temp_chars.push(Char::from_int((65 + (col - 1U).reinterpret_as_int())))
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
  let buf = FixedArray::make(10, ' ')  // Should be enough for most columns
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
/// Get function name by index
pub fn get_function_name(index : UInt) -> String? {
  // Excel function table for formula parsing  
  let ftab = [
    "COUNT", "IF", "ISNA", "ISERROR", "SUM", "AVERAGE", "MIN", "MAX",
    "ROW", "COLUMN", "NA", "NPV", "STDEV", "DOLLAR", "FIXED", "SIN",
    "COS", "TAN", "ATAN", "PI", "SQRT", "EXP", "LN", "LOG10",
    "ABS", "INT", "SIGN", "ROUND", "LOOKUP", "INDEX", "REPT", "MID",
    "LEN", "VALUE", "TRUE", "FALSE", "AND", "OR", "NOT", "MOD",
  ]
  if index.reinterpret_as_int() < ftab.length() {
    Some(ftab[index.reinterpret_as_int()])
  } else {
    None
  }
}

///|
/// Get function argument count by index
pub fn get_function_argc(index : UInt) -> UInt? {
  // Excel function argument counts
  let ftab_argc = [
    255U, 3U, 1U, 1U, 255U, 255U, 255U, 255U,  // COUNT, IF, ISNA, etc.
    1U, 1U, 0U, 254U, 255U, 2U, 3U, 1U,        // ROW, COLUMN, NA, etc.
    1U, 1U, 1U, 0U, 1U, 1U, 1U, 1U,           // COS, TAN, ATAN, etc.
    1U, 1U, 1U, 2U, 3U, 4U, 2U, 3U,           // ABS, INT, SIGN, etc.
    1U, 1U, 0U, 0U, 255U, 255U, 1U, 2U,       // LEN, VALUE, TRUE, etc.
  ]
  if index.reinterpret_as_int() < ftab_argc.length() {
    Some(ftab_argc[index.reinterpret_as_int()])
  } else {
    None
  }
}