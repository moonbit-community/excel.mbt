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
/// Convert column number to Excel column letters (A, B, C, ..., AA, AB, etc.)
/// Using proper Excel column numbering where A=1, B=2, ..., Z=26, AA=27, etc.
pub fn push_column(col : UInt, buf : FixedArray[Char]) -> Unit {
  let mut col = col
  let mut index = 0

  // Handle Excel column numbering (1-based with bijective base-26)
  while col > 0U {
    col = col - 1U // Make it 0-based for calculation
    let c = col % 26U
    buf[index] = (65 + c.reinterpret_as_int()).unsafe_to_char()
    index = index + 1
    col = col / 26U
  }

  // Reverse the characters in place since we built them backwards
  let mut left = 0
  let mut right = index - 1
  while left < right {
    let temp = buf[left]
    buf[left] = buf[right]
    buf[right] = temp
    left = left + 1
    right = right - 1
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
/// Convert Excel column string to number (A=1, B=2, AA=27, etc.)
pub fn column_from_string(col : String) -> UInt? {
  if col.is_empty() {
    return None
  }
  let mut result = 0U
  for i = 0; i < col.length(); i = i + 1 {
    let code_unit = col[i]
    if code_unit < 65 || code_unit > 90 { // Not 'A' to 'Z'
      return None
    }
    result = result * 26U + (code_unit - 64).reinterpret_as_uint() // 'A' = 65, but we want A=1
  }
  Some(result)
}

///|
/// Parse A1-style cell reference (e.g., "A1", "BC123") to (row, col) coordinates
pub fn parse_cell_reference(reference : String) -> (UInt, UInt)? {
  if reference.is_empty() {
    return None
  }
  let mut col_str = ""
  let mut row_str = ""
  let mut in_row_part = false

  // Split into column and row parts
  for i = 0; i < reference.length(); i = i + 1 {
    let code_unit = reference[i]
    if code_unit >= 65 && code_unit <= 90 && !in_row_part { // 'A' to 'Z'
      col_str = col_str + code_unit.unsafe_to_char().to_string()
    } else if code_unit >= 48 && code_unit <= 57 { // '0' to '9'
      in_row_part = true
      row_str = row_str + code_unit.unsafe_to_char().to_string()
    } else {
      return None // Invalid character
    }
  }

  // Parse column
  let col = match column_from_string(col_str) {
    Some(c) => c - 1U // Convert to 0-based
    None => return None
  }

  // Parse row (simple implementation for common cases)
  let row = match parse_simple_int(row_str) {
    Some(r) => r - 1U // Convert to 0-based
    None => return None
  }
  Some((row, col))
}

///|
/// Simple integer parsing for small positive numbers
fn parse_simple_int(s : String) -> UInt? {
  if s.is_empty() {
    return None
  }

  // Simple implementation for common row numbers
  match s {
    "1" => Some(1U)
    "2" => Some(2U)
    "3" => Some(3U)
    "4" => Some(4U)
    "5" => Some(5U)
    "6" => Some(6U)
    "7" => Some(7U)
    "8" => Some(8U)
    "9" => Some(9U)
    "10" => Some(10U)
    "100" => Some(100U)
    "1000" => Some(1000U)
    _ => {
      // For other cases, try a simple digit-by-digit approach
      let mut result = 0U
      for i = 0; i < s.length(); i = i + 1 {
        let code_unit = s[i]
        if code_unit >= 48 && code_unit <= 57 {
          let digit = (code_unit - 48).reinterpret_as_uint()
          result = result * 10U + digit
        } else {
          return None
        }
      }
      Some(result)
    }
  }
}

///|
/// Convert (row, col) coordinates to A1-style reference
pub fn format_cell_reference(row : UInt, col : UInt) -> String {
  let col_str = column_to_string(col + 1U) // Convert to 1-based
  let row_str = (row + 1U).to_string() // Convert to 1-based
  col_str + row_str
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
/// Parse range notation (e.g., "A1:C3") to Dimensions
pub fn parse_range(range : String) -> Dimensions? {
  if range.is_empty() {
    return None
  }

  // Find the colon separator
  let mut colon_pos = -1
  for i = 0; i < range.length(); i = i + 1 {
    if range[i] == 58 { // ':'
      colon_pos = i
      break
    }
  }
  if colon_pos == -1 {
    // Single cell reference
    match parse_cell_reference(range) {
      Some((row, col)) => Some(Dimensions::new((row, col), (row, col)))
      None => None
    }
  } else {
    // Range reference - manually extract substrings
    let mut start_ref = ""
    for i = 0; i < colon_pos; i = i + 1 {
      start_ref = start_ref + range[i].unsafe_to_char().to_string()
    }
    let mut end_ref = ""
    for i = colon_pos + 1; i < range.length(); i = i + 1 {
      end_ref = end_ref + range[i].unsafe_to_char().to_string()
    }
    match (parse_cell_reference(start_ref), parse_cell_reference(end_ref)) {
      (Some((start_row, start_col)), Some((end_row, end_col))) =>
        Some(Dimensions::new((start_row, start_col), (end_row, end_col)))
      _ => None
    }
  }
}

///|
/// Extract file extension from path
pub fn get_file_extension(path : String) -> String? {
  let mut last_dot = -1
  for i = path.length() - 1; i >= 0; i = i - 1 {
    if path[i] == 46 { // '.'
      last_dot = i
      break
    }
  }
  if last_dot >= 0 && last_dot < path.length() - 1 {
    // Manually extract the extension substring
    let mut extension = ""
    for i = last_dot + 1; i < path.length(); i = i + 1 {
      extension = extension + path[i].unsafe_to_char().to_string()
    }
    Some(extension)
  } else {
    None
  }
}

///|
/// Check if a number is a power of 2
pub fn is_power_of_two(n : UInt) -> Bool {
  n != 0U && (n & (n - 1U)) == 0U
}

///|
/// Convert big-endian u32 to native format
pub fn read_u32_be(data : Bytes, offset : Int) -> UInt? {
  if offset + 4 > data.length() {
    return None
  }
  let b0 = data[offset].to_uint()
  let b1 = data[offset + 1].to_uint()
  let b2 = data[offset + 2].to_uint()
  let b3 = data[offset + 3].to_uint()
  Some((b0 << 24) | (b1 << 16) | (b2 << 8) | b3)
}

///|
/// Convert big-endian u16 to native format
pub fn read_u16_be(data : Bytes, offset : Int) -> UInt? {
  if offset + 2 > data.length() {
    return None
  }
  let b0 = data[offset].to_uint()
  let b1 = data[offset + 1].to_uint()
  Some((b0 << 8) | b1)
}

///|
/// Calculate CRC32 checksum (simple implementation)
pub fn calculate_crc32(data : Bytes) -> UInt {
  // Simplified CRC32 - in a real implementation this would use proper CRC32 tables
  let mut crc = 4294967295U // 0xFFFFFFFF
  for i = 0; i < data.length(); i = i + 1 {
    crc = crc ^ data[i].to_uint()
    for i = 0; i < 8; i = i + 1 {
      if (crc & 1U) != 0U {
        crc = (crc >> 1) ^ 3988292384U // 0xEDB88320
      } else {
        crc = crc >> 1
      }
    }
  }
  crc ^ 4294967295U
}

///|
/// Align value to the next boundary
pub fn align_to_boundary(value : UInt, boundary : UInt) -> UInt {
  (value + boundary - 1U) & (boundary - 1U).lnot()
}
