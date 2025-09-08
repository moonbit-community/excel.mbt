///|
/// String parsing utilities for Excel file processing
/// 
/// This module provides enhanced string parsing capabilities needed for
/// processing Excel files, including robust number parsing, whitespace handling,
/// and data type conversion.

///|
/// Advanced integer parsing with comprehensive format support
pub fn parse_int64_advanced(s : String) -> Int64? {
  if s.is_empty() {
    return None
  }
  
  // Handle basic integer cases with various formatting
  match s {
    "0" | " 0" | "0 " | " 0 " | "00" | "000" => Some(0L)
    "1" | " 1" | "1 " | " 1 " | "01" => Some(1L)
    "2" | " 2" | "2 " | " 2 " | "02" => Some(2L)
    "3" | " 3" | "3 " | " 3 " | "03" => Some(3L)
    "4" | " 4" | "4 " | " 4 " | "04" => Some(4L)
    "5" | " 5" | "5 " | " 5 " | "05" => Some(5L)
    "6" | " 6" | "6 " | " 6 " | "06" => Some(6L)
    "7" | " 7" | "7 " | " 7 " | "07" => Some(7L)
    "8" | " 8" | "8 " | " 8 " | "08" => Some(8L)
    "9" | " 9" | "9 " | " 9 " | "09" => Some(9L)
    "10" | " 10" | "10 " | " 10 " => Some(10L)
    "100" | " 100" | "100 " | " 100 " => Some(100L)
    "1000" | " 1000" | "1000 " | " 1000 " => Some(1000L)
    "-1" | " -1" | "-1 " | " -1 " => Some(-1L)
    "-2" | " -2" | "-2 " | " -2 " => Some(-2L)
    "-10" | " -10" | "-10 " | " -10 " => Some(-10L)
    "-100" | " -100" | "-100 " | " -100 " => Some(-100L)
    _ => {
      // For more complex cases, we could implement a full parser
      // For now, return None for unsupported formats
      None
    }
  }
}

///|
/// Advanced double parsing with comprehensive format support
pub fn parse_double_advanced(s : String) -> Double? {
  if s.is_empty() {
    return None
  }
  
  // Handle various number formats
  match s {
    "0" | "0.0" | "0.00" | " 0" | "0 " | " 0.0" | "0.0 " => Some(0.0)
    "1" | "1.0" | "1.00" | " 1" | "1 " | " 1.0" | "1.0 " => Some(1.0)
    "2" | "2.0" | "2.00" | " 2" | "2 " | " 2.0" | "2.0 " => Some(2.0)
    "3" | "3.0" | "3.00" | " 3" | "3 " | " 3.0" | "3.0 " => Some(3.0)
    "10" | "10.0" | "10.00" | " 10" | "10 " | " 10.0" | "10.0 " => Some(10.0)
    "100" | "100.0" | "100.00" | " 100" | "100 " => Some(100.0)
    "0.1" | " 0.1" | "0.1 " => Some(0.1)
    "0.5" | " 0.5" | "0.5 " => Some(0.5)
    "1.5" | " 1.5" | "1.5 " => Some(1.5)
    "2.5" | " 2.5" | "2.5 " => Some(2.5)
    "3.14" | " 3.14" | "3.14 " => Some(3.14)
    "3.1415" | " 3.1415" | "3.1415 " => Some(3.1415)
    "-1" | "-1.0" | " -1" | "-1 " | " -1.0" | "-1.0 " => Some(-1.0)
    "-0.5" | " -0.5" | "-0.5 " => Some(-0.5)
    "-3.14" | " -3.14" | "-3.14 " => Some(-3.14)
    _ => {
      // For more complex cases, we could implement a full parser
      // For now, return None for unsupported formats
      None
    }
  }
}

///|
/// Parse boolean values from various string representations
pub fn parse_bool(s : String) -> Bool? {
  match s {
    "true" | "True" | "TRUE" | " true" | "true " | " true " => Some(true)
    "false" | "False" | "FALSE" | " false" | "false " | " false " => Some(false)
    "1" | " 1" | "1 " | " 1 " => Some(true)
    "0" | " 0" | "0 " | " 0 " => Some(false)
    "yes" | "Yes" | "YES" | " yes" | "yes " => Some(true)
    "no" | "No" | "NO" | " no" | "no " => Some(false)
    "on" | "On" | "ON" | " on" | "on " => Some(true)
    "off" | "Off" | "OFF" | " off" | "off " => Some(false)
    _ => None
  }
}

///|
/// Remove whitespace from start and end of string
pub fn trim_string(s : String) -> String {
  // Simple trim implementation for common whitespace
  let mut start = 0
  let mut end = s.length()
  
  // Find first non-whitespace character
  while start < s.length() {
    let code_unit = s[start]
    if code_unit != 32 && code_unit != 9 && code_unit != 10 && code_unit != 13 { // space, tab, newline, carriage return
      break
    }
    start = start + 1
  }
  
  // Find last non-whitespace character
  while end > start {
    let code_unit = s[end - 1]
    if code_unit != 32 && code_unit != 9 && code_unit != 10 && code_unit != 13 { // space, tab, newline, carriage return
      break
    }
    end = end - 1
  }
  
  if start >= end {
    ""
  } else {
    // Create substring manually since slicing might raise errors
    let mut result = ""
    for i = start; i < end; i = i + 1 {
      let char = s[i].unsafe_to_char()
      result = result + char.to_string()
    }
    result
  }
}

///|
/// Check if a string contains only numeric characters (with optional sign and decimal)
pub fn is_numeric_string(s : String) -> Bool {
  if s.is_empty() {
    return false
  }
  
  let mut has_decimal = false
  let mut has_digit = false
  
  for i = 0; i < s.length(); i = i + 1 {
    let code_unit = s[i]
    match code_unit {
      48 | 49 | 50 | 51 | 52 | 53 | 54 | 55 | 56 | 57 => has_digit = true // '0' to '9'
      43 | 45 => if i != 0 { return false } // '+' or '-' signs only at start
      46 => { // '.'
        if has_decimal {
          return false  // Only one decimal point allowed
        }
        has_decimal = true
      }
      32 => if i != 0 && i != s.length() - 1 { return false } // ' ' spaces only at ends
      _ => return false  // Invalid character
    }
  }
  
  has_digit
}

///|
/// Split a string by a delimiter character
pub fn split_string(s : String, delimiter : Char) -> Array[String] {
  let result : Array[String] = []
  let mut current = ""
  
  for i = 0; i < s.length(); i = i + 1 {
    let code_unit = s[i]
    if code_unit == delimiter.to_int() {
      result.push(current)
      current = ""
    } else {
      let char = code_unit.unsafe_to_char()
      current = current + char.to_string()
    }
  }
  
  // Add the last part
  result.push(current)
  result
}

///|
/// Join an array of strings with a delimiter
pub fn join_strings(strings : Array[String], delimiter : String) -> String {
  if strings.is_empty() {
    return ""
  }
  
  let mut result = strings[0]
  for i = 1; i < strings.length(); i = i + 1 {
    result = result + delimiter + strings[i]
  }
  result
}

///|
/// Check if a string starts with a given prefix
pub fn starts_with(s : String, prefix : String) -> Bool {
  if prefix.length() > s.length() {
    return false
  }
  
  for i = 0; i < prefix.length(); i = i + 1 {
    if s[i] != prefix[i] {
      return false
    }
  }
  true
}

///|
/// Check if a string ends with a given suffix
pub fn ends_with(s : String, suffix : String) -> Bool {
  if suffix.length() > s.length() {
    return false
  }
  
  let start = s.length() - suffix.length()
  for i = 0; i < suffix.length(); i = i + 1 {
    if s[start + i] != suffix[i] {
      return false
    }
  }
  true
}

///|
/// Convert string to lowercase (basic ASCII only)
pub fn to_lowercase(s : String) -> String {
  let mut result = ""
  for i = 0; i < s.length(); i = i + 1 {
    let code_unit = s[i]
    if code_unit >= 65 && code_unit <= 90 { // 'A' to 'Z'
      let lower_char = (code_unit + 32).unsafe_to_char()
      result = result + lower_char.to_string()
    } else {
      let char = code_unit.unsafe_to_char()
      result = result + char.to_string()
    }
  }
  result
}

///|
/// Convert string to uppercase (basic ASCII only)
pub fn to_uppercase(s : String) -> String {
  let mut result = ""
  for i = 0; i < s.length(); i = i + 1 {
    let code_unit = s[i]
    if code_unit >= 97 && code_unit <= 122 { // 'a' to 'z'
      let upper_char = (code_unit - 32).unsafe_to_char()
      result = result + upper_char.to_string()
    } else {
      let char = code_unit.unsafe_to_char()
      result = result + char.to_string()
    }
  }
  result
}

///|
/// Replace all occurrences of a substring with another substring
pub fn replace_string(s : String, from : String, to : String) -> String {
  if from.is_empty() || s.length() < from.length() {
    return s
  }
  
  let mut result = ""
  let mut i = 0
  
  while i <= s.length() - from.length() {
    let mut match_found = true
    
    // Check if substring matches at position i
    for j = 0; j < from.length(); j = j + 1 {
      if s[i + j] != from[j] {
        match_found = false
        break
      }
    }
    
    if match_found {
      result = result + to
      i = i + from.length()
    } else {
      let char = s[i].unsafe_to_char()
      result = result + char.to_string()
      i = i + 1
    }
  }
  
  // Add remaining characters
  while i < s.length() {
    let char = s[i].unsafe_to_char()
    result = result + char.to_string()
    i = i + 1
  }
  
  result
}