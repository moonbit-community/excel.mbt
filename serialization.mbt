///|
/// Serialization utilities for calamine data types
/// 
/// This module provides serialization and deserialization capabilities
/// for calamine data structures, enabling conversion to/from common formats.

///|
/// Convert Data value to JSON-like string representation
pub fn Data::to_json_string(self : Data) -> String {
  match self {
    Int(v) => v.to_string()
    Float(v) => v.to_string()
    String(v) => "\"" + v + "\""
    Bool(v) => v.to_string()
    DateTime(v) => "\"" + v.value.to_string() + "\""
    DateTimeIso(v) => "\"" + v + "\""
    DurationIso(v) => "\"" + v + "\""
    Error(e) => "\"" + e.to_string() + "\""
    Empty => "null"
  }
}

///|
/// Convert DataRef value to JSON-like string representation
pub fn DataRef::to_json_string(self : DataRef) -> String {
  match self {
    Int(v) => v.to_string()
    Float(v) => v.to_string()
    String(v) => "\"" + v + "\""
    SharedString(v) => "\"" + v + "\""
    Bool(v) => v.to_string()
    DateTime(v) => "\"" + v.value.to_string() + "\""
    DateTimeIso(v) => "\"" + v + "\""
    DurationIso(v) => "\"" + v + "\""
    Error(e) => "\"" + e.to_string() + "\""
    Empty => "null"
  }
}

///|
/// Convert Range to CSV format string
pub fn[T : Show] Range::to_csv(self : Range[T]) -> String {
  let rows = self.rows()
  let mut result = ""
  for i = 0; i < rows.length(); i = i + 1 {
    let row = rows[i]
    for j = 0; j < row.length(); j = j + 1 {
      if j > 0 {
        result = result + ","
      }
      let cell_value = row[j].to_string()
      // Simple CSV escaping - wrap in quotes if contains comma
      if string_contains(cell_value, ",") {
        result = result + "\"" + cell_value + "\""
      } else {
        result = result + cell_value
      }
    }
    result = result + "\n"
  }
  result
}

///|
/// Convert Range to tab-separated values format
pub fn[T : Show] Range::to_tsv(self : Range[T]) -> String {
  let rows = self.rows()
  let mut result = ""
  for i = 0; i < rows.length(); i = i + 1 {
    let row = rows[i]
    for j = 0; j < row.length(); j = j + 1 {
      if j > 0 {
        result = result + "\t"
      }
      result = result + row[j].to_string()
    }
    result = result + "\n"
  }
  result
}

///|
/// Convert Range to JSON array format
pub fn[T : Show] Range::to_json_array(self : Range[T]) -> String {
  let rows = self.rows()
  let mut result = "["
  for i = 0; i < rows.length(); i = i + 1 {
    if i > 0 {
      result = result + ","
    }
    result = result + "["
    let row = rows[i]
    for j = 0; j < row.length(); j = j + 1 {
      if j > 0 {
        result = result + ","
      }
      result = result + "\"" + row[j].to_string() + "\""
    }
    result = result + "]"
  }
  result = result + "]"
  result
}

///|
/// Simple helper to detect if string contains a substring
fn contains_substring(s : String, substring : String) -> Bool {
  if substring.length() > s.length() {
    return false
  }
  for i = 0; i <= s.length() - substring.length(); i = i + 1 {
    let mut match_found = true
    for j = 0; j < substring.length(); j = j + 1 {
      if s[i + j] != substring[j] {
        match_found = false
        break
      }
    }
    if match_found {
      return true
    }
  }
  false
}

///|
/// Check if string contains substring (helper function)
pub fn string_contains(s : String, substring : String) -> Bool {
  contains_substring(s, substring)
}

///|
/// Create a Range from CSV data (basic implementation)
pub fn Range::from_csv(csv_data : String) -> Range[Data] {
  // Use functions from the parsing module in the same package
  let lines = split_string(csv_data, '\n')
  let cells : Array[Cell[Data]] = []
  for row = 0; row < lines.length(); row = row + 1 {
    let line = trim_string(lines[row])
    if !line.is_empty() {
      let cols = split_string(line, ',')
      for col = 0; col < cols.length(); col = col + 1 {
        let cell_str = trim_string(cols[col])
        let data = if cell_str.is_empty() {
          Data::Empty
        } else {
          Data::String(cell_str)
        }
        cells.push(
          Cell::new(
            (row.reinterpret_as_uint(), col.reinterpret_as_uint()),
            data,
          ),
        )
      }
    }
  }
  Range::from_sparse(cells)
}

///|
/// Create helper functions for common data transformations
pub fn[T : Default] Range::transpose(self : Range[T]) -> Range[T] {
  if self.is_empty() {
    return Range::empty()
  }
  let original_cells = self.cells()
  let transposed_cells : Array[Cell[T]] = []
  for cell in original_cells {
    let original_row = cell.0
    let original_col = cell.1
    let value = cell.2
    // Swap row and column coordinates
    transposed_cells.push(Cell::new((original_col, original_row), value))
  }
  Range::from_sparse(transposed_cells)
}

///|
/// Merge two ranges horizontally (side by side)
pub fn[T : Default] Range::merge_horizontal(
  left : Range[T],
  right : Range[T],
) -> Range[T] {
  if left.is_empty() {
    return right
  }
  if right.is_empty() {
    return left
  }
  let left_height = left.height()
  let left_width = left.width()
  let right_height = right.height()
  let right_width = right.width()
  let merged_height = if left_height > right_height {
    left_height
  } else {
    right_height
  }
  let merged_width = left_width + right_width
  let merged = Range::new((0U, 0U), (merged_height - 1U, merged_width - 1U))

  // Copy left range
  let left_cells = left.cells()
  for cell in left_cells {
    let row = cell.0
    let col = cell.1
    let value = cell.2
    merged.set_cell((row, col), value)
  }

  // Copy right range with column offset
  let right_cells = right.cells()
  for cell in right_cells {
    let row = cell.0
    let col = cell.1 + left_width
    let value = cell.2
    merged.set_cell((row, col), value)
  }
  merged
}

///|
/// Merge two ranges vertically (one below the other)
pub fn[T : Default] Range::merge_vertical(
  top : Range[T],
  bottom : Range[T],
) -> Range[T] {
  if top.is_empty() {
    return bottom
  }
  if bottom.is_empty() {
    return top
  }
  let top_height = top.height()
  let top_width = top.width()
  let bottom_height = bottom.height()
  let bottom_width = bottom.width()
  let merged_width = if top_width > bottom_width {
    top_width
  } else {
    bottom_width
  }
  let merged_height = top_height + bottom_height
  let merged = Range::new((0U, 0U), (merged_height - 1U, merged_width - 1U))

  // Copy top range
  let top_cells = top.cells()
  for cell in top_cells {
    let row = cell.0
    let col = cell.1
    let value = cell.2
    merged.set_cell((row, col), value)
  }

  // Copy bottom range with row offset
  let bottom_cells = bottom.cells()
  for cell in bottom_cells {
    let row = cell.0 + top_height
    let col = cell.1
    let value = cell.2
    merged.set_cell((row, col), value)
  }
  merged
}
