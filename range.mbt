///|
/// A struct to hold a cell position and value.
pub(all) struct Cell[T] {
  /// The position for the cell (row, column).
  pos : (UInt, UInt)
  /// The value of the cell.
  val : T
} derive(Show, Eq)

///|
/// Creates a new Cell instance.
pub fn[T] Cell::new(position : (UInt, UInt), value : T) -> Cell[T] {
  { pos: position, val: value }
}

///|
/// Gets Cell position.
pub fn[T] Cell::get_position(self : Cell[T]) -> (UInt, UInt) {
  self.pos
}

///|
/// Gets Cell value.
pub fn[T] Cell::get_value(self : Cell[T]) -> T {
  self.val
}

///|
/// Gets Cell value as a reference.
pub fn[T] Cell::get_value_ref(self : Cell[T]) -> T {
  self.val
}

///|
/// Sets Cell value (creates new Cell with updated value).
pub fn[T] Cell::set_value(self : Cell[T], new_value : T) -> Cell[T] {
  { ..self, val: new_value }
}

///|
/// Gets the row number of the cell (zero-indexed).
pub fn[T] Cell::get_row(self : Cell[T]) -> UInt {
  self.pos.0
}

///|
/// Gets the column number of the cell (zero-indexed).
pub fn[T] Cell::get_column(self : Cell[T]) -> UInt {
  self.pos.1
}

///|
/// A struct which represents an area of cells and the data within it.
pub(all) struct Range[T] {
  start : (UInt, UInt)
  end : (UInt, UInt)
  inner : Array[T]
} derive(Show, Eq)

///|
/// Creates a new Range with default values.
pub fn[T : Default] Range::new(
  start : (UInt, UInt),
  end : (UInt, UInt),
) -> Range[T] {
  if start > end {
    abort("invalid range bounds")
  }
  let height = end.0 - start.0 + 1U
  let width = end.1 - start.1 + 1U
  let size = height * width
  let inner = Array::make(size.reinterpret_as_int(), T::default())
  { start, end, inner }
}

///|
/// Creates a new empty Range.
pub fn[T] Range::empty() -> Range[T] {
  { start: (0U, 0U), end: (0U, 0U), inner: [] }
}

///|
/// Get top left cell position of a Range.
pub fn[T] Range::start(self : Range[T]) -> (UInt, UInt)? {
  if self.is_empty() {
    None
  } else {
    Some(self.start)
  }
}

///|
/// Get bottom right cell position of a Range.
pub fn[T] Range::end(self : Range[T]) -> (UInt, UInt)? {
  if self.is_empty() {
    None
  } else {
    Some(self.end)
  }
}

///|
/// Get the column width of a Range.
pub fn[T] Range::width(self : Range[T]) -> UInt {
  if self.is_empty() {
    0U
  } else {
    self.end.1 - self.start.1 + 1U
  }
}

///|
/// Get the row height of a Range.
pub fn[T] Range::height(self : Range[T]) -> UInt {
  if self.is_empty() {
    0U
  } else {
    self.end.0 - self.start.0 + 1U
  }
}

///|
/// Get size of a Range in (height, width) format.
pub fn[T] Range::get_size(self : Range[T]) -> (UInt, UInt) {
  (self.height(), self.width())
}

///|
/// Check if a Range is empty.
pub fn[T] Range::is_empty(self : Range[T]) -> Bool {
  self.inner.is_empty()
}

///|
/// Get the total number of cells in the range.
pub fn[T] Range::cell_count(self : Range[T]) -> UInt {
  self.width() * self.height()
}

///|
/// Get a value at an absolute position in a Range.
pub fn[T] Range::get_value(
  self : Range[T],
  absolute_position : (UInt, UInt),
) -> T? {
  let (row, col) = absolute_position
  if row >= self.start.0 &&
    row <= self.end.0 &&
    col >= self.start.1 &&
    col <= self.end.1 {
    let relative_row = row - self.start.0
    let relative_col = col - self.start.1
    let index = relative_row * self.width() + relative_col
    self.inner.get(index.reinterpret_as_int())
  } else {
    None
  }
}

///|
/// Get a value at a relative position in a Range.
pub fn[T] Range::get(self : Range[T], relative_position : (UInt, UInt)) -> T? {
  let (row, col) = relative_position
  let (height, width) = self.get_size()
  if col >= width || row >= height {
    None
  } else {
    let index = row * width + col
    self.inner.get(index.reinterpret_as_int())
  }
}

///|
/// Check if a position is within the range bounds.
pub fn[T] Range::contains(self : Range[T], row : UInt, col : UInt) -> Bool {
  row >= self.start.0 &&
  row <= self.end.0 &&
  col >= self.start.1 &&
  col <= self.end.1
}

///|
/// Creates a Range from a sparse vector of cells.
pub fn[T : Default] Range::from_sparse(cells : Array[Cell[T]]) -> Range[T] {
  if cells.is_empty() {
    return Range::empty()
  }

  // Find bounds
  let mut row_start = 4294967295U // UInt max value
  let mut row_end = 0U
  let mut col_start = 4294967295U // UInt max value
  let mut col_end = 0U
  for cell in cells {
    let (r, c) = cell.pos
    if r < row_start {
      row_start = r
    }
    if r > row_end {
      row_end = r
    }
    if c < col_start {
      col_start = c
    }
    if c > col_end {
      col_end = c
    }
  }
  let width = col_end - col_start + 1U
  let height = row_end - row_start + 1U
  let size = (width * height).reinterpret_as_int()
  let inner = Array::make(size, T::default())

  // Fill the array with cell values
  for cell in cells {
    let row = (cell.pos.0 - row_start).reinterpret_as_int()
    let col = (cell.pos.1 - col_start).reinterpret_as_int()
    let idx = row * width.reinterpret_as_int() + col
    if idx >= 0 && idx < size {
      inner[idx] = cell.val
    }
  }
  { start: (row_start, col_start), end: (row_end, col_end), inner }
}

///|
/// Set a value at an absolute position in a Range.
pub fn[T] Range::set_value(
  self : Range[T],
  absolute_position : (UInt, UInt),
  value : T,
) -> Unit {
  if absolute_position.0 < self.start.0 || absolute_position.1 < self.start.1 {
    abort("absolute_position out of bounds")
  }

  // For now, only support setting within existing bounds
  if absolute_position.0 <= self.end.0 && absolute_position.1 <= self.end.1 {
    let pos = (
      absolute_position.0 - self.start.0,
      absolute_position.1 - self.start.1,
    )
    let idx = (pos.0 * self.width() + pos.1).reinterpret_as_int()
    if idx >= 0 && idx < self.inner.length() {
      self.inner[idx] = value
    }
  }
}

///|
/// Row to use as header
pub(all) enum HeaderRow {
  /// First non-empty row
  FirstNonEmptyRow
  /// Index of the header row
  Row(UInt)
} derive(Show, Eq)

///|
/// Common file metadata
pub(all) struct Metadata {
  sheets : Array[Sheet]
  /// Map of sheet names/sheet path within zip archive
  names : Array[(String, String)]
} derive(Show, Eq)

///|
/// Create default metadata
pub fn Metadata::default() -> Metadata {
  { sheets: [], names: [] }
}

///|
/// Add a sheet to metadata
pub fn Metadata::add_sheet(self : Metadata, sheet : Sheet) -> Unit {
  self.sheets.push(sheet)
}

///|
/// Add a name mapping to metadata
pub fn Metadata::add_name(
  self : Metadata,
  name : String,
  path : String,
) -> Unit {
  self.names.push((name, path))
}

///|
/// Get sheet names
pub fn Metadata::sheet_names(self : Metadata) -> Array[String] {
  self.sheets.map(fn(s) { s.name })
}

///|
/// The Table struct represents an Excel worksheet table.
pub(all) struct Table[T] {
  name : String
  sheet_name : String
  columns : Array[String]
  data : Range[T]
} derive(Show, Eq)

///|
/// Get the name of the table.
pub fn[T] Table::name(self : Table[T]) -> String {
  self.name
}

///|
/// Get the name of the parent worksheet for a table.
pub fn[T] Table::sheet_name(self : Table[T]) -> String {
  self.sheet_name
}

///|
/// Get the header names of the table columns.
pub fn[T] Table::columns(self : Table[T]) -> Array[String] {
  self.columns
}

///|
/// Get a range representing the data from the table
pub fn[T] Table::data(self : Table[T]) -> Range[T] {
  self.data
}

///|
/// Convert a Table[T] into a Range[T].
pub fn[T] Table::to_range(self : Table[T]) -> Range[T] {
  self.data
}

///|
/// Get an iterator over the rows of a Range (returns Array of rows).
/// Each row is an ArrayView over the range data.
pub fn[T] Range::rows(self : Range[T]) -> Array[ArrayView[T]] {
  if self.is_empty() {
    return []
  }
  let width = self.width().reinterpret_as_int()
  let rows = []
  let mut start = 0
  while start < self.inner.length() {
    let end = start + width
    if end <= self.inner.length() {
      rows.push(self.inner[start:end])
    }
    start = end
  }
  rows
}

///|
/// Get an iterator over all the cells in a Range.
/// Returns an Array of tuples (row, col, value) where row and col are relative indices.
pub fn[T] Range::cells(self : Range[T]) -> Array[(UInt, UInt, T)] {
  let result = []
  let width = self.width()
  for i = 0; i < self.inner.length(); i = i + 1 {
    let row = (i / width.reinterpret_as_int()).reinterpret_as_uint()
    let col = (i % width.reinterpret_as_int()).reinterpret_as_uint()
    result.push((row, col, self.inner[i]))
  }
  result
}

///|
/// Get an iterator over the used cells in a Range.
/// Returns an Array of tuples (row, col, value) for non-default values only.
pub fn[T : Default + Eq] Range::used_cells(
  self : Range[T],
) -> Array[(UInt, UInt, T)] {
  let result = []
  let width = self.width()
  let default_val = T::default()
  for i = 0; i < self.inner.length(); i = i + 1 {
    let value = self.inner[i]
    if value != default_val {
      let row = (i / width.reinterpret_as_int()).reinterpret_as_uint()
      let col = (i % width.reinterpret_as_int()).reinterpret_as_uint()
      result.push((row, col, value))
    }
  }
  result
}

///|
/// Build a new Range out of the current range (sub-range extraction).
/// Returns a new Range with cloned data for the specified bounds.
pub fn[T : Default] Range::range(
  self : Range[T],
  start : (UInt, UInt),
  end : (UInt, UInt),
) -> Range[T] {
  let other = Range::new(start, end)
  let (self_start_row, self_start_col) = self.start
  let (self_end_row, self_end_col) = self.end
  let (other_start_row, other_start_col) = other.start
  let (other_end_row, other_end_col) = other.end

  // Calculate intersection bounds
  let start_row = if self_start_row > other_start_row {
    self_start_row
  } else {
    other_start_row
  }
  let end_row = if self_end_row < other_end_row {
    self_end_row
  } else {
    other_end_row
  }
  let start_col = if self_start_col > other_start_col {
    self_start_col
  } else {
    other_start_col
  }
  let end_col = if self_end_col < other_end_col {
    self_end_col
  } else {
    other_end_col
  }
  if start_row > end_row || start_col > end_col {
    return other
  }
  let self_width = self.width()
  let other_width = other.width()

  // Copy data from self to other
  let self_row_start = (start_row - self_start_row).reinterpret_as_int()
  let self_row_end = (end_row + 1U - self_start_row).reinterpret_as_int()
  let self_col_start = (start_col - self_start_col).reinterpret_as_int()
  let self_col_end = (end_col + 1U - self_start_col).reinterpret_as_int()
  let other_row_start = (start_row - other_start_row).reinterpret_as_int()
  let other_row_end = (end_row + 1U - other_start_row).reinterpret_as_int()
  let other_col_start = (start_col - other_start_col).reinterpret_as_int()
  let other_col_end = (end_col + 1U - other_start_col).reinterpret_as_int()
  for r = 0; r < other_row_end - other_row_start; r = r + 1 {
    let self_row_idx = self_row_start + r
    let other_row_idx = other_row_start + r
    for c = 0; c < other_col_end - other_col_start; c = c + 1 {
      let self_col_idx = self_col_start + c
      let other_col_idx = other_col_start + c
      let self_idx = self_row_idx * self_width.reinterpret_as_int() +
        self_col_idx
      let other_idx = other_row_idx * other_width.reinterpret_as_int() +
        other_col_idx
      if self_idx >= 0 &&
        self_idx < self.inner.length() &&
        other_idx >= 0 &&
        other_idx < other.inner.length() {
        other.inner[other_idx] = self.inner[self_idx]
      }
    }
  }
  other
}

///|
/// Implementation of indexing for Range rows.
/// Get a row by its relative index as an ArrayView.
pub fn[T] Range::get_row(self : Range[T], index : UInt) -> ArrayView[T]? {
  if index >= self.height() {
    return None
  }
  let width = self.width().reinterpret_as_int()
  let start = (index * self.width()).reinterpret_as_int()
  let end = start + width
  if end <= self.inner.length() {
    Some(self.inner[start:end])
  } else {
    None
  }
}

///|
/// Implementation of indexing for Range cells.
/// Get a cell value by its relative position (row, column).
pub fn[T] Range::get_cell(self : Range[T], position : (UInt, UInt)) -> T? {
  self.get(position)
}

///|
/// Set a cell value by its relative position (row, column).
pub fn[T] Range::set_cell(
  self : Range[T],
  position : (UInt, UInt),
  value : T,
) -> Unit {
  let (row, col) = position
  let (height, width) = self.get_size()
  if col < width && row < height {
    let index = (row * width + col).reinterpret_as_int()
    if index >= 0 && index < self.inner.length() {
      self.inner[index] = value
    }
  }
}

///|
/// Get headers for a Range (first row as strings).
/// Returns the first row of the range as an optional vector of strings.
pub fn[T : Show] Range::headers(self : Range[T]) -> Array[String]? {
  let rows = self.rows()
  if rows.length() > 0 {
    let first_row = rows[0]
    let headers = []
    for i = 0; i < first_row.length(); i = i + 1 {
      headers.push(first_row[i].to_string())
    }
    Some(headers)
  } else {
    None
  }
}
