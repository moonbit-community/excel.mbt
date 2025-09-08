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
/// Get a value at an absolute position in a Range.
pub fn[T] Range::get_value(
  self : Range[T],
  absolute_position : (UInt, UInt),
) -> T? {
  let p = absolute_position
  if p.0 >= self.start.0 &&
    p.0 <= self.end.0 &&
    p.1 >= self.start.1 &&
    p.1 <= self.end.1 {
    let relative_pos = (
      (p.0 - self.start.0).reinterpret_as_int(),
      (p.1 - self.start.1).reinterpret_as_int(),
    )
    return self.get(relative_pos)
  }
  None
}

///|
/// Get a value at a relative position in a Range.
pub fn[T] Range::get(self : Range[T], relative_position : (Int, Int)) -> T? {
  let (row, col) = relative_position
  let (height, width) = self.get_size()
  if col.reinterpret_as_uint() >= width ||
    row.reinterpret_as_uint() >= height ||
    row < 0 ||
    col < 0 {
    None
  } else {
    let idx = row * width.reinterpret_as_int() + col
    if idx >= 0 && idx < self.inner.length() {
      Some(self.inner[idx])
    } else {
      None
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
