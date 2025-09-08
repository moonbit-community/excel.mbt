///|
fn main {
  // Create some sample Excel data using the Data type
  let cell_data = @lib.Data::Int(42L)
  let empty_cell = @lib.Data::Empty
  let text_cell = @lib.Data::String("Hello Excel")
  println("Cell data is int: \{cell_data.is_int()}")
  println("Empty cell is empty: \{empty_cell.is_empty()}")
  println("Text cell contains: \{text_cell.get_string()}")

  // Create dimensions using the correct constructor syntax
  let dims = @lib.Dimensions::{ start: (0U, 0U), end: (5U, 10U) }
  println("Dimensions contain (2,5): \{dims.contains(2U, 5U)}")
  println("Dimensions length: \{dims.len()}")
}
