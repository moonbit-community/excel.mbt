///|
/// Enhanced tests for Cell and Range functionality
test "Cell utility methods" {
  let cell = Cell::new((5U, 10U), Data::Int(42L))
  inspect(cell.get_row(), content="5")
  inspect(cell.get_column(), content="10")
  inspect(cell.get_position(), content="(5, 10)")
  inspect(cell.get_value(), content="Int(42)")
}

///|
test "Cell set_value creates new cell" {
  let original = Cell::new((2U, 3U), Data::String("original"))
  let updated = original.set_value(Data::String("updated"))

  // Original should be unchanged
  inspect(original.get_value(), content="String(\"original\")")
  // Updated should have new value but same position
  inspect(updated.get_value(), content="String(\"updated\")")
  inspect(updated.get_position(), content="(2, 3)")
}

///|
test "Range basic properties" {
  let range = Range::new((1U, 1U), (3U, 3U))
  inspect(range.width(), content="3")
  inspect(range.height(), content="3")
  inspect(range.cell_count(), content="9")
  inspect(range.get_size(), content="(3, 3)")
  inspect(range.is_empty(), content="false")
}

///|
test "Range empty" {
  let empty_range : Range[Data] = Range::empty()
  inspect(empty_range.is_empty(), content="true")
  inspect(empty_range.width(), content="0")
  inspect(empty_range.height(), content="0")
  inspect(empty_range.cell_count(), content="0")
  inspect(empty_range.start(), content="None")
  inspect(empty_range.end(), content="None")
}

///|
test "Range start and end positions" {
  let range = Range::new((2U, 5U), (4U, 8U))
  inspect(range.start(), content="Some((2, 5))")
  inspect(range.end(), content="Some((4, 8))")
}

///|
test "Range contains method" {
  let range = Range::new((2U, 3U), (5U, 7U))

  // Within bounds
  inspect(range.contains(2U, 3U), content="true") // Top-left corner
  inspect(range.contains(5U, 7U), content="true") // Bottom-right corner
  inspect(range.contains(3U, 5U), content="true") // Middle

  // Outside bounds
  inspect(range.contains(1U, 3U), content="false") // Above
  inspect(range.contains(2U, 2U), content="false") // Left
  inspect(range.contains(6U, 5U), content="false") // Below
  inspect(range.contains(3U, 8U), content="false") // Right
}

///|
test "Range get and get_value methods" {
  let range = Range::new((0U, 0U), (1U, 1U))

  // Test relative positioning
  let value_rel = range.get((0U, 0U))
  inspect(value_rel is Some(_), content="true")

  // Test absolute positioning  
  let value_abs = range.get_value((0U, 0U))
  inspect(value_abs is Some(_), content="true")

  // Test out of bounds
  let out_of_bounds = range.get((5U, 5U))
  inspect(out_of_bounds, content="None")
  let out_of_bounds_abs = range.get_value((10U, 10U))
  inspect(out_of_bounds_abs, content="None")
}
