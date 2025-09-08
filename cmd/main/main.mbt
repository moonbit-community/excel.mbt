///|
fn main {
  println("=== MoonBit Calamine Library Demo ===")

  // Demonstrate core data types
  println("\n1. Data Types Demo:")
  demo_data_types()

  // Demonstrate dimensions and ranges
  println("\n2. Dimensions and Ranges Demo:")
  demo_dimensions_and_ranges()

  // Demonstrate error handling
  println("\n3. Error Handling Demo:")
  demo_error_handling()

  // Demonstrate workbook type detection
  println("\n4. Workbook Type Detection Demo:")
  demo_workbook_detection()
  println("\n=== Demo Complete ===")
}

///|
fn demo_data_types() -> Unit {
  // Create various cell data types
  let int_data = @lib.Data::Int(42L)
  let float_data = @lib.Data::Float(3.14159)
  let string_data = @lib.Data::String("Hello Excel")
  let bool_data = @lib.Data::Bool(true)
  let empty_data = @lib.Data::Empty
  println("- Integer cell: \{int_data.is_int()}, value: \{int_data.get_int()}")
  println(
    "- Float cell: \{float_data.is_float()}, value: \{float_data.get_float()}",
  )
  println(
    "- String cell: \{string_data.is_string()}, value: \{string_data.get_string()}",
  )
  println(
    "- Boolean cell: \{bool_data.is_bool()}, value: \{bool_data.get_bool()}",
  )
  println("- Empty cell: \{empty_data.is_empty()}")

  // Demonstrate type conversion
  let converted_int = float_data.as_i64()
  let converted_float = int_data.as_f64()
  println("- Float to int: \{converted_int}")
  println("- Int to float: \{converted_float}")
}

///|
fn demo_dimensions_and_ranges() -> Unit {
  // Create dimensions
  let dims = @lib.Dimensions::new((0U, 0U), (5U, 10U))
  println("- Range dimensions: \{dims.width()} x \{dims.height()}")
  println("- Contains (2,5): \{dims.contains(2U, 5U)}")
  println("- Contains (6,11): \{dims.contains(6U, 11U)}")
  println("- Total cells: \{dims.len()}")

  // Create a simple range
  let range : @lib.Range[String] = @lib.Range::empty()
  println("- Empty range is empty: \{range.is_empty()}")

  // Create cells and sparse range
  let cells = [
    @lib.Cell::new((0U, 0U), "A1"),
    @lib.Cell::new((0U, 1U), "B1"),
    @lib.Cell::new((1U, 0U), "A2"),
  ]
  let sparse_range = @lib.Range::from_sparse(cells)
  println(
    "- Sparse range size: \{sparse_range.width()} x \{sparse_range.height()}",
  )
  println("- Cell at (0,0): \{sparse_range.get_value((0U, 0U))}")
  println("- Cell at (0,1): \{sparse_range.get_value((0U, 1U))}")
}

///|
fn demo_error_handling() -> Unit {
  // Create various error types
  let xlsx_error = @lib.XlsxError::Password
  let xls_error = @lib.XlsError::InvalidBOM
  let io_error = @lib.CalamineError::Io("File not found")
  println("- XLSX Error: \{xlsx_error}")
  println("- XLS Error: \{xls_error}")
  println("- General Error: \{io_error}")

  // Convert to general error
  let general_xlsx = xlsx_error.to_error()
  let general_xls = xls_error.to_error()
  println("- Converted XLSX Error: \{general_xlsx}")
  println("- Converted XLS Error: \{general_xls}")
}

///|
fn demo_workbook_detection() -> Unit {
  // Test workbook type detection
  let xlsx_data = b"PK\x03\x04"
  let xls_data = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
  let unknown_data = b"INVALID"
  let xlsx_type = @lib.detect_workbook_type(xlsx_data)
  let xls_type = @lib.detect_workbook_type(xls_data)
  let unknown_type = @lib.detect_workbook_type(unknown_data)
  println("- ZIP signature detected as: \{xlsx_type}")
  println("- OLE2 signature detected as: \{xls_type}")
  println("- Unknown signature detected as: \{unknown_type}")

  // Demonstrate VBA project handling
  let vba_data = b"VBA Project Data"
  let vba = @lib.VbaProject::new(vba_data)
  println("- VBA project data length: \{vba.data().length()}")
}
