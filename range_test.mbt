///|
test "cell_creation" {
  let cell = Cell::new((1U, 2U), "test")
  inspect(cell.get_position(), content="(1, 2)")
  inspect(cell.get_value(), content="test")
}

///|
test "range_creation" {
  let range : Range[String] = Range::new((0U, 0U), (2U, 2U))
  inspect(range.width(), content="3")
  inspect(range.height(), content="3")
  inspect(range.get_size(), content="(3, 3)")
  inspect(range.is_empty(), content="false")
}

///|
test "empty_range" {
  let range : Range[String] = Range::empty()
  inspect(range.is_empty(), content="true")
  inspect(range.start(), content="None")
  inspect(range.end(), content="None")
}

///|
test "range_from_sparse" {
  let cells = [
    Cell::new((1U, 1U), "A"),
    Cell::new((1U, 2U), "B"),
    Cell::new((2U, 1U), "C"),
  ]
  let range = Range::from_sparse(cells)
  
  inspect(range.width(), content="2")
  inspect(range.height(), content="2")
  inspect(range.get_value((1U, 1U)), content="Some(\"A\")")
  inspect(range.get_value((1U, 2U)), content="Some(\"B\")")
  inspect(range.get_value((2U, 1U)), content="Some(\"C\")")
}

///|
test "metadata_operations" {
  let metadata = Metadata::default()
  let sheet = Sheet::{
    name: "Sheet1",
    typ: WorkSheet,
    visible: Visible,
  }
  
  metadata.add_sheet(sheet)
  metadata.add_name("Sheet1", "/xl/worksheets/sheet1.xml")
  
  let names = metadata.sheet_names()
  inspect(names.length(), content="1")
  inspect(names[0], content="Sheet1")
}

///|
test "table_operations" {
  let range : Range[String] = Range::empty()
  let table = Table::{
    name: "Table1",
    sheet_name: "Sheet1",
    columns: ["Col1", "Col2"],
    data: range,
  }
  
  inspect(table.name(), content="Table1")
  inspect(table.sheet_name(), content="Sheet1")
  inspect(table.columns().length(), content="2")
}