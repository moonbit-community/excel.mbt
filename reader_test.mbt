///|
test "vba_project_creation" {
  let vba = VbaProject::create_empty()
  inspect(vba.get_references().length(), content="0")
  inspect(vba.get_module_names().length(), content="0")
}

///|
test "workbook_type_detection" {
  // Test ZIP signature detection (XLSX)
  let xlsx_data = b"PK\x03\x04"
  let detected = detect_workbook_type(xlsx_data)
  inspect(detected, content="Unknown")

  // Test OLE2 signature detection (XLS) 
  let xls_data = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
  let detected_xls = detect_workbook_type(xls_data)
  inspect(detected_xls, content="Xls")

  // Test unknown format
  let unknown_data = b"INVALID"
  let unknown = detect_workbook_type(unknown_data)
  inspect(unknown, content="Unknown")
}

///|
test "xlsx_reader_struct" {
  let data = b"PK\x03\x04test_data"
  let reader = XlsxReader::{
    data,
    header_row: FirstNonEmptyRow,
    metadata: Metadata::default(),
  }
  inspect(reader.data.length(), content="13")
  inspect(reader.header_row, content="FirstNonEmptyRow")
}
