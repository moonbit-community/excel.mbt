///|
test "calamine_error_types" {
  let io_error = CalamineError::Io("IO error occurred")
  let xlsx_error = CalamineError::Xlsx(XlsxError::Password)
  let xls_error = CalamineError::Xls(XlsError::InvalidBOM)
  
  inspect(io_error, content="Io(\"IO error occurred\")")
  inspect(xlsx_error, content="Xlsx(Password)")
  inspect(xls_error, content="Xls(InvalidBOM)")
}

///|
test "xlsx_error_types" {
  let parse_error = XlsxError::Parse("Invalid XML")
  let file_not_found = XlsxError::FileNotFound("sheet1.xml")
  let dimension_error = XlsxError::DimensionCount(5U)
  
  inspect(parse_error, content="Parse(\"Invalid XML\")")
  inspect(file_not_found, content="FileNotFound(\"sheet1.xml\")")
  inspect(dimension_error, content="DimensionCount(5)")
}

///|
test "error_conversion" {
  let xlsx_error = XlsxError::Password
  let general_error = xlsx_error.to_error()
  
  inspect(general_error, content="Xlsx(Password)")
  
  let vba_error = VbaError::InvalidDirectory
  let converted_vba = vba_error.to_error()
  
  inspect(converted_vba, content="Vba(InvalidDirectory)")
}