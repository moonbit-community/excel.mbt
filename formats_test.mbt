///|
/// Tests for format detection functionality

test "detect_custom_number_format date formats" {
  let result1 = detect_custom_number_format("DD/MM/YY")
  inspect(result1, content="DateTime")
  
  let result2 = detect_custom_number_format("H:MM:SS;@")
  inspect(result2, content="DateTime")
  
  let result3 = detect_custom_number_format("m\"M\"d\"D\";@")
  inspect(result3, content="DateTime")
}

test "detect_custom_number_format duration formats" {
  let result1 = detect_custom_number_format("[h]:mm:ss")
  inspect(result1, content="TimeDelta")
  
  let result2 = detect_custom_number_format("[h]")
  inspect(result2, content="TimeDelta")
  
  let result3 = detect_custom_number_format("[ss]")
  inspect(result3, content="TimeDelta")
  
  let result4 = detect_custom_number_format("[s].000")
  inspect(result4, content="TimeDelta")
}

test "detect_custom_number_format other formats" {
  let result1 = detect_custom_number_format("#,##0\\ [$\\u20bd-46D]")
  inspect(result1, content="Other")
  
  let result2 = detect_custom_number_format("\"Y: \"0.00\"m\";\"Y: \"-0.00\"m\";\"Y: <num>m\";@")
  inspect(result2, content="Other")
  
  let result3 = detect_custom_number_format("\"$\"#,##0_);[Red](\"$\"#,##0)")
  inspect(result3, content="Other")
}

test "builtin_format_by_id date formats" {
  let result1 = builtin_format_by_id("14")
  inspect(result1, content="DateTime")
  
  let result2 = builtin_format_by_id("18")
  inspect(result2, content="DateTime")
  
  let result3 = builtin_format_by_id("46")
  inspect(result3, content="TimeDelta")
}

test "builtin_format_by_code date formats" {
  let result1 = builtin_format_by_code(14U)
  inspect(result1, content="DateTime")
  
  let result2 = builtin_format_by_code(46U)
  inspect(result2, content="TimeDelta")
  
  let result3 = builtin_format_by_code(100U)
  inspect(result3, content="Other")
}

test "format_excel_i64 conversion" {
  let result1 = format_excel_i64(25569L, Some(DateTime), false)
  match result1 {
    DateTime(_) => inspect(true, content="true")
    _ => inspect(false, content="true")
  }
  
  let result2 = format_excel_i64(123L, None, false)
  inspect(result2, content="Int(123)")
}

test "format_excel_f64 conversion" {
  let result1 = format_excel_f64(25569.5, Some(DateTime), false)
  match result1 {
    DateTime(_) => inspect(true, content="true")
    _ => inspect(false, content="true")
  }
  
  let result2 = format_excel_f64(123.45, None, false)
  inspect(result2, content="Float(123.45)")
}