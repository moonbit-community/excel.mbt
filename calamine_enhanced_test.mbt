///|
/// Enhanced tests for Data type with improved string parsing

test "Data type detection methods" {
  let int_data = Data::Int(42L)
  inspect(int_data.is_int(), content="true")
  inspect(int_data.is_float(), content="false")
  
  let float_data = Data::Float(3.14)
  inspect(float_data.is_float(), content="true")
  inspect(float_data.is_int(), content="false")
  
  let bool_data = Data::Bool(true)
  inspect(bool_data.is_bool(), content="true")
  
  let string_data = Data::String("hello")
  inspect(string_data.is_string(), content="true")
  
  let empty_data = Data::Empty
  inspect(empty_data.is_empty(), content="true")
}

test "Data as_string conversion" {
  let int_data = Data::Int(42L)
  inspect(int_data.as_string(), content="Some(\"42\")")
  
  let float_data = Data::Float(3.14)
  inspect(float_data.as_string(), content="Some(\"3.14\")")
  
  let bool_data = Data::Bool(true)
  inspect(bool_data.as_string(), content="Some(\"true\")")
  
  let string_data = Data::String("hello")
  inspect(string_data.as_string(), content="Some(\"hello\")")
  
  let empty_data = Data::Empty
  inspect(empty_data.as_string(), content="Some(\"\")")
}

test "Data as_i64 conversion with string parsing" {
  let int_data = Data::Int(42L)
  inspect(int_data.as_i64(), content="Some(42)")
  
  let float_data = Data::Float(3.14)
  inspect(float_data.as_i64(), content="Some(3)")
  
  let bool_true = Data::Bool(true)
  inspect(bool_true.as_i64(), content="Some(1)")
  
  let bool_false = Data::Bool(false)
  inspect(bool_false.as_i64(), content="Some(0)")
  
  let string_zero = Data::String("0")
  inspect(string_zero.as_i64(), content="Some(0)")
  
  let string_one = Data::String("1")
  inspect(string_one.as_i64(), content="Some(1)")
  
  let string_invalid = Data::String("abc")
  inspect(string_invalid.as_i64(), content="None")
}

test "Data as_f64 conversion with string parsing" {
  let int_data = Data::Int(42L)
  inspect(int_data.as_f64(), content="Some(42.0)")
  
  let float_data = Data::Float(3.14)
  inspect(float_data.as_f64(), content="Some(3.14)")
  
  let string_zero = Data::String("0.0")
  inspect(string_zero.as_f64(), content="Some(0.0)")
  
  let string_one = Data::String("1.0")
  inspect(string_one.as_f64(), content="Some(1.0)")
  
  let string_invalid = Data::String("not_a_number")
  inspect(string_invalid.as_f64(), content="None")
}

test "CellErrorType to_string" {
  let div0 = CellErrorType::Div0
  inspect(div0.to_string(), content="#DIV/0!")
  
  let na = CellErrorType::NA
  inspect(na.to_string(), content="#N/A")
  
  let name = CellErrorType::Name
  inspect(name.to_string(), content="#NAME?")
  
  let ref_err = CellErrorType::Ref
  inspect(ref_err.to_string(), content="#REF!")
}

test "DataRef string methods" {
  let int_data = DataRef::Int(42L)
  inspect(int_data.is_int(), content="true")
  inspect(int_data.as_i64(), content="Some(42)")
  
  let shared_string = DataRef::SharedString("shared")
  inspect(shared_string.is_string(), content="true")
  inspect(shared_string.get_string(), content="Some(\"shared\")")
  inspect(shared_string.as_string(), content="Some(\"shared\")")
  
  let string_data = DataRef::String("regular")
  inspect(string_data.is_string(), content="true")
  inspect(string_data.get_string(), content="Some(\"regular\")")
}

test "Excel datetime handling" {
  let excel_dt = ExcelDateTime::new(25569.0, ExcelDateTimeType::DateTime, false)
  inspect(excel_dt.is_datetime(), content="true")
  inspect(excel_dt.is_duration(), content="false")
  inspect(excel_dt.as_f64(), content="25569.0")
  
  let excel_dur = ExcelDateTime::new(0.5, ExcelDateTimeType::TimeDelta, false)
  inspect(excel_dur.is_duration(), content="true")
  inspect(excel_dur.is_datetime(), content="false")
}

test "Data datetime methods" {
  let excel_dt = ExcelDateTime::new(25569.0, ExcelDateTimeType::DateTime, false)
  let dt_data = Data::DateTime(excel_dt)
  inspect(dt_data.is_datetime(), content="true")
  inspect(dt_data.get_datetime().is_some(), content="true")
  
  let iso_dt = Data::DateTimeIso("2023-01-01T12:00:00")
  inspect(iso_dt.is_datetime_iso(), content="true")
  inspect(iso_dt.get_datetime_iso(), content="Some(\"2023-01-01T12:00:00\")")
  
  let iso_dur = Data::DurationIso("PT1H30M")
  inspect(iso_dur.is_duration_iso(), content="true")
  inspect(iso_dur.get_duration_iso(), content="Some(\"PT1H30M\")")
}