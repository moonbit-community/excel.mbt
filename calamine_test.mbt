///|
test "data_types" {
  let empty_data = Data::Empty
  let int_data = Data::Int(42L)
  let float_data = Data::Float(3.14)
  let string_data = Data::String("test")
  let bool_data = Data::Bool(true)
  inspect(empty_data.is_empty(), content="true")
  inspect(int_data.is_int(), content="true")
  inspect(float_data.is_float(), content="true")
  inspect(string_data.is_string(), content="true")
  inspect(bool_data.is_bool(), content="true")
}

///|
test "dimensions" {
  let dims = Dimensions::{ start: (0U, 0U), end: (2U, 3U) }
  inspect(dims.contains(1U, 2U), content="true")
  inspect(dims.contains(3U, 4U), content="false")
  inspect(dims.len(), content="12")
}

///|
test "excel_datetime" {
  let dt = ExcelDateTime::{
    value: 42000.0,
    datetime_type: DateTime,
    is_1904: false,
  }
  inspect(dt.is_datetime(), content="true")
  inspect(dt.is_duration(), content="false")
  inspect(dt.as_f64(), content="42000")
}
