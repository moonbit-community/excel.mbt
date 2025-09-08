//|

///|
/// Tests for utility functions
test "read_u32 little endian" {
  let data = b"\x44\x43\x42\x41" // ABCD in little endian
  let result = read_u32(data, 0)
  inspect(result, content="Some(1094861636)")
}

///|
test "read_u32 bounds check" {
  let data = b"\x01\x02" // Only 2 bytes
  let result = read_u32(data, 0)
  inspect(result, content="None")
}

///|
test "read_u16 little endian" {
  let data = b"\x42\x41" // AB in little endian
  let result = read_u16(data, 0)
  inspect(result, content="Some(16706)")
}

///|
test "read_i32 conversion" {
  let data = b"\xFF\xFF\xFF\xFF" // -1 in two's complement
  let result = read_i32(data, 0)
  inspect(result, content="Some(-1)")
}

///|
test "column_to_string basic" {
  let result = column_to_string(0U)
  inspect(result, content="A")
  let result2 = column_to_string(25U)
  inspect(result2, content="Z")
}

///|
test "column_to_string multi char" {
  let result = column_to_string(26U)
  inspect(result, content="AA")
  let result2 = column_to_string(51U)
  inspect(result2, content="AZ")
}

///|
// TODO: Re-implement function table and then uncomment this test
// test "get_function_name" {
//   let result = get_function_name(0U)
//   inspect(result, content="Some(\"COUNT\")")
//   let result2 = get_function_name(1000U)
//   inspect(result2, content="None")
// }

///|
test "read_f64 double conversion" {
  // Test a simple double value (1.0)
  let data = b"\x00\x00\x00\x00\x00\x00\xF0\x3F" // 1.0 in IEEE 754 little endian
  let result = read_f64(data, 0)
  inspect(result is Some(_), content="true")
}
