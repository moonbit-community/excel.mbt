///|
/// Tests for string parsing utilities

test "parse_int64_advanced basic numbers" {
  inspect(parse_int64_advanced("0"), content="Some(0)")
  inspect(parse_int64_advanced("1"), content="Some(1)")
  inspect(parse_int64_advanced("10"), content="Some(10)")
  inspect(parse_int64_advanced("100"), content="Some(100)")
  inspect(parse_int64_advanced("-1"), content="Some(-1)")
  inspect(parse_int64_advanced("-10"), content="Some(-10)")
}

test "parse_int64_advanced with whitespace" {
  inspect(parse_int64_advanced(" 0"), content="Some(0)")
  inspect(parse_int64_advanced("1 "), content="Some(1)")
  inspect(parse_int64_advanced(" 10 "), content="Some(10)")
  inspect(parse_int64_advanced(" -1 "), content="Some(-1)")
}

test "parse_int64_advanced with leading zeros" {
  inspect(parse_int64_advanced("00"), content="Some(0)")
  inspect(parse_int64_advanced("01"), content="Some(1)")
  inspect(parse_int64_advanced("000"), content="Some(0)")
}

test "parse_int64_advanced invalid cases" {
  inspect(parse_int64_advanced(""), content="None")
  inspect(parse_int64_advanced("abc"), content="None")
  inspect(parse_int64_advanced("12.34"), content="None")
}

test "parse_double_advanced basic numbers" {
  inspect(parse_double_advanced("0"), content="Some(0)")
  inspect(parse_double_advanced("1.0"), content="Some(1)")
  inspect(parse_double_advanced("3.14"), content="Some(3.14)")
  inspect(parse_double_advanced("-1.0"), content="Some(-1)")
  inspect(parse_double_advanced("0.5"), content="Some(0.5)")
}

test "parse_double_advanced various formats" {
  inspect(parse_double_advanced("1"), content="Some(1)")
  inspect(parse_double_advanced("1.00"), content="Some(1)")
  inspect(parse_double_advanced(" 3.14 "), content="None")
  inspect(parse_double_advanced("100.0"), content="Some(100)")
}

test "parse_bool various representations" {
  inspect(parse_bool("true"), content="Some(true)")
  inspect(parse_bool("false"), content="Some(false)")
  inspect(parse_bool("TRUE"), content="Some(true)")
  inspect(parse_bool("False"), content="Some(false)")
  inspect(parse_bool("1"), content="Some(true)")
  inspect(parse_bool("0"), content="Some(false)")
  inspect(parse_bool("yes"), content="Some(true)")
  inspect(parse_bool("no"), content="Some(false)")
  inspect(parse_bool("on"), content="Some(true)")
  inspect(parse_bool("off"), content="Some(false)")
}

test "parse_bool with whitespace" {
  inspect(parse_bool(" true "), content="Some(true)")
  inspect(parse_bool("false "), content="Some(false)")
  inspect(parse_bool(" 1"), content="Some(true)")
}

test "parse_bool invalid cases" {
  inspect(parse_bool("maybe"), content="None")
  inspect(parse_bool(""), content="None")
  inspect(parse_bool("2"), content="None")
}

test "trim_string functionality" {
  inspect(trim_string("  hello  "), content="hello")
  inspect(trim_string("world"), content="world")
  inspect(trim_string("   "), content="")
  inspect(trim_string(""), content="")
  inspect(trim_string(" test"), content="test")
  inspect(trim_string("test "), content="test")
}

test "is_numeric_string detection" {
  inspect(is_numeric_string("123"), content="true")
  inspect(is_numeric_string("12.34"), content="true")
  inspect(is_numeric_string("-123"), content="true")
  inspect(is_numeric_string("+123"), content="true")
  inspect(is_numeric_string("0"), content="true")
  inspect(is_numeric_string("abc"), content="false")
  inspect(is_numeric_string("12a34"), content="false")
  inspect(is_numeric_string("12..34"), content="false")
  inspect(is_numeric_string(""), content="false")
}

test "split_string functionality" {
  let result1 = split_string("a,b,c", ',')
  inspect(result1.length(), content="3")
  inspect(result1[0], content="a")
  inspect(result1[1], content="b")
  inspect(result1[2], content="c")
  
  let result2 = split_string("hello", ',')
  inspect(result2.length(), content="1")
  inspect(result2[0], content="hello")
}

test "join_strings functionality" {
  let strings = ["a", "b", "c"]
  let result = join_strings(strings, ",")
  inspect(result, content="a,b,c")
  
  let empty : Array[String] = []
  let result2 = join_strings(empty, ",")
  inspect(result2, content="")
  
  let single = ["hello"]
  let result3 = join_strings(single, ",")
  inspect(result3, content="hello")
}

test "starts_with and ends_with" {
  inspect(starts_with("hello world", "hello"), content="true")
  inspect(starts_with("hello world", "world"), content="false")
  inspect(starts_with("hello", "hello world"), content="false")
  
  inspect(ends_with("hello world", "world"), content="true")
  inspect(ends_with("hello world", "hello"), content="false")
  inspect(ends_with("world", "hello world"), content="false")
}

test "case conversion" {
  inspect(to_lowercase("Hello World"), content="hello world")
  inspect(to_lowercase("HELLO"), content="hello")
  inspect(to_lowercase("hello"), content="hello")
  
  inspect(to_uppercase("Hello World"), content="HELLO WORLD")
  inspect(to_uppercase("hello"), content="HELLO")
  inspect(to_uppercase("HELLO"), content="HELLO")
}

test "replace_string functionality" {
  inspect(replace_string("hello world", "world", "moon"), content="hello moon")
  inspect(replace_string("abc abc abc", "abc", "xyz"), content="xyz xyz xyz")
  inspect(replace_string("hello", "world", "moon"), content="hello")
  inspect(replace_string("", "a", "b"), content="")
}