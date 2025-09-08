///|
/// Tests for VBA project functionality

///|
test "VbaReference creation and methods" {
  let vba_ref = VbaReference::new("TestRef", "Test Description", "/path/to/ref")
  inspect(vba_ref.name, content="TestRef")
  inspect(vba_ref.description, content="Test Description")
  inspect(vba_ref.path, content="/path/to/ref")
  inspect(vba_ref.is_missing(), content="false")
}

///|
test "VbaProject creation and basic methods" {
  let project = VbaProject::create_empty()
  inspect(project.get_references().length(), content="0")
  inspect(project.get_module_names().length(), content="0")
  inspect(project.get_module_names()[0], content="Module1")
}

///|
test "VbaProject error handling" {
  // Test VbaError types
  let io_error = VbaError::Io("Test IO error")
  let module_not_found = VbaError::ModuleNotFound("TestModule")
  let invalid_record = VbaError::InvalidRecordId(0x0016U, 0x0017U)
  inspect(io_error, content="Io(\"Test IO error\")")
  inspect(module_not_found, content="ModuleNotFound(\"TestModule\")")
  inspect(invalid_record, content="InvalidRecordId(22, 23)")
}
