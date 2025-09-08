//|

///|
/// Tests for VBA project functionality
test "VbaReference creation and methods" {
  let vba_ref = VbaReference::new("TestRef", "Test Description", "/path/to/ref")
  inspect(vba_ref.name, content="TestRef")
  inspect(vba_ref.description, content="Test Description")
  inspect(vba_ref.path, content="/path/to/ref")
  inspect(vba_ref.is_missing(), content="false")
}

///|
test "VbaProject creation and basic methods" {
  let data = b"test vba data"
  let project = VbaProject::new(data)
  inspect(project.data().to_string(), content=(
    #|b"\x74\x65\x73\x74\x20\x76\x62\x61\x20\x64\x61\x74\x61"
  ))
  inspect(project.get_references().length(), content="0")
  inspect(project.get_module_names().length(), content="0")
}

///|
test "VbaProject with metadata" {
  let data = b"test vba data"
  let refs = [VbaReference::new("Ref1", "Description 1", "/path1")]
  let modules = Map::new()
  modules["Module1"] = b"module content 1"
  let project = VbaProject::new_with_metadata(data, refs, modules, "UTF-8")
  inspect(project.get_references().length(), content="1")
  inspect(project.get_module_names().length(), content="1")
  inspect(project.encoding, content="UTF-8")
}

///|
test "VbaProject add_reference" {
  let data = b"test data"
  let project = VbaProject::new(data)
  let vba_ref = VbaReference::new("NewRef", "New Description", "/new/path")
  let updated_project = project.add_reference(vba_ref)
  inspect(updated_project.get_references().length(), content="1")
  inspect(updated_project.get_references()[0].name, content="NewRef")
}

///|
test "VbaProject add_module" {
  let data = b"test data"
  let project = VbaProject::new(data)
  let module_content = b"Sub Test()\nEnd Sub"
  let updated_project = project.add_module("TestModule", module_content)
  inspect(updated_project.get_module_names().length(), content="1")
  inspect(updated_project.get_module_names()[0], content="TestModule")
}

///|
test "VbaProject get_module operations" {
  let data = b"test data"
  let project = VbaProject::new(data)
  let module_content = b"Sub Test()\nEnd Sub"
  let updated_project = project.add_module("TestModule", module_content)

  // Test raw module access
  let raw_module = updated_project.get_module_raw("TestModule")
  inspect(raw_module is Some(_), content="true")

  // Test string module access
  let string_module = updated_project.get_module("TestModule")
  inspect(string_module is Some(_), content="true")

  // Test non-existent module
  let missing_module = updated_project.get_module("NonExistent")
  inspect(missing_module, content="None")
}
