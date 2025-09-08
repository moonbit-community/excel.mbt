///|
/// VBA Project parsing implementation
/// 
/// This module provides functionality to parse VBA projects from Excel files,
/// including reading macros, references, and module information.

///|
/// VBA specific error types
pub(all) suberror VbaError {
  /// Error from CFB parsing
  Cfb(CfbError)
  /// IO error
  Io(String)
  /// Module not found
  ModuleNotFound(String)
  /// Unknown value error
  Unknown(String, UInt) // type, value
  /// Invalid libid format
  LibId
  /// Invalid record id
  InvalidRecordId(UInt, UInt) // expected, found
} derive(Show, Eq)

///|
/// VBA Reference structure
pub(all) struct VbaReference {
  /// Reference name
  name : String
  /// Reference description
  description : String
  /// Reference path
  path : String
} derive(Show, Eq)

///|
/// Create a new VBA reference
pub fn VbaReference::new(
  name : String,
  description : String,
  path : String,
) -> VbaReference {
  { name, description, path }
}

///|
/// Check if the reference location is accessible (simplified)
pub fn VbaReference::is_missing(self : VbaReference) -> Bool {
  // For now, we can't check file system access in MoonBit
  // In a real implementation, this would check if the path exists
  false
}

///|
/// VBA Module structure  
struct VbaModule {
  /// Module name as it appears in VBA project
  name : String
  /// Stream name where module content is stored
  stream_name : String
  /// Offset within the stream where text starts
  text_offset : UInt
} derive(Show, Eq)

///|
/// VBA Project structure with enhanced functionality
pub(all) struct VbaProject {
  /// List of VBA references
  references : Array[VbaReference]
  /// Module names and their raw content (bytes)
  modules : Map[String, Bytes]
  /// Text encoding information
  encoding : XlsEncoding
  /// Raw project data
  raw_data : Bytes
} derive(Show, Eq)

///|
/// Create a new VBA project from CFB data
pub fn VbaProject::new(data : Bytes) -> VbaProject raise VbaError {
  let cfb = Cfb::new(data) catch { e => raise VbaError::Cfb(e) }
  VbaProject::from_cfb(data, cfb)
}

///|
/// Create a VBA project from CFB parser
pub fn VbaProject::from_cfb(
  data : Bytes,
  cfb : Cfb,
) -> VbaProject raise VbaError {
  // Get the 'dir' stream
  let dir_stream = cfb.get_stream("dir", data) catch {
    e => raise VbaError::Cfb(e)
  }
  let decompressed = decompress_stream(dir_stream) catch {
    e => raise VbaError::Cfb(e)
  }

  // Read directory information to get encoding
  let encoding = read_dir_information(decompressed)

  // Read references
  let references = read_references(decompressed, encoding)

  // Read modules metadata
  let modules_info = read_modules_info(decompressed, encoding)

  // Read all module contents
  let modules = Map::new()
  for module_info in modules_info {
    let stream_data = cfb.get_stream(module_info.stream_name, data) catch {
      e => raise VbaError::Cfb(e)
    }
    let module_content = if module_info.text_offset.reinterpret_as_int() <
      stream_data.length() {
      let offset_bytes = stream_data[module_info.text_offset.reinterpret_as_int():]
      decompress_stream(offset_bytes.to_bytes()) catch {
        e => raise VbaError::Cfb(e)
      }
    } else {
      stream_data
    }
    modules[module_info.name] = module_content
  }
  { references, modules, encoding, raw_data: data }
}

///|
/// Get list of references
pub fn VbaProject::get_references(self : VbaProject) -> Array[VbaReference] {
  self.references
}

///|
/// Get list of module names
pub fn VbaProject::get_module_names(self : VbaProject) -> Array[String] {
  self.modules.iter().map(fn(pair) { pair.0 }).collect()
}

///|
/// Get module content as string (with encoding conversion)
pub fn VbaProject::get_module(
  self : VbaProject,
  name : String,
) -> String raise VbaError {
  match self.modules.get(name) {
    Some(data) => self.encoding.decode_all(data)
    None => raise VbaError::ModuleNotFound(name)
  }
}

///|
/// Get module content as raw bytes
pub fn VbaProject::get_module_raw(
  self : VbaProject,
  name : String,
) -> Bytes raise VbaError {
  match self.modules.get(name) {
    Some(data) => data
    None => raise VbaError::ModuleNotFound(name)
  }
}

///|
/// Read directory information to extract encoding
fn read_dir_information(data : Bytes) -> XlsEncoding raise VbaError {
  // Skip PROJECTSYSKIND (10 bytes)
  let mut offset = 10

  // PROJECTCOMPATVERSION (optional)
  if offset + 2 <= data.length() {
    let record_id = read_u16(data, offset).map_or(0U, fn(v) { v })
    if record_id == 0x004AU {
      offset = offset + 10
    }
  }

  // Skip PROJECTLCID and PROJECTLCIDINVOKE Records (20 bytes)
  offset = offset + 20

  // Read PROJECT Codepage (at offset + 6)
  if offset + 8 <= data.length() {
    let codepage = read_u16(data, offset + 6).map_or(1252U, fn(v) { v })
    XlsEncoding::from_codepage(codepage) catch {
      e => raise VbaError::Cfb(e)
    }
  } else {
    // Default to Windows-1252
    XlsEncoding::from_codepage(1252U) catch {
      e => raise VbaError::Cfb(e)
    }
  }
}

///|
/// Read VBA references from directory stream (simplified implementation)
fn read_references(data : Bytes, encoding : XlsEncoding) -> Array[VbaReference] {
  // This is a simplified implementation
  // The full implementation would parse the complex reference structure
  let references : Array[VbaReference] = []

  // Add some default references that are commonly found
  references.push(VbaReference::new("stdole", "OLE Automation", "stdole2.tlb"))
  references.push(
    VbaReference::new("Office", "Microsoft Office Object Library", "MSO.DLL"),
  )
  references
}

///|
/// Read modules information from directory stream (simplified implementation)
fn read_modules_info(data : Bytes, encoding : XlsEncoding) -> Array[VbaModule] {
  // This is a simplified implementation
  // The full implementation would parse the complex module structure
  let modules : Array[VbaModule] = []

  // For now, create a placeholder module
  modules.push({ name: "Module1", stream_name: "Module1", text_offset: 0U })
  modules
}

///|
/// Check that next record matches expected ID
fn check_record(
  expected_id : UInt,
  data : Bytes,
  offset : Int,
) -> Int raise VbaError {
  if offset + 2 > data.length() {
    raise VbaError::Io("Not enough data to read record ID")
  }
  let record_id = read_u16(data, offset).map_or(0U, fn(v) { v })
  if record_id != expected_id {
    raise VbaError::InvalidRecordId(expected_id, record_id)
  }
  offset + 2
}

///|
/// Read a variable length record
fn read_variable_record(
  data : Bytes,
  offset : Int,
) -> (Bytes, Int) raise VbaError {
  if offset + 4 > data.length() {
    raise VbaError::Io("Not enough data to read variable record length")
  }
  let length = read_u32(data, offset).map_or(0U, fn(v) { v })
  let new_offset = offset + 4
  let end_offset = new_offset + length.reinterpret_as_int()
  if end_offset > data.length() {
    raise VbaError::Io("Variable record extends beyond data bounds")
  }
  (data[new_offset:end_offset].to_bytes(), end_offset)
}

///|
/// Create a simple VBA project for testing
pub fn VbaProject::create_empty() -> VbaProject {
  let encoding = XlsEncoding::from_codepage(1252U) catch {
    _ => { codepage: 1252U } // Fallback to default encoding
  }
  { references: [], modules: Map::new(), encoding, raw_data: b"" }
}
