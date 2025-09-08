///|
/// Compound File Binary format implementation for MS-CFB
/// 
/// This module handles the parsing of Microsoft Compound File Binary format,
/// which is used by older Excel files (.xls) and other Office documents.

///|
/// CFB specific error types
pub(all) suberror CfbError {
  /// IO error
  Io(String)
  /// Invalid OLE signature
  Ole
  /// Empty root directory
  EmptyRootDir
  /// Stream not found
  StreamNotFound(String)
  /// Invalid value error
  Invalid(String, String, UInt) // name, expected, found
  /// Codepage not found
  CodePageNotFound(UInt)
} derive(Show, Eq)

///|
/// Special sector types in CFB format
const RESERVED_SECTORS : UInt = 0xFFFFFFFAU

///|
const DIFSECT : UInt = 0xFFFFFFFCU

///|
const ENDOFCHAIN : UInt = 0xFFFFFFFEU

///|
/// A struct for managing Compound File Binary format
pub(all) struct Cfb {
  directories : Array[Directory]
  sectors : Sectors
  fats : Array[UInt]
  mini_sectors : Sectors
  mini_fats : Array[UInt]
} derive(Show, Eq)

///|
/// CFB header structure
struct Header {
  version : UInt
  sector_size : UInt
  dir_len : UInt
  dir_start : UInt
  fat_len : UInt
  mini_fat_len : UInt
  mini_fat_start : UInt
  difat_start : UInt
} derive(Show, Eq)

///|
/// Sector management structure
struct Sectors {
  data : Bytes
  size : UInt
} derive(Show, Eq)

///|
/// Directory entry structure
pub(all) struct Directory {
  name : String
  start : UInt
  len : UInt
} derive(Show, Eq)

///|
/// Create a new CFB parser from bytes
pub fn Cfb::new(data : Bytes) -> Cfb raise CfbError {
  // Parse header
  let (header, initial_difat) = Header::from_bytes(data)
  let sectors = Sectors::new(header.sector_size, b"")

  // Load DIFAT sectors
  let mut sector_id = header.difat_start
  let difat : Array[UInt] = []
  for val in initial_difat {
    difat.push(val)
  }
  while sector_id < RESERVED_SECTORS {
    let sector_data = sectors.get(sector_id, data)
    let sector_u32s = to_u32_from_bytes(sector_data)
    for val in sector_u32s {
      difat.push(val)
    }
    sector_id = match difat.pop() {
      Some(id) => id
      None => break
    }
  }

  // Load the FATs
  let fats : Array[UInt] = []
  for id in difat {
    if id < DIFSECT {
      let sector_data = sectors.get(id, data)
      let sector_u32s = to_u32_from_bytes(sector_data)
      for val in sector_u32s {
        fats.push(val)
      }
    }
  }

  // Get directory sectors
  let dirs_data = sectors.get_chain(
    header.dir_start,
    fats,
    data,
    header.dir_len * header.sector_size,
  )
  let directories = parse_directories(dirs_data, header.sector_size)
  if directories.is_empty() ||
    (header.version != 3U && directories[0].start == ENDOFCHAIN) {
    raise CfbError::EmptyRootDir
  }

  // Load mini streams
  let (mini_fats, ministream) = if header.mini_fat_len > 0U {
    let ministream = sectors.get_chain(
      directories[0].start,
      fats,
      data,
      directories[0].len,
    )
    let minifat_data = sectors.get_chain(
      header.mini_fat_start,
      fats,
      data,
      header.mini_fat_len * header.sector_size,
    )
    let minifat = to_u32_from_bytes(minifat_data)
    (minifat, ministream)
  } else {
    ([], b"")
  }
  {
    directories,
    sectors,
    fats,
    mini_sectors: Sectors::new(64U, ministream),
    mini_fats,
  }
}

///|
/// Check if directory exists
pub fn Cfb::has_directory(self : Cfb, name : String) -> Bool {
  for d in self.directories {
    if d.name == name {
      return true
    }
  }
  false
}

///|
/// Get a stream by name from directories
pub fn Cfb::get_stream(
  self : Cfb,
  name : String,
  data : Bytes,
) -> Bytes raise CfbError {
  let mut found_dir : Directory? = None
  for d in self.directories {
    if d.name == name {
      found_dir = Some(d)
      break
    }
  }
  match found_dir {
    None => raise CfbError::StreamNotFound(name)
    Some(d) =>
      if d.len < 4096U {
        self.mini_sectors.get_chain(d.start, self.mini_fats, data, d.len)
      } else {
        self.sectors.get_chain(d.start, self.fats, data, d.len)
      }
  }
}

///|
/// Parse CFB header from bytes
fn Header::from_bytes(data : Bytes) -> (Header, Array[UInt]) raise CfbError {
  if data.length() < 512 {
    raise CfbError::Io("Not enough data for CFB header")
  }

  // Check OLE signature
  let signature = read_u64(data, 0)
  match signature {
    Some(0xE11AB1A1E011CFD0UL) => ()
    _ => raise CfbError::Ole
  }
  let version = read_u16(data, 26).map_or(0U, fn(v) { v })
  let sector_shift = read_u16(data, 30).map_or(0U, fn(v) { v })
  let sector_size = match sector_shift {
    0x0009U => 512U
    0x000CU => 4096U
    _ => raise CfbError::Invalid("sector shift", "0x09 or 0x0C", sector_shift)
  }
  let mini_sector_shift = read_u16(data, 32).map_or(0U, fn(v) { v })
  if mini_sector_shift != 0x0006U {
    raise CfbError::Invalid("minisector shift", "0x06", mini_sector_shift)
  }
  let dir_len = read_u32(data, 40).map_or(0U, fn(v) { v })
  let fat_len = read_u32(data, 44).map_or(0U, fn(v) { v })
  let dir_start = read_u32(data, 48).map_or(0U, fn(v) { v })
  let mini_fat_start = read_u32(data, 60).map_or(0U, fn(v) { v })
  let mini_fat_len = read_u32(data, 64).map_or(0U, fn(v) { v })
  let difat_start = read_u32(data, 68).map_or(0U, fn(v) { v })

  // Extract initial DIFAT from header
  let difat : Array[UInt] = []
  for i = 76; i < 512; i = i + 4 {
    match read_u32(data, i) {
      Some(val) => difat.push(val)
      None => break
    }
  }
  let header = {
    version,
    sector_size,
    dir_len,
    dir_start,
    fat_len,
    mini_fat_len,
    mini_fat_start,
    difat_start,
  }
  (header, difat)
}

///|
/// Create new Sectors structure
fn Sectors::new(size : UInt, data : Bytes) -> Sectors {
  { size, data }
}

///|
/// Get a sector by ID
fn Sectors::get(
  self : Sectors,
  id : UInt,
  full_data : Bytes,
) -> Bytes raise CfbError {
  let start = id * self.size
  let end = start + self.size
  let start_int = start.reinterpret_as_int()
  let end_int = end.reinterpret_as_int()
  if end_int > full_data.length() {
    raise CfbError::Io("Sector extends beyond file bounds")
  }
  full_data[start_int:end_int].to_bytes()
}

///|
/// Get a chain of sectors
fn Sectors::get_chain(
  self : Sectors,
  sector_id : UInt,
  fats : Array[UInt],
  data : Bytes,
  len : UInt,
) -> Bytes raise CfbError {
  let mut chain = b""
  let mut current_sector = sector_id
  while current_sector != ENDOFCHAIN {
    let sector = self.get(current_sector, data)
    chain = chain + sector
    if current_sector.reinterpret_as_int() >= fats.length() {
      break
    }
    current_sector = fats[current_sector.reinterpret_as_int()]
  }
  if len > 0U && chain.length() > len.reinterpret_as_int() {
    chain[0:len.reinterpret_as_int()].to_bytes()
  } else {
    chain
  }
}

///|
/// Parse directories from bytes
fn parse_directories(data : Bytes, sector_size : UInt) -> Array[Directory] {
  let directories : Array[Directory] = []
  let entry_size = 128
  for i = 0; i < data.length(); i = i + entry_size {
    if i + entry_size <= data.length() {
      let dir_bytes = data[i:i + entry_size].to_bytes()
      let dir = Directory::from_bytes(dir_bytes, sector_size)
      directories.push(dir)
    }
  }
  directories
}

///|
/// Parse directory from byte slice
fn Directory::from_bytes(data : Bytes, sector_size : UInt) -> Directory {
  // Extract name (UTF-16LE encoded, first 64 bytes)
  let mut name = ""
  for i = 0; i < 64; i = i + 2 {
    if i + 1 < data.length() {
      let low = data[i].to_uint()
      let high = data[i + 1].to_uint()
      let code_point = low | (high << 8)
      if code_point == 0U {
        break
      }
      // Simple ASCII conversion (for now)
      if code_point < 128U {
        name = name +
          code_point.reinterpret_as_int().unsafe_to_char().to_string()
      }
    }
  }
  let start = read_u32(data, 116).map_or(0U, fn(v) { v })
  let len = if sector_size == 512U {
    read_u32(data, 120).map_or(0U, fn(v) { v })
  } else {
    // For 4096 byte sectors, read 64-bit length
    read_u64(data, 120).map_or(0UL, fn(v) { v }).to_uint()
  }
  { name, start, len }
}

///|
/// Convert bytes to array of u32 values (little-endian)
fn to_u32_from_bytes(data : Bytes) -> Array[UInt] {
  let result : Array[UInt] = []
  for i = 0; i < data.length(); i = i + 4 {
    match read_u32(data, i) {
      Some(val) => result.push(val)
      None => break
    }
  }
  result
}

///|
/// Decompress stream using the CFB compression algorithm
pub fn decompress_stream(data : Bytes) -> Bytes raise CfbError {
  if data.length() == 0 || data[0] != 0x01 {
    raise CfbError::Invalid(
      "signature",
      "0x01",
      if data.length() > 0 {
        data[0].to_uint()
      } else {
        0U
      },
    )
  }
  let mut result = b""
  let mut i = 1
  while i < data.length() {
    if i + 1 >= data.length() {
      break
    }
    let chunk_header = read_u16(data, i).map_or(0U, fn(v) { v })
    i = i + 2
    let chunk_size = chunk_header & 0x0FFFU
    let chunk_signature = (chunk_header & 0x7000U) >> 12
    let chunk_flag = (chunk_header & 0x8000U) >> 15
    if chunk_signature != 0b011U {
      raise CfbError::Invalid("chunk signature", "0b011", chunk_signature)
    }
    if chunk_flag == 0U {
      // Uncompressed chunk
      let chunk_end = i + 4096
      if chunk_end <= data.length() {
        result = result + data[i:chunk_end].to_bytes()
        i = chunk_end
      }
    } else {
      // Compressed chunk - simplified implementation
      let mut chunk_len = 0U
      while chunk_len < chunk_size && i < data.length() {
        let bit_flags = data[i]
        i = i + 1
        chunk_len = chunk_len + 1U
        for bit_index = 0; bit_index < 8; bit_index = bit_index + 1 {
          if chunk_len > chunk_size {
            break
          }
          if (bit_flags.to_uint() & (1U << bit_index)) == 0U {
            // Literal token
            if i < data.length() {
              result = result + Bytes::from_array([data[i]])
              i = i + 1
              chunk_len = chunk_len + 1U
            }
            // Copy token - simplified implementation
          } else if i + 1 < data.length() {
            i = i + 2
            chunk_len = chunk_len + 2U
            // For now, just skip copy tokens
          }
        }
      }
    }
  }
  result
}

///|
/// XLS encoding handler (simplified)
pub(all) struct XlsEncoding {
  codepage : UInt
} derive(Show, Eq)

///|
/// Create XLS encoding from codepage
pub fn XlsEncoding::from_codepage(
  codepage : UInt,
) -> XlsEncoding raise CfbError {
  // For now, we support only basic codepages
  match codepage {
    1252U | 65001U => { codepage, } // Windows-1252 or UTF-8
    _ => raise CfbError::CodePageNotFound(codepage)
  }
}

///|
/// Decode bytes to string (simplified)
pub fn XlsEncoding::decode_all(self : XlsEncoding, data : Bytes) -> String {
  // Simple ASCII/UTF-8 conversion for now
  let mut result = ""
  for i = 0; i < data.length(); i = i + 1 {
    let byte_val = data[i]
    if byte_val < 128 {
      result = result +
        byte_val.to_uint().reinterpret_as_int().unsafe_to_char().to_string()
    }
  }
  result
}

///|
/// Decode bytes to string with length and high byte flag
pub fn XlsEncoding::decode_to(
  self : XlsEncoding,
  data : Bytes,
  len : UInt,
  high_byte : Bool?,
) -> (String, UInt, UInt) {
  // Simplified implementation
  let actual_len = if len.reinterpret_as_int() > data.length() {
    data.length().reinterpret_as_uint()
  } else {
    len
  }
  let mut result = ""
  let used_bytes = match high_byte {
    Some(true) => {
      // UTF-16 style - read pairs of bytes
      let pairs = actual_len / 2U
      for i = 0; i < pairs.reinterpret_as_int(); i = i + 2 {
        if i + 1 < data.length() {
          let low = data[i]
          let high = data[i + 1]
          let code_point = low.to_uint() | (high.to_uint() << 8)
          if code_point < 128U {
            result = result +
              code_point.reinterpret_as_int().unsafe_to_char().to_string()
          }
        }
      }
      pairs * 2U
    }
    _ => {
      // Single byte encoding
      for i = 0; i < actual_len.reinterpret_as_int(); i = i + 1 {
        let byte_val = data[i]
        if byte_val < 128 {
          result = result +
            byte_val.to_uint().reinterpret_as_int().unsafe_to_char().to_string()
        }
      }
      actual_len
    }
  }
  (result, actual_len, used_bytes)
}
