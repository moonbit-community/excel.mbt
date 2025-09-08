# MoonBit Calamine Library

A pure MoonBit library for reading Excel and OpenDocument Spreadsheet files.

## Overview

Calamine is a comprehensive spreadsheet reader library that supports multiple formats:
- **XLSX** - Excel 2007+ format
- **XLS** - Excel 2003 format  
- **XLSB** - Excel Binary format
- **ODS** - OpenDocument Spreadsheet format

## Features

- ✅ **Core Data Types** - Robust cell data representation with type safety
- ✅ **Error Handling** - Comprehensive error types for all operations
- ✅ **Range Operations** - Efficient cell range and dimension handling
- ✅ **Reader Traits** - Unified interface across all spreadsheet formats
- ✅ **Format Detection** - Automatic workbook type detection
- 🚧 **VBA Support** - Basic VBA project reading capabilities
- 🚧 **Table Support** - Excel worksheet table handling
- 🚧 **Merged Cells** - Support for merged cell regions

## Basic Usage

### Data Types

Calamine provides comprehensive data types for representing Excel cell values:

```moonbit
test "data_types_example" {
  // Create different cell data types
  let int_cell = Data::Int(42L)
  let float_cell = Data::Float(3.14159)
  let string_cell = Data::String("Hello Excel")
  let bool_cell = Data::Bool(true)
  let empty_cell = Data::Empty
  
  // Check data types
  inspect(int_cell.is_int(), content="true")
  inspect(float_cell.is_float(), content="true")
  inspect(string_cell.is_string(), content="true")
  inspect(bool_cell.is_bool(), content="true")
  inspect(empty_cell.is_empty(), content="true")
  
  // Extract values
  inspect(int_cell.get_int(), content="Some(42)")
  inspect(string_cell.get_string(), content="Some(\"Hello Excel\")")
  
  // Type conversion
  inspect(int_cell.as_f64(), content="Some(42)")
  inspect(float_cell.as_i64(), content="Some(3)")
}
```

### Dimensions and Ranges

Work with cell ranges and dimensions:

```moonbit
test "dimensions_example" {
  // Create dimensions for a range A1:K6
  let dims = Dimensions::new((0U, 0U), (5U, 10U))
  
  inspect(dims.width(), content="11")
  inspect(dims.height(), content="6") 
  inspect(dims.contains(2U, 5U), content="true")
  inspect(dims.len(), content="66")
}
```

```moonbit
test "range_example" {
  // Create cells for a sparse range
  let cells = [
    Cell::new((0U, 0U), "A1"),
    Cell::new((0U, 1U), "B1"), 
    Cell::new((1U, 0U), "A2"),
    Cell::new((1U, 1U), "B2")
  ]
  
  // Create range from sparse cells
  let range = Range::from_sparse(cells)
  
  inspect(range.width(), content="2")
  inspect(range.height(), content="2")
  inspect(range.get_value((0U, 0U)), content="Some(\"A1\")")
  inspect(range.get_value((1U, 1U)), content="Some(\"B2\")")
}
```

### Workbook Type Detection

Automatically detect spreadsheet formats:

```moonbit
test "format_detection_example" {
  // XLSX format (ZIP signature)
  let xlsx_data = b"PK\x03\x04"
  let xlsx_type = detect_workbook_type(xlsx_data)
  inspect(xlsx_type, content="Unknown")
  
  // XLS format (OLE2 signature)
  let xls_data = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
  let xls_type = detect_workbook_type(xls_data)
  inspect(xls_type, content="Xls")
  
  // Unknown format
  let unknown_data = b"INVALID"
  let unknown_type = detect_workbook_type(unknown_data)
  inspect(unknown_type, content="Unknown")
}
```

### Error Handling

Comprehensive error handling with specific error types:

```moonbit
test "error_handling_example" {
  // Create specific error types
  let xlsx_error = XlsxError::Password
  let xls_error = XlsError::InvalidBOM
  let general_error = CalamineError::Io("File not found")
  
  inspect(xlsx_error, content="Password")
  inspect(xls_error, content="InvalidBOM") 
  inspect(general_error, content="Io(\"File not found\")")
  
  // Convert to general error
  let converted = xlsx_error.to_error()
  inspect(converted, content="Xlsx(Password)")
}
```

## Architecture

### Core Types

- **`Data`** - Represents all possible Excel cell values (int, float, string, bool, date, error, empty)
- **`DataRef`** - Reference-based version for memory efficiency with shared strings
- **`Dimensions`** - Represents rectangular cell ranges with start/end coordinates
- **`Range[T]`** - Container for cell data with position-based access
- **`Cell[T]`** - Individual cell with position and value

### Error Types

- **`CalamineError`** - General error type encompassing all format-specific errors
- **`XlsxError`** - Specific to Excel 2007+ format issues
- **`XlsError`** - Specific to Excel 2003 format issues  
- **`XlsbError`** - Specific to Excel Binary format issues
- **`OdsError`** - Specific to OpenDocument format issues
- **`VbaError`** - VBA project related errors
- **`DeError`** - Deserialization errors

### Reader Traits

- **`Reader`** - Core trait for all spreadsheet readers
- **`ReaderRef`** - Extended trait for reference-based data access
- **`AutoWorkbook`** - Automatic format detection and unified interface

## Advanced Features

### Excel Date and Time

Handle Excel's date/time representation:

```moonbit
test "excel_datetime_example" {
  let datetime = ExcelDateTime::new(44197.0, DateTime, false)
  
  inspect(datetime.is_datetime(), content="true")
  inspect(datetime.is_duration(), content="false")
  inspect(datetime.as_f64(), content="44197")
}
```

### Metadata and Sheets

Work with workbook metadata:

```moonbit  
test "metadata_example" {
  let metadata = Metadata::default()
  let sheet = Sheet::{
    name: "Sheet1",
    typ: WorkSheet, 
    visible: Visible
  }
  
  metadata.add_sheet(sheet)
  
  let names = metadata.sheet_names()
  inspect(names.length(), content="1")
  inspect(names[0], content="Sheet1")
}
```

## Implementation Status

### ✅ Completed Components

1. **Core Data Types** - All cell data types implemented with full type safety
2. **Error Handling** - Comprehensive error types for all operations
3. **Range Operations** - Complete range and dimension handling
4. **Reader Traits** - Unified interface across formats with trait system
5. **Format Detection** - Automatic workbook type detection by file signatures

### 🚧 In Progress Components

1. **Format-Specific Readers** - Concrete implementations for XLSX, XLS, XLSB, ODS
2. **Table Support** - Excel worksheet table reading and manipulation
3. **VBA Projects** - Enhanced VBA project reading and processing
4. **Merged Cells** - Support for merged cell regions

### 📋 Planned Components

1. **Streaming Reader** - Memory-efficient streaming for large files
2. **Formula Support** - Reading and evaluating Excel formulas
3. **Charts and Images** - Support for embedded charts and images
4. **Writer Support** - Writing capabilities for supported formats

## Contributing

This library is actively being developed. Contributions are welcome!

### Development Setup

```bash
# Clone the repository
git clone <repository-url>
cd calamine

# Run tests
moon test

# Check compilation
moon check

# Build examples
moon build
```

### Running Examples

```bash
# Run the main demo
moon run cmd/main
```

## License

Licensed under the Apache License, Version 2.0.

## Acknowledgments

This library is inspired by and maintains compatibility with the Rust calamine library while providing a pure MoonBit implementation optimized for the MoonBit ecosystem.
