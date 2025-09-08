///|
/// Picture format enumeration for supported image types
pub(all) enum PictureFormat {
  /// Enhanced Metafile Format
  Emf
  /// Windows Metafile Format  
  Wmf
  /// Mac PICT Format
  Pict
  /// JPEG Image Format
  Jpeg
  /// Portable Network Graphics
  Png
  /// Device Independent Bitmap
  Dib
  /// Graphics Interchange Format
  Gif
  /// Tagged Image File Format
  Tiff
  /// Encapsulated PostScript
  Eps
  /// Bitmap Image
  Bmp
  /// WordPerfect Graphics
  Wpg
  /// Unknown format
  Unknown
} derive(Show, Eq)

///|
/// Picture data structure representing an embedded image
pub(all) struct Picture {
  /// File extension/format
  format : PictureFormat
  /// Raw image data
  data : Bytes
  /// Optional name/identifier
  name : String?
} derive(Show, Eq)

///|
/// Create a new Picture from extension and data
pub fn Picture::new(
  extension : String,
  data : Bytes,
  name : String?,
) -> Picture {
  let format = extension_to_format(extension)
  { format, data, name }
}

///|
/// Get the file extension for a picture format
pub fn PictureFormat::extension(self : PictureFormat) -> String {
  match self {
    Emf => "emf"
    Wmf => "wmf"
    Pict => "pict"
    Jpeg => "jpg"
    Png => "png"
    Dib => "dib"
    Gif => "gif"
    Tiff => "tiff"
    Eps => "eps"
    Bmp => "bmp"
    Wpg => "wpg"
    Unknown => "bin"
  }
}

///|
/// Convert file extension to picture format
fn extension_to_format(extension : String) -> PictureFormat {
  match extension {
    "emf" => Emf
    "wmf" => Wmf
    "pict" => Pict
    "jpeg" | "jpg" => Jpeg
    "png" => Png
    "dib" => Dib
    "gif" => Gif
    "tiff" | "tif" => Tiff
    "eps" => Eps
    "bmp" => Bmp
    "wpg" => Wpg
    _ => Unknown
  }
}