using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.Collections.ObjectModel;

namespace WiaInterop
{
	/// <summary>
	/// These GUIDs define image formats that scanners can send information
	/// across to our application in.  
	/// We only support bitmaps and jpegs (which, to be fair, should cover
	/// pretty much all devices apart from the fancy-schancy ones that 
	/// insist on using PNGs).
	/// 
	/// They are defined in GDI+ and also the WIA section of the registry, here
	/// HKEY_CLASSES_ROOT/CLSID/{D2923B86-15F1-46FF-A19A-DE825F919576}/SupportedExtension/.bmp
	/// and
	/// HKEY_CLASSES_ROOT/CLSID/{D2923B86-15F1-46FF-A19A-DE825F919576}/SupportedExtension/.jpg
	/// and
	/// HKEY_CLASSES_ROOT/CLSID/{D2923B86-15F1-46FF-A19A-DE825F919576}/SupportedExtension/.png
	/// </summary>
	public static class WiaFormats
	{
		#region Constants

		/// <summary>
		/// Extension for bitmaps
		/// </summary>
		public const string BITMAP_EXTENSION = "bmp";

		/// <summary>
		/// Extension for jpegs
		/// </summary>
		public const string JPG_EXTENSION = "jpg";

		/// <summary>
		/// Extension for pngs
		/// </summary>
		public const string PNG_EXTENSION = "png";

		#endregion

		#region Defined Formats

		/// <summary>
		/// In Memory bitmap format
		/// </summary>
		public static readonly WiaFormat InMemoryBmp = new WiaFormat(BITMAP_EXTENSION, "{B96B3CAA-0728-11D3-9D7B-0000F81EF32E}");

		/// <summary>
		/// Bitmap format
		/// </summary>
		public static readonly WiaFormat Bmp = new WiaFormat(BITMAP_EXTENSION, "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}");

		/// <summary>
		/// Jpeg format
		/// </summary>
		public static readonly WiaFormat Jpg = new WiaFormat(JPG_EXTENSION, "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}");

		/// <summary>
		/// Png format
		/// </summary>
		public static readonly WiaFormat Png = new WiaFormat(PNG_EXTENSION, "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}");

		#endregion

		#region Members

		/// <summary>
		/// The list of formats that we support
		/// </summary>
		private static List<WiaFormat> _formats = null;

		#endregion

		#region Static Constructor

		/// <summary>
		/// Static constructor, used to build all of the cachable items
		/// </summary>
		static WiaFormats()
		{
			_formats = new List<WiaFormat>();

			FieldInfo[] fields = typeof(WiaFormats).GetFields(BindingFlags.DeclaredOnly | BindingFlags.Public | BindingFlags.Static);

			foreach (FieldInfo fi in fields)
			{
				if (fi.FieldType == typeof(WiaFormat))
				{
					_formats.Add((WiaFormat)fi.GetValue(null));
				}
			}
		}

		#endregion

		#region Public properties

		/// <summary>
		/// Gets a list of formats
		/// </summary>
		public static ReadOnlyCollection<WiaFormat> FormatList
		{
			get
			{
				return _formats.AsReadOnly();
			}
		}

		#endregion

		#region Public methods

		/// <summary>
		/// Whether the supplied GUID is the InMemoryBmp format guid
		/// </summary>
		public static bool IsInMemoryBitmap(string formatGuid)
		{
			return InMemoryBmp.ComGuid.Equals(NormalizeGuid(formatGuid), StringComparison.OrdinalIgnoreCase);
		}

		/// <summary>
		/// Whether the supplied GUID is the jpg format guid
		/// </summary>
		public static bool IsJpg(string formatGuid)
		{
			return Jpg.ComGuid.Equals(NormalizeGuid(formatGuid), StringComparison.OrdinalIgnoreCase);
		}

		/// <summary>
		/// Whether the supplied GUID is the bmp format guid
		/// </summary>
		public static bool IsBmp(string formatGuid)
		{
			return Bmp.ComGuid.Equals(NormalizeGuid(formatGuid), StringComparison.OrdinalIgnoreCase);
		}

		/// <summary>
		/// Whether the supplied GUID is the png format guid
		/// </summary>
		public static bool IsPng(string formatGuid)
		{
			return Png.ComGuid.Equals(NormalizeGuid(formatGuid), StringComparison.OrdinalIgnoreCase);
		}

		/// <summary>
		/// Gets a WiaFormat for the supplied format GUID
		/// </summary>
		public static WiaFormat GetWiaFormat(string formatGuid)
		{
			if (string.IsNullOrEmpty(formatGuid))
			{
				throw new ArgumentNullException("formatGuid");
			}

			foreach(WiaFormat format in _formats)
			{
				if(format.ComGuid.Equals(NormalizeGuid(formatGuid), StringComparison.OrdinalIgnoreCase))
				{
					return format;
				}
			}
			
			return null;
		}

		/// <summary>
		/// Checks whether or not the supplied formatGuid is one we can work with
		/// </summary>
		public static bool IsFormatSupported(string formatGuid)
		{
			WiaFormat format = GetWiaFormat(formatGuid);

			return (format != null);
		}

		#endregion

		#region Private methods

		/// <summary>
		/// Converts the supplied guid string into a guid that COM can recognize
		/// (wrapped in curly braces and in uppercase).
		/// </summary>
		private static string NormalizeGuid(string guid)
		{
			guid = guid.Trim().ToUpper();

			if (!guid.StartsWith("{"))
			{
				guid = "{" + guid;
			}

			if (!guid.EndsWith("}"))
			{
				guid += "}";
			}

			return guid;
		}

		#endregion
	}
}
