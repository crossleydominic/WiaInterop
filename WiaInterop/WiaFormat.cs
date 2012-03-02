using System;
using System.Collections.Generic;
using System.Text;

namespace WiaInterop
{
	/// <summary>
	/// A simple class used to hold a file extension and format guid.
	/// </summary>
	public class WiaFormat
	{
		#region Private members

		/// <summary>
		/// The file extension for this format
		/// </summary>
		private string _fileExtension;

		/// <summary>
		/// The Windows defined guid for this format
		/// </summary>
		private Guid _guid;

		/// <summary>
		/// A COM prepared string representing the guid
		/// </summary>
		private string _comGuid;

		#endregion

		#region Constructors

		/// <summary>
		/// Constructs a new instance of the WiaFormat class
		/// </summary>
		public WiaFormat(string fileExtension, string guid)
		{
			if (string.IsNullOrEmpty(fileExtension))
			{
				throw new ArgumentNullException("fileExtension");
			}

			if (string.IsNullOrEmpty(guid))
			{
				throw new ArgumentNullException("guid");
			}

			_fileExtension = fileExtension;
			_guid = new Guid(guid);

			//Watch for escaped { and } !
			_comGuid = string.Format("{{{0}}}", _guid.ToString().ToUpper());
		}

		#endregion

		#region Public properties

		/// <summary>
		/// Get the file extesion
		/// </summary>
		public string FileExtension
		{
			get
			{
				return _fileExtension;
			}
		}

		/// <summary>
		/// Gets the raw guid
		/// </summary>
		public Guid Guid
		{
			get
			{
				return _guid;
			}
		}

		/// <summary>
		/// Gets the guid that has been formatted for COM.
		/// </summary>
		public string ComGuid
		{
			get
			{
				return _comGuid;
			}
		}

		#endregion

	}
}
