using System;
using System.Collections.Generic;
using System.Text;

namespace WiaInterop
{
	/// <summary>
	/// A structure containing a list of WIA defined constants
	/// </summary>
	public struct WiaConstants
	{
		/// <summary>
		/// Contains some constants we'll use for logging.
		/// </summary>
		public struct LoggingConstants
		{
			public const string MethodCalled = "Method Called";
			public const string ExceptionOccurred = "Exception Occurred";
		}

		/// <summary>
		/// A set of constants used to query a WiaItem for their properties.
		/// </summary>
		public struct WiaProperties
		{
			/// <summary>
			/// Used to request the image format the device uses
			/// </summary>
			public const int FORMAT_ID = 4106;

			/// <summary>
			/// Uses to request the preferred image format that the device wants to use.
			/// </summary>
			public const int PREFERRED_FORMAT_ID = 4105;
		}
	}
}
