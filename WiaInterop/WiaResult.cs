using System;
using System.Collections.Generic;
using System.Text;

namespace WiaInterop
{
	/// <summary>
	/// A simple class used to represent the outcome of a Wia operation
	/// along with any error messages.
	/// </summary>
	public class WiaResult
	{
		#region Private members

		/// <summary>
		/// Whether or not the operation succeeded
		/// </summary>
		private bool _succeeded;

		/// <summary>
		/// The error message if the operation failed.
		/// </summary>
		private string _errorMessage;

		#endregion

		#region Private static members

		/// <summary>
		/// A single instance of the WiaResult that represents success.
		/// Clients can access this through the Success property.
		/// </summary>
		private static readonly WiaResult _success = new WiaResult();

		#endregion

		#region Constructors

		/// <summary>
		/// Private constructor.  The only way of creating a success WiaResult is from
		/// inside this class.  Outside code should get a reference to the success WiaResult
		/// by using the Success property.
		/// </summary>
		private WiaResult()
		{
			_succeeded = true;
			_errorMessage = string.Empty;
		}

		/// <summary>
		/// A public constructor for constructing results that have an error message and therefore
		/// failed.
		/// </summary>
		public WiaResult(string errorMessage)
		{
			if (string.IsNullOrEmpty(errorMessage))
			{
				throw new ArgumentNullException("errorMessage");
			}
			_succeeded = false;
			_errorMessage = errorMessage;
		}

		#endregion

		#region Public properties

		/// <summary>
		/// Gets whether or not the operation succeeded.
		/// </summary>
		public bool Succeeded
		{
			get
			{
				return _succeeded;
			}
		}

		/// <summary>
		/// Gets the error message associated with this result.
		/// Will only be set if Succeeded == false.
		/// </summary>
		public string ErrorMessage
		{
			get
			{
				return _errorMessage;
			}
		}

		#endregion

		#region Public static properties

		/// <summary>
		/// Gets the only instance of the WiaResult that represents success
		/// </summary>
		public static WiaResult Success
		{
			get
			{
				return _success;
			}
		}

		#endregion

	}
}
