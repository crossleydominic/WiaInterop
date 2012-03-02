using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace WiaInterop
{
	/// <summary>
	/// A simple class that maps a Wia returned HRESULT into a meaningful
	/// error message.
	/// Defined in WiaDef.h
	/// </summary>
	public static class WiaError
	{
		public const int SEVERITY_SUCCESS = 0;
		public const int SEVERITY_ERROR = 1;
		public const int FACILITY_WIA = 33;
		public const int FACILITY_ITA = 4;

		public static readonly int WIA_ERROR_GENERAL_ERROR;
		public static readonly int WIA_ERROR_PAPER_JAM;
		public static readonly int WIA_ERROR_PAPER_EMPTY;
		public static readonly int WIA_ERROR_PAPER_PROBLEM;
		public static readonly int WIA_ERROR_OFFLINE;
		public static readonly int WIA_ERROR_BUSY;
		public static readonly int WIA_ERROR_WARMING_UP;
		public static readonly int WIA_ERROR_USER_INTERVENTION;
		public static readonly int WIA_ERROR_ITEM_DELETED;
		public static readonly int WIA_ERROR_DEVICE_COMMUNICATION;
		public static readonly int WIA_ERROR_INVALID_COMMAND;
		public static readonly int WIA_ERROR_INCORRECT_HARDWARE_SETTING;
		public static readonly int WIA_ERROR_DEVICE_LOCKED;
		public static readonly int WIA_ERROR_EXCEPTION_IN_DRIVER;
		public static readonly int WIA_ERROR_INVALID_DRIVER_RESPONSE;
		public static readonly int WIA_ERROR_COVER_OPEN;
		public static readonly int WIA_ERROR_LAMP_OFF;
		public static readonly int WIA_ERROR_DESTINATION;
		public static readonly int WIA_ERROR_NETWORK_RESERVATION_FAILED;
		public static readonly int WIA_STATUS_END_OF_MEDIA;
		public static readonly int WIA_S_NO_DEVICE_AVAILABLE;
		public static readonly int COM_ERROR_CLASS_NOT_REGISTERED;

		static WiaError()
		{
			WIA_ERROR_GENERAL_ERROR = MakeHresult(SEVERITY_ERROR, FACILITY_WIA, 1);
			WIA_ERROR_PAPER_JAM = MakeHresult(SEVERITY_ERROR, FACILITY_WIA, 2);
			WIA_ERROR_PAPER_EMPTY = MakeHresult(SEVERITY_ERROR, FACILITY_WIA, 3);
			WIA_ERROR_PAPER_PROBLEM = MakeHresult(SEVERITY_ERROR, FACILITY_WIA, 4);
			WIA_ERROR_OFFLINE = MakeHresult(SEVERITY_ERROR, FACILITY_WIA, 5);
			WIA_ERROR_BUSY = MakeHresult(SEVERITY_ERROR, FACILITY_WIA, 6);
			WIA_ERROR_WARMING_UP = MakeHresult(SEVERITY_ERROR, FACILITY_WIA, 7);
			WIA_ERROR_USER_INTERVENTION = MakeHresult(SEVERITY_ERROR, FACILITY_WIA, 8);
			WIA_ERROR_ITEM_DELETED = MakeHresult(SEVERITY_ERROR, FACILITY_WIA, 9);
			WIA_ERROR_DEVICE_COMMUNICATION = MakeHresult(SEVERITY_ERROR, FACILITY_WIA, 10);
			WIA_ERROR_INVALID_COMMAND = MakeHresult(SEVERITY_ERROR, FACILITY_WIA, 11);
			WIA_ERROR_INCORRECT_HARDWARE_SETTING = MakeHresult(SEVERITY_ERROR, FACILITY_WIA, 12);
			WIA_ERROR_DEVICE_LOCKED = MakeHresult(SEVERITY_ERROR, FACILITY_WIA, 13);
			WIA_ERROR_EXCEPTION_IN_DRIVER = MakeHresult(SEVERITY_ERROR, FACILITY_WIA, 14);
			WIA_ERROR_INVALID_DRIVER_RESPONSE = MakeHresult(SEVERITY_ERROR, FACILITY_WIA, 15);
			WIA_ERROR_COVER_OPEN = MakeHresult(SEVERITY_ERROR, FACILITY_WIA, 16);
			WIA_ERROR_LAMP_OFF = MakeHresult(SEVERITY_ERROR, FACILITY_WIA, 17);
			WIA_ERROR_DESTINATION = MakeHresult(SEVERITY_ERROR, FACILITY_WIA, 18);
			WIA_ERROR_NETWORK_RESERVATION_FAILED = MakeHresult(SEVERITY_ERROR, FACILITY_WIA, 19);
			WIA_STATUS_END_OF_MEDIA = MakeHresult(SEVERITY_SUCCESS, FACILITY_WIA, 1);
			WIA_S_NO_DEVICE_AVAILABLE = MakeHresult(SEVERITY_ERROR, FACILITY_WIA, 21);

			COM_ERROR_CLASS_NOT_REGISTERED = MakeHresult(SEVERITY_ERROR, FACILITY_ITA, 340);
		}

		public static string GetErrorMessage(COMException ce)
		{
			//These error messages come from the WIA documenation.
			if (ce.ErrorCode == WIA_ERROR_GENERAL_ERROR) { return "An unknown error has occurred with the Windows Image Acquisition (WIA) device."; }
			if (ce.ErrorCode == WIA_ERROR_PAPER_JAM) { return "Paper is jammed in the scanner's document feeder."; }
			if (ce.ErrorCode == WIA_ERROR_PAPER_EMPTY) { return "The user requested a scan and there are no documents left in the document feeder."; }
			if (ce.ErrorCode == WIA_ERROR_PAPER_PROBLEM) { return "An unspecified problem occurred with the scanner's document feeder."; }
			if (ce.ErrorCode == WIA_ERROR_OFFLINE) { return "The WIA device is not online."; }
			if (ce.ErrorCode == WIA_ERROR_BUSY) { return "The WIA device is busy."; }
			if (ce.ErrorCode == WIA_ERROR_WARMING_UP) { return "The WIA device is warming up."; }
			if (ce.ErrorCode == WIA_ERROR_USER_INTERVENTION) { return "An unspecified error has occurred with the WIA device that requires user intervention. The user should ensure that the device is turned on, online, and any cables are properly connected."; }
			if (ce.ErrorCode == WIA_ERROR_ITEM_DELETED) { return "The WIA device was deleted. It can no longer be accessed."; }
			if (ce.ErrorCode == WIA_ERROR_DEVICE_COMMUNICATION) { return "	An unspecified error occurred during an attempted communication with the WIA device."; }
			if (ce.ErrorCode == WIA_ERROR_INVALID_COMMAND) { return "	The device does not support this command."; }
			if (ce.ErrorCode == WIA_ERROR_INCORRECT_HARDWARE_SETTING) { return "There is an incorrect setting on the WIA device."; }
			if (ce.ErrorCode == WIA_ERROR_DEVICE_LOCKED) { return "The scanner head is locked."; }
			if (ce.ErrorCode == WIA_ERROR_EXCEPTION_IN_DRIVER) { return "The device driver threw an exception."; }
			if (ce.ErrorCode == WIA_ERROR_INVALID_DRIVER_RESPONSE) { return "The response from the driver is invalid."; }
			if (ce.ErrorCode == WIA_S_NO_DEVICE_AVAILABLE) { return "No WIA device of the selected type is available."; }

			//These error messages are totally made up because they don't appear in the
			//documentation but are defined in WiaDef.h
			if (ce.ErrorCode == WIA_ERROR_COVER_OPEN) { return "The device cover is open."; }
			if (ce.ErrorCode == WIA_ERROR_LAMP_OFF) { return "The devices lamp is off."; }
			if (ce.ErrorCode == WIA_ERROR_DESTINATION) { return "The destination is invalid"; }
			if (ce.ErrorCode == WIA_ERROR_NETWORK_RESERVATION_FAILED) { return "A network error occurred."; }
			if (ce.ErrorCode == WIA_STATUS_END_OF_MEDIA) { return "End of media reached."; }

			//Watch for the specific error that means that the Wia Automation library is not installed.
			if (ce.ErrorCode == COM_ERROR_CLASS_NOT_REGISTERED) { return ce.Message + "\r\nEnsure that the WIA Automation Library 2.0 is installed and registered."; }

			//If it's not a WIA error then just return the message we got from COM.
			return ce.Message;
		}

		private static int MakeHresult(int severity, int facility, int code)
		{
			//Code ripped straight out of the MAKEHRESULT macro defined in WinErr.h
			return (int)(((ulong)(severity) << 31) | ((ulong)(facility) << 16) | ((ulong)(code)));
		}
	}
}
