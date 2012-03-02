using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using WIA;
using System.Drawing;
using System.Runtime.InteropServices;
using log4net;

namespace WiaInterop
{
	public static class WiaManager
	{
		#region Constants

		/// <summary>
		/// A generic error message template.
		/// </summary>
		private const string ERROR_MESSAGE_FORMAT = "Could not {0}.\r\nThe error message was '{1}'.\r\nSee logs for full error information";

		#endregion

		#region Logging

		/// <summary>
		/// For logging
		/// </summary>
		private static ILog logger = LogManager.GetLogger(typeof(WiaManager));

		#endregion

		#region Public methods

		/// <summary>
		/// Checks to see if WIA is available for use on this machine.
		/// </summary>
		public static bool IsWiaAvailable()
		{
			if (logger.IsDebugEnabled) { logger.Debug(WiaConstants.LoggingConstants.MethodCalled); }

			try
			{
				//We'll see if the WIA automation library is available by trying to
				//create an instance of the CommonDialogClass.  Internally, the CLR
				//will be attempting to create an instance of the WIA Automation library.
				//If this instance is not registered then we'll get COM hand us a
				//CLASS_NOT_REGISTERED HRESULT which the CLR converts into a COMException for
				//us.  If an exception is thrown then the it means that COM couldn't locate 
				//the automation library so we can't scan using WIA.

				CommonDialogClass cdc = new CommonDialogClass();
				return true;
			}
			catch (COMException)
			{
				return false;
			}
		}

		/// <summary>
		/// Presents the user with multiple peices of WIA UI so that they
		/// can select a capture device, then perform the image acquisition and
		/// will finally save the image to the HD (and transcode to JPG is required).
		/// </summary>
		/// <param name="outputDirectory">The directory the image should be placed into</param>
		public static WiaResult SelectAndCapture(string outputDirectory, out string outputtedFile)
		{
			if (logger.IsDebugEnabled) { logger.Debug(WiaConstants.LoggingConstants.MethodCalled); }

			outputtedFile = string.Empty;

			#region Input validation

			if (string.IsNullOrEmpty(outputDirectory) ||
				Directory.Exists(outputDirectory) == false)
			{
				return new WiaResult("The specified output directory does not exist");
			}

			#endregion

			WiaResult result;

			//1. Get the capture device
			Device device = null;
			result = GetDevice(out device);

			if (!result.Succeeded)
			{
				return result;
			}

			//2. Get the capture item that we can use to transfer the image from.
			Item captureItem = null;
			result = GetCaptureItem(device, out captureItem);

			if (!result.Succeeded)
			{
				return result;
			}

			//3. Obtain the format information we need to transfer the image.
			WiaFormat format = null;
			result = GetFormat(captureItem, out format);
			if (!result.Succeeded)
			{
				return result;
			}

			//4. Transfer the image
			string fileName = null;
			result = TransferImage(captureItem, format, outputDirectory, out fileName);
			if (!result.Succeeded)
			{
				return result;
			}

			//If we've gotten here then everything went ok
			outputtedFile = fileName;
			return WiaResult.Success;
		}

		#endregion

		#region Private methods

		/// <summary>
		/// Transcodes the supplied file from whatever format it's in into a jpg.
		/// Will only work if the supplied filename points to a valid image that is
		/// natively supported by the .Net Image class.
		/// </summary>
		private static WiaResult TranscodeImageToJpg(string fileName, out string newFileName)
		{
			if (logger.IsDebugEnabled) { logger.Debug(WiaConstants.LoggingConstants.MethodCalled); }

			newFileName = null;
			try
			{
				//Load the image
				Image img = Image.FromFile(fileName);

				FileInfo fi = new FileInfo(fileName);
				string tempFileName = fi.DirectoryName + "\\" + Guid.NewGuid() + ".jpg";

				//Save and transcode
				img.Save(tempFileName, System.Drawing.Imaging.ImageFormat.Jpeg);
				img.Dispose();

				newFileName = tempFileName;

				//Once we've successfully written the new file we can remove the old one.
				File.Delete(fileName);

				return WiaResult.Success;
			}
			catch (Exception e)
			{
				if (logger.IsDebugEnabled) 
				{ 
					logger.Debug(WiaConstants.LoggingConstants.ExceptionOccurred, e); 
				}
				return new WiaResult(string.Format(ERROR_MESSAGE_FORMAT, "transcode image to jpg", e.Message));
			}
		}

		/// <summary>
		/// Transfers an image from the capture device
		/// </summary>
		private static WiaResult TransferImage(Item item, WiaFormat format, string outputPath, out string fileName)
		{
			fileName = string.Empty;
			try
			{
				//Initiate transfer
				ImageFile image = item.Transfer(format.ComGuid) as ImageFile;

				if (image != null)
					//Ensure something was obtained.
				{
					string tempFileName = string.Format("{0}.{1}", Guid.NewGuid().ToString(), format.FileExtension);

					if (!outputPath.EndsWith("\\"))
					{
						outputPath += "\\";
					}

					tempFileName = outputPath + tempFileName;

					image.SaveFile(tempFileName);

					//An exception will be thrown if the image cannot be saved, by getting here
					//we know that the save succeeded.
					fileName = tempFileName;

					//If the image was not saved in Jpg format then transcode it here
					if (!WiaFormats.IsJpg(format.ComGuid))
					{
						string transcodedFileName;
						WiaResult transcodeResult = TranscodeImageToJpg(fileName, out transcodedFileName);

						if (!transcodeResult.Succeeded)
						{
							return transcodeResult;
						}
						else
						{
							fileName = transcodedFileName;
						}
					}
					return WiaResult.Success;
				}
				else
				{
					return new WiaResult(string.Format(ERROR_MESSAGE_FORMAT, 
						"transfer the image from the device",
						"The image returned could not be cast to an ImageFile interface"));
				}
			}
			catch (COMException ce)
			{
				if (logger.IsDebugEnabled)
				{
					logger.Debug(WiaConstants.LoggingConstants.ExceptionOccurred, ce);
				}
				return new WiaResult(string.Format(ERROR_MESSAGE_FORMAT,
					"transfer the image from the device",
					WiaError.GetErrorMessage(ce)));
			}
		}

		/// <summary>
		/// Gets a WiaFormat object that represents an image format that the device
		/// and our app both understand.
		/// </summary>
		private static WiaResult GetFormat(Item item, out WiaFormat format)
		{
			format = null;
			try
			{
				//Devices can have supported formats and preferred formats.
				string deviceFormatGuid, devicePreferredFormatGuid;
				deviceFormatGuid = devicePreferredFormatGuid = null;

				//Get the format information from the device
				foreach (Property prop in item.Properties)
				{
					System.Diagnostics.Debug.WriteLine(prop.PropertyID + " - " + prop.Name + " - " + prop.get_Value());
					switch (prop.PropertyID)
					{
						case WiaConstants.WiaProperties.FORMAT_ID:
							deviceFormatGuid = (prop.get_Value() == null ? null : prop.get_Value().ToString());
							break;
						case WiaConstants.WiaProperties.PREFERRED_FORMAT_ID:
							devicePreferredFormatGuid = (prop.get_Value() == null ? null : prop.get_Value().ToString());
							break;
					}
				}

				//We'll try and match on the preferred format first.

				if (string.IsNullOrEmpty(devicePreferredFormatGuid) == false)
				//Check to see if we support the preferred format.
				{
					format = WiaFormats.GetWiaFormat(devicePreferredFormatGuid);

					if (format != null)
					{
						return WiaResult.Success;
					}

				}
				
				if (string.IsNullOrEmpty(deviceFormatGuid) == false )
				//Check to see if we support the format
				{
					format = WiaFormats.GetWiaFormat(deviceFormatGuid);

					if (format != null)
					{
						return WiaResult.Success;
					}
				}

				//If we've gotten here then the device does not support a format we can
				//work with
				return new WiaResult(string.Format(ERROR_MESSAGE_FORMAT,
					"get the image format from the device",
					"No supported file format can be found on the device"));

			}
			catch (COMException ce)
			{
				if (logger.IsDebugEnabled)
				{
					logger.Debug(WiaConstants.LoggingConstants.ExceptionOccurred, ce);
				}
				return new WiaResult(string.Format(ERROR_MESSAGE_FORMAT,
					"get the image format from the device",
					WiaError.GetErrorMessage(ce)));
			}
		}

		/// <summary>
		/// Queries a device for the Wia Item that we will use to do the capture and
		/// transfer.
		/// </summary>
		private static WiaResult GetCaptureItem(Device device, out Item item)
		{
			item = null;
			try
			{
				//Show UI.
				CommonDialogClass cdc = new CommonDialogClass();
				Items items = cdc.ShowSelectItems(
					device,
					WiaImageIntent.UnspecifiedIntent,
					WiaImageBias.MaximizeQuality,
					true,
					true,
					false);

				if (items == null || items.Count == 0)
					//Cannot obtain an item to acquire the image OR the user cancelled.
				{

					return new WiaResult(string.Format(ERROR_MESSAGE_FORMAT,
						"get the capture item",
						"No items were selected to be captured"));
				}
				else
				{
					//Wia arrays are 1-based-indexed.  1 here represents the first item.
					item = items[1];

					return WiaResult.Success;
				}
			}
			catch (COMException ce)
			{
				if (logger.IsDebugEnabled)
				{
					logger.Debug(WiaConstants.LoggingConstants.ExceptionOccurred, ce);
				}
				return new WiaResult(string.Format(ERROR_MESSAGE_FORMAT,
						"get the capture item",
						WiaError.GetErrorMessage(ce)));
			}
		}

		/// <summary>
		/// Gets the device that will be used to acquire an image.
		/// </summary>
		private static WiaResult GetDevice(out Device device)
		{
			device = null;
			try
			{
				CommonDialogClass cdc = new CommonDialogClass();
				device = cdc.ShowSelectDevice(WiaDeviceType.UnspecifiedDeviceType, false, false);

				if (device == null)
					//Either the user cancelled OR we have no device we can use to get an image.
				{
					return new WiaResult(string.Format(ERROR_MESSAGE_FORMAT,
						"get the capture device",
						"No device was selected"));
				}
				else
				{
					return WiaResult.Success;
				}
			}
			catch (COMException ce)
			{
				if (logger.IsDebugEnabled)
				{
					logger.Debug(WiaConstants.LoggingConstants.ExceptionOccurred, ce);
				}
				return new WiaResult(string.Format(ERROR_MESSAGE_FORMAT,
					"get the capture device",
					WiaError.GetErrorMessage(ce)));
			}
		}

		#endregion

	}
}
