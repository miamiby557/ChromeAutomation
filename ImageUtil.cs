using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChromeAutomation
{
    class ImageUtil
    {
		public static string GetBase64FromImage(Bitmap image)
		{
			string result = "";
			try
			{
				using (MemoryStream memoryStream = new MemoryStream())
				{
					image.Save(memoryStream, ImageFormat.Jpeg);
					byte[] array = new byte[memoryStream.Length];
					memoryStream.Position = 0L;
					memoryStream.Read(array, 0, (int)memoryStream.Length);
					memoryStream.Close();
					result = "data:image/png;base64," + Convert.ToBase64String(array);
				}
			}
			catch (Exception)
			{
				throw new Exception("Something wrong during convert!");
			}
			return result;
		}
	}
}
