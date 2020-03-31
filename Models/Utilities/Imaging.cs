using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Drawing;
using System.Drawing.Imaging;

namespace RoutledgeAssignmentFour.Models.Utilities
{
    public class Imaging
    {
        public static string ConvertandSaveImage (byte[] byteArrayIn)
        {
            Image returnImage = null;
            string base64String = new string("");
            try
            {
                MemoryStream ms = new MemoryStream(byteArrayIn, 0, byteArrayIn.Length);
                ms.Write(byteArrayIn, 0, byteArrayIn.Length);                            
                returnImage = Image.FromStream(ms, true);
                //base64String = Convert.ToBase64String(byteArrayIn);
                FileInfo newfileinfo = new FileInfo(Constants.Locations.ImageSaveFile);
                returnImage.Save(newfileinfo.FullName, ImageFormat.Jpeg);
                return base64String;
            }
            catch { }

            return base64String;
        }
    }
}
