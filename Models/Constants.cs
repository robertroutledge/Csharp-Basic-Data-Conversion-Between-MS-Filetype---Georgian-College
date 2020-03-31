using System;
using System.Collections.Generic;
using System.Text;

namespace RoutledgeAssignmentFour.Models
{
    public class Constants
    {
        public readonly Student Student = new Student { StudentId = "200449068", FirstName = "Robert", LastName = "Routledge" };

        public class Locations
        {
        ///public readonly static string DesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

        public readonly static string ExePath = Environment.CurrentDirectory;

        public readonly static string ContentFolder = $"{ExePath}\\..\\..\\..\\Content";
        public readonly static string DataFolder = $"{ContentFolder}\\Data";
        public readonly static string ImagesFolder = $"{ContentFolder}\\Images";
        public readonly static string ExcelFile = $"{ContentFolder}\\info.xlsx";             
        public readonly static string DOCXFile = $"{ContentFolder}\\info.docx";
        public readonly static string PPTXFile = $"{ContentFolder}\\info.pptx";
        public const string InfoFile = "info.csv";
        public const string ImageFile = "myimage.jpg";
        public readonly static string ImageSaveFile = $"{ContentFolder}\\myimage.jpg";
            
        }

        public class FTP
        {
            public const string UserName = @"bdat100119f\bdat1001";
            public const string Password = "bdat1001";
            public const string remoteUploadFileDestinationdocx = BaseUrl + "/200449068 Robert Routledge/info.docx";
            public const string remoteUploadFileDestinationxlsx = BaseUrl + "/200449068 Robert Routledge/info.xlsx";
            //public const string remoteUploadFileDestinationxml = BaseUrl + "/200449068 Robert Routledge/students.xml";
            

            public const string BaseUrl = "ftp://waws-prod-dm1-127.ftp.azurewebsites.windows.net/bdat1001-20914";

            public const int OperationPauseTime = 5000;
        }
    }
}

