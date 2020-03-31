using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using RoutledgeAssignmentFour.Models;
using RoutledgeAssignmentFour.Models.Utilities;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;


namespace RoutledgeAssignmentFour
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Student> students = new List<Student>();
            byte[] myimagebytes = new byte[0];

            ///Connects to FTP and builds list of directories, Populates Student list, gets my image
            List<string> directories = FTP.GetDirectory(Constants.FTP.BaseUrl);
            foreach (var directory in directories)
            {
                Student student = new Student() { AbsoluteUrl = Constants.FTP.BaseUrl };
                student.FromDirectory(directory);
                string infofilepath = student.FullPathUrl + "/" + Constants.Locations.InfoFile;
                string imagefilepath = student.FullPathUrl + "/" + Constants.Locations.ImageFile;
                bool InfofileExists = FTP.FileExists(infofilepath);
                bool ImagefileExists = FTP.FileExists(imagefilepath);
                if (InfofileExists == true)
                {
                    ///Downloads the csv files into bytes
                    byte[] bytes = FTP.DownloadFileBytes(infofilepath);
                    string csvData = Encoding.Default.GetString(bytes);
                    string[] csvlines = csvData.Split("\r\n", StringSplitOptions.RemoveEmptyEntries);
                    if (csvlines.Length != 2)
                    {
                        Console.WriteLine("Error in CSV format.");
                    }
                    else
                    {
                        student.FromCSV(csvlines[1]);
                        students.Add(student);
                        if (ImagefileExists == true)
                        {
                            if (student.IsMe == true)
                            {
                                byte[] imagebytes = FTP.DownloadFileBytes(imagefilepath);
                                Imaging.ConvertandSaveImage(imagebytes);
                            }
                            else continue;
                        }
                        else continue;
                    }
                }
            }
            
        XMLDocument.CreateWordprocessingDocument(Constants.Locations.DOCXFile, myimagebytes, students);
        XMLSpreadsheet.CreateSpreadsheetWorkbook(Constants.Locations.ExcelFile, students);
        FTPUpload.UploadFile(Constants.Locations.DOCXFile, Constants.FTP.remoteUploadFileDestinationdocx);
        FTPUpload.UploadFile(Constants.Locations.ExcelFile, Constants.FTP.remoteUploadFileDestinationxlsx);

        }   
    }
}