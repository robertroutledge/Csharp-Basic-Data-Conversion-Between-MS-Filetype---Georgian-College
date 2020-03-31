using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace RoutledgeAssignmentFour.Models
{
    public class Student
    {
         
        public static string HeaderRow = "{nameof(Student.UniqueId)},{nameof(Student.StudentId)},{nameof(Student.FirstName)},{nameof(Student.LastName)},{nameof(Student.DateOfBirth)},nameof(Student.ImageData),nameof(Student.IsMe)}";
        public Guid UniqueID { get; set; }
        public string StudentId { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        
        private string _DateOfBirth;
        public string DateOfBirth
        {
            get { return _DateOfBirth; }
            set
            {
                _DateOfBirth = value;

                //Convert DateOfBirth to DateTime
                DateTime dtOut;
                DateTime.TryParse(_DateOfBirth, out dtOut);
                DateOfBirthDT = dtOut;
            }
        }
        public DateTime DateOfBirthDT { get; internal set; }
        public bool IsMe { get; set; }

        public virtual int Age
        {
            get
            {
                DateTime Now = DateTime.Now;

                int age = Now.Year - DateOfBirthDT.Year;
                if (DateOfBirthDT > Now.AddYears(-age))
                    age--;
             return age;
            }
        }

        public string ImageString { get; set; }
        public byte[] ImageBytes { get; set; }

    //public Image image { get {  return Utilities.Imaging.Base64ToImage(ImageBase64); } }
    public string AbsoluteUrl { get; set; }

        public string FullPathUrl
        {
            get
            {
                return AbsoluteUrl + "/" + Directory;
            }
        }
        public string Directory { get; set; }

        public List<string> Exceptions { get; set; } = new List<string>();

        public void FromCSV(string csvdata)
        {
            string[] data = csvdata.Split(",", StringSplitOptions.None);
            try
            {
                
                StudentId = data[0];
                FirstName = data[1];
                LastName = data[2];
                DateOfBirth = data[3];
                ImageString = data[4];                
                IsMe = false;
                UniqueID = Guid.NewGuid();
            }
            catch (Exception e)
            {
                Exceptions.Add(e.Message);
            }
            if (StudentId == "200449068")
            {
                IsMe = true;
            }
        }
                 
        public void FromDirectory(string directory)
        {
            Directory = directory;

            if (String.IsNullOrEmpty(directory.Trim()))
            {
                return;
            }
           
            string[] data = directory.Trim().Split(" ", StringSplitOptions.None);
            StudentId = data[0];
            FirstName = data[1];
            LastName = data[2];
        }        

    }
        
}

