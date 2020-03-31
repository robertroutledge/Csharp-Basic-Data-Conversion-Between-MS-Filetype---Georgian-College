using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;


namespace RoutledgeAssignmentFour.Models.Utilities
{
    public class XMLDocument
    {
        
        /// <summary>
        /// Take the filepath, and students data, add it to a word document
        /// </summary>
        public static void CreateWordprocessingDocument(string filepath, List<Student> students)
        {
            // Create a document by supplying the filepath. 
            using (WordprocessingDocument wordDocument =
                WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                
                
                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run1 = para.AppendChild(new Run());
                run1.AppendChild(new Text("Here is a list of students from the ftp directory, with each student getting their own page: "));

                foreach (var student in students)
                {
                    ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Bmp);
                    using (MemoryStream stream = new MemoryStream(student.ImageBytes))
                    {
                        imagePart.FeedData(stream);
                    }
                    var temp = mainPart.GetIdOfPart(imagePart);
                    Paragraph para2 = body.AppendChild(new Paragraph());
                    Run run2 = para2.AppendChild(new Run(
                    new Break() { Type = BreakValues.Page }));
                    run2.AppendChild(new Text(student.FirstName + " " + student.LastName)); 
                    AddImageToBody(wordDocument, student,temp);                  

                }
            }
        }
        /// <summary>
        /// Take the image from above and defines it in xml format for insertion into .docx.
        /// </summary>  
        private static void AddImageToBody(WordprocessingDocument wordDoc, Student student, string relationshipId)
        {
            int imageWidth = 0;
            int imageHeight = 0;
            using (Bitmap bmp = Imaging.BytesToBitmap(student.ImageBytes) as Bitmap)
            {
                imageWidth = bmp.Width * 9525;
                imageHeight = bmp.Height * 9525;
            }

            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = imageWidth, Cy = imageHeight },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture of Student"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "myimage.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                       new PIC.BlipFill(
                                 new A.Blip(
                                     new A.BlipExtensionList(
                                         new A.BlipExtension()
                                         {
                                             Uri =
                                               "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                         })
                                 )
                                 {
                                     Embed = relationshipId,
                                     CompressionState =
                                     A.BlipCompressionValues.Print
                                 },
                                 new A.Stretch(
                                     new A.FillRectangle())),
                             new PIC.ShapeProperties(
                                 new A.Transform2D(
                                     new A.Offset() { X = 0L, Y = 0L },
                                     new A.Extents() { Cx = imageWidth, Cy = imageHeight }),
                                 new A.PresetGeometry(
                                     new A.AdjustValueList()
                                 )
                                 { Preset = A.ShapeTypeValues.Rectangle }))
                     )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
             )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            // Append the reference to the body. The element should be in 
            // a DocumentFormat.OpenXml.Wordprocessing.Run.
            wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
        }
    }
}

