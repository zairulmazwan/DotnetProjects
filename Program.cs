﻿using MigraDoc.Rendering;
using MigraDoc.DocumentObjectModel;
using System;

namespace DotnetProjects
{
    class Program
    {
        static void Main(string[] args)
        {


            Document doc = new Document(); // creating a document object
            Section section = new Section(); //creating a section in the document object

            Paragraph para = section.AddParagraph(); //adding a paragraph for the section
            para.Format.Font.Name = "Arial";
            para.Format.Font.Size = 26;
            para.Format.Font.Color = Color.FromCmyk(0,0,0,100);
            para.AddFormattedText("Hello Zairul Mazwan", TextFormat.Italic);
            para.Format.SpaceAfter = "1.0cm";

            Paragraph para1 = section.AddParagraph(); //adding a paragraph for the section
            para1.Format.Font.Name = "Arial";
            para1.Format.Font.Size = 18;
            para1.Format.Font.Color = Color.FromCmyk(0, 0, 0, 100);
            para1.AddFormattedText("Welcome to my first PDF file", TextFormat.Bold);
            para1.Format.SpaceAfter = "1.0cm";


            //Rendering
            PdfDocumentRenderer pdfRend = new PdfDocumentRenderer();
            pdfRend.Document = doc;
            pdfRend.RenderDocument();
            pdfRend.PdfDocument.Save("MyFile.pdf");

            Console.WriteLine("File Saved");

        }
    }
}