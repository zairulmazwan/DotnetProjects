using MigraDoc.Rendering;
using MigraDoc.DocumentObjectModel;
using System;
using MigraDoc.DocumentObjectModel.Tables;


namespace DotnetProjects
{
    class Program
    {
        static void Main(string[] args)
        {


            Document doc = new Document(); // creating a document object
            Section section = doc.AddSection(); //creating a section in the document object

            Paragraph para = section.AddParagraph(); //adding a paragraph for the section
            para.Format.Font.Name = "Arial";
            para.Format.Font.Size = 26;
            para.Format.Font.Color = Color.FromCmyk(0,0,0,100);
            para.AddFormattedText("Hello Zairul Mazwan", TextFormat.Bold);
            para.Format.SpaceAfter = "1.0cm";

            Paragraph para1 = section.AddParagraph(); //adding a paragraph for the section
            para1.Format.Font.Name = "Arial";
            para1.Format.Font.Size = 18;
            para1.Format.Font.Color = Color.FromCmyk(0, 0, 0, 100);
            para1.Format.Font.Italic = true;
            para1.AddFormattedText("Welcome to my first PDF file");
            para1.Format.SpaceAfter = "1.0cm";


            //Table
            Table table = new Table();
            table.Borders.Width = 0.8;
            table.TopPadding = 5;
            table.BottomPadding = 5;
            table.LeftPadding = 10;

            //the ordering is important! create column first
            Column col = table.AddColumn(Unit.FromCentimeter(2)); //first column
            col.Format.Alignment = ParagraphAlignment.Left;
            table.AddColumn(Unit.FromCentimeter(7)); //Second column

            Row row = table.AddRow();
            row.Shading.Color = Colors.PaleGoldenrod;

            Cell cell = new Cell(); //create a cell object
            cell = row.Cells[0]; //adding a cell into the a row
            cell.AddParagraph("Number");

            cell = row.Cells[1]; //adding the second cell
            cell.AddParagraph("Student Name");

            int len = 10;

            for (int i=0; i<len; i++)
            {
                row = table.AddRow();
                cell = row.Cells[0];
                cell.AddParagraph(Convert.ToString(i + 1));
                cell = row.Cells[1];
                cell.AddParagraph("Name "+(i+1));
            }

            table.SetEdge(0,0,2,(len+1), Edge.Box, BorderStyle.Single, 1.5, Colors.Black);
            section.Add(table);

            //Rendering
            PdfDocumentRenderer pdfRend = new PdfDocumentRenderer();
            pdfRend.Document = doc;
            pdfRend.RenderDocument();
            pdfRend.PdfDocument.Save("C:\\Users\\zm1142\\DotnetProjects\\MyFile.pdf");

            Console.WriteLine("File Saved");
            Console.ReadLine();

        }
    }
}
