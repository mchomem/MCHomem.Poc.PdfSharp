using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using System;
using System.Diagnostics;
using System.Text;

namespace MCHomem.Poc.PdfSharp
{
    class Program
    {
        static void Main(string[] args)
        {
            ShowMenu();
            Console.ReadKey();
        }

        private static void ShowMenu()
        {
            while (true)
            {
                Console.Clear();
                Console.WriteLine("Menu\r\n");
                Console.WriteLine("1 - Hello Word");
                Console.WriteLine("2 - Box");
                Console.WriteLine("3 - Box, inner word and paper orientation");
                Console.WriteLine("4 - Create PDF with long text");
                Console.WriteLine("5 - Report layout");
                Console.WriteLine("6 - Online samples");
                Console.WriteLine("0 - Exit\r\n");

                Console.Write("Choose a option: ");
                String op = Console.ReadLine();
                Console.WriteLine();

                switch (op)
                {
                    case "1":
                        CreateTestPDF();
                        break;

                    case "2":
                        CreateBox();
                        break;

                    case "3":
                        CreateTestSizePagePDF();
                        break;

                    case "4":
                        CreatePDFWithLongText();
                        break;

                    case "5":
                        ReportLayout();
                        break;

                    case "6":
                        OpenOnlineSamples();
                        break;

                    case "0":
                        Exit();
                        break;

                    default:
                        ShowMessage(MessageType.WARNING, String.Format("The value \"{0}\" is incorrect! Choose a correct option from menu.", op));
                        break;
                }
            }
        }

        private static void CreateTestPDF()
        {
            try
            {
                PdfDocument pdf = new PdfDocument();
                PdfPage page = pdf.AddPage();
                XGraphics gfx = XGraphics.FromPdfPage(page);

                XFont font = new XFont("Verdana", 20, XFontStyle.Bold);
                gfx.DrawString("Hello, World!", font, XBrushes.Black, new XRect(0, 0, page.Width, page.Height), XStringFormats.Center);

                String fileName = @"D:\Test.pdf";
                pdf.Save(fileName);
                Process.Start(String.Format("file:///{0}", fileName));
                ShowMessage(MessageType.SUCCESS, "Pdf done!");
            }
            catch (Exception e)
            {
                ShowMessage(MessageType.ERROR, String.Format("Error!\r\nMessage {0}", e.Message));
            }
        }

        private static void CreateBox()
        {
            try
            {
                PdfDocument pdf = new PdfDocument();
                PdfPage page = pdf.AddPage();
                XGraphics gfx = XGraphics.FromPdfPage(page);

                XPen pen = new XPen(XColors.Black, 1);
                XRect rectangle = new XRect(20, 20, page.Width - 50, page.Height - 50);
                gfx.DrawRectangle(pen, rectangle);

                String fileName = @"D:\Test.pdf";
                pdf.Save(fileName);
                Process.Start(String.Format("file:///{0}", fileName));
                ShowMessage(MessageType.SUCCESS, "Pdf done!");
            }
            catch (Exception e)
            {
                ShowMessage(MessageType.ERROR, String.Format("Error!\r\nMessage {0}", e.Message));
            }
        }

        private static void CreateTestSizePagePDF()
        {
            try
            {
                PdfDocument pdf = new PdfDocument();
                PdfPage page = pdf.AddPage();
                page.Size = PageSize.A5;
                page.Orientation = PageOrientation.Landscape;
                XGraphics gfx = XGraphics.FromPdfPage(page);

                XFont font = new XFont("Verdana", 20, XFontStyle.Bold);
                gfx.DrawString("Hello, World!", font, XBrushes.Black, new XRect(0, 0, page.Width, page.Height), XStringFormats.Center);

                XPen pen = new XPen(XColors.Black, 1);
                XRect rectangle = new XRect(20, 20, page.Width - 50, page.Height - 50);
                gfx.DrawRectangle(pen, rectangle);

                String fileName = @"D:\Test.pdf";
                pdf.Save(fileName);
                Process.Start(String.Format("file:///{0}", fileName));
                ShowMessage(MessageType.SUCCESS, "Pdf done!");
            }
            catch (Exception e)
            {
                ShowMessage(MessageType.ERROR, String.Format("Error!\r\nMessage {0}", e.Message));
            }
        }

        private static void CreatePDFWithLongText()
        {
            PdfDocument pdf = new PdfDocument();

            StringBuilder longText = new StringBuilder();
            Random r = new Random();
            Int32 insertSpace = 0;
            Int32 maxCharacters = 200000;
            Int32 maxCharactersPerPage = 5000;

            for (int i = 0, count = 0; i < maxCharacters; i++, count++)
            {
                insertSpace++;

                Char c = Convert.ToChar(r.Next(127));

                if (char.IsLetter(c) || char.IsNumber(c))
                {
                    if (insertSpace.Equals(10))
                    {
                        longText.Append(" ");
                        insertSpace = 0;
                    }
                    else
                    {
                        longText.Append(c.ToString());
                    }
                }
                else
                {
                    // In this point the "c" char is not a letter or number, but insertSpace is higher 10, then receive zero.
                    if (insertSpace > 10)
                        insertSpace = 0;
                    i--;
                }

                if (count > maxCharactersPerPage)
                {
                    PdfPage page = pdf.AddPage();
                    XGraphics gfx = XGraphics.FromPdfPage(page);
                    XTextFormatter tf = new XTextFormatter(gfx);
                    tf.Alignment = XParagraphAlignment.Justify;
                    XRect rect = new XRect(20, 20, page.Width - 50, page.Height - 50);
                    gfx.DrawRectangle(XBrushes.White, rect);
                    XFont font = new XFont("Verdana", 12, XFontStyle.Regular);
                    tf.DrawString(longText.ToString(), font, XBrushes.Black, rect);
                    longText.Clear();
                    count = 0;
                }
            }

            String fileName = @"D:\Test.pdf";
            pdf.Save(fileName);
            Process.Start(String.Format("file:///{0}", fileName));
            ShowMessage(MessageType.SUCCESS, "Pdf done!");
        }

        private static void OpenOnlineSamples()
        {
            try
            {
                Process.Start("http://www.pdfsharp.net/wiki/PDFsharpSamples.ashx");
                ShowMessage(MessageType.SUCCESS, "Done.");
            }
            catch (Exception e)
            {
                ShowMessage(MessageType.ERROR, String.Format("Error!\r\nMessage {0}", e.Message));
            }
        }

        private static void ReportLayout()
        {
            try
            {
                PdfDocument pdf = new PdfDocument();
                PdfPage page = pdf.AddPage();
                page.Size = PageSize.A4;
                page.Orientation = PageOrientation.Portrait;
                XGraphics gfx = XGraphics.FromPdfPage(page);

                #region Header

                Int32 xLocationHeader = 20;
                Int32 yLocationHeader = 20;
                Double wHeader = page.Width - 40;
                Double hHeader = page.Height * 0.05;

                XFont fontTitle = new XFont("Verdana", 20, XFontStyle.Bold);
                XRect rectPlaceTitleHeader = new XRect(xLocationHeader, yLocationHeader, wHeader, hHeader);
                gfx.DrawString("Report title", fontTitle, XBrushes.Black, rectPlaceTitleHeader, XStringFormats.Center);

                XPen penHeader = new XPen(XColors.Black, 0.5);
                XRect rectBorderHeader = new XRect(xLocationHeader, yLocationHeader, wHeader, hHeader);
                gfx.DrawRectangle(penHeader, rectBorderHeader);

                #endregion

                #region Body

                Int32 xLocationBody = xLocationHeader;
                Int32 yLocationBody = yLocationHeader + Convert.ToInt32(hHeader);
                Double wBody = wHeader;
                Double hBody = page.Height * 0.85;

                XFont fontBody = new XFont("Verdana", 12, XFontStyle.Regular);
                XRect rectPlaceBody = new XRect(xLocationBody, yLocationBody, wBody, hBody);
                gfx.DrawString("Text for body", fontBody, XBrushes.Black, rectPlaceBody, XStringFormats.TopLeft);

                XPen penBody = new XPen(XColors.Black, 0.5);
                XRect rectBorderBody = new XRect(xLocationBody, yLocationBody, wBody, hBody);
                gfx.DrawRectangle(penBody, rectBorderBody);

                #endregion

                #region Footer

                Int32 xLocationFooter = xLocationHeader;
                Int32 yLocationFooter = yLocationBody + Convert.ToInt32(hBody);
                Double wFooter = wHeader;
                Double hFooter = page.Height * 0.05;

                XFont fontFooter = new XFont("Verdana", 12, XFontStyle.Regular);
                XRect rectPlaceFooter = new XRect(xLocationFooter, yLocationFooter, wFooter, hFooter);
                gfx.DrawString("Text for footer", fontFooter, XBrushes.Black, rectPlaceFooter, XStringFormats.TopLeft);

                XPen penFooter = new XPen(XColors.Black, 0.5);
                XRect rectBorderFooter = new XRect(xLocationFooter, yLocationFooter, wFooter, hFooter);
                gfx.DrawRectangle(penFooter, rectBorderFooter);

                #endregion

                String fileName = @"D:\Test.pdf";
                pdf.Save(fileName);
                Process.Start(String.Format("file:///{0}", fileName));
                ShowMessage(MessageType.SUCCESS, "Pdf done!");
            }
            catch (Exception e)
            {
                ShowMessage(MessageType.ERROR, String.Format("Error!\r\nMessage {0}", e.Message));
            }
        }

        private static void Exit()
        {
            ShowMessage(MessageType.INFORMATION, "Press a key");
            Environment.Exit(0);
        }

        private static void ShowMessage(MessageType messageType, String message)
        {
            switch (messageType)
            {
                case MessageType.ERROR:
                    Console.ForegroundColor = ConsoleColor.Red;
                    break;
                case MessageType.INFORMATION:
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    break;
                case MessageType.SUCCESS:
                    Console.ForegroundColor = ConsoleColor.Green;
                    break;
                case MessageType.WARNING:
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    break;
                default:
                    break;
            }
            Console.WriteLine(message);
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.ReadKey();
        }
    }

    public enum MessageType
    {
        ERROR
        , INFORMATION
        , SUCCESS
        , WARNING
    }
}
