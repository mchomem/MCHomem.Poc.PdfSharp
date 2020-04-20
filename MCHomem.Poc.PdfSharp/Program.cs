using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using System;
using System.Configuration;
using System.Diagnostics;
using System.Text;

namespace MCHomem.Poc.PdfSharp
{
    class Program
    {
        public static String LogFilesDirPath
        {
            get
            {
                String defaultAppKey = "LOG_FILES_DIR_PATH";
                String path = @"c:\PdfShartpTests\Logs";

                if (ConfigurationManager.AppSettings[defaultAppKey] != null)
                {
                    if (!String.IsNullOrEmpty(ConfigurationManager.AppSettings[defaultAppKey]))
                    {
                        path = ConfigurationManager.AppSettings[defaultAppKey];
                    }
                }

                return path;
            }
        }

        public static String SampleFilesDirPath
        {
            get
            {
                String defaultAppKey = "SAMPLE_FILES_DIR_PATH";
                String path = @"c:\PdfShartpTests\Samples";

                if (ConfigurationManager.AppSettings[defaultAppKey] != null)
                {
                    if (!String.IsNullOrEmpty(ConfigurationManager.AppSettings[defaultAppKey]))
                    {
                        path = ConfigurationManager.AppSettings[defaultAppKey];
                    }
                }

                return path;
            }
        }

        #region Main method

        static void Main(string[] args)
        {
            Console.Title = "PDF Sharp tests";
            ShowMenu();
            Console.ReadKey();
        }

        #endregion

        #region Static methods

        private static void ShowMenu()
        {
            while (true)
            {
                Console.Clear();
                Console.WriteLine("Menu\r\n");
                ShowMessage(MessageType.INFORMATION, "\tChoose a number from below menu to create a pdf file\r\n\tand show in default browser:\r\n", false);
                Console.WriteLine("\t1 - Hello Word.");
                Console.WriteLine("\t2 - Box.");
                Console.WriteLine("\t3 - Box, inner word and paper orientation.");
                Console.WriteLine("\t4 - Create PDF with long text.");
                Console.WriteLine("\t5 - Report layout.");
                Console.WriteLine("\t6 - Online samples.");
                Console.WriteLine("\t0 - Exit\r\n");
                Console.Write("Choose a option: ");
                String op = Console.ReadLine();
                Console.WriteLine();

                String fullFilePath = String.Empty;

                switch (op)
                {
                    case "1":
                        fullFilePath = String.Format(@"{0}\{1}", FileDirectoryHelper.GetDirPath(SampleFilesDirPath), "hello-world-example.pdf");
                        CreateHelloWorldPDF(fullFilePath);
                        break;

                    case "2":
                        fullFilePath = String.Format(@"{0}\{1}", FileDirectoryHelper.GetDirPath(SampleFilesDirPath), "box-example.pdf");
                        CreateBox(fullFilePath);
                        break;

                    case "3":
                        fullFilePath = String.Format(@"{0}\{1}", FileDirectoryHelper.GetDirPath(SampleFilesDirPath), "page-example.pdf");
                        CreatePagePDF(fullFilePath);
                        break;

                    case "4":
                        fullFilePath = String.Format(@"{0}\{1}", FileDirectoryHelper.GetDirPath(SampleFilesDirPath), "pages-example.pdf");
                        CreatePagesPDF(fullFilePath);
                        break;

                    case "5":
                        fullFilePath = String.Format(@"{0}\{1}", FileDirectoryHelper.GetDirPath(SampleFilesDirPath), "report-layout-example.pdf");
                        ReportLayout(fullFilePath);
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

        private static void CreateHelloWorldPDF(String fullFilePath)
        {
            try
            {
                PdfDocument pdf = new PdfDocument();
                PdfPage page = pdf.AddPage();
                XGraphics gfx = XGraphics.FromPdfPage(page);

                XFont font = new XFont("Verdana", 20, XFontStyle.Bold);
                gfx.DrawString("Hello, World!", font, XBrushes.Black, new XRect(0, 0, page.Width, page.Height), XStringFormats.Center);

                pdf.Save(fullFilePath);
                Process.Start(String.Format("file:///{0}", fullFilePath));
                ShowMessage(MessageType.SUCCESS, "Pdf done!");
            }
            catch (Exception e)
            {
                ShowMessage(MessageType.ERROR, String.Format("Error!\r\nMessage {0}", e.Message));
                CreateLogError(e);
            }
        }

        private static void CreateBox(String fullFilePath)
        {
            try
            {
                PdfDocument pdf = new PdfDocument();
                PdfPage page = pdf.AddPage();
                XGraphics gfx = XGraphics.FromPdfPage(page);

                XPen pen = new XPen(XColors.Black, 1);
                XRect rectangle = new XRect(20, 20, page.Width - 50, page.Height - 50);
                gfx.DrawRectangle(pen, rectangle);

                pdf.Save(fullFilePath);
                Process.Start(String.Format("file:///{0}", fullFilePath));
                ShowMessage(MessageType.SUCCESS, "Pdf done!");
            }
            catch (Exception e)
            {
                ShowMessage(MessageType.ERROR, String.Format("Error!\r\nMessage {0}", e.Message));
                CreateLogError(e);
            }
        }

        private static void CreatePagePDF(String fullFilePath)
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

                pdf.Save(fullFilePath);
                Process.Start(String.Format("file:///{0}", fullFilePath));
                ShowMessage(MessageType.SUCCESS, "Pdf done!");
            }
            catch (Exception e)
            {
                ShowMessage(MessageType.ERROR, String.Format("Error!\r\nMessage {0}", e.Message));
                CreateLogError(e);
            }
        }

        private static void CreatePagesPDF(String fullFilePath)
        {
            try
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

                pdf.Save(fullFilePath);
                Process.Start(String.Format("file:///{0}", fullFilePath));
                ShowMessage(MessageType.SUCCESS, "Pdf done!");
            }
            catch (Exception e)
            {
                ShowMessage(MessageType.ERROR, String.Format("Error!\r\nMessage {0}", e.Message));
                CreateLogError(e);
            }
        }

        private static void ReportLayout(String fullFilePath)
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

                pdf.Save(fullFilePath);
                Process.Start(String.Format("file:///{0}", fullFilePath));
                ShowMessage(MessageType.SUCCESS, "Pdf done!");
            }
            catch (Exception e)
            {
                ShowMessage(MessageType.ERROR, String.Format("Error!\r\nMessage {0}", e.Message));
                CreateLogError(e);
            }
        }

        private static void Exit()
        {
            ShowMessage(MessageType.INFORMATION, "Press a key to exit.");
            Environment.Exit(0);
        }

        private static void ShowMessage(MessageType messageType, String message, Boolean useReadyKey = true)
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

            if (useReadyKey)
            {
                Console.ReadKey();
            }
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

        private static void CreateLogError(Exception e)
        {
            FileDirectoryHelper.CreateFile
                    (
                        String.Format("Message: {0}\r\nStacktrace: {1}\r\n", e.Message, e.StackTrace)
                        , String.Format(@"{0}\{1}", FileDirectoryHelper.GetDirPath(LogFilesDirPath), String.Format("Error_{0}.log", DateTime.Now.ToString("ddMMyyyy.HHmmss")))
                    );
        }

        #endregion
    }

    public enum MessageType
    {
        ERROR
        , INFORMATION
        , SUCCESS
        , WARNING
    }
}
