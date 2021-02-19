using System;
using System.Collections.Generic;
using System.Text;
using FilesToPDFConvertor.Infrastructure;
using System.Configuration;
using System.IO;
using System.Drawing;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;
using Syncfusion.XlsIO;
using Syncfusion.ExcelToPdfConverter;
using Syncfusion.Presentation;
using Syncfusion.PresentationToPdfConverter;
using Syncfusion.Pdf.Graphics;
using Syncfusion.Pdf.Parsing;
using Syncfusion.Office;

namespace FilesToPDFConvertor
{
    public class ConvertFiles
    {
        public static void Main(string[] args)
        {
            String companyId = String.Empty;
            try
            {
                //String action = "OTHER";
                //companyId = "8";
                //String meetingId = "3024";
                //String agendaId = "10302";
                //String annuxerId = "4420,4421";
                ////string supportingDocId = "2003";

                String action = Convert.ToString(args[0]);
                String meetingId = String.Empty;
                String agendaId = String.Empty;
                String annuxerId = String.Empty;
                String supportingDocId = String.Empty;

                //Console.WriteLine(action);
                //Console.ReadLine();
                //Console.WriteLine(companyId);
                //Console.ReadLine();
                //Console.WriteLine(meetingId);
                //Console.ReadLine();
                //Console.WriteLine(agendaId);
                //Console.ReadLine();

                if (action.ToUpper() == "OTHER")
                {
                    companyId = Convert.ToString(args[1]);
                    meetingId = Convert.ToString(args[2]);
                    agendaId = Convert.ToString(args[3]);
                    annuxerId = String.Empty;
                    if (args.Length > 4)
                    {
                        annuxerId = Convert.ToString(args[4]);
                    }
                    //Console.WriteLine("GetAgendaFilesToConvert");
                    //Console.ReadLine();
                    GetAgendaFilesToConvert(Convert.ToInt32(companyId), meetingId, agendaId);
                    if (!String.IsNullOrEmpty(meetingId) && !String.IsNullOrEmpty(agendaId) && !String.IsNullOrEmpty(annuxerId))
                    {
                        GetAnnuxersFilesToConvert(Convert.ToInt32(companyId), meetingId, agendaId, annuxerId);
                    }
                }
                else if (action.ToUpper() == "WITHDRAW")
                {
                    String mode = Convert.ToString(args[1]);
                    companyId = Convert.ToString(args[2]);
                    meetingId = Convert.ToString(args[3]);
                    agendaId = Convert.ToString(args[4]);
                    //annuxerId = Convert.ToString(args[4]);
                    if (mode == "AGENDA")
                    {
                        AddWatermarktoAgendaPdf(Convert.ToInt32(companyId), meetingId, agendaId);
                    }
                    else if (mode == "ANNUXER")
                    {
                        AddWatermarktoAnnuxerPdf(Convert.ToInt32(companyId), meetingId, agendaId, String.Empty);
                    }
                }
                else if (action.ToUpper() == "CONVERTSUPPORTINGDOCUMENT")
                {
                    companyId = Convert.ToString(args[1]);
                    meetingId = Convert.ToString(args[2]);
                    agendaId = Convert.ToString(args[3]);
                    supportingDocId = Convert.ToString(args[4]);
                    GetAgendaSupportingDocToConvert(Convert.ToInt32(companyId), meetingId, agendaId, supportingDocId);
                }
                else if (action.ToUpper() == "PUBLISHITEMS")
                {
                    companyId = Convert.ToString(args[1]);
                    meetingId = Convert.ToString(args[2]);
                    agendaId = Convert.ToString(args[3]);
                    InsertPDFTextToSQL(Convert.ToInt32(companyId), meetingId, agendaId);
                }
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "Main Method", "FilesToPDFConvertor Scheduler", 1, Convert.ToInt32(companyId));
            }
        }

        private static void GetAgendaFilesToConvert(Int32 companyId, String meetingId, String agendaId)
        {
            try
            {
                MeetingRepository meetingRepo = new MeetingRepository();

                Meeting objMeeting = new Meeting();
                objMeeting = meetingRepo.GetAgendaDocs(companyId, meetingId, agendaId);
                if (objMeeting != null)
                {
                    if (objMeeting.agendaItems != null)
                    {
                        if (objMeeting.agendaItems.Count > 0)
                        {
                            String itemdir = CryptorEngine.Decrypt(ConfigurationManager.AppSettings["FilesPath"], true);
                            String targetDirectory = itemdir + "Meeting_" + meetingId;

                            if (Directory.Exists(targetDirectory))
                            {
                                String[] fileEntries = Directory.GetFiles(targetDirectory);

                                foreach (AgendaItems objAgendaItem in objMeeting.agendaItems)
                                {
                                    foreach (String fileName in fileEntries)
                                    {
                                        if (objAgendaItem.agendaDoc.Trim().ToUpper() == Path.GetFileName(fileName.Trim().ToUpper()))
                                        {
                                            //Console.WriteLine("FILE MATCH");
                                            //Console.ReadLine();
                                            String extension = Path.GetExtension(objAgendaItem.agendaDoc);
                                            switch (extension.Split('.')[1].ToUpper())
                                            {
                                                case "DOCX":
                                                case "DOC":
                                                    ConvertWordToPdf(targetDirectory, Path.GetFileName(fileName), companyId, extension);
                                                    break;
                                                case "XLSX":
                                                case "XLS":
                                                    ConvertExcelToPdf(targetDirectory, Path.GetFileName(fileName), companyId);
                                                    break;
                                                case "PPTX":
                                                    ConvertPPtToPdf(targetDirectory, Path.GetFileName(fileName), companyId);
                                                    break;
                                                case "JPEG":
                                                case "BMP":
                                                case "JPG":
                                                case "PNG":
                                                case "GIF":
                                                    ConvertImageToPdf(targetDirectory, Path.GetFileName(fileName), companyId);
                                                    break;
                                                case "TXT":
                                                    ConvertTextToPdf(targetDirectory, Path.GetFileName(fileName), companyId);
                                                    break;
                                                default:
                                                    break;
                                            }
                                            //Console.WriteLine("fileName=" + fileName);
                                            //Console.ReadLine();
                                            objAgendaItem.pageFrom = 1;
                                            objAgendaItem.pageTo = ReadPageCount(targetDirectory, Path.GetFileNameWithoutExtension(fileName) + ".pdf", companyId);
                                            //Console.WriteLine("pageTo=" + objAgendaItem.pageTo);
                                            //Console.ReadLine();
                                        }
                                    }
                                }
                            }
                            objMeeting.moduleDatabase = "PROCS_BOARD_MEETING";
                            bool status = meetingRepo.UpdateAgendaFilesPageCount(objMeeting, companyId);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "GetAgendaFilesToConvert", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
        }

        private static void GetAgendaSupportingDocToConvert(Int32 companyId, String meetingId, String agendaId, String documentId)
        {
            try
            {
                MeetingRepository meetingRepo = new MeetingRepository();

                Meeting objMeeting = new Meeting();
                objMeeting = meetingRepo.GetAgendaSupportingDoc(companyId, meetingId, agendaId, documentId);
                if (objMeeting != null)
                {
                    if (objMeeting.agendaItems != null)
                    {
                        if (objMeeting.agendaItems.Count > 0)
                        {
                            AgendaItems objAgendaItem = objMeeting.agendaItems[0];

                            String itemdir = CryptorEngine.Decrypt(ConfigurationManager.AppSettings["FilesPath"], true);
                            String targetDirectory = itemdir + "Meeting_" + meetingId;

                            if (Directory.Exists(targetDirectory))
                            {
                                String[] fileEntries = Directory.GetFiles(targetDirectory);

                                foreach (AgendaItemSupportingDocument objDoc in objAgendaItem.listSupportingDocument)
                                {
                                    foreach (String fileName in fileEntries)
                                    {
                                        if (objDoc.originalDocumentName.Trim().ToUpper() == Path.GetFileName(fileName.Trim().ToUpper()))
                                        {
                                            String extension = Path.GetExtension(objDoc.originalDocumentName);
                                            switch (extension.Split('.')[1].ToUpper())
                                            {
                                                case "DOCX":
                                                case "DOC":
                                                    ConvertWordToPdf(targetDirectory, Path.GetFileName(fileName), objDoc.documentName, companyId, extension);
                                                    break;
                                                case "XLSX":
                                                case "XLS":
                                                    ConvertExcelToPdf(targetDirectory, Path.GetFileName(fileName), objDoc.documentName, companyId);
                                                    break;
                                                case "PPTX":
                                                    ConvertPPtToPdf(targetDirectory, Path.GetFileName(fileName), objDoc.documentName, companyId);
                                                    break;
                                                case "JPEG":
                                                case "BMP":
                                                case "JPG":
                                                case "PNG":
                                                case "GIF":
                                                    ConvertImageToPdf(targetDirectory, Path.GetFileName(fileName), objDoc.documentName, companyId);
                                                    break;
                                                case "TXT":
                                                    ConvertTextToPdf(targetDirectory, Path.GetFileName(fileName), objDoc.documentName, companyId);
                                                    break;
                                                case "PDF":
                                                    SavePdfSupportingDoc(targetDirectory, Path.GetFileName(fileName), objDoc.documentName, companyId);
                                                    break;
                                                default:
                                                    break;
                                            }
                                            MergeFiles(objAgendaItem.agendaDoc, objDoc.documentName, objMeeting.ID, objAgendaItem.ID, objAgendaItem.isItemMeredTosupportingDocument, companyId);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "GetAgendaSupportingDocToConvert", "FilesToPDFConvertor Scheduler", 1, Convert.ToInt32(companyId));
            }
        }

        private static void GetAnnuxersFilesToConvert(Int32 companyId, String meetingId, String agendaId, String annuxerId)
        {
            try
            {
                MeetingRepository meetingRepo = new MeetingRepository();
                Meeting objMeeting = new Meeting();
                objMeeting = meetingRepo.GetAnnuxerDocForReplace(companyId, meetingId, agendaId, annuxerId);
                if (objMeeting != null)
                {
                    if (objMeeting.agendaItems != null)
                    {
                        if (objMeeting.agendaItems.Count > 0)
                        {
                            if (objMeeting.agendaItems[0].agendaAnnuxers != null)
                            {
                                if (objMeeting.agendaItems[0].agendaAnnuxers.Count > 0)
                                {
                                    String itemdir = CryptorEngine.Decrypt(ConfigurationManager.AppSettings["FilesPath"], true);
                                    String targetDirectory = itemdir + "Meeting_" + meetingId;

                                    if (Directory.Exists(targetDirectory))
                                    {
                                        String[] fileEntries = Directory.GetFiles(targetDirectory);
                                        foreach (AgendaAnnuxers objAgendaAnnuxers in objMeeting.agendaItems[0].agendaAnnuxers)
                                        {
                                            foreach (String fileName in fileEntries)
                                            {
                                                if (objAgendaAnnuxers.annuxerDoc.Trim().ToUpper() == Path.GetFileName(fileName.Trim().ToUpper()))
                                                {
                                                    String extension = Path.GetExtension(objAgendaAnnuxers.annuxerDoc);
                                                    switch (extension.Split('.')[1].ToUpper())
                                                    {
                                                        case "DOCX":
                                                        case "DOC":
                                                            ConvertWordToPdf(targetDirectory, Path.GetFileName(fileName), companyId, extension);
                                                            break;
                                                        case "XLSX":
                                                        case "XLS":
                                                            ConvertExcelToPdf(targetDirectory, Path.GetFileName(fileName), companyId);
                                                            break;
                                                        case "PPTX":
                                                            ConvertPPtToPdf(targetDirectory, Path.GetFileName(fileName), companyId);
                                                            break;
                                                        case "JPEG":
                                                        case "BMP":
                                                        case "JPG":
                                                        case "PNG":
                                                        case "GIF":
                                                            ConvertImageToPdf(targetDirectory, Path.GetFileName(fileName), companyId);
                                                            break;
                                                        case "TXT":
                                                            ConvertTextToPdf(targetDirectory, Path.GetFileName(fileName), companyId);
                                                            break;
                                                        default:
                                                            break;
                                                    }
                                                    objAgendaAnnuxers.pageFrom = 1;
                                                    objAgendaAnnuxers.pageTo = ReadPageCount(targetDirectory, Path.GetFileNameWithoutExtension(fileName) + ".pdf", companyId);
                                                }
                                            }
                                        }
                                    }
                                    objMeeting.moduleDatabase = "PROCS_BOARD_MEETING";
                                    bool status = meetingRepo.UpdateAnnuxerFilesPageCount(objMeeting, companyId);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "GetAnnuxersFilesToConvert", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
        }

        private static void AddWatermarktoAgendaPdf(Int32 companyId, String meetingId, String agendaId)
        {
            try
            {
                MeetingRepository meetingRepo = new MeetingRepository();

                Meeting objMeeting = new Meeting();
                objMeeting = meetingRepo.GetWithdrawnAgendaDocs(companyId, meetingId, agendaId);
                if (objMeeting != null)
                {
                    if (objMeeting.agendaItems != null)
                    {
                        if (objMeeting.agendaItems.Count > 0)
                        {
                            String itemdir = CryptorEngine.Decrypt(ConfigurationManager.AppSettings["FilesPath"], true);
                            String targetDirectory = itemdir + "Meeting_" + meetingId + "\\PublishPdfFiles\\";

                            if (Directory.Exists(targetDirectory))
                            {
                                String[] fileEntries = Directory.GetFiles(targetDirectory);

                                foreach (AgendaItems objAgendaItem in objMeeting.agendaItems)
                                {
                                    foreach (String fileName in fileEntries)
                                    {
                                        if (objAgendaItem.agendaDoc.Trim().ToUpper() == Path.GetFileName(fileName.Trim().ToUpper()))
                                        {
                                            AddWatermark(targetDirectory, fileName, companyId);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "AddWatermarktoPdf", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
        }

        private static void AddWatermarktoAnnuxerPdf(Int32 companyId, String meetingId, String agendaId, String annuxerId)
        {
            try
            {
                List<Tuple<Int32, Int32>> listIds = new List<Tuple<Int32, Int32>>();
                String newfileName = String.Empty;
                MeetingRepository meetingRepo = new MeetingRepository();
                Meeting objMeeting = new Meeting();
                objMeeting = meetingRepo.GetWithdrawnAnnuxerDocs(companyId, meetingId, agendaId, annuxerId);
                if (objMeeting != null)
                {
                    if (objMeeting.agendaItems != null)
                    {
                        if (objMeeting.agendaItems.Count > 0)
                        {
                            if (objMeeting.agendaItems[0].agendaAnnuxers != null)
                            {
                                if (objMeeting.agendaItems[0].agendaAnnuxers.Count > 0)
                                {
                                    String itemdir = CryptorEngine.Decrypt(ConfigurationManager.AppSettings["FilesPath"], true);
                                    String targetDirectory = itemdir + "Meeting_" + meetingId + "\\PublishPdfFiles\\";

                                    if (Directory.Exists(targetDirectory))
                                    {
                                        String[] fileEntries = Directory.GetFiles(targetDirectory);
                                        foreach (AgendaAnnuxers objAgendaAnnuxers in objMeeting.agendaItems[0].agendaAnnuxers)
                                        {
                                            foreach (String fileName in fileEntries)
                                            {
                                                if (objAgendaAnnuxers.annuxerDoc.Trim().ToUpper() == Path.GetFileName(fileName.Trim().ToUpper()))
                                                {
                                                    newfileName = fileName;
                                                    listIds.Add(new Tuple<Int32, Int32>(objAgendaAnnuxers.pageFrom, objAgendaAnnuxers.pageTo));
                                                }
                                            }
                                        }
                                        if (listIds.Count > 0)
                                        {
                                            AddWatermark(targetDirectory, newfileName, listIds, companyId);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "GetAnnuxersFilesToConvert", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
        }

        private static void ConvertWordToPdf(String filePath, String fileName, Int32 companyId, String fileExtension)
        {
            ConvertWordToPdf(filePath, fileName, null, companyId, fileExtension);
        }

        private static void ConvertWordToPdf(String filePath, String fileName, String supportingdocName, Int32 companyId, String fileExtension)
        {
            try
            {
                //Loads an existing Word document
                WordDocument wordDocument;
                if (fileExtension.ToUpper() == "DOCX")
                {
                    wordDocument = new WordDocument(Path.Combine(filePath, fileName), Syncfusion.DocIO.FormatType.Docx);
                }
                else
                {
                    wordDocument = new WordDocument(Path.Combine(filePath, fileName), Syncfusion.DocIO.FormatType.Doc);
                }

                //WordDocument wordDocument = new WordDocument(Path.Combine(filePath, fileName), Syncfusion.DocIO.FormatType.Docx);

                ////Initialize chart to image converter for converting charts during Word to pdf conversion
                //wordDocument.ChartToImageConverter = new Syncfusion.OfficeChartToImageConverter.ChartToImageConverter();

                ////Set the scaling mode for charts
                //wordDocument.ChartToImageConverter.ScalingMode = Syncfusion.OfficeChart.ScalingMode.Normal;

                //create an instance of DocToPDFConverter - responsible for Word to PDF conversion
                DocToPDFConverter converter = new DocToPDFConverter();
                converter.Settings.EmbedFonts = true;
                converter.Settings.EmbedCompleteFonts = true;

                ////Set the image quality
                //converter.Settings.ImageQuality = 100;

                ////Set the image resolution
                //converter.Settings.ImageResolution = 640;

                ////Set true to optimize the memory usage for identical images
                //converter.Settings.OptimizeIdenticalImages = true;

                //Convert Word document into PDF document
                PdfDocument pdfDocument = converter.ConvertToPDF(wordDocument);
                //pdfDocument.PageSettings.Size = PdfPageSize.Letter;

                //Save the PDF file to file system
                String newName = String.Empty;
                if (!String.IsNullOrEmpty(Convert.ToString(supportingdocName)))
                {
                    newName = supportingdocName;
                }
                else
                {
                    newName = Path.GetFileNameWithoutExtension(fileName) + ".pdf";
                }
                pdfDocument.Save(Path.Combine(filePath, newName));

                //close the instance of document objects
                pdfDocument.Close(true);

                wordDocument.Close();
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "ConvertWordToPdf", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
        }

        public static void ConvertExcelToPdf(String filePath, String fileName, Int32 companyId)
        {
            ConvertExcelToPdf(filePath, fileName, null, companyId);
        }

        public static void ConvertExcelToPdf(String filePath, String fileName, String supportingdocName, Int32 companyId)
        {
            try
            {
                //ExcelEngine excelEngine = new ExcelEngine();
                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    IApplication application = excelEngine.Excel;
                    application.DefaultVersion = ExcelVersion.Excel2016;

                    //Instantiating the ChartToImageConverter and assigning the ChartToImageConverter instance of XlsIO application
                    application.ChartToImageConverter = new Syncfusion.ExcelChartToImageConverter.ChartToImageConverter();

                    //Tuning chart image quality
                    application.ChartToImageConverter.ScalingMode = Syncfusion.XlsIO.ScalingMode.Best;

                    IWorkbook workbook = application.Workbooks.Open(Path.Combine(filePath, fileName), ExcelOpenType.Automatic);

                    //bool isMacroEnabled = workbook.HasMacros;

                    //if (isMacroEnabled)
                    //{
                    //Accessing Vba project
                    IVbaProject project = workbook.VbaProject;

                    //Accessing vba modules collection
                    IVbaModules vbaModules = project.Modules;

                    //Remove all macros
                    vbaModules.Clear();

                    //Saving as macro document
                    workbook.SaveAs(Path.Combine(filePath, fileName));
                    workbook.Close();

                    workbook = application.Workbooks.Open(Path.Combine(filePath, fileName), ExcelOpenType.Automatic);

                   // foreach (IWorksheet worksheet in workbook.Worksheets)
                  //  {
                   //     worksheet.PageSetup.PaperSize = ExcelPaperSize.PaperA4;
                   // }

                    //Initialize warning class to capture warnings during the conversion.
                    Warning warning = new Warning();

                    //Open the Excel Document to Convert
                    ExcelToPdfConverter converter = new ExcelToPdfConverter(workbook);

                    //Intialize the ExcelToPdfconverterSettings
                    ExcelToPdfConverterSettings settings = new ExcelToPdfConverterSettings();

                    //Set the warning class that is implemented.
                    settings.Warning = warning;

                    //Set the Layout Options for the output Pdf page.
                    settings.LayoutOptions = LayoutOptions.FitAllColumnsOnOnePage;

                    //Initialize PDF Document
                    PdfDocument pdfDocument = new PdfDocument();
                    //pdfDocument.PageSettings.Size = PdfPageSize.A4;

                    //Convert Excel Document into PDF document
                    pdfDocument = converter.Convert(settings);

                    //Save the pdf file
                    String newName = String.Empty;
                    if (!String.IsNullOrEmpty(Convert.ToString(supportingdocName)))
                    {
                        newName = supportingdocName;
                    }
                    else
                    {
                        newName = Path.GetFileNameWithoutExtension(fileName) + ".pdf";
                    }
                    pdfDocument.Save(Path.Combine(filePath, newName));

                    //Dispose the objects
                    pdfDocument.Close(true);
                    converter.Dispose();
                    workbook.Close();
                    excelEngine.Dispose();
                }
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "ConvertExcelToPdf", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
        }

        public static void ConvertPPtToPdf(String filePath, String fileName, Int32 companyId)
        {
            ConvertPPtToPdf(filePath, fileName, null, companyId);
        }

        public static void ConvertPPtToPdf(String filePath, String fileName, String supportingdocName, Int32 companyId)
        {
            try
            {
                //Opens a PowerPoint Presentation
                IPresentation presentation = Presentation.Open(Path.Combine(filePath, fileName));

                ////Creates an instance of ChartToImageConverter and assigns it to ChartToImageConverter property of Presentation
                //presentation.ChartToImageConverter = new Syncfusion.OfficeChartToImageConverter.ChartToImageConverter();

                ////Sets the scaling mode of the chart to best.
                //presentation.ChartToImageConverter.ScalingMode = Syncfusion.OfficeChart.ScalingMode.Best;

                //Instantiates the Presentation to pdf converter settings instance.
                PresentationToPdfConverterSettings settings = new PresentationToPdfConverterSettings();
                settings.EmbedCompleteFonts = true;

                ////Set the image resolution
                //settings.ImageResolution = 100;

                ////Set the image quality
                //settings.ImageQuality = 100;

                //Sets the option for adding hidden slides to converted pdf
                settings.ShowHiddenSlides = false;

                //Sets the slide per page settings; this is optional.
                //settings.SlidesPerPage = SlidesPerPage.One;

                //Sets the settings to enable notes pages while conversion.
                //settings.PublishOptions = PublishOptions.NotesPages;

                //Converts the PowerPoint Presentation into PDF document
                PdfDocument pdfDocument = PresentationToPdfConverter.Convert(presentation, settings);
                //pdfDocument.PageSettings.Size = PdfPageSize.A4;
                //pdfDocument.PageSettings.Margins.All = 0;
                //pdfDocument.PageSettings.Orientation = PdfPageOrientation.Landscape;

                //Saves the PDF document
                String newName = String.Empty;
                if (!String.IsNullOrEmpty(Convert.ToString(supportingdocName)))
                {
                    newName = supportingdocName;
                }
                else
                {
                    newName = Path.GetFileNameWithoutExtension(fileName) + ".pdf";
                }
                pdfDocument.Save(Path.Combine(filePath, newName));

                //Closes the PDF document
                pdfDocument.Close(true);

                //Closes the Presentation
                presentation.Close();
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.Message);
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "ConvertPPtToPdf", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
        }

        public static void ConvertImageToPdf(String filePath, String fileName, Int32 companyId)
        {
            ConvertImageToPdf(filePath, fileName, null, companyId);
        }

        public static void ConvertImageToPdf(String filePath, String fileName, String supportingdocName, Int32 companyId)
        {
            try
            {
                //Create a new PDF document
                PdfDocument doc = new PdfDocument();
                doc.PageSettings.Size = PdfPageSize.A4;

                //Add a page to the document
                PdfPage page = doc.Pages.Add();

                //Create PDF graphics for the page
                PdfGraphics graphics = page.Graphics;

                //Load the image from the disk
                PdfBitmap image = new PdfBitmap(Path.Combine(filePath, fileName));

                //Draw the image
                graphics.DrawImage(image, 0, 0);

                //Save the document
                String newName = String.Empty;
                if (!String.IsNullOrEmpty(Convert.ToString(supportingdocName)))
                {
                    newName = supportingdocName;
                }
                else
                {
                    newName = Path.GetFileNameWithoutExtension(fileName) + ".pdf";
                }
                doc.Save(Path.Combine(filePath, newName));

                //Close the document
                doc.Close(true);
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "ConvertImageToPdf", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
        }

        public static void ConvertTextToPdf(String filePath, String fileName, Int32 companyId)
        {
            ConvertTextToPdf(filePath, fileName, null, companyId);
        }

        public static void ConvertTextToPdf(String filePath, String fileName, String supportingdocName, Int32 companyId)
        {
            try
            {
                //Create a PDF document instance
                PdfDocument document = new PdfDocument();
                document.PageSettings.Size = PdfPageSize.A4;

                //Add page to the document
                PdfPage page = document.Pages.Add();
                PdfGraphics graphics = page.Graphics;

                //Read the long text from the text file.
                StreamReader reader = new StreamReader(Path.Combine(filePath, fileName), Encoding.ASCII);
                string text = reader.ReadToEnd();
                reader.Close();

                const int paragraphGap = 10;

                //Create a text element with the text and font
                PdfTextElement textElement = new PdfTextElement(text, new PdfStandardFont(PdfFontFamily.TimesRoman, 14));
                PdfLayoutFormat layoutFormat = new PdfLayoutFormat();
                layoutFormat.Layout = PdfLayoutType.Paginate;
                layoutFormat.Break = PdfLayoutBreakType.FitPage;

                //Draw the first paragraph
                PdfLayoutResult result = textElement.Draw(page, new RectangleF(0, 0, page.GetClientSize().Width / 2, page.GetClientSize().Height), layoutFormat);

                //Draw the second paragraph from the first paragraph end position
                result = textElement.Draw(page, new RectangleF(0, result.Bounds.Bottom + paragraphGap, page.GetClientSize().Width / 2, page.GetClientSize().Height), layoutFormat);

                String newName = String.Empty;
                if (!String.IsNullOrEmpty(Convert.ToString(supportingdocName)))
                {
                    newName = supportingdocName;
                }
                else
                {
                    newName = Path.GetFileNameWithoutExtension(fileName) + ".pdf";
                }
                document.Save(Path.Combine(filePath, newName));
                document.Close(true);
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "ConvertTextToPdf", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
        }

        public static void SavePdfSupportingDoc(String filePath, String fileName, String supportingdocName, Int32 companyId)
        {
            try
            {
                //Create a PDF document instance
                PdfLoadedDocument loadedDocument = new PdfLoadedDocument(Path.Combine(filePath, fileName));

                String newName = String.Empty;
                if (!String.IsNullOrEmpty(Convert.ToString(supportingdocName)))
                {
                    newName = supportingdocName;
                }
                else
                {
                    newName = Path.GetFileNameWithoutExtension(fileName) + ".pdf";
                }
                loadedDocument.Save(Path.Combine(filePath, newName));
                loadedDocument.Close(true);
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "SavePdfSupportingDoc", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
        }

        public static Int32 ReadPageCount(String filePath, String fileName, Int32 companyId)
        {
            Int32 pageCount = 0;
            try
            {
                //Load the PDF document.
                PdfLoadedDocument loadedDocument = new PdfLoadedDocument(Path.Combine(filePath, fileName));

                //Get the page count.
                pageCount = loadedDocument.Pages.Count;

                //Close the document.
                loadedDocument.Close(true);
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "ReadPageCount", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
            return pageCount;
        }

        private static bool AddWatermark(String filePath, String fileName, Int32 companyId)
        {
            try
            {
                //Loads the pdf document
                PdfLoadedDocument ldoc = new PdfLoadedDocument(Path.Combine(filePath, fileName));

                //itereate through the pages of loaded document
                foreach (PdfLoadedPage lpage in ldoc.Pages)
                {
                    PdfGraphics graphics = lpage.Graphics;

                    //set the font
                    PdfFont font = new PdfStandardFont(PdfFontFamily.Helvetica, 70F);

                    //Create template from the loaded page
                    PdfTemplate template = lpage.CreateTemplate();

                    // watermark text.
                    PdfGraphicsState state = graphics.Save();

                    graphics.TranslateTransform((lpage.Size.Width / 2) - 200, lpage.Size.Height / 2 + 100);

                    graphics.SetTransparency(0.25f);

                    graphics.RotateTransform(-45);

                    graphics.DrawString("WITHDRAWN", font, PdfPens.Red, PdfBrushes.Red, new PointF(0, 0));
                }

                //Save the document and dispose it   
                String newName = "Watermark_" + Path.GetFileNameWithoutExtension(fileName) + ".pdf";
                ldoc.Save(Path.Combine(filePath, newName));
                ldoc.Close(true);

                return true;
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "AddWatermark", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
            return false;
        }

        private static bool AddWatermark(String filePath, String fileName, List<Tuple<Int32, Int32>> listIds, Int32 companyId)
        {
            try
            {
                Int32 mPageFrom = 0;
                Int32 mPageTo = 0;

                //Loads the pdf document
                PdfLoadedDocument ldoc = new PdfLoadedDocument(Path.Combine(filePath, fileName));

                foreach (var pages in listIds)
                {
                    mPageFrom = pages.Item1 - 1;
                    mPageTo = pages.Item2 - 1;

                    //itereate through the pages of loaded document
                    while (mPageFrom <= mPageTo)
                    {
                        PdfPageBase lpage = ldoc.Pages[mPageFrom];

                        PdfGraphics graphics = lpage.Graphics;

                        //set the font
                        PdfFont font = new PdfStandardFont(PdfFontFamily.Helvetica, 70F);

                        //Create template from the loaded page
                        PdfTemplate template = lpage.CreateTemplate();

                        // watermark text.
                        PdfGraphicsState state = graphics.Save();

                        graphics.TranslateTransform((lpage.Size.Width / 2) - 200, lpage.Size.Height / 2 + 100);

                        graphics.SetTransparency(0.25f);

                        graphics.RotateTransform(-45);

                        graphics.DrawString("WITHDRAWN", font, PdfPens.Red, PdfBrushes.Red, new PointF(0, 0));

                        mPageFrom += mPageFrom + 1;
                    }
                }

                //Save the document and dispose it   
                String newName = "Watermark_" + Path.GetFileNameWithoutExtension(fileName) + ".pdf";
                ldoc.Save(Path.Combine(filePath, newName));
                ldoc.Close(true);

                return true;
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message.ToString(), ex.Source, ex.StackTrace, "FilesToPDFConvertor", "AddWatermark", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
            return false;
        }

        private static bool MergeFiles(String agendaDoc, String pdfDocumentName, Int32 meetingId, Int32 agendaId, bool isAlreadyMerged, Int32 companyId)
        {
            try
            {
                String itemdir = CryptorEngine.Decrypt(ConfigurationManager.AppSettings["FilesPath"], true);
                String targetDirectory = itemdir + "Meeting_" + meetingId;

                if (Directory.Exists(targetDirectory))
                {
                    String[] fileEntries = Directory.GetFiles(targetDirectory);

                    if (File.Exists(Path.Combine(targetDirectory, Path.GetFileNameWithoutExtension(agendaDoc) + ".pdf")))
                    {
                        if (!isAlreadyMerged)
                        {
                            //Load the existing PDF document
                            PdfLoadedDocument loadedDoc = new PdfLoadedDocument(Path.Combine(targetDirectory, Path.GetFileNameWithoutExtension(agendaDoc) + ".pdf"));

                            //Save the PDF document                                        
                            loadedDoc.Save(Path.Combine(targetDirectory, Path.GetFileNameWithoutExtension(agendaDoc) + "_BM.pdf"));

                            //Close the document.
                            loadedDoc.Close(true);
                        }

                        Stream[] streams = new Stream[2];

                        FileStream file = new FileStream(Path.Combine(targetDirectory, Path.GetFileNameWithoutExtension(agendaDoc) + ".pdf"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                        streams[0] = file;

                        if (File.Exists(Path.Combine(targetDirectory, pdfDocumentName)))
                        {
                            file = new FileStream(Path.Combine(targetDirectory, pdfDocumentName), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                            streams[1] = file;

                            PdfDocument finalDoc = new PdfDocument();

                            //Optimizes memory while merging PDF documents. 
                            finalDoc.EnableMemoryOptimization = true;

                            PdfDocument.Merge(finalDoc, streams);

                            //Save the PDF document                            
                            finalDoc.Save(Path.Combine(targetDirectory, Path.GetFileNameWithoutExtension(agendaDoc) + ".pdf"));

                            //Close the document
                            finalDoc.Close(true);

                            Meeting objMeeting = new Meeting();
                            objMeeting.ID = meetingId;
                            objMeeting.moduleDatabase = "PROCS_BOARD_MEETING";

                            AgendaItems objAgendaItem = new AgendaItems();
                            objAgendaItem.ID = agendaId;
                            objAgendaItem.pageFrom = 1;
                            objAgendaItem.pageTo = ReadPageCount(targetDirectory, Path.GetFileNameWithoutExtension(agendaDoc) + ".pdf", companyId);
                            objMeeting.agendaItems = new List<AgendaItems>();
                            objMeeting.agendaItems.Add(objAgendaItem);

                            MeetingRepository meetingRepo = new MeetingRepository();
                            bool status = meetingRepo.UpdateAgendaFilesPageCount(objMeeting, companyId);
                            meetingRepo.UpdateItemMergedOrNot(companyId, meetingId, agendaId);
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message, ex.Source, ex.StackTrace, "FilesToPDFConvertor", "MergeFiles", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
            return false;
        }

        private static bool InsertPDFTextToSQL(Int32 companyId, String meetingId, String agendaId)
        {            
            try
            {
                MeetingRepository meetingRepo = new MeetingRepository();
                List<AgendaItems> lstAgendaItems = new List<AgendaItems>();

                Meeting objMeeting = new Meeting();
                objMeeting = meetingRepo.GetAgendaDocsToReadFileText(companyId, meetingId, agendaId);
                if (objMeeting != null)
                {
                    if (objMeeting.agendaItems != null)
                    {
                        if (objMeeting.agendaItems.Count > 0)
                        {                            
                            String itemdir = CryptorEngine.Decrypt(ConfigurationManager.AppSettings["FilesPath"], true);
                            String targetDirectory = itemdir + "Meeting_" + meetingId + "/PublishPdfFiles/";

                            if (Directory.Exists(targetDirectory))
                            {
                                String[] fileEntries = Directory.GetFiles(targetDirectory);

                                foreach (AgendaItems objAgendaItem in objMeeting.agendaItems)
                                {
                                    foreach (String fileName in fileEntries)
                                    {
                                        if (objAgendaItem.agendaDoc.Trim().ToUpper() == Path.GetFileName(fileName.Trim().ToUpper()))
                                        {
                                            objAgendaItem.fileContent = ExtractTextFromPdf(targetDirectory, objAgendaItem.agendaDoc, companyId);
                                            lstAgendaItems.Add(objAgendaItem);
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if(lstAgendaItems.Count > 0)
                    {
                        objMeeting.agendaItems = new List<AgendaItems>();
                        objMeeting.agendaItems = lstAgendaItems;
                        objMeeting.companyId = companyId;
                        meetingRepo.InsertPublishItemsTextIntoSQL(objMeeting);
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {                
                new LogHelper().AddExceptionLogs(ex.Message, ex.Source, ex.StackTrace, "FilesToPDFConvertor", "InsertPDFTextToSQL", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
            return false;
        }

        public static string ExtractTextFromPdf(String filePath, String fileName, Int32 companyId)
        {
            try
            {
                // Load an existing PDF document.
                PdfLoadedDocument loadedDocument = new PdfLoadedDocument(Path.Combine(filePath, fileName));

                // Loading page collections
                PdfLoadedPageCollection loadedPages = loadedDocument.Pages;

                string extractedText = string.Empty;

                // Extract text from existing PDF document pages
                foreach (PdfLoadedPage loadedPage in loadedPages)
                {
                    extractedText += loadedPage.ExtractText(true);
                }

                //Close the document
                loadedDocument.Close(true);

                return extractedText;
            }
            catch(Exception ex)
            {
                new LogHelper().AddExceptionLogs(ex.Message, ex.Source, ex.StackTrace, "FilesToPDFConvertor", "ExtractTextFromPdf", "FilesToPDFConvertor Scheduler", 1, companyId);
            }
            return String.Empty;
        }
    }    

    public class Warning : Syncfusion.XlsIO.IWarning
    {
        public void ShowWarning(Syncfusion.XlsIO.WarningInfo warning)
        {
            //Cancel the converion process if the warning type is conditional formatting.
            if (warning.Type == Syncfusion.XlsIO.WarningType.PageSettings)
                Cancel = true;
            //else if (warning.Type == Syncfusion.XlsIO.WarningType.ConditionalFormatting)
            //    Cancel = true;

            //To view or log the warning, you can make use of warning.Description.
        }
        public bool Cancel { get; set; }
    }
}
