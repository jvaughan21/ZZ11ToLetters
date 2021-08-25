using System;
using System.IO;
using System.Text;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.Diagnostics;
using System.Drawing;

namespace ZZ11ToLetter
{
    class LetterDetails
    {

        private string loanNumber;
        private string letterID;
        private string letterVersion;
        private string letterDesc;
        private string attachmentCode;
        private string letterContent;
        public const string DEFAULTLETTERDESCRIPTION = "SF Notices";

        public LetterDetails(string inLoanNumber, string inLetterID, string inLetterVersion, string inLetterDesc, string inAttachmentCode)
        {
            loanNumber = inLoanNumber;
            letterID = inLetterID;
            letterVersion = inLetterVersion;
            letterDesc = DEFAULTLETTERDESCRIPTION;
            attachmentCode = inAttachmentCode;
            letterContent = "";
        }

        public LetterDetails(string inParseLine)
        {
            loanNumber = inParseLine.Substring(4, 10);
            letterID = inParseLine.Substring(14, 5);
            attachmentCode = inParseLine.Substring(32, 2);
            letterDesc = DEFAULTLETTERDESCRIPTION;
            letterVersion = "0";
        }

        public LetterDetails()
        {
            loanNumber = "header"; // default to create filename
            letterID = "";
            letterVersion = "";
            letterDesc = "";
            attachmentCode = "";
            letterContent = "";
        }

        public void AddContentLine(Config c, string letterLine)
        {
            if (c.OutFileType.Equals("TEXT"))
            {
                if (letterLine.Substring(0, 1).Equals("+")) // found a OLLW line that is to be boldface. append differently
                    letterContent = letterContent.Substring(0, letterContent.Length - 1) + Convert.ToChar((byte)0x0d) + " " + letterLine.Substring(1);
                else letterContent += letterLine;
            }
            else letterContent += letterLine;
        }

        public void WriteLetter(Config c, string fileNumber)
        {
            if (c.OutFileType.Equals("TEXT"))
                WriteTextLetter(c, fileNumber);
            else WritePDFLetter(c, fileNumber);
        } // public void WriteLetter(...

        public void WriteTextLetter(Config c, string fileNumber)
        {
            //string fileN = "";

            if (loanNumber.Equals("header"))  // ignore the print alignment letters in the files
                return;

            //fileN = GetHeader().Substring(3);
            //fileN = fileN.Substring(0, fileN.Length - 3);
            //fileN += "_" + fileNumber;

            try
            {
                System.IO.File.WriteAllText(c.TempDirectory + "\\" + OutputFilename() + ".ltr", LetterContent, Encoding.GetEncoding(1252));
            }
            catch (Exception e)
            {
                System.Console.WriteLine("\n\n Error: " + e.Message);
                System.Console.WriteLine(" Unable to write letter file.  An error occured: ");
                System.Environment.Exit(102);
            }
        } // public void WriteLetter(...

        public void WritePDFLetter(Config c, string fileNumber)
        {
            //string fileN = "";

            if (loanNumber.Equals("header"))  // ignore the print alignment letters in the files
                return;

            //fileN = GetHeader().Substring(3);
            //fileN = fileN.Substring(0, fileN.Length - 3);
            //fileN += "_" + fileNumber;

            //System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            FileStream output = new FileStream(c.TempDirectory + "\\" + OutputFilename() + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            Document doc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(doc, output);
            doc.SetPageSize(new iTextSharp.text.Rectangle(PageSize.LETTER));

            BaseFont bfCourier = BaseFont.CreateFont( c.PDFFontLocation + "\\" + Config.COURIERFONTFILE , BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);
            //iTextSharp.text.Font courier = new iTextSharp.text.Font(bfCourier, c.PDFFontPitch);
            BaseFont bfCourierBold = BaseFont.CreateFont(c.PDFFontLocation + "\\" + Config.COURIERFONTBOLDFILE, BaseFont.WINANSI, BaseFont.EMBEDDED);
            //iTextSharp.text.Font courierBold = new iTextSharp.text.Font(bfCourierBold, c.PDFFontPitch);

            

            doc.Open();

            String letterText = LetterContent;
            string prevLine = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"; // intentional jibberish

            PdfContentByte cb = writer.DirectContent;
            // load logo files tiff image 
            System.Drawing.Bitmap bm = new System.Drawing.Bitmap(c.PDFLogoFileName);
            iTextSharp.text.Image imgRLMSLogo = iTextSharp.text.Image.GetInstance(bm, System.Drawing.Imaging.ImageFormat.Tiff);
            imgRLMSLogo.ScalePercent(72f / imgRLMSLogo.DpiX * c.PDFLogoScale);
            imgRLMSLogo.SetAbsolutePosition(c.PDFLogoXPosition, c.PDFLogoYPosition);
            

            System.Drawing.Bitmap bm2 = new System.Drawing.Bitmap(c.PDFEHOFileName);
            iTextSharp.text.Image imgEHOLogo = iTextSharp.text.Image.GetInstance(bm2, System.Drawing.Imaging.ImageFormat.Tiff);
            imgEHOLogo.ScalePercent(72f / imgEHOLogo.DpiX * c.PDFLogoScale);
            imgEHOLogo.SetAbsolutePosition(c.PDFEHOXPosition, c.PDFEHOYPosition);
            cb.AddImage(imgEHOLogo);

            cb.AddImage(imgRLMSLogo);
            cb.AddImage(imgEHOLogo);
            cb.BeginText();
            cb.SetFontAndSize(bfCourier, c.PDFFontPitch);
            float ypos = c.PDFTopPageOne;
            //string[] txt = letterText.Split('\n');
            foreach (string txt in letterText.Split('\n'))
            {
                
                if (prevLine.Equals("ABCDEFGHIJKLMNOPQRSTUVWXYZ"))  // this ignores the first LINE (tag line) for the letter
                {
                    prevLine = txt;
                }
                else if ((txt != "") && (txt.Length == 1) && (txt.IndexOf((char)0x0c) > -1)) // form feed character found, start new page
                {
                    cb.EndText();
                    ypos = c.PDFTopPageTwoPlus;
                    doc.NewPage();

                    cb.AddImage(imgRLMSLogo);
                    cb.AddImage(imgEHOLogo);

                    cb.BeginText();
                    cb.SetFontAndSize(bfCourier, c.PDFFontPitch);
                }
                else if ((txt.Length > 0) && txt.Substring(0,1).Equals("+"))
                {
                    //Console.WriteLine("BOLD FOUND");
                    ypos += c.PDFFontPitch;
                    cb.SetFontAndSize(bfCourierBold, c.PDFFontPitch);
                    cb.SetTextMatrix(c.PDFLeftTextMargin, ypos);
                    cb.ShowText(" " + txt.Substring(1));
                    cb.SetFontAndSize(bfCourier, c.PDFFontPitch);
                    ypos -= c.PDFFontPitch;
                }
                else if (!txt.Equals(prevLine) || txt.Equals(""))  // normal text line.  Write it and move on.
                {
                    cb.SetTextMatrix(c.PDFLeftTextMargin, ypos);
                    cb.ShowText(txt);
                    ypos -= c.PDFFontPitch;
                    prevLine = txt;
                }

            } // for Each
            cb.EndText();

            // Add Attachments
            PdfReader attachmentReader = c.GetPDFAttachment(LetterID);
            if (attachmentReader != null)  // if not null, attachments needed
            {
                for (int i=1; i <= attachmentReader.NumberOfPages; i++)
                {
                    PdfImportedPage page = writer.GetImportedPage(attachmentReader, i);
                    doc.NewPage();
                    iTextSharp.text.Image Img = iTextSharp.text.Image.GetInstance(page);
                    Img.ScaleAbsolute((float)(page.Width * 1.0), (float)(page.Height * 1.0));
                    Img.SetAbsolutePosition(0f, 0f);
                    doc.Add(Img);
                }
            }

            doc.Close();

        } // WritePDFLetter

   
        public String GetHeader()
        {
            return "***" + loanNumber + "_" + letterDesc + "_" + letterID + "_" + attachmentCode + "_" + letterVersion + "***";
        }

        public string LetterContent
        {
            get { return "***" + loanNumber + "_" + letterDesc + "_" + letterID + "_" + attachmentCode + "_" + letterVersion + "***" + "\n" + letterContent; }
            set { LetterContent = value; }
        } 

        public string OutputFilename()
        {
            return loanNumber + "_" + DateTime.Now.ToString("MMddyyyy") + "_" + letterID + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff");
        }

        

        public string LetterID
        {
            get { return letterID; }
            set { letterID = value; }
        }

        public string LetterDesc
        {
            get { return LetterDesc; }
            set { letterDesc = value; }
        }

        public string LoanNumber
        {
            get { return loanNumber; }
            set { loanNumber = value; }
        }

        public string LetterVersion
        {
            get { return LetterVersion; }
            set { letterVersion = value;}
        }

        public string AttachmentCode
        {
            get { return attachmentCode; }
            set { attachmentCode = value; }
        }

    }  // Class Letter Details
}
