using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Policy;

namespace ZZ11ToLetter
{
    class Config
    {
        private string configFilename = null;
        private string inputDirectory = null;
        private string outputDirectory = null;
        private string logFileDirectory = null;
        private string logFileName = null;
        private string tempDirectory = null;
        private string archiveDirectory = null;
        private string outFileType = "TEXT";  // Text File Type is the fefault
        private List<KeyValuePair<byte, byte>> xRefs = new List<KeyValuePair<byte, byte>>();
        private float pdfLeftTextMargin = 72.0f;
        private float pdfTopPageOne = 680.0f;
        private float pdfTopPageTwoPlus = 756f;
        private float pdfFontPitch = 10.5f;
        private string pdfRLMSLogoFileName = "";
        private string pdfEHOLogoFileName = "";
        private float pdfLogoXPosition = 39.0f;
        private float pdfLogoYPosition = 680.0f;
        private float pdfEHOXPosition = 475.0f;
        private float pdfEHOYPosition = 25.0f;
        private string pdfFontLocation = "C:\\Windows\\Fonts";
        private float pdfLogoScale = 33.0f;
        public const string COURIERFONTFILE = "cour.ttf";
        public const string COURIERFONTBOLDFILE = "courbd.ttf";
        private string attachmentDirectory = null;
        private List<LetterAttachment> attachmentsList = new List<LetterAttachment>();

        public Config(string inFilename)
        {
            configFilename = inFilename;
        }

        public Config()
        {
            
        }

        public void ReadConfiguration()
        {
            String[] lines = System.IO.File.ReadAllLines(configFilename);
            String line;

            if (File.Exists(configFilename))
                lines = System.IO.File.ReadAllLines(configFilename);
            else
            {
                System.Console.WriteLine(" EXCEPTION: CONFIGURATION FILE not found: " + configFilename);
                System.Environment.Exit(97);
            }

            foreach (string lne in lines)
            {
                line = lne.ToUpper().Trim();

                if (line.Trim().Length > 0 && line.IndexOf("=") > -1)
                {
                    if (line.Substring(0, line.IndexOf("=")).Equals("INPUTDIR"))
                        InputDirectory = line.Substring(line.IndexOf("=") + 1).Trim();
                    else if (line.Substring(0, line.IndexOf("=")).Equals("OUTPUTDIR"))
                        OutputDirectory = line.Substring(line.IndexOf("=") + 1).Trim();
                    else if (line.Substring(0, line.IndexOf("=")).Equals("LOGFILEDIR"))
                        LogFileDirectory = line.Substring(line.IndexOf("=") + 1).Trim();
                    else if (line.Substring(0, line.IndexOf("=")).Equals("LOGFILENAME"))
                        LogFileName = line.Substring(line.IndexOf("=") + 1).Trim();
                    else if (line.Substring(0, line.IndexOf("=")).Equals("TEMPDIR"))
                        TempDirectory = line.Substring(line.IndexOf("=") + 1).Trim();
                    else if (line.Substring(0, line.IndexOf("=")).Equals("ARCHIVEDIR"))
                        ArchiveDirectory = line.Substring(line.IndexOf("=") + 1).Trim();
                    else if (line.Substring(0, line.IndexOf("=")).Equals("FILETYPE"))
                        OutFileType = line.Substring(line.IndexOf("=") + 1).Trim();
                    else if (line.Substring(0, line.IndexOf("=")).Equals("PDFMARGINS"))
                        SetPDFMargins(line.Substring(line.IndexOf("=") + 1).Trim());
                    else if (line.Substring(0, line.IndexOf("=")).Equals("PDFRLMSLOGOFILE"))
                        RLMSLogoFileName = line.Substring(line.IndexOf("=") + 1).Trim();
                    else if (line.Substring(0, line.IndexOf("=")).Equals("PDFEHOLOGOFILE"))
                        EHOLogoFileName = line.Substring(line.IndexOf("=") + 1).Trim();
                    else if (line.Substring(0, line.IndexOf("=")).Equals("PDFRLMSLOGOPOSITION"))
                        SetPDFRLMSLogoPosition(line.Substring(line.IndexOf("=") + 1).Trim());
                    else if (line.Substring(0, line.IndexOf("=")).Equals("PDFEHOLOGOPOSITION"))
                        SetPDFEHOLogoPosition(line.Substring(line.IndexOf("=") + 1).Trim());
                    else if (line.Substring(0, line.IndexOf("=")).Equals("PDFFONTPITCH"))
                        SetFontPitch(line.Substring(line.IndexOf("=") + 1).Trim());
                    else if (line.Substring(0, line.IndexOf("=")).Equals("TRUETYPEFONTLOCATION"))
                        PDFFontLocation = line.Substring(line.IndexOf("=") + 1).Trim();
                    else if (line.Substring(0, line.IndexOf("=")).Equals("PDFLOGOSCALE"))
                        SetLogoScale(line.Substring(line.IndexOf("=") + 1).Trim());
                    else if (line.Substring(0, line.IndexOf("=")).Equals("ATTACHMENTDIR"))
                        AttachmentDirectory = line.Substring(line.IndexOf("=") + 1).Trim();
                    else if (line.Substring(0, line.IndexOf("=")).Equals("ATTACH"))
                        SetLetterAttachments(line.Substring(line.IndexOf("=") + 1).Trim());
                    else if (line.Substring(0, line.IndexOf("=")).Equals("REPLACE"))
                    {
                        string[] parms = line.Substring(line.IndexOf("=") + 1).Split('|');
                        if (parms.Length == 2)
                            xRefs.Add(new KeyValuePair<byte, byte>(Convert.ToByte(parms[0]), Convert.ToByte(parms[1])));
                        else Console.WriteLine(" REPLACE IGNORED IN CONFIG FILE: " + line);
                    }
                    else
                    {
                        Console.WriteLine("**UNKNOWN LINE IN CONFIG FILE: " + line + "\n\n");
                    }
                }
                else if (line.Length > 0)
                {
                    Console.WriteLine("**UNKNOWN LINE IN CONFIG FILE: " + line + "\n\n");
                }

            } // foreach (string line ....

            if (inputDirectory != null &&
                 outputDirectory != null &&
                 logFileDirectory != null &&
                 archiveDirectory != null &&
                 tempDirectory != null)
                return;
            else // *** raise an error here because the config file was incomplete.....
            {
                System.Console.WriteLine(" **ERROR: Configuration file invalid, missing an item, or invalid formatting.");
                System.Environment.Exit(101);
            }

        } //ReadConfiguration

        public byte FixSpanishCharacter(byte bt)
        {
            var matches = from val in xRefs where val.Key == bt select val.Value;
            foreach (KeyValuePair<byte, byte> kvp in xRefs)
                if (kvp.Key == bt)
                    return kvp.Value;
            return bt;
        }

        private void SetPDFMargins(string paramText)
        {
            string[] pdfParams = paramText.Split('|');

            if (pdfParams.Length == 3)
            {
                Console.WriteLine(" PDF Margins: " + paramText);
                pdfLeftTextMargin = float.Parse(pdfParams[0]);
                Console.WriteLine(" Left Text Margin: " + pdfLeftTextMargin);
                pdfTopPageOne = float.Parse(pdfParams[1]);
                Console.WriteLine(" Top Margin Page One: " + pdfTopPageOne);
                pdfTopPageTwoPlus = float.Parse(pdfParams[2]);
                Console.WriteLine(" Top Margin Page Two+: " + pdfTopPageTwoPlus);
            }
            else Console.WriteLine("Invalid PDF Margins configuration. Config file ignored. Using defaults {0}, {1}, {2}", pdfLeftTextMargin, pdfTopPageOne, pdfTopPageTwoPlus);
        } // SetPDFMargins
       
        private void SetLetterAttachments(string paramText)
        {
            string[] attachTxt = paramText.Split('|');
            LetterAttachment LTRA;
            if (attachTxt.Length == 2)
            {
                LTRA = attachmentsList.Find(x => x.LetterID == attachTxt[0]);
                if (LTRA == null)
                {
                    if (File.Exists(attachmentDirectory + "\\" + attachTxt[1]))
                    {
                        LTRA = new LetterAttachment(attachTxt[0], attachmentDirectory + "\\" + attachTxt[1]);
                        attachmentsList.Add(LTRA);
                        Console.WriteLine(" attachments added: " + paramText);
                    }
                    else
                    {
                        Console.WriteLine("ERROR: attachments file not found: " + attachmentDirectory + "\\" + attachTxt[1]);
                        Environment.Exit(99);
                    }
                }
                else Console.WriteLine(" IGNORED Duplicate Attachment List: " + paramText);
            }
            else Console.WriteLine(" IGNORED Invalid Attachment List: " + paramText);
        } // Set Letter Attachments


        private void SetPDFRLMSLogoPosition(string paramText)
        {
            string[] pdfParams = paramText.Split('|');

            if (pdfParams.Length == 2)
            {
                Console.WriteLine(" PDF Logo Position: " + paramText);
                pdfLogoXPosition = float.Parse(pdfParams[0]);
                Console.WriteLine(" Logo X Position: " + pdfLogoXPosition);
                pdfLogoYPosition = float.Parse(pdfParams[1]);
                Console.WriteLine(" Logo Y Position: " + pdfLogoYPosition);
            }
            else Console.WriteLine("Invalid Logo Position configuration. Config file ignored. Using defaults {0}, {1}", pdfLogoXPosition, pdfLogoYPosition);
        } // SetPDFRLMSLogoPosition

        private void SetPDFEHOLogoPosition(string paramText)
        {
            string[] pdfParams = paramText.Split('|');

            if (pdfParams.Length == 2)
            {
                Console.WriteLine(" PDF Logo Position: " + paramText);
                pdfEHOXPosition = float.Parse(pdfParams[0]);
                Console.WriteLine(" EHO X Position: " + pdfEHOXPosition);
                pdfEHOYPosition = float.Parse(pdfParams[1]);
                Console.WriteLine(" EHO Y Position: " + pdfEHOYPosition);
            }
            else Console.WriteLine("Invalid EHO Logo Position configuration. Config file ignored. Using defaults {0}, {1}", pdfEHOXPosition, pdfEHOYPosition);
        } // SetPDFMargins

        private void SetFontPitch(string paramText)
        {
            pdfFontPitch = float.Parse(paramText);
            Console.WriteLine(" PDF Font Pitch: " + pdfFontPitch);
        } // SetFontPitch

        private void SetLogoScale(string paramText)
        {
            pdfLogoScale = float.Parse(paramText);
            Console.WriteLine(" PDF Logo Scale: " + pdfLogoScale);
        } // SetFontPitch

        public string ConfigFilename
        {
            get { return configFilename; }
            set { configFilename = value; }
        }

        public PdfReader GetPDFAttachment(string inLetterID)
        {
            LetterAttachment LA = attachmentsList.Find(x => x.LetterID == inLetterID);
            if (LA != null)
                return LA.PDFReader;
            else return null;
        } // GetPDFAttachment

        public string InputDirectory
        {
            get { return inputDirectory; }
            set {
                
                if (value != null && !Directory.Exists(value))
                {
                    System.Console.WriteLine("ERROR: Input Path not found: " + value);
                    System.Environment.Exit(99);
                }
                else inputDirectory = value;
                Console.WriteLine(" Input Directory: " + inputDirectory);
            }
        }

        public string OutputDirectory
        {
            get { return outputDirectory; }
            set
            {
                if (value != null && !Directory.Exists(value))
                {
                    System.Console.WriteLine("ERROR: Output Path not found: " + value);
                    System.Environment.Exit(99);
                }
                else outputDirectory = value;
                Console.WriteLine(" Output Directory: " + outputDirectory);
            }
        }

        public string LogFileDirectory
        {
            get { return logFileDirectory; }
            set
            {
                if (value != null && !Directory.Exists(value))
                {
                    System.Console.WriteLine("ERROR: Logfile Path not found: " + value);
                    System.Environment.Exit(99);
                }
                else logFileDirectory = value;
                Console.WriteLine(" Log File Directory: " + logFileDirectory);
            }
        }

        public string LogFileName
        {
            get { return logFileName; }
            set
            {
                string tempName = value;
                int startPos = tempName.IndexOf("%D");
                int endPos = -1;
                if (startPos > -1)
                {
                    endPos = tempName.Substring(startPos + 1).IndexOf("%D");
                    if (endPos > -1)
                        logFileName = tempName.Substring(0, startPos) + DateTime.Now.ToString(tempName.Substring(startPos + 2, endPos - 1).Replace('Y', 'y').Replace('D', 'd')) + ".logs";
                    else
                    {
                        logFileName = tempName.Substring(0, startPos - 1) + ".logs";
                        Console.WriteLine("CONFIG FILE Log File Name Date Parameter not value: " + value);
                    }
                }
                else logFileName = value;
                Console.WriteLine(" Log Filename: " + logFileName);
            }
        }

        public string TempDirectory
        {
            get { return tempDirectory; }
            set
            {
                if (value != null && !Directory.Exists(value))
                {
                    System.Console.WriteLine("ERROR: Temporary Path not found: " + value);
                    System.Environment.Exit(99);
                }
                else tempDirectory = value;
                Console.WriteLine(" Temp Directory: " + tempDirectory);
            }
        }

        public string ArchiveDirectory
        {
            get { return archiveDirectory; }
            set
            {
                if (value != null && !Directory.Exists(value))
                {
                    System.Console.WriteLine("ERROR: Archive Path not found: " + value);
                    System.Environment.Exit(99);
                }
                else archiveDirectory = value;
                Console.WriteLine(" Archive Directory: " + archiveDirectory);
            }
        }

        public string AttachmentDirectory
        {
            get { return attachmentDirectory; }
            set
            {
                if (value != null && !Directory.Exists(value))
                {
                    System.Console.WriteLine("ERROR: Attachment Path not found: " + value);
                    System.Environment.Exit(99);
                }
                else attachmentDirectory = value;
                Console.WriteLine(" Attachment Directory: " + attachmentDirectory);
            }
        }

        public string PDFFontLocation
        {
            get { return pdfFontLocation; }
            set
            {
                if (value != null && !Directory.Exists(value))
                {
                    System.Console.WriteLine("ERROR: True Type Font Path not found: " + value);
                    System.Environment.Exit(99);
                }
                else if(!File.Exists(value + "\\" + COURIERFONTFILE) || !File.Exists(value + "\\" + COURIERFONTBOLDFILE))
                {
                    System.Console.WriteLine("ERROR: Unable to find True Type Font Files: \n   " + value + "\\" + COURIERFONTFILE + "\n            " + value + "\\" + COURIERFONTBOLDFILE);
                    System.Environment.Exit(99);
                }
                else pdfFontLocation = value;
                Console.WriteLine(" Font Location: " + pdfFontLocation);
            }
        }

        public string RLMSLogoFileName
        {
            get { return pdfRLMSLogoFileName; }
            set
            {
                if (value != null && !File.Exists(value))
                {
                    System.Console.WriteLine("ERROR: Logo File not found. Unable to create PDFs : " + value);
                    System.Environment.Exit(99);
                }
                else pdfRLMSLogoFileName = value;
                Console.WriteLine(" Logo File Name: " + pdfRLMSLogoFileName);
            }
        }

        public string EHOLogoFileName
        {
            get { return pdfEHOLogoFileName; }
            set
            {
                if (value != null && !File.Exists(value))
                {
                    System.Console.WriteLine("ERROR: Logo File not found. Unable to create PDFs : " + value);
                    System.Environment.Exit(99);
                }
                else pdfEHOLogoFileName = value;
                Console.WriteLine(" Logo File Name: " + pdfEHOLogoFileName);
            }
        }

        public string OutFileType
        {
            get { return outFileType; }
            set
            {
                if (!value.Equals("PDF") && !value.Equals("TEXT"))
                {
                    outFileType = "TEXT";
                    Console.WriteLine("Invalid File Type Parameter in Config file.  FILETYPE=PDF or TEXT");
                }
                else outFileType = value;
                Console.WriteLine(" Output Type: " + outFileType);
            }
        }

        public float PDFLeftTextMargin => pdfLeftTextMargin;

        public float PDFTopPageOne => pdfTopPageOne;

        public float PDFTopPageTwoPlus => pdfTopPageTwoPlus;

        public float PDFFontPitch => pdfFontPitch;

        public string PDFLogoFileName => pdfRLMSLogoFileName;

        public string PDFEHOFileName => pdfEHOLogoFileName;

        public float PDFLogoXPosition => pdfLogoXPosition;

        public float PDFLogoYPosition => pdfLogoYPosition;

        public float PDFEHOXPosition => pdfEHOXPosition;

        public float PDFEHOYPosition => pdfEHOYPosition;

        public float PDFLogoScale => pdfLogoScale;

    } // end of class Config
}
