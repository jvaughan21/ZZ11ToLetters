using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ZZ11ToLetter
{
    public class ZZ11ToLetter
    {
        public const string ZZ11FILETYPE = "ZZ11";
        public const string Z0192FILETYPE = "0192";

        static Config c = new Config();
        static int letterCount = 0;

        static int Main(string[] args)
        {
            string inFile = "";
            string logDetails = "";

            Console.WriteLine("ZZ11 To Letter form");
            Console.WriteLine("RLMS");
			Console.WriteLIne("This utility contains and uses iTextSharp. Copyright by iText Group NV. Affero General Public License.\n");

            if (args.Length == 1)
            {
                if (args[0].Substring(0, 6).ToUpper().Equals("CONFIG"))
                {
                    c.ConfigFilename = args[0].Substring(args[0].IndexOf("=") + 1);
                    c.ReadConfiguration();
                }
                else
                {
                    Console.WriteLine("\n USAGE: ZZ11toLetter CONFIG=(config filename)");
                    Console.WriteLine("    or: ZZ11toLetter (filename) (output directory)");
                    Environment.Exit(99);
                }
            }
            else if (args.Length != 2)
            {
                Console.WriteLine("\n USAGE: ZZ11toLetter (filename) (output directory)");
                Console.WriteLine("\n        This utility reads an MSP ZZ11 or 0192 file.");
                Console.WriteLine("          The utility formats the file and seperates each letter into");
                Console.WriteLine("          individual files that can then be used.");
                Environment.Exit(99);
            }
            else if (args.Length == 2 && ArgsAreValid(args))
            {
                inFile = args[0];
                c.OutputDirectory = args[1];
            }
            else
            {
                Environment.Exit(99);
            }

            // Register Text Encoder because we need to manage to an ANSI Code Page 1252
            // Using the Encoding library downloaded through Nugget
            //
            Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            
            // loop to read file and write letters
            String[] fileList = Directory.GetFiles( c.InputDirectory );

            foreach (string f in fileList)
            {
                Console.WriteLine(" " + f);
                try
                {
                    if (ParseFile(f))
                    {
                        Console.Write(".. Moving letters to output directory");
                        logDetails = "DateTime|filename|message\n";
                        foreach (string ltr in Directory.GetFiles( c.TempDirectory))
                        {

                            if (!ltr.Trim().ToUpper().Substring(0, 6).Equals("HEADER"))
                            {
                                File.Move(ltr, c.OutputDirectory + "\\" + Path.GetFileName(ltr));
                                logDetails += DateTime.Now + "|" + Path.GetFileName(ltr) + "|created\n";
                            }
                            else File.Delete(ltr);
                        }
                        File.Move( f, c.ArchiveDirectory + "\\" + Path.GetFileName(f));
                        Console.WriteLine("");
                        logDetails += DateTime.Now + "|" + Path.GetFileName(f) + "|" + (letterCount-1) + "\n";
                        File.AppendAllText(c.LogFileDirectory + "\\" + c.LogFileName, logDetails);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("\n Exception: " + e.Message);
                    logDetails += DateTime.Now + "|" + Path.GetFileName(f) + "|" + e.Message + "\n";
                    File.AppendAllText(c.LogFileDirectory + "\\" + c.LogFileName, logDetails);
                }
            }

            return 0;
        } // end of main

        // Parsefile
        // reads all of the bytes from the ZZ11 or the 0192 or (an invalid file
        // the utility ignores the print alignment "letters" - as if this is dot matrix.... welcome to the 1980s....
        // the utility creats individual letter files by splitting the ZZ11/0192
        // letter files placed into the tempLetter directory
        //
        public static bool ParseFile(string inFile)
        {
            int counter = 0;
            LetterDetails LD = null;
            List<string> letterLines = new List<string>();
            IEnumerable<string> s;
            string secs;
            string text = "";
            bool ignoreLine = true;
            
            try
            {
                
                byte[] bts = File.ReadAllBytes(inFile);
                for (int i=0; i<bts.Length; i++)
                {
                    if (bts[i] > 127)
                        bts[i] = c.FixSpanishCharacter(bts[i]);
                }
                File.WriteAllBytes(c.TempDirectory + "\\" + Path.GetFileName(inFile), bts);
                text = File.ReadAllText(c.TempDirectory + "\\" + Path.GetFileName(inFile), Encoding.GetEncoding(1252));
                File.Delete(c.TempDirectory + "\\" + Path.GetFileName(inFile));
            }
            catch (Exception e)
            {
                Console.WriteLine("\n\nERROR: " + e.Message);
                Console.WriteLine("An error occurred reading the data file: " + inFile);
                return false;
            }

            letterCount = 0;

            if ((text.Length >= 135) && (text.Substring(0, 135).IndexOf('\n') < 0)) //** if true, then this is a ZZ11 file. 
                s = text.Split(133);
            else                                        //** else assume This is a 0192 file.
                s = text.Split('\n');

            counter = 0;

            foreach (string line in s)                       // loop traverses each line found in file. Removes unwanted lines.
                if (line.Length > 8 &&
                       (line.Substring(0, 8).Equals("EDI_DATA") ||  //** This logic ignores header in the ZZ11 and 0192
                        line.Substring(0, 4).Equals("1MM ") ||
                        line.Substring(0, 4).Equals(" MM ") ||
                        line.Substring(0, 5).Equals(" MMH ") ||
                        line.Substring(0, 8).Equals("1102XXXX")))
                    ignoreLine = true;
                else if (line.Length >=4 && ignoreLine && !line.Substring(0, 4).Equals("1102")) //** Continued logic to skip print alignment header (letter)
                    ignoreLine = true;
                else if (line.Length >= 4 && ignoreLine && line.Substring(0, 4).Equals("1102")) //** if true, then we have found a printable letter after the 
                {                                                           //** print alignment header (stop skipping the lines)
                    ignoreLine = false;
                    letterLines.Add(line);
                }
                else letterLines.Add(line);   // Keep this line....ignore all others

            LD = new LetterDetails();

            foreach (string line in letterLines)
            {
                if ((line.Length > 4) && line.Substring(0, 4).Equals("1102"))  //** This logic finds the letter header produced by OLLW
                {
                    if (LD == null)
                        LD = new LetterDetails(line.TrimEnd());
                    else if ((line.Length>20) && !line.Substring(19, 2).Equals("01"))        // is this page 2+ of a letter - meaning pagenumber <> 1
                        LD.AddContentLine( c, ((char)12).ToString());   // put a FF to tell Brooksnet to pagebreak
                    else
                    {
                        secs = DateTime.Now.ToString("ss");
                        secs = (letterCount < 100000) ? secs += letterCount.ToString("00000") : secs += letterCount.ToString("00000");
                        LD.WriteLetter( c, secs);
                        letterCount++;
                        LD = new LetterDetails(line.TrimEnd());
                    }
                }
                else
                {
                    if ((line.Trim().Length >= 5) && line.Trim().Substring(0, 5).Equals(LD.LetterID))
                    {
                        string[] lineTokens = line.Trim().Split(" ");
                        if (lineTokens.Length >= 3)
                            LD.LetterVersion = lineTokens[1];
                        else LD.LetterVersion = "000"; // version not found....
                    }

                    if ((LD != null) && (!line.Equals("")))
                        LD.AddContentLine( c, line.TrimEnd() + "\n");
                }

                if (++counter == 1000)
                {
                    Console.Write(".");
                    counter = 0;
                }
            } // foreach (...

            if (LD != null)
            {
                secs = DateTime.Now.ToString("ss");
                letterCount++;
                secs = (letterCount < 100000) ? secs += letterCount.ToString("00000") : secs += letterCount.ToString("00000");
                LD.WriteLetter( c, secs);
            }

            Console.Write("  DONE.\n  Letters Written: " + (letterCount-1));
            return true;

        } // End fo ParseFile


        public static bool ArgsAreValid(string[] args)
        {
            if (!File.Exists(args[0]))
            {
                Console.WriteLine("ERROR: File not found: " + args[0]);
                return false;
            }

            if (!Directory.Exists(args[1]))
            {
                Console.WriteLine("ERROR: Path not found: " + args[1]);
                return false;
            }

            return true;
        } // ArgsAreValid ...

    }  // class ZZ11toLetter

} // namespace


