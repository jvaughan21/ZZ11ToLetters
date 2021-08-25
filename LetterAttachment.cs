
using iTextSharp.text.pdf;

namespace ZZ11ToLetter
{
    class LetterAttachment
    {
        private string letterID;
        private string attachmentFileName;
        private PdfReader attachmentReader;

        public LetterAttachment(string inletterID, string inattachmentFilename)
        {
            LetterID = inletterID;
            AttachmentFilename = inattachmentFilename;
            attachmentReader = new PdfReader(inattachmentFilename);
        }

        public string LetterID
        {
            get { return letterID; }
            set { letterID = value; }
        }

        public string AttachmentFilename
        {
            get { return attachmentFileName; }
            set { attachmentFileName = value; }
        }

        public PdfReader PDFReader
        {
            get { return attachmentReader; }
        }

    } // class Letter Attachment
}
