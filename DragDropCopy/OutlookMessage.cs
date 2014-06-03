using System;
using System.IO;

namespace DragDropSF
{
    public class OutlookMessage {

        private string targetPath;
        Form1 form1;

        // Construct OutlookMessage with form and targetPath
        public OutlookMessage(Form1 form1, string targetPath)
        {
            this.form1 = form1;
            this.targetPath = targetPath;
        }

        // Get the names and data streams of the files dropped and send to CopyFiles
        public void GetFileFromDrag(OutlookDataObject dataObject)
        {
            string[] filenames = (string[])dataObject.GetData("FileGroupDescriptorW");
            MemoryStream[] filestreams = (MemoryStream[])dataObject.GetData("FileContents");
            CopyFiles(filenames, filestreams);
        }

        // Copy Files to Temp Directory at targetPath
        private void CopyFiles(String[] filenames, MemoryStream[] filestreams)
        {
            string filename;

            for (int fileIndex = 0; fileIndex < filenames.Length; fileIndex++)
            {
                //use the fileindex to get the name and data stream
                filename = filenames[fileIndex];
                MemoryStream filestream = filestreams[fileIndex];

                //save the file stream using its name to the application path
                FileStream outputStream = File.Create(filename);
                filestream.WriteTo(outputStream);
                outputStream.Close();

                //Generate new Files
                //Create new Outlook Message object from filename
                OutlookStorage.Message outlookMsg = new OutlookStorage.Message(filename);
                
                //Save the message as a text file to target path
                toPlainText(outlookMsg, this.targetPath, filename.Substring(0, filename.Length - 4));

                //Save each attachment to target path
                foreach (OutlookStorage.Attachment attach in outlookMsg.Attachments)
                {
                    byte[] attachBytes = attach.Data;
                    FileStream attachStream = File.Create(this.targetPath + attach.Filename);
                    attachStream.Write(attachBytes, 0, attachBytes.Length);
                    attachStream.Close();
                    form1.appendTextBox1("Attachment " + attach.Filename + "\n\r");
                }
                filestream.Close();
            }
        }

        // Convert all .msg files to plain text
        public void toPlainText(OutlookStorage.Message outlookMsg, string path, string filename)
        {

            System.IO.StreamWriter file = new System.IO.StreamWriter(path + filename + ".txt", true);
            //this.textBox1.Text += "Message " + filename + " copied to\n\r\n\r\t" + path + "\n\r\n\r";
            form1.appendTextBox1("Message " + filename + "\n\r");
            file.WriteLine("Subject: {0}", outlookMsg.Subject);
            file.WriteLine("Body: {0}", outlookMsg.BodyText);

            file.WriteLine("{0} Recipients", outlookMsg.Recipients.Count);
            foreach (OutlookStorage.Recipient recip in outlookMsg.Recipients)
            {
                file.WriteLine(" {0}:{1}", recip.Type, recip.Email);
            }

            file.WriteLine("{0} Attachments", outlookMsg.Attachments.Count);
            foreach (OutlookStorage.Attachment attach in outlookMsg.Attachments)
            {
                file.WriteLine(" {0}, {1}b", attach.Filename, attach.Data.Length);
            }

            file.WriteLine("{0} Messages", outlookMsg.Messages.Count);
            foreach (OutlookStorage.Message subMessage in outlookMsg.Messages)
            {
                int msgSub = 1;
                toPlainText(subMessage, path, filename + "_sub_" + msgSub);
                form1.appendTextBox1("Sub-Message of " + filename + "\n\r");
            }
            file.Close();
        }
        
    }
}
