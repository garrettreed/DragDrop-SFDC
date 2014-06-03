using DragDropSF.sforce;
using System.IO;
using System.Web.Services.Protocols;

namespace DragDropSF
{
    class SalesForceUpload
    {
        private Form1 form1;
        private string targetPath;
        private SalesForceLogin sflogin;

        public SalesForceUpload(Form1 form1, SalesForceLogin sflogin, string targetPath)
        {
            this.form1 = form1;
            this.targetPath = targetPath;
            this.sflogin = sflogin;
        }

        public void Upload()
        {
            LoginResult lr = sflogin.getLoginResult();
            SforceService binding = new SforceService();
            binding.SessionHeaderValue = new SessionHeader();
            binding.SessionHeaderValue.sessionId = lr.sessionId;
            binding.Url = lr.serverUrl;
            bool errorFlag = false;
            
            try
            {
                DirectoryInfo dir = new DirectoryInfo(this.targetPath);
                FileInfo[] files = dir.GetFiles();

                foreach (FileInfo file in files)
                {
                    if (form1.CheckState() == false)
                    {
                        if (Path.GetExtension(Path.GetFullPath(file.FullName)) == ".txt")
                        {
                            file.Delete();
                        }
                    }
                    else
                    {
                            Attachment oppattach = new Attachment();
                            string filename = file.FullName;
                            byte[] bytes = System.IO.File.ReadAllBytes(filename);
                            oppattach.Body = bytes;
                            oppattach.Name = file.Name;
                            oppattach.ParentId = this.form1.getOppID();
                            sObject[] attachments = new Attachment[1];
                            attachments[0] = oppattach;
                            SaveResult[] result = binding.create(attachments);

                            for (int i = 0; i < attachments.Length; i++)
                            {
                                if (result[i].success)
                                {
                                    form1.appendTextBox5("A Document was created");
                                }
                                else
                                {
                                    form1.appendTextBox5("Item had an error updating");
                                    errorFlag = true;
                                }
                            }
                        }
                    }
                if (!(errorFlag))
                {
                    form1.clearDirectory();
                    form1.updateTextBox1("Upload Successful. Files removed from " + this.targetPath);
                }
            }
            catch (SoapException e)
            {
                form1.appendTextBox1("Upload failed");
                form1.appendTextBox1(e.Detail.ToString());
            }
        }
    }
}