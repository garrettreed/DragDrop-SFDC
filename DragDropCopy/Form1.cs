using System;
using System.IO;
using System.Windows.Forms;

namespace DragDropSF
{
    public partial class Form1 : Form
    {
        private string targetPath;
        private Inputs inputs = new Inputs();
        private SalesForceLogin salesforce;
        private OutlookMessage outlookmessage;
        private SalesForceUpload sfUpload;

        public Form1()
        {
            InitializeComponent();
            this.targetPath = inputs.getTargetPath();
            this.outlookmessage = new OutlookMessage(this, this.targetPath); 
            this.salesforce = new SalesForceLogin(this, this.targetPath);
            this.FormClosing += Form1_FormClosing;
            this.sfUpload = new SalesForceUpload(this, this.salesforce, this.targetPath);
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            //wrap standard IDataObject in OutlookDataObject
            OutlookDataObject dataObject = new OutlookDataObject(e.Data);
            outlookmessage.GetFileFromDrag(dataObject);
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            //ensure FileGroupDescriptor is present before allowing drop
            if (e.Data.GetDataPresent("FileGroupDescriptor"))
            {
                e.Effect = DragDropEffects.All;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Check to see if Target Path directory exists. If not, create it.
            DirectoryInfo dir = new DirectoryInfo(inputs.getTargetPath());
            if (!dir.Exists)
            {
                dir.Create();
                MessageBox.Show("A new Directory has been created to store outlook files at" + this.targetPath);
            }
            else
            {
                clearDirectory();
            }
        }

        private void Form1_FormClosing(Object sender, FormClosingEventArgs e)
        {
            DirectoryInfo dir = new DirectoryInfo(this.targetPath);
            if (dir.GetFileSystemInfos().Length != 0)
            {
                MessageBox.Show("Exit and Delete Temp Outlook Files?");

                FileInfo[] files = dir.GetFiles();
                foreach (FileInfo file in files)
                {
                    file.Delete();
                }
            }
        }

        public void clearDirectory()
        {
            DirectoryInfo dir = new DirectoryInfo(this.targetPath);
            if (dir.GetFileSystemInfos().Length != 0)
            {
                FileInfo[] files = dir.GetFiles();
                foreach (FileInfo file in files)
                {
                    file.Delete();
                }
            }
        }

        public string getUsername()
        {
            return this.textBox2.Text;
        }

        public string getPassword()
        {
            return this.textBox3.Text;
        }

        public string getOppID()
        {
            return this.textBox6.Text;
        }

        public bool CheckState()
        {
            if (checkBox1.Checked)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public void appendTextBox1(string text)
        {
            this.textBox1.AppendText(Environment.NewLine + text);
        }

        public void updateTextBox1(string text)
        {
            this.textBox1.Clear();
            this.textBox1.AppendText(Environment.NewLine + text);
        }

        public void updateTextBox4(string text)
        {
            this.textBox4.Clear();
            this.textBox4.AppendText(text);
        }

        public void updateTextBox5(string text)
        {
            this.textBox5.Clear();
            this.textBox5.AppendText(text);
        }

        public void appendTextBox5(string text)
        {
            this.textBox5.AppendText(Environment.NewLine + text);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            salesforce.run();
            if (salesforce.getLoginFlag() == true)
            {
                this.textBox2.ReadOnly = true;
                this.textBox3.Clear();
                this.textBox3.ReadOnly = true;
                this.button2.Visible = true;
                this.button1.Visible = false;
                this.button3.Enabled = true;
            }
            else if (salesforce.getLoginFlag() == false)
            {
                this.textBox2.ReadOnly = false;
                this.textBox3.Clear();
                this.textBox3.ReadOnly = false;
                this.button3.Enabled = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (salesforce.getLoginFlag() == true)
            {
                salesforce.logout();
                this.textBox2.Update();
                this.textBox2.ReadOnly = false;
                this.textBox3.Update();
                this.textBox3.Clear();
                this.textBox3.ReadOnly = false;
                this.button2.Visible = false;
                this.button1.Visible = true;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.sfUpload = new SalesForceUpload(this, this.salesforce, this.targetPath);
            this.sfUpload.Upload();
        }
        
        private void label1_Click(object sender, EventArgs e) {}

        private void textBox1_TextChanged(object sender, EventArgs e){}

        private void label1_Click_1(object sender, EventArgs e){}

        private void label3_Click(object sender, EventArgs e){}

        private void textBox2_TextChanged(object sender, EventArgs e){}

        private void textBox3_TextChanged(object sender, EventArgs e){}

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            
        }
    }
}