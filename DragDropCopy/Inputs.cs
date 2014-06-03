using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DragDropSF
{
    class Inputs
    {
        private string targetPath = (Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\DragDropDirectory\\");

        public string getTargetPath()
        {
            return this.targetPath;
        }
    }
}
