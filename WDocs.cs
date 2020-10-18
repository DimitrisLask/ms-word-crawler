using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication1
{
    public class WDocs
    {

        private string _fileName;
        public string FileName
        {
            get
            {
                return _fileName;
            }

            set
            {
                if (_fileName == value) return;
                _fileName = value;
            }
        }
        private string _filePath;
        public string FilePath
        {
            get { return _filePath; }
            set
            {
                if (_filePath == value) return;

                _filePath = value;
            }
        }
        private int _count;
        public int Count
        {
            get
            {
                return _count;
            }
            set
            {
                if (_count == value) return;
                _count = value;
            }
        }
        private bool _isChecked;
        public bool IsChecked
        {
            get { return _isChecked; }
            set
            {
                if (_isChecked == value) return;

                _isChecked = value;
            }
        }

        public WDocs(string fileName, string filepath, int count, bool ischecked)
        {
            _fileName = fileName;
            _filePath = filepath;
            _count = count;
            _isChecked = ischecked;
        }

        public WDocs(string filename, string filepath)
        {
            _fileName = filename;
            _filePath = filepath;
            _count = 0;
            _isChecked = false;
        }
    }
}
