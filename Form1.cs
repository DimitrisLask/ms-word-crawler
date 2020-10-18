using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WindowsFormsApplication1
{

public partial class Form1 : Form
{
    private BindingList<WDocs> _docList;
    public BindingList<WDocs> DocList
    {
        get
        {
            return _docList;
        }
        set
        {
            if (_docList == value) return;

            _docList = value;
        }
    }
    private string _log;
    public string Log
    {
        get
        {
            return _log;
        }
        set
        {
            if (_log == value) return;

            _log = value;

            }
        }
    private int _countLinks;
    public int CountLinks
    {
        get { return _countLinks; }
        set
        {
            if (_countLinks == value) return;

            _countLinks = value;
            }
   
    }
    private WDocs _selectedDoc;
    public WDocs SelectedDoc
    {
        get { return _selectedDoc; }
        set
        {
            if (_selectedDoc == value) return;

            _selectedDoc = value;
        }
    }
    Word.Application app;
    Word.Documents docs;
    Word.Document currentDoc;
    Word.Table table;
    Word.Column col;
    Word.Range range;
    Stopwatch watch;
    Stopwatch total;
    string filename, path;
        
    public Form1()
    {
        InitializeComponent();
            
        _docList = new BindingList<WDocs>();
        dataGridView1.DataSource = DocList;
        _selectedDoc = null;
        button2.Enabled=button4.Enabled = false;
        _docList.ListChanged += new ListChangedEventHandler(docList_ListChanged);
        Bitmap bmp = new Bitmap(Properties.Resources.delos);
        pictureBox1.Image = bmp;

        }

    private void InsertLinks(object sender, EventArgs e)
    {
            textBox1.Clear();
            app = new Word.Application();
            docs = app.Documents;
            app.Visible = false;
            Cursor = Cursors.WaitCursor;

            total = Stopwatch.StartNew();
            foreach (WDocs doc in _docList)
            {
                docs.Add(doc.FilePath);
                currentDoc = docs.Open(doc.FilePath);
                textBox1.AppendText("Document: " + currentDoc.Name);

                if (doc.IsChecked)
                {
                    watch = Stopwatch.StartNew();
                    SearchWithTables();
                    watch.Stop();
                }
                else
                {
                    watch = Stopwatch.StartNew();
                    foreach (Word.Range r in currentDoc.Words)
                    {
                        if (r.Text.TrimEnd().Length == 9)
                        {
                            int n;
                            if (int.TryParse(r.Text, out n))
                            {
                                r.Hyperlinks.Add(r, "http://pbsales" + "//" + r.Text);
                                r.Font.Bold = 0;
                                r.Underline = Word.WdUnderline.wdUnderlineNone;
                                r.Font.Color = Word.WdColor.wdColorBlack;
                                _countLinks++;

                            }
                        }
                    }

                    watch.Stop();
                }
                
                _log ="    Links: " + _countLinks + "      Time :" + watch.Elapsed + "\n";
                textBox1.AppendText(_log);
                textBox1.AppendText("*********************************************************************************************\n");
                currentDoc.Close(true);
                _countLinks = 0;
                watch.Reset();

            }

            total.Stop();
            Cursor = Cursors.Arrow;
            textBox1.AppendText("\n\n");
            textBox1.AppendText("Total time ellapsed : "+total.Elapsed);

            docs.Close();
            app.Quit();

        }

    private void Add(object sender, EventArgs e)
    {
        OpenFileDialog dialog = new OpenFileDialog();
        dialog.InitialDirectory = "Desktop";
        dialog.Filter = "Word files (*.doc)|*.doc|Word files (*.docx)|*.docx";
        dialog.FilterIndex = 2;
        dialog.RestoreDirectory = true;
        dialog.Multiselect = true;

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                
                for (int i=0; i<dialog.SafeFileNames.Length; i++)
                {
                    filename = dialog.SafeFileNames[i];
                    path = dialog.FileNames[i];
                    if (!_docList.Any(x => x.FileName == filename))
                    {
                        WDocs doc = new WDocs(filename, path);
                        _docList.Add(doc);
                        continue;
                    }

                    var result = MessageBox.Show("File already exists!", "Info",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                }

            }

        dataGridView1.ClearSelection();
        _selectedDoc = null;

    }

    private void ClearList(object sender, EventArgs e)
    {
        var result = MessageBox.Show("Clear entire list?", "Clear",
                            MessageBoxButtons.OKCancel,
                            MessageBoxIcon.Question);

        if (result != DialogResult.OK) return;

        _selectedDoc = null;
        _docList.Clear();
        textBox1.Clear();
    }

    private void Remove(object sender, EventArgs e)
    {
            if (_selectedDoc != null)
            {
                var result = MessageBox.Show("Remove selected item?", "Remove",
                                    MessageBoxButtons.OKCancel,
                                    MessageBoxIcon.Question);

                if (result != DialogResult.OK) return;

                _docList.Remove(_selectedDoc);
                dataGridView1.ClearSelection();
                _selectedDoc = null;
            }
            else
            {
                var result = MessageBox.Show("No item is selected!", "Error",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error);
            }
    }

    private void docList_ListChanged(object sender, ListChangedEventArgs e)
    {
        button2.Enabled=button4.Enabled = (_docList.Count != 0);
    }

    private void selectedItemChanged(object sender, EventArgs e)
    {
        if (_docList.Count == 0) return;

        _selectedDoc = (WDocs)dataGridView1.CurrentRow.DataBoundItem;

    }

    private void SearchWithTables()
        {
            foreach (Word.Shape shape in currentDoc.Shapes)
            {
                if (shape.Type == MsoShapeType.msoTextBox)
                {
                    try
                    {
                        table = shape.TextFrame.TextRange.Tables[1];
                        col = table.Columns[1];
                        foreach (Word.Cell cell in col.Cells)
                        {
                            range = cell.Range;
                            range.Hyperlinks.Add(range, "http://pbsales" + "//" + range.Text);
                            range.Font.Bold = 0;
                            range.Underline = Word.WdUnderline.wdUnderlineNone;
                            range.Font.Color = Word.WdColor.wdColorBlack;
                            _countLinks++;
                        }
                    }
                    catch (Exception e)
                    {
                        
                    }
                }
            }

            foreach (Word.Range r in currentDoc.Words)
            {
                if (r.Text.TrimEnd().Length == 9)
                {
                    int n;
                    if (int.TryParse(r.Text, out n))
                    {
                        r.Hyperlinks.Add(r, "http://pbsales" + "//" + r.Text);
                        r.Font.Bold = 0;
                        r.Underline = Word.WdUnderline.wdUnderlineNone;
                        r.Font.Color = Word.WdColor.wdColorBlack;
                        _countLinks++;

                    }
                }
            }


        }

    }
}
