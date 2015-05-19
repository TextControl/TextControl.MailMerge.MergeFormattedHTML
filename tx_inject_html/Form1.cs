using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TXTextControl;
using TXTextControl.DocumentServer.Fields;

namespace tx_inject_html
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void mergeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // create a new DataSet and load the XML file
            DataSet ds = new DataSet();
            ds.ReadXml("data.xml");

            // start the merge process
            mailMerge1.Merge(ds.Tables[0]);
        }

        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // load the template
            TXTextControl.LoadSettings ls = new TXTextControl.LoadSettings();
            ls.ApplicationFieldFormat = TXTextControl.ApplicationFieldFormat.MSWord;

            textControl1.Load(TXTextControl.StreamType.WordprocessingML, ls);
        }

        private void mailMerge1_FieldMerged(object sender, TXTextControl.DocumentServer.MailMerge.FieldMergedEventArgs e)
        {
            // check whether the field is a MergeField
            if(e.MailMergeFieldAdapter.TypeName != "MERGEFIELD")
                return;

            MergeField field = e.MailMergeFieldAdapter as MergeField;

            // if switch "Preserve formatting during updates" is set
            // load HTML data into temporary ServerTextControl and return
            // formatted text
            if (field.PreserveFormatting == true)
            {
                byte[] data;

                using (ServerTextControl tx = new ServerTextControl())
                {
                    tx.Create();

                    tx.Load(field.Text, StringStreamType.HTMLFormat);
                    tx.Save(out data, BinaryStreamType.InternalUnicodeFormat);
                }

                e.MergedField = data;
            }
        }

    }
}
