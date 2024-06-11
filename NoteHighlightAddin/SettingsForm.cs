using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NoteHighlightAddin
{
    public partial class SettingsForm : Form
    {

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        public SettingsForm()
        {
            InitializeComponent();
            
            fontDialog1.Font = new Font(NoteHighlightAddin.Properties.Settings.Default.Font, NoteHighlightAddin.Properties.Settings.Default.FontSize);
            btnFont.Text = "Font:" + fontDialog1.Font.Name + ", Size:" + fontDialog1.Font.Size;
            btnFont.Font = fontDialog1.Font;
            cbShowTableBorder.Checked = NoteHighlightAddin.Properties.Settings.Default.ShowTableBorder;
            chkOneRowPerLine.Checked = NoteHighlightAddin.Properties.Settings.Default.OneRowPerLine;
        }

        private void BtnFont_Click(object sender, EventArgs e)
        {
            fontDialog1.Font = new Font(NoteHighlightAddin.Properties.Settings.Default.Font, NoteHighlightAddin.Properties.Settings.Default.FontSize);
            if (fontDialog1.ShowDialog() == DialogResult.OK)
            {
                int roundedSize = (int)Math.Round(fontDialog1.Font.Size);
                // Display rounded (rather than exact) size on the button to avoid confusion.
                btnFont.Text = "Font:"+fontDialog1.Font.Name + ", Size:" + roundedSize.ToString();
                btnFont.Font = fontDialog1.Font;

                NoteHighlightAddin.Properties.Settings.Default.Font = fontDialog1.Font.Name;
                NoteHighlightAddin.Properties.Settings.Default.FontSize = roundedSize;

                NoteHighlightAddin.Properties.Settings.Default.Save();
            }

            
        }

        private void ChShowTableBorder_CheckedChanged(object sender, EventArgs e)
        {
            NoteHighlightAddin.Properties.Settings.Default.ShowTableBorder = cbShowTableBorder.Checked;

            NoteHighlightAddin.Properties.Settings.Default.Save();
        }

        private void SettingsForm_Shown(object sender, EventArgs e)
        {
            // This is necessary in order for SetForegroundWindow to work consistently
            this.WindowState = FormWindowState.Minimized;
            this.WindowState = FormWindowState.Normal;

            SetForegroundWindow(this.Handle);
        }

        private void chkOneRowPerLine_CheckedChanged(object sender, EventArgs e)
        {
            NoteHighlightAddin.Properties.Settings.Default.OneRowPerLine = chkOneRowPerLine.Checked;
            NoteHighlightAddin.Properties.Settings.Default.Save();
        }
    }
}
