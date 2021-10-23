using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ListaCumparaturi
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Text = "Adauga";
            f3.button1.Text = "Adauga";
            if (DialogResult.OK == f3.ShowDialog())
            {
                listView1.Items.Add((listView1.Items.Count+1).ToString());
                listView1.Items[listView1.Items.Count - 1].SubItems.Add(f3.tbp.Text);
                listView1.Items[listView1.Items.Count - 1].SubItems.Add(f3.tbum.Text);
                listView1.Items[listView1.Items.Count - 1].SubItems.Add(f3.tbc.Text);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.Text = "Modifica";
            f3.button1.Text = "Modifica";
            ListView.SelectedListViewItemCollection cels = listView1.SelectedItems;
            if (cels.Count == 0) MessageBox.Show("Nu ati selectat niciun produs! \nNu se poate efectua modificarea!");
            else
            {
                f3.tbp.Text = cels[0].SubItems[1].Text;
                f3.tbum.Text = cels[0].SubItems[2].Text;
                f3.tbc.Text = cels[0].SubItems[3].Text;
                if(DialogResult.OK==f3.ShowDialog())
                {
                    cels[0].SubItems[1].Text = f3.tbp.Text;
                    cels[0].SubItems[2].Text = f3.tbum.Text;
                    cels[0].SubItems[3].Text = f3.tbc.Text;
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ListView.SelectedListViewItemCollection cels = listView1.SelectedItems;
            if (cels.Count == 0) MessageBox.Show("Nu ati selectat niciun produs! \nNimic de sters!");
            else
            {
                listView1.Items.Remove(cels[0]);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx", ValidateNames = true })
            {
                if (listView1.Items.Count != 0)
                {
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                        Workbook wb = app.Workbooks.Add(XlSheetType.xlWorksheet);
                        Worksheet ws = (Worksheet)app.ActiveSheet;
                        app.Visible = false;
                        ws.Cells[1, 1] = "Nr. Crt.";
                        ws.Cells[1, 2] = "Produs";
                        ws.Cells[1, 3] = "U.M.";
                        ws.Cells[1, 4] = "Cantitate";
                        int i = 2;
                        foreach (ListViewItem item in listView1.Items)
                        {
                            ws.Cells[i, 1] = item.SubItems[0].Text;
                            ws.Cells[i, 2] = item.SubItems[1].Text;
                            ws.Cells[i, 3] = item.SubItems[2].Text;
                            ws.Cells[i, 4] = item.SubItems[3].Text;
                            i++;
                        }
                        wb.SaveAs(sfd.FileName, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                        app.Quit();
                        MessageBox.Show("Lista a fost salvata cu succes.", "Succes", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Lista este goala!", "Eroare", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
