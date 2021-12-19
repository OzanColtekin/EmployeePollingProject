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
using Excel = Microsoft.Office.Interop.Excel;

namespace EgeYapı
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            isciList.ColumnCount = 1;
            isciList.Columns[0].Name = "Adı Soyadı";
            isciList.Columns[0].Width = 450;
            ListMembers();
        }

        public void ListMembers()
        {
            string path = Application.StartupPath;
            string[] kisiler = System.IO.File.ReadAllLines(path + @"\isimler.txt");
            foreach (string kisi in kisiler)
            {
                isciList.Rows.Add(kisi);
            }

        }
        public void ListMembersWithName(string text)
        {
            int x = 1;
            string path = Application.StartupPath;
            string[] kisiler = System.IO.File.ReadAllLines(path + @"\isimler.txt");
            foreach (string kisi in kisiler)
            {
                string lower = kisi.ToLower();
                if (lower.StartsWith(text))
                {
                    isciList.Rows.Add(kisi);
                    x++;
                }
                
            }
        }

        private void isciList_DoubleClick(object sender, EventArgs e)
        {

            string seciliKisi = isciList.CurrentRow.Cells[0].Value.ToString();
            int _indexOf = Gelenler.Items.IndexOf(seciliKisi);
            if(_indexOf == -1)
            {
                Gelenler.Items.Add(seciliKisi);
            }
            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(textBox1.Text))
            {
                isciList.Rows.Clear();
                ListMembers();
            }
            else
            {
                string kisi = textBox1.Text.ToLower();
                isciList.Rows.Clear();
                ListMembersWithName(kisi);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Gelenler.Items.Clear();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if(Gelenler.SelectedItem == null)
            {
                MessageBox.Show("Silmek istediğiniz isimi yan taraftan seçin.");
            }
            else
            {
                Gelenler.Items.Remove(Gelenler.SelectedItem);
            }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Gelenler.Items.Count == 0) { 
                MessageBox.Show("Kaydedilecek isimler seçili değil."); 
            }
            else
            {
            
            string _Date = DateTime.Now.ToString("d-M-yyyy");
            string path = Application.StartupPath + @"\Puantaj Kayıtları\" + _Date + " Gelenler.txt";
            if (!File.Exists(path))
            {
                using (StreamWriter sw = File.CreateText(path));
            }
            using (StreamWriter sw = File.AppendText(path))
            {
                
                for(int i =0; i<Gelenler.Items.Count; i++)
                {
                    sw.WriteLine(Gelenler.Items[i]);
                }
            }
            }
            MessageBox.Show("Metin Belgesi dosyası kaydedildi.");
        }

        private void Sayfa_Ayarla(Excel.Worksheet Sayfa1, string tam_tarih)
        {

            Excel.Range A1 = (Excel.Range)Sayfa1.Cells[1, 1];
            A1.Value = tam_tarih + @" MARMARAY CR3 KAPSAMINDA YAPIMI DEVAM EDEN HAYDAPAŞA GAR SAHASI ARKELOJİ KAZI ÇALIŞMASINDA ÇALIŞAN PERSONEL LİSTESİ";
            A1.Interior.Color = Excel.XlRgbColor.rgbYellow;
            A1.Font.Size = 18;
            A1.Font.Bold = true;
            ((Excel.Range)Sayfa1.Range[(Excel.Range)Sayfa1.Cells[1, 1], (Excel.Range)Sayfa1.Cells[1, 6]]).Cells.Merge();
            A1.WrapText = true;
            A1.ColumnWidth = 8.11;
            A1.RowHeight = 51;
            A1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            A1.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
            A1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;


            Excel.Range A2 = (Excel.Range)Sayfa1.Cells[2, 1];
            A2.Interior.Color = Excel.XlRgbColor.rgbYellow;

            Excel.Range B2 = (Excel.Range)Sayfa1.Cells[2, 2];
            B2.Value = "ADI SOYADI";
            B2.Interior.Color = Excel.XlRgbColor.rgbYellow;
            B2.Font.Size = 12;
            B2.ColumnWidth = 24.56;
            B2.Font.Bold = true;
            B2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            Excel.Range C2 = (Excel.Range)Sayfa1.Cells[2, 3];
            C2.Value = "SIRA NO";
            C2.Interior.Color = Excel.XlRgbColor.rgbYellow;
            C2.Font.Size = 12;
            C2.ColumnWidth = 8.11;
            C2.Font.Bold = true;
            C2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            Excel.Range D2 = (Excel.Range)Sayfa1.Cells[2, 4];
            D2.Value = "ADI SOYADI";
            D2.Interior.Color = Excel.XlRgbColor.rgbYellow;
            D2.Font.Size = 12;
            D2.ColumnWidth = 11.56;
            D2.Font.Bold = true;
            D2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            Excel.Range E2 = (Excel.Range)Sayfa1.Cells[2, 5];
            E2.Value = "SIRA NO";
            E2.Interior.Color = Excel.XlRgbColor.rgbYellow;
            E2.Font.Size = 12;
            E2.ColumnWidth = 8.11;
            E2.Font.Bold = true;
            E2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            Excel.Range F2 = (Excel.Range)Sayfa1.Cells[2, 6];
            F2.Value = "ADI SOYADI";
            F2.Interior.Color = Excel.XlRgbColor.rgbYellow;
            F2.Font.Size = 12;
            F2.ColumnWidth = 47.67;
            F2.Font.Bold = true;
            F2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            Excel.Range F3 = (Excel.Range)Sayfa1.Cells[3, 6];
            F3.Value = "İZİNLİ OLAN PERSONEL LİSTESİ";
            F3.Interior.Color = Excel.XlRgbColor.rgbYellow;
            F3.Font.Size = 12;
            F3.Font.Bold = true;
            F3.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            Excel.Range F5 = (Excel.Range)Sayfa1.Cells[5, 6];
            F5.Value = "İŞE GELMEYEN PERSONEL LİSTESİ";
            F5.Interior.Color = Excel.XlRgbColor.rgbYellow;
            F5.Font.Size = 12;
            F5.Font.Bold = true;
            F5.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

        }

        private void button4_Click(object sender, EventArgs e)
        {
            string tam_tarih = DateTime.Now.ToString("d");
            string[] gün_ay = DateTime.Now.ToString("d M").Split(' ');
            string ay   = DateTime.Now.ToString("MMMM").ToUpper();
            string yıl  = DateTime.Now.ToString("yyyy");
            string file_path = Application.StartupPath + @"\Puantaj Kayıtları\" + ay + " AYI PUANTAJ ( " + yıl + " ).xlsx";
            Excel.Application uygulama = new Excel.Application();
            uygulama.Visible = true;

            int a = 1;
            int b = 3;

            if (!File.Exists(file_path))
            {

                Excel.Workbook kitap = uygulama.Workbooks.Add(System.Reflection.Missing.Value);
                Excel.Worksheet Sayfa1 = (Excel.Worksheet)kitap.Sheets[1];
                Sayfa1.Name = "1";

                Sayfa_Ayarla(Sayfa1, tam_tarih);

                

                for (int i = 0; i <Gelenler.Items.Count; i++)
                {
                    bool failed = false;
                    do
                    {
                        try
                        {
                            Excel.Range alan = (Excel.Range)Sayfa1.Cells[b, 2];
                            Excel.Range alan2 = (Excel.Range)Sayfa1.Cells[b, 1];
                            Excel.Range C3 = (Excel.Range)Sayfa1.Cells[b, 3];
                            Excel.Range D3 = (Excel.Range)Sayfa1.Cells[b, 4];
                            Excel.Range E3 = (Excel.Range)Sayfa1.Cells[b, 5];
                            Excel.Range F3 = (Excel.Range)Sayfa1.Cells[b, 6];
                            C3.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            D3.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            E3.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            F3.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            alan.Value = Gelenler.Items[i].ToString().ToUpper();
                            alan.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            alan2.Value = a;
                            alan2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            alan2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            b++;
                            a++;
                            failed = false;
                        }
                        catch (System.Runtime.InteropServices.COMException err)
                        {
                            failed = true;
                        }
                        System.Threading.Thread.Sleep(10);
                    } while (failed);
                    
                }
            }
            else
            {
                Excel.Workbook kitap = uygulama.Workbooks.Open(file_path);
                Excel.Worksheet Sayfa1 = (Excel.Worksheet)kitap.Sheets.Add();
                Sayfa1.Name = gün_ay[0];

                Sayfa_Ayarla(Sayfa1, tam_tarih);

                for (int i = 0; i < Gelenler.Items.Count; i++)
                {
                    bool failed = false;
                    do
                    {
                        try
                        {
                            Excel.Range alan = (Excel.Range)Sayfa1.Cells[b, 2];
                            Excel.Range alan2 = (Excel.Range)Sayfa1.Cells[b, 1];
                            Excel.Range C3 = (Excel.Range)Sayfa1.Cells[b, 3];
                            Excel.Range D3 = (Excel.Range)Sayfa1.Cells[b, 4];
                            Excel.Range E3 = (Excel.Range)Sayfa1.Cells[b, 5];
                            Excel.Range F3 = (Excel.Range)Sayfa1.Cells[b, 6];
                            C3.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            D3.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            E3.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            F3.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            alan.Value = Gelenler.Items[i].ToString().ToUpper();
                            alan.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            alan2.Value = a;
                            alan2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            alan2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            b++;
                            a++;
                            failed = false;
                        }
                        catch (System.Runtime.InteropServices.COMException err)
                        {
                            failed = true;
                        }
                        System.Threading.Thread.Sleep(10);
                    } while (failed);
                    
                }

            }
            

            
        }

        private void isciList_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar ==(char) Keys.Enter)
            {
                int i = isciList.CurrentRow.Index;
                if(i != 0)
                {
                    string seciliKisi = isciList.Rows[i - 1].Cells[0].Value.ToString();
                    int _indexOf = Gelenler.Items.IndexOf(seciliKisi);
                    if (_indexOf == -1)
                    {
                        Gelenler.Items.Add(seciliKisi);
                    }
                }
                else
                {
                    string seciliKisi = isciList.CurrentRow.Cells[0].Value.ToString();
                    int _indexOf = Gelenler.Items.IndexOf(seciliKisi);
                    if (_indexOf == -1)
                    {
                        Gelenler.Items.Add(seciliKisi);
                    }
                }

            }
        }
    }
}
