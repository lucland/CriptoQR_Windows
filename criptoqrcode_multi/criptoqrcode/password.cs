using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace criptoqrcode
{
    public partial class password : Form
    {
        Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
        public password()
        {
            InitializeComponent();
        }

        private void password_Load(object sender, EventArgs e)
        {
            ler_linha_projeto();
        }


        private void atualizar_file()
        {
            try
            {

                File.WriteAllText(@"C:\compartilhamento\data_txt\PROJETO.txt", textBox2.Text.Trim() + "\r\n" + textBox3.Text.Trim());
                // File.WriteAllText(@"C:\compartilhamento\data_txt\count.txt", "0");
                File.WriteAllText(@"C:\compartilhamento\data_txt\data.txt", String.Empty);
                // File.WriteAllText(@"C:\data_txt\data2.txt", String.Empty);

                // var file2 = new DirectoryInfo(@"C:\data_base\").GetFiles().OrderBy(o => o.CreationTime).LastOrDefault();
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = app.Workbooks.Add(misValue);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                // xlWorkSheet = pasta.Worksheets["Planilha1"];
                /*
                xlWorkSheet.Cells[1, 1] = "NUMBER";
                xlWorkSheet.Cells[1, 2] = "NAME";
                xlWorkSheet.Cells[1, 3] = "COMPANY";
                xlWorkSheet.Cells[1, 4] = "FUNCTION";
                xlWorkSheet.Cells[1, 5] = "ID";
                xlWorkSheet.Cells[1, 6] = "EMAIL";
                xlWorkSheet.Cells[1, 7] = "VESSEL";
                xlWorkSheet.Cells[1, 8] = "CHECK-IN VALIDATION";
                xlWorkSheet.Cells[1, 9] = "CHECK-OUT VALIDATION";
                xlWorkSheet.Cells[1, 10] = "CHECK-IN";
                xlWorkSheet.Cells[1, 11] = "CHECK-OUT";

                xlWorkSheet.Cells[1, 12] = "PROJECT";
                xlWorkSheet.Cells[1, 13] = "ASO";
                xlWorkSheet.Cells[1, 14] = "NR-35";
                xlWorkSheet.Cells[1, 15] = "VACCINE-1";
                xlWorkSheet.Cells[1, 16] = "VACCINE-2";
                xlWorkSheet.Cells[1, 17] = "BOOST VACCINE";
                xlWorkSheet.Cells[1, 18] = "LOCAL";
                xlWorkSheet.Cells[1, 19] = "LEVEL";
                xlWorkSheet.Cells[1, 20] = "ESTADO";
                xlWorkSheet.Cells[1, 21] = "MOTIVO";
                xlWorkSheet.Cells[1, 22] = "USUARIO";
                */
                xlWorkSheet.Cells[1, 1] = "NUMBER";
                xlWorkSheet.Cells[1, 2] = "NAME";
                xlWorkSheet.Cells[1, 3] = "COMPANY";
                xlWorkSheet.Cells[1, 4] = "FUNCTION";
                xlWorkSheet.Cells[1, 5] = "ID";
                xlWorkSheet.Cells[1, 6] = "EMAIL";
                xlWorkSheet.Cells[1, 7] = "VESSEL";
                xlWorkSheet.Cells[1, 8] = "CHECK-IN VALIDATION";
                xlWorkSheet.Cells[1, 9] = "CHECK-OUT VALIDATION";

                xlWorkSheet.Cells[1, 10] = "CHECK-IN  DATA";
                xlWorkSheet.Cells[1, 11] = "CHECK-IN  HORA";

                xlWorkSheet.Cells[1, 12] = "CHECK-OUT DATA";
                xlWorkSheet.Cells[1, 13] = "CHECK-OUT HORA";

                xlWorkSheet.Cells[1, 14] = "PROJECT";
                xlWorkSheet.Cells[1, 15] = "ASO";
                xlWorkSheet.Cells[1, 16] = "NR-35";
                xlWorkSheet.Cells[1, 17] = "VACCINE-1";
                xlWorkSheet.Cells[1, 18] = "VACCINE-2";
                xlWorkSheet.Cells[1, 19] = "BOOST VACCINE";
                xlWorkSheet.Cells[1, 20] = "LOCAL";
                xlWorkSheet.Cells[1, 21] = "LEVEL";
                xlWorkSheet.Cells[1, 22] = "ESTADO";
                xlWorkSheet.Cells[1, 23] = "MOTIVO";
                xlWorkSheet.Cells[1, 24] = "USUARIO";



                // xlWorkSheet.Pictures.Add(1, 1, @"E:\work\sample.jpg");
                string mydate = DateTime.Today.ToString("yyyy/MM/dd");
                mydate = mydate.Replace("/", "_");

                xlWorkBook.SaveAs(@"C:\compartilhamento\data_base\" + textBox2.Text.Trim() + "_" + textBox3.Text.Trim() + "_" + mydate + ".xls", misValue);
                xlWorkBook.Close(true, misValue, misValue);
                app.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(app);

                //  MessageBox.Show("New date created");
                System.Windows.Forms.Application.Restart();
                //  }

                String rede = System.IO.File.ReadAllText(@"C:\compartilhamento\rede.txt");
                File.Copy(@"C:\compartilhamento\data_base\" + textBox2.Text.Trim() + "_" + textBox3.Text.Trim() + "_" + mydate + ".xls", rede.Trim() + "data_base\\" + textBox2.Text.Trim() + "_" + textBox3.Text.Trim() + "_" + mydate + ".xls", true);
                File.Copy(@"C:\compartilhamento\data_txt\PROJETO.txt", rede.Trim() + "data_txt\\PROJETO.txt", true);
                textBox2.ReadOnly = true;
                textBox3.ReadOnly = true;
            }
            catch
            {

            }
        }
        private void ler_linha_projeto()
        {

            // string FileToRead = @"C:\data_txt\data.txt";
            string[] arquivo = File.ReadAllLines(@"C:\compartilhamento\data_txt\PROJETO.txt");
            // TextReader Leitor = new StreamReader(@"C:\data_txt\ROJETO.txt", true);//Inicializa o Leitor


            textBox2.Text = arquivo[0].ToString();
            textBox3.Text = arquivo[1].ToString();
            // Leitor.Close(); //Fecha o Leitor, dando acesso ao arquivo para outros programas....


        }
        int click = 0;
        private void reset_Click(object sender, EventArgs e)
        {
            click++;
            if (click == 1)
            {
                if (textBox1.Text == "")
                {
                    MessageBox.Show("PLease enter Password");
                    click = 0;
                }
                if (textBox1.Text == "gustavoadmin")
                {

                    textBox2.ReadOnly = false;
                    textBox3.ReadOnly = false;
                    textBox2.Text = " ";
                    textBox3.Text = " ";
                }
            }
            if (click == 2)
            {

                // string mydate = DateTime.Today.ToString("yyyy/MM/dd");

                atualizar_file();
                /// MessageBox.Show("Change concluded with success");

                click = 0;
            }
        }

        private void password_Load_1(object sender, EventArgs e)
        {
            ler_linha_projeto();
        }
    }
}
