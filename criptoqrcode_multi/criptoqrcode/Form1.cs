using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System;
using System.Data;
using System.Drawing;
using System.Linq;
using AForge.Video;
using AForge.Video.DirectShow;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using Microsoft.Office.Interop.Excel;
using ZXing;
using System.IO;
using System.Runtime.InteropServices;
using Spire.Barcode;
using System.Drawing.Printing;
using System.Drawing.Imaging;
using System.Data.OleDb;
using System.Collections;
using System.Text.RegularExpressions;
using SimpleWifi;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Net;
using System.Net.NetworkInformation;
using System.Threading;
using System.Management;
using System.Diagnostics;
using System.Globalization;
using System.Text;
using Font = System.Drawing.Font;
using System.Security.Permissions;
using System.Security;
using ClosedXML.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Reflection;
using DocumentFormat.OpenXml.Bibliography;
using uPLibrary.Networking.M2Mqtt;
using uPLibrary.Networking.M2Mqtt.Messages;
using ZXing.QrCode.Internal;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Spreadsheet;
using Color = System.Drawing.Color;
using DocumentFormat.OpenXml.Drawing;
using Path = System.IO.Path;

namespace criptoqrcode
{
    public partial class Form1 : Form
    {

        delegate void SetTextCallback(string text);

        MqttClient client;
        string clientId;
        string vintequatro;
        string vintedois;
        string vinte;
        string vinteum;
        public Form1()
        {

            InitializeComponent();



        }
        //internal static password Password = new password();
        static AutoResetEvent reconnectEvent = new AutoResetEvent(false);
        //static S22.Imap.ImapClient client;
        //Form mainFormHandler;
        //Microsoft.Office.Interop.Excel.Application XcelApp = new Microsoft.Office.Interop.Excel.Application();
        private static Wifi wifi;
        public Microsoft.Office.Interop.Excel.XlCutCopyMode CutCopyMode { get; set; }
        // public static extern int GetWindowThreadProcessId(int handle, out int processId);
        Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
        Microsoft.Office.Interop.Excel.Workbook wb;

        Microsoft.Office.Interop.Excel.Worksheet ws;
        //  Excel.Application app = new Excel.Application();
        Microsoft.Office.Interop.Excel.Workbook pasta;
        Microsoft.Office.Interop.Excel.Worksheet plan;
        Microsoft.Office.Interop.Excel.Application XcelApp = new Microsoft.Office.Interop.Excel.Application();
        StreamReader rdr = null;

        String[] lMessage1 = { "Favor marcar o local e o nivel de acesso para finalizar o cadastro", "Check Local and Access level Please" };
        String[] label_nome = { "Número/Nome", "Number/Name" };//onboard People
        String[] label_onboard = { "Pessoas a bordo", "People Onboard" };//onboard People
        String[] label_emp = { "Empresa/Trip", "Company/Crew" };
        // registration only
        //Número de pessoas a bordo:
        String[] onboard = { "Número de pessoas a bordo:", "                    Onboard People:" };
        String[] label_reg = { "Cadastrar", "Registration" };
        String[] label_vessel = { "Local", "Place" };
        String[] label_porj = { "Projeto", "Project" };
        String[] label_Function = { "Função", "Function" };
        String[] label_Id = { "Identidade", "ID number" };
        String[] label_vaccine1 = { "NR-10", "NR-10" };
        String[] label_vaccine2 = { "NR-33", "NR-33" };
        String[] label_reforco = { "NR-35", "NR-35" };
        String[] label_level = { "Nivel", "Level" };
        String[] label_acc = { "Nivel de Acesso", "Access Level" };
        String[] label_yellow = { "Amarelo", "Yellow" };
        String[] label_green = { "Verde", "Green" };
        String[] label_red = { "Vermelho", "Red" };
        String[] place1 = { "Convés", "Main Deck" };
        String[] place2 = { "Praça de Maquina", "Engine Roon" };
        String[] place3 = { "Tijupá", "monkey island" };
        String[] place4 = { "Casario", "Acomodation Place" };
        String[] Label_initial = { "Data Inicial", "Check-in" };
        String[] Label_final = { "Data Limite ", "Check-out" };
        String[] Label_Read_QRcode_On = { "Modo ler Qrcode ", "Read QRcode mode" };
        String[] Label_Read_QRcode_Off = { "Modo Cadastro", "Registration mode" };
        String[] Label_Create_QRcode = { "Imprimir Qrcode:", "Print QRcode" };
        String[] Label_Show_data = { "Mostrar banco de dados:", "Show DataBase" };
        String[] Label_close_data = { "Fecha banco de dados:", "Close DataBase" };
        String[] Label_Save_data = { "Salvar banco de dados:", "Save Database Backup" };
        String[] Label_Config = { "Configurações:", "Settings" };
        String[] Label_wifi = { "Conexão Wi-Fi:", "Wi-Fi connection" };
        String[] Label_email = { "Enviar Qrcode por E-mail:", "Send Qr Code  by E-mail" };
        String[] Label_Mostrar_checkin = { "Pessoas a bordo:", "Show Onboard" };
        String[] Label_fechar = { "Fechar:", "Exit" };
        String[] Label_reset = { "Reiniciar:", "Reset" };
        String[] Label_cancel = { "Cancelar", "Cancel" };
        String[] Label_entrada = { "Entrada", "Check-in" };
        String[] Label_saida = { "Saida", "Check-out" };
        String[] id_check = { "ESTA IDENTIDADE JÁ ESTÁ CADASTRADA", "THIS ID NUMBER ALREADY EXIST" };
        String[] Label_reset_project = { "Novo Projeto", "New Project" };
        String[] Label_cadastro = { "Cadastro Concluido com sucesso", "Register concluded with success" };
        String[] label_cad1 = { "Cadastrados", "Registed" };
        String[] Label_53 = { "Cadastro Embarcação", "Vessel Register" };
        String[] bt_41 = { "Sair", "Exit" };
        String[] bt_42 = { "Cadastrar", "Register" };
        String[] bt_43 = { "Editar", "Edit" };
        String[] bt_44 = { "Cadastrar", "Register" };
        String[] bt_45 = { "Sair", "Exit" };
        String[] bt_regis = { "CADASTRAR", "REGISTER" };
        String[] cad_mode = { "Favor selecionar o local antes de colocar em modo cadastro", "Please select the place option before use register mode" };
        String[] read_mode = { "Favor selecionar o local antes de colocar em modo ler qrcode", "Please select the place option before use read mode" };
        DateTime fdataa;
        DateTime fdatab;
        string _cad;
        string _read;
        string nb;

        string passall;
        bool libera = false;
        Boolean entrou = false;
        Boolean online_ = false;
        // Print QRcode Create Show DataBase Save Backup Settings Wi-Fi connection Send Qr Code  by E-mail
        //Show Check-in
        string hostName = System.Net.Dns.GetHostName();
        string _ipstart;
        string _ipstop;
        string MyhostName;
        int tempo = 0;
        string bb;
        string rede10;
        string rede1;
        string rede2;
        string rede3;
        string localname;
        string comboname = "";
        string qr_generate = "";
        string data2 = "NOME: CRISTIANO";
        string number = "Number";
        String nome = "NOME:";
        string emp = "COMPANY:";
        string function = "FUNCTION:";
        string id = "ID:";
        string vessel = "VESSEL:";
        string initial = "START";
        string final = "END";
        string input_data = "";
        string input_hora = "";
        string output_data = "";
        string output_hora = "";
        string email = "";
        string criterio;
        string path2;
        string path3;
        string[] subs2;
        DateTime fproj;
        int band = 0;
        int confere = 0;
        int teste = 0;
        int resultado = 0;
        int checado = 0;
        int checado2 = 0;
        int zzz = 0;
        int comp =0;
        int compr = 0;
        int lista2 = 0;
        int lista3_;
        String st;
        int loc = 0;
        int lav = 0;
        bool rec = false;
        int ping_local = 0;
        int count = 2; //Count the number of successful pings
        bool grava_number = false;
        Boolean pega=false;
        String varPalavra = "teste";
        string path = @"C:\compartilhamento\data_base\2022_02_19.xls";
        int Linhas = 0;
        string z = "";
        int inside = 0;
        int s4 = 0;
        int okay = 0;
        int cri = 0;
        int aso_1 = 0;
        int id_1 = 0;
        bool online = false;
        string text1 = "";
        string text2 = "";
        string text3 = "";
        string text4 = "";
        string text5 = "";
        string text6 = "";
        string text7 = "";
        string text8 = "";
        string ip;
        string ip1;
        string ip2;
        string ip3;
        string rich5 = "";
        string rich9 = "";
        int local1val = 0;
        int local2val = 0;
        int local3val = 0;
        int local4val = 0;
        int levelyellow = 0;
        int levelgreen = 0;
        int levelred = 0;
        String myping2 = "";
        int p = 0;
        Boolean inout = false;
        int mmm = 0;
        string textbloc;
        string textbloc2;
        int l = 0;
        public string caminhoImagemSalva = null;
        public string caminhoImagemSalva2 = null;
        bool r7 = false;
        bool company_loc = false;
        bool id_exist = false;
        bool ok_but2 = false;
        bool valor_ch = false;
        bool id_onboard = false;
        bool id_onboard2 = false;
        bool cad_change = false;
        bool alter = false;
        bool cad = false;
        FilterInfoCollection filterInfoCollection;
        VideoCaptureDevice videoCaptureDevice;
        Bitmap bmp;
        // DateTime flogo3;

        DateTime flogo3 = DateTime.Now;

        int tr = 0;

        public static string GetWIFISignalStrength()
        {
            try
            {
                ObjectQuery query = new ObjectQuery("SELECT * FROM MSNdis_80211_ReceivedSignalStrength Where active = true");
                ManagementScope scope = new ManagementScope("root\\wmi");
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
                string result = "";
                foreach (ManagementObject obj in searcher.Get())
                {
                    if ((bool)obj["Active"] == true)
                    {
                        result += (string)obj["Ndis80211ReceivedSignalStrength"].ToString() + Environment.NewLine;
                    }
                }
                if (result == "")
                {
                    result = "No active WI-FI adapters found!";
                }

                return result.Trim();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Search Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return string.Empty;
            }
        }

        private static void CopyFilesRecursively(string sourcePath, string targetPath)
        {
            //Now Create all of the directories
            foreach (string dirPath in Directory.GetDirectories(sourcePath, "*", SearchOption.AllDirectories))
            {
                Directory.CreateDirectory(dirPath.Replace(sourcePath, targetPath));
            }

            //Copy all the files & Replaces any files with the same name
            foreach (string newPath in Directory.GetFiles(sourcePath, "*.*", SearchOption.AllDirectories))
            {
                File.Copy(newPath, newPath.Replace(sourcePath, targetPath), true);
            }
        }
        // Move a file into another file, delete the original, and create a backup of the replaced file.
        public static void ReplaceFile(string fileToMoveAndDelete, string fileToReplace, string backupOfFileToReplace)
        {
            // Create a new FileInfo object.
            FileInfo fInfo = new FileInfo(fileToMoveAndDelete);

            // replace the file.
            fInfo.Replace(fileToReplace, backupOfFileToReplace, true);
        }


        public static void copyAll(string SourcePath, string DestinationPath)
        {
            try
            {
                //Now Create all of the directories
                foreach (string dirPath in Directory.GetDirectories(SourcePath, "*", SearchOption.AllDirectories))
                    Directory.CreateDirectory(Path.Combine(DestinationPath, dirPath.Remove(0, SourcePath.Length)));

                //Copy all the files & Replaces any files with the same name
                foreach (string newPath in Directory.GetFiles(SourcePath, "*.*", SearchOption.AllDirectories))
                    File.Copy(newPath, Path.Combine(DestinationPath, newPath.Remove(0, SourcePath.Length)), true);
            }
            catch
            {

            }
        }

        private void clearFolder(string FolderName)
        {
            DirectoryInfo dir = new DirectoryInfo(FolderName);

            foreach (FileInfo fi in dir.GetFiles())
            {
                fi.Delete();
            }

            foreach (DirectoryInfo di in dir.GetDirectories())
            {
                clearFolder(di.FullName);
                di.Delete();
            }
        }
        int vas = 0;
        int valor = 0;

      

        private void atualiza_received()
        {






            p = 0;

            vas = 0;
            l = 1;
            try
            {


                DateTime fvessel = File.GetLastWriteTime(@"C:\compartilhamento\data_txt\PROJETO.txt");
                DateTime fvesse2 = File.GetLastWriteTime(rede1.Trim() + "\\data_txt\\PROJETO.txt");
                if (fvessel > fvesse2)
                {


                    //  copyAll(@"C:\compartilhamento\data_new_picture\", @"C:\compartilhamento\data_picture\");
                    copyAll(@"C:\compartilhamento\data_txt\PROJETO.txt", rede1.Trim() + "data_txt\\PROJETO.txt");
                    // clearFolder(@"C:\compartilhamento\data_new_picture\");

                }

                if (fvessel < fvesse2)
                {


                    //  copyAll(@"C:\compartilhamento\data_new_picture\", @"C:\compartilhamento\data_picture\");
                    // copyAll(@"C:\compartilhamento\logo_criptoqrcode\", rede1.Trim() + "data_txt\\PROJETO.txt");

                    copyAll(rede1.Trim() + "data_txt\\PROJETO.txt", @"C:\compartilhamento\data_txt\PROJETO.txt");
                    // clearFolder(@"C:\compartilhamento\data_new_picture\");

                }



                //  MessageBox.Show(p.ToString());
                DateTime flogo = File.GetLastWriteTime(@"C:\compartilhamento\logo_criptoqrcode\");
                DateTime flogo1 = File.GetLastWriteTime(rede1.Trim() + "logo_criptoqrcode\\");
                if (flogo > flogo1)
                {


                    //  copyAll(@"C:\compartilhamento\data_new_picture\", @"C:\compartilhamento\data_picture\");
                    copyAll(@"C:\compartilhamento\logo_criptoqrcode\", rede1.Trim() + "logo_criptoqrcode\\");
                    // clearFolder(@"C:\compartilhamento\data_new_picture\");

                }



                //  MessageBox.Show(p.ToString());
                DateTime fpicture = File.GetLastWriteTime(@"C:\compartilhamento\data_new_picture\");
                DateTime fpicture1 = File.GetLastWriteTime(rede1.Trim() + "data_new_picture\\");
                if (fpicture > fpicture1)
                {


                    copyAll(@"C:\compartilhamento\data_new_picture\", @"C:\compartilhamento\data_picture\");
                    copyAll(@"C:\compartilhamento\data_new_picture\", rede1.Trim() + "data_picture\\");
                    clearFolder(@"C:\compartilhamento\data_new_picture\");

                }
                //  MessageBox.Show("data_picture ok");



                DateTime ftime = File.GetLastWriteTime(@"C:\compartilhamento\data_base\" + label18.Text.Trim());
                DateTime ftime1 = File.GetLastWriteTime(rede1.Trim() + "data_base\\" + label18.Text.Trim());


                //MessageBox.Show(ftime2.ToString());
                if (ftime > ftime1)
                {
                    // MessageBox.Show(ftime.ToString());
                    // File.Copy(rede.Trim() + "data_base\\" + label18.Text.Trim(), rede.Trim() + "data_base\\" + label18.Text.Trim() + ".backup", true);
                    //  File.Delete(rede.Trim() + "data_base\\" + label18.Text.Trim());
                    // File.Copy(@"C:\compartilhamento\data_base\" + label18.Text.Trim(), rede1.Trim() + "data_base\\" + label18.Text.Trim(), true);

                }

                if (ftime1 > ftime)
                {

                    // File.Copy(@"C:\compartilhamento\data_base\" + label18.Text.Trim(), @"C:\compartilhamento\data_base\" + label18.Text.Trim() + ".backup", true);
                    //  File.Delete(@"C:\compartilhamento\data_base\" + label18.Text.Trim());
                    File.Copy(rede1.Trim() + "data_base\\" + label18.Text.Trim(), @"C:\compartilhamento\data_base\" + label18.Text.Trim(), true);

                }



                ///  MessageBox.Show("data_base ok");

                ////////////////////////////////////////
                ///


                DateTime fdata = File.GetLastWriteTime(@"C:\compartilhamento\data_txt\data.txt");








                if (vas == 0)
                {

                    DateTime f1data = File.GetLastWriteTime(rede1.Trim() + "data_txt\\data.txt");

                    if (fdata > f1data)
                    {
                        // vas = 1;
                        // File.Copy(@"C:\compartilhamento\data_txt\data.txt", rede1.Trim() + "data_txt\\data.txt", true);
                        // ler_linha();

                    }


                }



                if (vas == 0)
                {

                    // DateTime fdata = File.GetLastWriteTime(@"C:\compartilhamento\data_txt\data.txt");
                    DateTime f1data = File.GetLastWriteTime(rede1.Trim() + "data_txt\\data.txt");
                    if (f1data > fdata)
                    {
                        vas = 1;
                        File.Copy(rede1.Trim() + "data_txt\\data.txt", @"C:\compartilhamento\data_txt\data.txt", true);
                      //  ler_linha();

                    }


                }







                DateTime fdata0 = File.GetLastWriteTime(@"C:\compartilhamento\data_txt\data2.txt");


                DateTime fdata1 = File.GetLastWriteTime(rede1.Trim() + "data_txt\\data2.txt");

                if (fdata0 > fdata1)
                {

                    // File.Copy(@"C:\compartilhamento\data_txt\data2.txt", rede1.Trim() + "data_txt\\data2.txt", true);
                    // int count = File.ReadAllLines(@"C:\compartilhamento\data_txt\data2.txt").Length;
                    // label3.Text = count.ToString().Trim();
                }


                if (fdata1 > fdata0)
                {
                    //   File.Delete(@"C:\compartilhamento\data_txt\data2.txt");
                    File.Copy(rede1.Trim() + "data_txt\\data2.txt", @"C:\compartilhamento\data_txt\data2.txt", true);
                    int count = File.ReadAllLines(@"C:\compartilhamento\data_txt\data2.txt").Length;
                    label3.Text = count.ToString().Trim();
                }



                //  MessageBox.Show("data2.txt ok");





            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                // MessageBox.Show("ok");
                // valor++;
            }

            finally
            {



                //  atualiza_compartilhamento();
            }
            l = 0;
            // ler_linha();

            //  }


            timer8.Start();
        }

        private void atualiza_compartilhamento()
        {
           

            string linha;
            using (StreamReader reader = new StreamReader(@"C:\\compartilhamento\recebido.txt"))
            {
                linha = reader.ReadLine();
            }

            soma = 1;
            // timer9.Enabled= false;
            // somando = 0;

            //  if (p == 1)
            //  {
            //   MessageBox.Show(p.ToString());

            //  MessageBox.Show(rede1.ToString());

            rec = false;
            p = 0;

            vas = 0;
            l = 1;
            try
            {
              
          
                    DateTime fvessel = File.GetLastWriteTime(@"C:\compartilhamento\data_txt\PROJETO.txt");
                    DateTime fvesse2 = File.GetLastWriteTime(rede1.Trim() + "\\data_txt\\PROJETO.txt");
                    if (fvessel > fvesse2)
                    {
                 
                        File.Copy(@"C:\compartilhamento\data_txt\PROJETO.txt", rede1.Trim() + "data_txt\\PROJETO.txt", true);
              

                        //  copyAll(@"C:\compartilhamento\data_new_picture\", @"C:\compartilhamento\data_picture\");
                        //  copyAll(@"C:\compartilhamento\data_txt\PROJETO.txt", rede1.Trim() + "data_txt\\PROJETO.txt");
                        // clearFolder(@"C:\compartilhamento\data_new_picture\");





                    }
               // 

                if (fvessel < fvesse2)
                    {
                    string caminhoArquivo3 = rede1.Trim() + "\\recebido.txt";
                    // Abre o arquivo para escrita
                    using (StreamWriter sw = new StreamWriter(caminhoArquivo3))
                    {
                        sw.WriteLine("0"); // Escreve o número 1 na primeira linha do arquivo
                    }

                    File.Copy(rede1.Trim() + "data_txt\\PROJETO.txt", @"C:\compartilhamento\data_txt\PROJETO.txt", true);
                        //  copyAll(@"C:\compartilhamento\data_new_picture\", @"C:\compartilhamento\data_picture\");
                        // copyAll(@"C:\compartilhamento\logo_criptoqrcode\", rede1.Trim() + "data_txt\\PROJETO.txt");

                        //   copyAll(rede1.Trim() + "data_txt\\PROJETO.txt", @"C:\compartilhamento\data_txt\PROJETO.txt");
                        // clearFolder(@"C:\compartilhamento\data_new_picture\");

                    }





                    DateTime flogo = File.GetLastWriteTime(@"C:\compartilhamento\logo_criptoqrcode\");
                    DateTime flogo1 = File.GetLastWriteTime(rede1.Trim() + "logo_criptoqrcode\\");
                    if (flogo > flogo1)
                    {


                        //  copyAll(@"C:\compartilhamento\data_new_picture\", @"C:\compartilhamento\data_picture\");
                        copyAll(@"C:\compartilhamento\logo_criptoqrcode\", rede1.Trim() + "logo_criptoqrcode\\");
                        // clearFolder(@"C:\compartilhamento\data_new_picture\");

                    }
               


                DateTime fpicture = File.GetLastWriteTime(@"C:\compartilhamento\data_new_picture\");
                    if (rede1 != null)
                    {
                        DateTime fpicture1 = File.GetLastWriteTime(rede1.Trim() + "data_new_picture\\");
                        if (fpicture > fpicture1)
                        {

                            copyAll(@"C:\compartilhamento\data_new_picture\", @"C:\compartilhamento\data_picture\");
                            copyAll(@"C:\compartilhamento\data_new_picture\", rede1.Trim() + "data_picture\\");
                            clearFolder(@"C:\compartilhamento\data_new_picture\");

                        }
                    //  MessageBox.Show("data_picture ok");

                   

                    DateTime ftime = File.GetLastWriteTime(@"C:\compartilhamento\data_base\" + label18.Text.Trim());
                        DateTime ftime1 = File.GetLastWriteTime(rede1.Trim() + "data_base\\" + label18.Text.Trim());

                
                    //MessageBox.Show(ftime2.ToString());
                    if (ftime > ftime1)
                        {
                        // MessageBox.Show(ftime.ToString());
                        // File.Copy(rede.Trim() + "data_base\\" + label18.Text.Trim(), rede.Trim() + "data_base\\" + label18.Text.Trim() + ".backup", true);
                        //  File.Delete(rede.Trim() + "data_base\\" + label18.Text.Trim());
                      //  if (linha == "1")
                      //  {
                          

                            File.Copy(@"C:\compartilhamento\data_base\" + label18.Text.Trim(), rede1.Trim() + "data_base\\" + label18.Text.Trim(), true);
                            // MessageBox.Show("data_base maior");
                         //   MessageBox.Show("ok");
                        //    linha = "0";
                      //  }

                        }
                 
                    if (ftime1 > ftime)
                        {

                            // File.Copy(@"C:\compartilhamento\data_base\" + label18.Text.Trim(), @"C:\compartilhamento\data_base\" + label18.Text.Trim() + ".backup", true);
                            //  File.Delete(@"C:\compartilhamento\data_base\" + label18.Text.Trim());
                            File.Copy(rede1.Trim() + "data_base\\" + label18.Text.Trim(), @"C:\compartilhamento\data_base\" + label18.Text.Trim(), true);
                            // MessageBox.Show("data_base menor");
                        }

                   

                    ///  MessageBox.Show("data_base ok");

                    ////////////////////////////////////////
                    ///


                    DateTime fdata = File.GetLastWriteTime(@"C:\compartilhamento\data_txt\data.txt");

                        if (vas == 0)
                        {

                            DateTime f1data = File.GetLastWriteTime(rede1.Trim() + "data_txt\\data.txt");

                            if (fdata > f1data)
                            {
                                vas = 1;
                            //if (linha == "1")
                           // {
                                File.Copy(@"C:\compartilhamento\data_txt\data.txt", rede1.Trim() + "data_txt\\data.txt", true);
                             //  saida_manual();
                           //     linha="0";
                           //  }

                            // client.Publish("switch1", Encoding.UTF8.GetBytes("ok 1"), MqttMsgBase.QOS_LEVEL_AT_MOST_ONCE, false);
                            // File.
                            // ler_linha();
                            //  MessageBox.Show(number);
                            }


                        }



                        if (vas == 0)
                        {

                            // DateTime fdata = File.GetLastWriteTime(@"C:\compartilhamento\data_txt\data.txt");
                            DateTime f1data = File.GetLastWriteTime(rede1.Trim() + "data_txt\\data.txt");
                            if (f1data > fdata)
                            {
                                vas = 1;
                                File.Copy(rede1.Trim() + "data_txt\\data.txt", @"C:\compartilhamento\data_txt\data.txt", true);
                               // ler_linha();


                                // MessageBox.Show("data menor");

                            }


                        }







                        DateTime fdata0 = File.GetLastWriteTime(@"C:\compartilhamento\data_txt\data2.txt");
                        DateTime fdata1 = File.GetLastWriteTime(rede1.Trim() + "data_txt\\data2.txt");

                        if (fdata0 > fdata1)
                        {
                      //  if (linha == "1")
                      //  {

                            File.Copy(@"C:\compartilhamento\data_txt\data2.txt", rede1.Trim() + "data_txt\\data2.txt", true);
                            int count = File.ReadAllLines(@"C:\compartilhamento\data_txt\data2.txt").Length;
                            label3.Text = count.ToString().Trim();
                         //   linha = "0";
                       // }  
                        }


                        if (fdata1 > fdata0)
                        {
                            //   File.Delete(@"C:\compartilhamento\data_txt\data2.txt");
                            File.Copy(rede1.Trim() + "data_txt\\data2.txt", @"C:\compartilhamento\data_txt\data2.txt", true);
                            int count = File.ReadAllLines(@"C:\compartilhamento\data_txt\data2.txt").Length;
                            label3.Text = count.ToString().Trim();


                        }




                         

                    }

                    string caminhoArquivo = rede1.Trim() + "data_txt\\data4.txt";
                // Abre o arquivo para escrita
                //   using (StreamWriter sw = new StreamWriter(caminhoArquivo))
                //   {
                //      sw.WriteLine("1"); // Escreve o número 1 na primeira linha do arquivo
                //   }

                // somando++;

                //   string caminhoArquivo2 = rede1.Trim() + "\\recebido.txt";
                // Abre o arquivo para escrita
                //  using (StreamWriter sw = new StreamWriter(caminhoArquivo2))
                // {
                // sw.WriteLine("1"); // Escreve o número 1 na primeira linha do arquivo
                // }

            

            }
            catch (Exception ex)
            {
                // somando--;

                //   MessageBox.Show(ex.Message);
                ///   MessageBox.Show("falha");
                // valor++;
               // panel12.BackColor = Color.Black;
               
            }
    
            //  MessageBox.Show("ok");
            l = 0;

            comp = 1;
            //  string nomeArquivo3 = rede1.Trim() + "lock.txt";
            //  using (StreamWriter writer = new StreamWriter(nomeArquivo3, false)) //@"\data_full\vessel\local_list.txt("C:\\data_full\\local.txt", true))
            //  {
            //    writer.WriteLine("0");
            //    writer.Close();
            // }



            rec = true;
            soma = 0;
            /// timer9.Enabled = true;
         //   timer9.Enabled = true;
        }


        System.Data.DataTable dt = new System.Data.DataTable();
        private DataSet ds;
        int number2 = 0;
        private void print_qrcode()
        {

            try
            {
                if (richTextBox1.Text != "" && richTextBox2.Text != "" && richTextBox3.Text != "" && richTextBox4.Text != "" && comboBox1.Text != "" && richTextBox8.Text != "" && checado == 1)
                {


                    String data_new;
                    String data2_new;
                    if (dateTimePicker1.Visible == true)
                    {
                        data_new = dateTimePicker1.Value.ToString("dd/MM/yyyy").ToString().Trim();
                    }
                    else
                    {
                        data_new = richTextBox6.Text.Trim();
                    }
                    if (dateTimePicker2.Visible == true)
                    {
                        data2_new = dateTimePicker2.Value.ToString("dd/MM/yyyy").ToString().Trim();
                    }
                    else
                    {
                        data2_new = richTextBox7.Text.Trim();
                    }
                    if (richTextBox16.Text == "")
                    {



                        int count = File.ReadAllLines(@"C:\compartilhamento\data_txt\data2.txt").Length;
                        

  

                        string number = count.ToString().Trim();

                        // string number = System.IO.File.ReadAllText(@"C:\compartilhamento\data_txt\count.txt");
                        label3.Text = number;
                        number2 = Int32.Parse(label3.Text);
                        lb4.Visible = true;
                        number2++;
                        // File.WriteAllText(@"C:\compartilhamento\data_txt\count.txt", number2.ToString());
                        label3.Text = number2.ToString();
                        lb4.Text = label3.Text;




                        label5.Text = " ";
                        label28.Text = " ";
                        label30.Text = " ";
                        label31.Text = " ";
                        panel4.BackColor = Color.White;
                        panel11.Visible = false;
                        panel4.Visible = true;
                        label5.Text = richTextBox2.Text;
                        label28.Text = comboBox1.Text;
                        label30.Text = richTextBox9.Text;
                        label31.Text = "De: " + data_new;
                        label32.Text = "A:    " + data2_new;
                    }

                    else
                    {

                        label5.Text = " ";
                        label28.Text = " ";
                        label30.Text = " ";
                        label31.Text = " ";
                        panel4.BackColor = Color.White;
                        panel11.Visible = false;
                        panel4.Visible = true;
                        label5.Text = richTextBox2.Text;
                        label28.Text = comboBox1.Text;
                        label30.Text = richTextBox9.Text;
                        label31.Text = "De: " + data_new;
                        label32.Text = "A:    " + data2_new;

                    }


                    int width = panel4.Size.Width;
                    int height = panel4.Size.Height;
                    Bitmap bm = new Bitmap(width, height);
                    panel4.DrawToBitmap(bm, new System.Drawing.Rectangle(0, 0, width, height));
                    bm.Save(@"C:\compartilhamento\data_picture\qr\Qrcode10.png", ImageFormat.Png);


                    PrintDocument pd = new PrintDocument();
                    comboPaperSize.DisplayMember = "Etiqueta";
                    System.Drawing.Printing.PaperSize pkSize;
                    pkSize = pd.PrinterSettings.PaperSizes[57];
                    //   pd.DefaultPageSettings.PrinterResolution = new PrinterResolution() { Kind = PrinterResolutionKind.Medium };

                    comboPaperSize.Items.Add(pkSize);
                    pd.DefaultPageSettings.Landscape = true;
                    pd.PrintPage += (sender, args) =>
                    {
                        Image i = bm;
                        System.Drawing.Rectangle m = args.PageBounds;
                        args.Graphics.DrawImage(i, 20, 5, 296, 216);
                    };

                    //  bm.Save(@"C:\data_picture\qr\Qrcode10.png", ImageFormat.Png);
                    pd.Print();
                    // label5.Visible = false;
                    /// label4.Visible = false;

                }
                else
                {
                    if (band == 0)
                    {
                        MessageBox.Show("Favor preencher todos os campos");
                    }

                    if (band == 1)
                    {
                        MessageBox.Show("Please complete all informations places");
                    }
                }
            }
            catch
            {
                MessageBox.Show("A IMPRESSORA BROTHER QL810 ou QL800 NÃO ESTÁ DEFINIDA COMO IMPRESSORA PADRÃO, FAVOR DEFINIR NO PAINEL DE CONTROLE DO WINDOWS NA OPÇÃO (Dispositivos e Impressoras)!");
            }
        }
        private void out_by_user()
        {
            try
            {
                // richTextBox16.Text = secondLine.Split(':')[1].Trim();
                string bio = listBox1.SelectedItem.ToString().Trim();
                // string secondLine2 = File.ReadLines(@"C:\compartilhamento\data_txt\data2.txt").ElementAtOrDefault(Int16.Parse(bio.Split(':')[0]));
                //  Number: 1 : Name: Cristiano: Compay: Googlemarine: Funcition: Engenheiro: Id: 111098414 : E - mail : 1 : Vessel: Googlemarine: Project: 190603 : ASO: 22 / 02 / 2023 : NR - 34 : 22 / 02 / 2023 : Vaccine - 1 : 22 / 02 / 2023 : Vaccine - 2 : 22 / 02 / 2023 : Booster vaccine : 22 / 02 / 2023 : Bloqueado: GUSTAVO: Falta da quarta dose da vacina
                int rich2 = Int32.Parse(label3.Text);
                //  int lab3 = Int16.Parse(label3.Text);
                for (int i = 0; i < Int32.Parse(label3.Text); i++)
                {

                    string secondLine = File.ReadLines(@"C:\compartilhamento\data_txt\data2.txt").ElementAtOrDefault(i);
                    if (secondLine != null)
                    {
                        if (secondLine.Split(':')[1].Trim() == bio.Split(':')[1].Trim())
                        {

                            //NUMBER	NAME	COMPANY	FUNCTION	ID	EMAIL	VESSEL	CHECK-IN VALIDATION	  CHECK-OUT VALIDATION	 CHECK-IN	     CHECK-OUT	PROJECT	ASO	NR-35	VACCINE-1	VACCINE-2	BOOST VACCINE	LOCAL	LEVEL	ESTADO	MOTIVO	USUARIO
                            try
                            {    //  1        2          3     4        5    6       7          8                           9                  10              11             12              13         14      15      16        17            18             19          20     21      22     23       24

                                // NUMBER    NAME    COMPANY FUNCTION    ID EMAIL   VESSEL CHECK-IN VALIDATION     CHECK-OUT VALIDATION   CHECK-IN  DATA  CHECK-IN  HORA CHECK-OUT DATA CHECK-OUT HORA PROJECT    ASO    NR-35   VACCINE - 1   VACCINE - 2   BOOST VACCINE   LOCAL LEVEL   ESTADO MOTIVO  USUARIO

                                //Number: 1  Name: CRISTIANO CALHEIROS  Compay: GOOGLEMARINE Id: 111098414  :E - Mail: cristiano.engenharia.ac @gmail.com
                                // string bio = listBox1.SelectedItem.ToString().Trim();

                                // 0      1    2              3             4         5             6           7      8       9          10                    11                      12         13          14      15      16        17           18            19            20              22              23               24               25              26         27   28
                                // Number: 1 : Name: CRISTIANO CALHEIROS : Compay: GOOGLEMARINE: Funcition: ENGENHEIRO: Id: 111098414 : E - mail : cristiano.engenharia.ac @gmail.com: Vessel: Googlemarine: Project: 190605 : ASO: 02 / 02 / 2025 : NR - 34 : 02 / 02 / 2025 : Vaccine - 1 : 02 / 02 / 2025 : Vaccine - 2 : 02 / 02 / 2025 : Booster vaccine : 02 / 02 / 2025 :  : COMUM:
                                pasta = app.Workbooks.Open(@"C:\compartilhamento\data_base\" + label18.Text);
                                plan = pasta.Worksheets["Planilha1"];
                                int lastRow = plan.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;

                                txtCodigoFunci.Text = lastRow.ToString();
                                lastRow++;
                                output_data = DateTime.Now.ToString("MM/dd/yyyy").Trim();
                                output_hora = DateTime.Now.ToString("HH:mm:ss tt").Trim();

                                plan.Cells[lastRow, 1] = secondLine.Split(':')[1].Trim();
                                plan.Cells[lastRow, 2] = secondLine.Split(':')[3].Trim();
                                plan.Cells[lastRow, 3] = secondLine.Split(':')[5].Trim();
                                plan.Cells[lastRow, 4] = secondLine.Split(':')[7].Trim();
                                plan.Cells[lastRow, 5] = secondLine.Split(':')[9].Trim();
                                plan.Cells[lastRow, 6] = secondLine.Split(':')[11].Trim();
                                plan.Cells[lastRow, 7] = secondLine.Split(':')[13].Trim();
                                //  plan.Cells[lastRow, 9] = secondLine.Split(':')[17];
                                plan.Cells[lastRow, 10] = ""; // input;
                                plan.Cells[lastRow, 11] = "";
                                plan.Cells[lastRow, 12] = output_data;
                                plan.Cells[lastRow, 13] = output_hora + ": MANUAL PELO USUÁRIO";
                                plan.Cells[lastRow, 14] = secondLine.Split(':')[15].Trim();
                                // plan.Cells[lastRow, 15] = richTextBox12.Text;
                                // plan.Cells[lastRow, 16] = richTextBox13.Text;
                                // plan.Cells[lastRow, 17] = richTextBox14.Text;
                                plan.Cells[lastRow, 24] = comuser.Text;
                                pasta.Save();
                                pasta.Close();
                                app.Quit();
                                //  pasta.Close(true, misValue, misValue);
                                //  xlApp.Quit();
                                Marshal.ReleaseComObject(pasta);
                                Marshal.ReleaseComObject(pasta);
                                Marshal.ReleaseComObject(pasta);
                            }
                            catch
                            {

                            }
                            // */
                            break;
                        }

                    }

                }
            }
            catch
            {
                /// MessageBox.Show("NÃO A DADOS CADASTRADOS!");
            }




        }
        // [DllImport("user32.dll"), CharSet= CharSet.Auto, SetLastError = true)]
        //  public static extern int GetWindowThreadProcessId(int handle,out int processId);



        private void carrega_planilha_txt()
        {
            string filePath = "C:\\compartilhamento\\data_base\\novo.txt";



            // localname = local1.Text + " " + local2.Text + " " + local3.Text + " " + local4.Text;
            if (local1.Checked == true)
            {
                localname = localname + " " + local1.Text;
            }
            if (local2.Checked == true)
            {
                localname = localname + " " + local2.Text;
            }
            if (local3.Checked == true)
            {
                localname = localname + " " + local3.Text;
            }
            if (local4.Checked == true)
            {
                localname = localname + " " + local4.Text;
            }


            vinte = localname;

            localname = "";
            // localname = local1.Text + " " + local2.Text + " " + local3.Text + " " + local4.Text;
            if (level_yellow.Checked == true)
            {
                localname = localname + " " + level_yellow.Text;

            }
            if (level_green.Checked == true)
            {
                localname = localname + " " + level_green.Text;
            }
            if (level_red.Checked == true)
            {
                localname = localname + " " + level_red.Text;
            }
            vinteum = localname;

            string contentToAppend = richTextBox16.Text.Trim() + "," + richTextBox2.Text.Trim() + "," + richTextBox1.Text.Trim() + "," + richTextBox3.Text.Trim() + "," + richTextBox4.Text.Trim() + "," + richTextBox8.Text.Trim() + "," + comboBox1.Text.Trim() + "," + richTextBox6.Text.Trim() + "  Até  " + richTextBox7.Text.Trim() + "," + "." + richTextBox7.Text.Trim() + "," + input_data.Trim() + "," + input_hora.Trim() + "," + output_data.Trim() + "," + output_hora.Trim() + "," + richTextBox9.Text.Trim() + "," + richTextBox10.Text.Trim() + "," + richTextBox11.Text.Trim() + "," + richTextBox12.Text.Trim() + "," + richTextBox13.Text.Trim() + "," + richTextBox14.Text.Trim() + "," + vinte + ","+vinteum+"," + vintedois + "," + vintequatro + ",";

            // Verifica se o arquivo existe antes de tentar adicionar o conteúdo
            if (File.Exists(filePath))
            {
                // Abre o arquivo em modo de anexação (append)
                using (StreamWriter writer = File.AppendText(filePath))
                {
                    // Escreve o conteúdo no arquivo
                    writer.WriteLine(contentToAppend);
                }

                Console.WriteLine("Conteúdo adicionado com sucesso.");
            }
        }
        private void CarregarPlanilha()
        {

            //  int lastRow = plan.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;

            pasta = app.Workbooks.Open(@"C:\compartilhamento\data_base\" + label18.Text);
            plan = pasta.Worksheets["Planilha1"];
            int lastRow = plan.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;

            txtCodigoFunci.Text = lastRow.ToString();
            lastRow++;
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
                         output_data = DateTime.Now.ToString("dd-MM-yyyy");
                output_hora = DateTime.Now.ToString("hh:mm:ss tt");
*/

            string contentToAppend = nb + "," + richTextBox2.Text + "," + richTextBox1.Text + "," + richTextBox3.Text + "," + richTextBox4.Text + "," + richTextBox8.Text + "," + comboBox1.Text + "," + richTextBox6.Text.Trim() + "  Até  " + richTextBox7.Text.Trim() + "," + "." + richTextBox7.Text + ","+ input_data+","+ input_hora+","+ output_data+","+ output_hora+"," + richTextBox9.Text.Trim() + "," + richTextBox10.Text.Trim() + "," + richTextBox11.Text.Trim() + "," + richTextBox12.Text.Trim() + "," + richTextBox13.Text.Trim() + "," + richTextBox14.Text.Trim() + "," + vinte + ",," + vintedois + ",," + vintequatro + ",";
            plan.Cells[lastRow, 1] = richTextBox16.Text;
            plan.Cells[lastRow, 2] = richTextBox2.Text;
            plan.Cells[lastRow, 3] = richTextBox1.Text;
            plan.Cells[lastRow, 4] = richTextBox3.Text;
            plan.Cells[lastRow, 5] = richTextBox4.Text;
            plan.Cells[lastRow, 6] = richTextBox8.Text;
            plan.Cells[lastRow, 7] = comboBox1.Text;
            plan.Cells[lastRow, 8] = richTextBox6.Text.Trim() + "  Até  " + richTextBox7.Text.Trim();
            plan.Cells[lastRow, 9] = "." + richTextBox7.Text.Trim(); //DateTime.Now.ToString("hh:mm:ss tt");
            plan.Cells[lastRow, 10] = input_data;
            plan.Cells[lastRow, 11] = input_hora;
            plan.Cells[lastRow, 12] = output_data;
            plan.Cells[lastRow, 13] = output_hora;
            plan.Cells[lastRow, 14] = richTextBox9.Text.Trim();
            plan.Cells[lastRow, 15] = richTextBox10.Text.Trim();
            plan.Cells[lastRow, 16] = richTextBox11.Text.Trim();
            plan.Cells[lastRow, 17] = richTextBox12.Text.Trim();
            plan.Cells[lastRow, 18] = richTextBox13.Text.Trim();
            plan.Cells[lastRow, 19] = richTextBox14.Text.Trim();
            plan.Cells[lastRow, 24] = comuser.Text;

            localname = "";

            if (local1.Checked == true)
            {
                localname = localname + " " + local1.Text;
            }
            if (local2.Checked == true)
            {
                localname = localname + " " + local2.Text;
            }
            if (local3.Checked == true)
            {
                localname = localname + " " + local3.Text;
            }
            if (local4.Checked == true)
            {
                localname = localname + " " + local4.Text;
            }


            plan.Cells[lastRow, 20] = localname;

            localname = "";
            // localname = local1.Text + " " + local2.Text + " " + local3.Text + " " + local4.Text;
            if (level_yellow.Checked == true)
            {
                localname = localname + " " + level_yellow.Text;
            }
            if (level_green.Checked == true)
            {
                localname = localname + " " + level_green.Text;
            }
            if (level_red.Checked == true)
            {
                localname = localname + " " + level_red.Text;
            }

            plan.Cells[lastRow, 21] = localname;

            pasta.Save();
            pasta.Close();
            Marshal.ReleaseComObject(pasta);
            Marshal.ReleaseComObject(pasta);
            Marshal.ReleaseComObject(pasta);

            checa_host();


            try
            {
                  checa_host();
                // atualiza_compartilhamento();

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.ToString(), ex.Message, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            //  MessageBox.Show("CarregarPlanilha - leitura");
        }

        private void bloqueio()
        {
            try
            {
                string[] lines = File.ReadAllLines(@"C:\compartilhamento\data_txt\data2.txt");

                // Start at line number 2 because there is a header
                for (int i = 0; i < lines.Length; i++)
                {
                    if (richTextBox16.Text.Trim() != "")
                    {
                        if (lines[i].Contains(richTextBox4.Text.Trim()))
                        {
                            // Copy it where you want
                            //  MessageBox.Show(lines[i].ToString());
                            string text3 = lines[i];

                            string text4 = "Number : " + richTextBox16.Text + " : Name : " + richTextBox2.Text + " : Compay : " + richTextBox1.Text + " :Funcition:  " + richTextBox3.Text + "  :Id: " + richTextBox4.Text + " : E-mail : " + richTextBox8.Text + " : Vessel : " + comboBox1.Text + " : Project : " + richTextBox9.Text + " : ASO : " + maskedTextBox1.Text + " : NR-34 : " + maskedTextBox2.Text + " : NR-10 : " + maskedTextBox3.Text + " : NR-33 : " + maskedTextBox4.Text + " : NR-35 : " + maskedTextBox5.Text + " : " + bb + " : " + comuser.Text + " :" + richTextBox17.Text;
                            string text = File.ReadAllText(@"C:\compartilhamento\data_txt\data2.txt");
                            text = text.Replace(text3, text4);
                            File.WriteAllText(@"C:\compartilhamento\data_txt\data2.txt", text);
                        }
                    }
                }




                //////////////////////////

                pasta = app.Workbooks.Open(@"C:\compartilhamento\data_base\" + label18.Text);
                plan = pasta.Worksheets["Planilha1"];
                int lastRow = plan.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;

                txtCodigoFunci.Text = lastRow.ToString();
                lastRow++;

                if (richTextBox16.Text == "")
                {
                    plan.Cells[lastRow, 1] = number2;
                }
                if (richTextBox16.Text != "")
                {
                    plan.Cells[lastRow, 1] = Int32.Parse(richTextBox16.Text);
                }
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
             output_data = DateTime.Now.ToString("dd-MM-yyyy");
    output_hora = DateTime.Now.ToString("hh:mm:ss tt");
*/
                // qr_generate = label37.Text;
                plan.Cells[lastRow, 2] = richTextBox2.Text;
                plan.Cells[lastRow, 3] = richTextBox1.Text;
                plan.Cells[lastRow, 4] = richTextBox3.Text;
                plan.Cells[lastRow, 5] = richTextBox4.Text;
                plan.Cells[lastRow, 6] = richTextBox8.Text;
                plan.Cells[lastRow, 7] = comboBox1.Text;
                // plan.Cells[lastRow, 8] = DateTime.Now;

                //xlWorkSheet.Cells[1, 20] = "ESTADO";
                // xlWorkSheet.Cells[1, 21] = "MOTIVO";
                // xlWorkSheet.Cells[1, 22] = "USUARIO";
                plan.Cells[lastRow, 14] = richTextBox9.Text;
                plan.Cells[lastRow, 15] = richTextBox10.Text;
                plan.Cells[lastRow, 16] = richTextBox11.Text;
                plan.Cells[lastRow, 17] = richTextBox12.Text;
                plan.Cells[lastRow, 18] = richTextBox13.Text;
                plan.Cells[lastRow, 19] = richTextBox14.Text;
                plan.Cells[lastRow, 22] = bb;
                plan.Cells[lastRow, 23] = richTextBox17.Text;
                plan.Cells[lastRow, 24] = comuser.Text + ": " + DateTime.Now;

                localname = "";
                bb = "";
                // localname = local1.Text + " " + local2.Text + " " + local3.Text + " " + local4.Text;
                if (local1.Checked == true)
                {
                    localname = localname + " " + local1.Text;
                }
                if (local2.Checked == true)
                {
                    localname = localname + " " + local2.Text;
                }
                if (local3.Checked == true)
                {
                    localname = localname + " " + local3.Text;
                }
                if (local4.Checked == true)
                {
                    localname = localname + " " + local4.Text;
                }


                plan.Cells[lastRow, 20] = localname;

                localname = "";
                // localname = local1.Text + " " + local2.Text + " " + local3.Text + " " + local4.Text;
                if (level_yellow.Checked == true)
                {
                    localname = localname + " " + level_yellow.Text;

                }
                if (level_green.Checked == true)
                {
                    localname = localname + " " + level_green.Text;
                }
                if (level_red.Checked == true)
                {
                    localname = localname + " " + level_red.Text;
                }

                plan.Cells[lastRow, 21] = localname;




                pasta.Save();
                //pasta.Close();
                app.Quit();
                pasta.Close();
                Marshal.ReleaseComObject(pasta);
                Marshal.ReleaseComObject(pasta);
                Marshal.ReleaseComObject(pasta);
                //CarregarPlanilha();
                //   atualiza_compartilhamento();

            }
            catch (Exception ex)
            {

                // MessageBox.Show(ex.ToString(), ex.Message, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        public void txt_to_excel()
        {
            string[] InputNamesLines = System.IO.File.ReadAllLines(@"c:\teste\excel.txt");
            string secondLine = File.ReadLines(@"c:\teste\data2.txt").ElementAtOrDefault(1);
            //if()
            Excel.Application oXl;
            Excel._Workbook oWB;
            Excel._Worksheet xlWorkSheet;
            Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;
            try
            {
                oXl = new Excel.Application();
                oXl.Visible = true;

                oWB = (Excel._Workbook)(oXl.Workbooks.Add(""));
                xlWorkSheet = (Excel.Worksheet)oWB.ActiveSheet;
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
                xlWorkSheet.Cells[1, 16] = "NR-34";
                xlWorkSheet.Cells[1, 17] = "NR-10";
                xlWorkSheet.Cells[1, 18] = "NR-33";
                xlWorkSheet.Cells[1, 19] = "NR-35";
                xlWorkSheet.Cells[1, 20] = "LOCAL";
                xlWorkSheet.Cells[1, 21] = "LEVEL";
                xlWorkSheet.Cells[1, 22] = "ESTADO";
                xlWorkSheet.Cells[1, 23] = "MOTIVO";
                xlWorkSheet.Cells[1, 24] = "USUARIO";
                xlWorkSheet.get_Range("A1", "C1").Font.Bold = true;
                xlWorkSheet.get_Range("A1", "C1").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                for (int i = 1; i <= InputNamesLines.Length; i++)
                {
                    // oSheet.Cells[1][i + 1] = i;
                    // oSheet.Cells[2][i + 1] =   InputNamesLines[i - 1];
                    // oSheet.Cells[3]= secondLine.Split(':')[1].Trim();
                }
                xlWorkSheet.Cells[1, 3] = secondLine.Split(':')[3].Trim();
                Thread.Sleep(5000);
                oRng = xlWorkSheet.get_Range("A1", "C1");
                oRng.EntireColumn.AutoFit();
                oXl.Visible = false;
                oWB.SaveAs(@"c:\teste\teste.xls", Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                oWB.Close();
                Marshal.ReleaseComObject(oWB);
                Marshal.ReleaseComObject(oWB);
                Marshal.ReleaseComObject(oWB);
            }
            catch
            {

            }
        }

        private void carrega_planilha2_txt()
        {
            string filePath = "C:\\compartilhamento\\data_base\\novo.txt";
        
            if (richTextBox16.Text == "")
            {
                nb = number2.ToString();
            }
            if (richTextBox16.Text != "")
            {
                nb= richTextBox16.Text.Trim();
            }
            if (alter == true)
            {
                vintequatro = comuser.Text + ": " + DateTime.Now;
                vintedois = "CADASTRO ALTERADO";
            }
            if (alter == false)
            {
                vintedois = qr_generate + ": " + DateTime.Now;
                vintequatro = comuser.Text;
            }
            if (cad == true)
            {
                vintequatro = comuser.Text + ": " + DateTime.Now;
                vintedois = "CADASTRADO";
                cad = false;
            }

            // localname = local1.Text + " " + local2.Text + " " + local3.Text + " " + local4.Text;
            if (local1.Checked == true)
            {
                localname = localname + " " + local1.Text;
            }
            if (local2.Checked == true)
            {
                localname = localname + " " + local2.Text;
            }
            if (local3.Checked == true)
            {
                localname = localname + " " + local3.Text;
            }
            if (local4.Checked == true)
            {
                localname = localname + " " + local4.Text;
            }


            vinte = localname;

            localname = "";
            // localname = local1.Text + " " + local2.Text + " " + local3.Text + " " + local4.Text;
            if (level_yellow.Checked == true)
            {
                localname = localname + " " + level_yellow.Text;

            }
            if (level_green.Checked == true)
            {
                localname = localname + " " + level_green.Text;
            }
            if (level_red.Checked == true)
            {
                localname = localname + " " + level_red.Text;
            }


            string contentToAppend = nb+","+ richTextBox2.Text+","+ richTextBox1.Text+","+ richTextBox3.Text+","+ richTextBox4.Text+","+ richTextBox8.Text+","+ comboBox1.Text+","+ richTextBox6.Text+","+ "." + richTextBox7.Text+ ",,,,,"+richTextBox9.Text.Trim()+","+ richTextBox10.Text.Trim()+","+ richTextBox11.Text.Trim()+","+ richTextBox12.Text.Trim()+","+ richTextBox13.Text.Trim()+","+ richTextBox14.Text.Trim()+","+vinte+",,"+vintedois+",,"+ vintequatro + "," ;

            // Verifica se o arquivo existe antes de tentar adicionar o conteúdo
            if (File.Exists(filePath))
            {
                // Abre o arquivo em modo de anexação (append)
                using (StreamWriter writer = File.AppendText(filePath))
                {
                    // Escreve o conteúdo no arquivo
                    writer.WriteLine(contentToAppend);
                }

                Console.WriteLine("Conteúdo adicionado com sucesso.");
            }
        }
        private void CarregarPlanilha2()
        {
            try
            {

                okay = 1;
                pasta = app.Workbooks.Open(@"C:\compartilhamento\data_base\" + label18.Text);

                plan = pasta.Worksheets["Planilha1"];
                int lastRow = plan.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;


                txtCodigoFunci.Text = lastRow.ToString();
                lastRow++;

                if (richTextBox16.Text == "")
                {
                    plan.Cells[lastRow, 1] = number2;
                }
                if (richTextBox16.Text != "")
                {
                    plan.Cells[lastRow, 1] = Int16.Parse(richTextBox16.Text.Trim());
                }


                plan.Cells[lastRow, 2] = richTextBox2.Text;
                plan.Cells[lastRow, 3] = richTextBox1.Text;
                plan.Cells[lastRow, 4] = richTextBox3.Text;
                plan.Cells[lastRow, 5] = richTextBox4.Text;
                plan.Cells[lastRow, 6] = richTextBox8.Text;
                plan.Cells[lastRow, 7] = comboBox1.Text;
                plan.Cells[lastRow, 8] = richTextBox6.Text;
                plan.Cells[lastRow, 9] = "." + richTextBox7.Text;
                plan.Cells[lastRow, 14] = richTextBox9.Text;


                plan.Cells[lastRow, 15] = richTextBox10.Text.Trim();
                plan.Cells[lastRow, 16] = richTextBox11.Text.Trim();
                plan.Cells[lastRow, 17] = richTextBox12.Text.Trim();
                plan.Cells[lastRow, 18] = richTextBox13.Text.Trim();
                plan.Cells[lastRow, 19] = richTextBox14.Text.Trim();
                // MessageBox.Show("ok");


                // plan.Cells[lastRow, 24] = comuser.Text;

                localname = "";




                if (alter == true)
                {
                    plan.Cells[lastRow, 24] = comuser.Text + ": " + DateTime.Now;
                    plan.Cells[lastRow, 22] = "CADASTRO ALTERADO";
                }
                if (alter == false)
                {
                    plan.Cells[lastRow, 22] = qr_generate + ": " + DateTime.Now;
                    plan.Cells[lastRow, 24] = comuser.Text;
                }
                if (cad == true)
                {
                    plan.Cells[lastRow, 24] = comuser.Text + ": " + DateTime.Now;
                    plan.Cells[lastRow, 22] = "CADASTRADO";
                    cad = false;
                }

                // localname = local1.Text + " " + local2.Text + " " + local3.Text + " " + local4.Text;
                if (local1.Checked == true)
                {
                    localname = localname + " " + local1.Text;
                }
                if (local2.Checked == true)
                {
                    localname = localname + " " + local2.Text;
                }
                if (local3.Checked == true)
                {
                    localname = localname + " " + local3.Text;
                }
                if (local4.Checked == true)
                {
                    localname = localname + " " + local4.Text;
                }


                plan.Cells[lastRow, 20] = localname;

                localname = "";
                // localname = local1.Text + " " + local2.Text + " " + local3.Text + " " + local4.Text;
                if (level_yellow.Checked == true)
                {
                    localname = localname + " " + level_yellow.Text;

                }
                if (level_green.Checked == true)
                {
                    localname = localname + " " + level_green.Text;
                }
                if (level_red.Checked == true)
                {
                    localname = localname + " " + level_red.Text;
                }

                plan.Cells[lastRow, 21] = localname;




                pasta.Save();
                pasta.Close();
                Marshal.ReleaseComObject(pasta);
                Marshal.ReleaseComObject(pasta);
                Marshal.ReleaseComObject(pasta);
                //  atualiza_compartilhamento();
               // myThread.Abort();
                checa_host();
                //  MessageBox.Show("CarregarPlanilha - atualiza excel");

            }
            catch
            {

            }
            //CarregarPlanilha();

        }
        private static void CloseExcel(Excel.Application ExcelApplication = null)
        {
            if (ExcelApplication != null)
            {
                //ExcelApplication.Workbooks.Close();
                // ExcelApplication.Quit();
            }

            System.Diagnostics.Process[] PROC = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (System.Diagnostics.Process PK in PROC)
            {
                //  if (PK.MainWindowTitle.Length == 0) { PK.Kill(); }
            }


            var valor = 1;
            //  var array[] = 1;
            int x = 2;
            var processes = Process.GetProcessesByName("Excel");

            foreach (var p in processes)
            {


                //  x++;
                if (x >= 2)
                {
                    valor = p.Id;

                    //MessageBox.Show(valor.ToString());
                    Process processes2 = Process.GetProcessById(valor);
                    processes2.Kill();
                    x--;
                }
            }
            x = 2;


        }
        private void excel_close()
        {

            int ourPID = 0;
            int tmpX = 0;
            int indexRow = 1;
            int indexCol = 1;
            int[] existingPIDs;
            existingPIDs = new int[100];

            Process[] localByName = Process.GetProcessesByName("excel");
            // user didnt have any excels open, kill excel
            if (tmpX == 0)
            {
                foreach (Process proc in localByName)
                {
                    proc.Kill();
                }
            }
            // user does have excel(s) already open, only kill our apps excel
            else if (tmpX > 0 && ourPID != 0)
            {
                foreach (Process proc in localByName)
                {
                    if (proc.Id == ourPID)
                    {
                        proc.Kill();
                    }
                }
            }
        }

        private void GerarCabecalho()
        {

            pasta = app.Workbooks.Open("label17.Text");
            plan = pasta.Worksheets["Planilha1"];
            plan.Cells[1, 1].Text = "NAME";
            plan.Cells[1, 2].Text = "COMPANY";
            plan.Cells[1, 3].Text = "FUNCTION";
            plan.Cells[1, 4].Text = "ID";
            plan.Cells[1, 5].Text = "EMAIL:";
            plan.Cells[1, 6].text = "VESSEL";
            plan.Cells[1, 7].Text = "INITIAL";
            plan.Cells[1, 8].Text = "FINAL";
            plan.Cells[1, 9].Text = "INPUT";
            plan.Cells[1, 10].Text = "OUTPU";
            pasta.Save();
            pasta.Close();
            app.Quit();
            Marshal.ReleaseComObject(pasta);
            Marshal.ReleaseComObject(pasta);
            Marshal.ReleaseComObject(pasta);
        }

        int wi = 0;
        // VideoCaptureDevice videoCaptureDevice;
        private void button6_Click(object sender, EventArgs e)
        {

        }
        private void atualizar_file()
        {

            var file2 = new DirectoryInfo(@"C:\compartilhamento\data_base\").GetFiles().OrderBy(o => o.CreationTime).LastOrDefault();
            label17.Text = file2.ToString();


            //  if (label17.Text != label18.Text)
            //  {
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = app.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            // xlWorkSheet = pasta.Worksheets["Planilha1"];
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
            xlWorkSheet.Cells[1, 16] = "NR-34";
            xlWorkSheet.Cells[1, 17] = "NR-10";
            xlWorkSheet.Cells[1, 18] = "NR-33";
            xlWorkSheet.Cells[1, 19] = "NR-35";
            xlWorkSheet.Cells[1, 20] = "LOCAL";
            xlWorkSheet.Cells[1, 21] = "LEVEL";
            xlWorkSheet.Cells[1, 22] = "ESTADO";
            xlWorkSheet.Cells[1, 23] = "MOTIVO";
            xlWorkSheet.Cells[1, 24] = "USUARIO";


            // xlWorkSheet.Pictures.Add(1, 1, @"E:\work\sample.jpg");
            string mydate = DateTime.Today.ToString("yyyy/MM/dd");
            label18.Text = mydate.Replace("/", "_") + ".xls";

            xlWorkBook.SaveAs(@"C:\compartilhamento\data_base\" + label18.Text, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            app.Quit();

            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlWorkBook);

            MessageBox.Show("New date created");
            System.Windows.Forms.Application.Restart();
            //  }

        }

        private void inicialize_qrreader()
        {
            vid = 1;
            richTextBox1.ReadOnly = true;
            richTextBox2.ReadOnly = true;
            richTextBox3.ReadOnly = true;
            richTextBox4.ReadOnly = true;
            //richTextBox5.ReadOnly = true;
            richTextBox6.ReadOnly = true;
            richTextBox7.ReadOnly = true;
            richTextBox8.ReadOnly = true;
            // button2.Enabled = false;

            button1.Text = "Read QRcode On";
            videoCaptureDevice = new VideoCaptureDevice(filterInfoCollection[cboCamera.SelectedIndex].MonikerString);
            videoCaptureDevice.NewFrame += FinalFrame_NewFrame;
            //   videoCaptureDevice.Start();
            dateTimePicker1.Visible = false;
            dateTimePicker2.Visible = false;



            richTextBox1.Text = "";
            richTextBox2.Text = "";
            richTextBox3.Text = "";
            richTextBox4.Text = "";
            //richTextBox5.Text = "";
            richTextBox6.Text = "";
            richTextBox7.Text = "";
            richTextBox8.Text = "";


            richTextBox6.Visible = true;
            richTextBox7.Visible = true;
            if (timer4.Enabled == false)
            {
                // button2.Enabled = false;
                timer4.Start();
            }
            if (textBox7.SelectionLength >= 0)
            {
                textBox7.Focus();
                textBox7.Text = "";
            }
        }

        private static string GetFiles(string path)
        {
            
            var file = new DirectoryInfo(path).GetFiles().OrderByDescending(o => o.LastWriteTime).FirstOrDefault();
            return file.Name;
            // label18.Text =
        }
        private void ckecked_false()
        {
            richTextBox1.Text = "";
            richTextBox2.Text = "";
            richTextBox3.Text = "";
            richTextBox4.Text = "";
            // richTextBox5.Text = "";
            richTextBox6.Text = "";
            richTextBox7.Text = "";
            richTextBox8.Text = "";
            // richTextBox9.Text = "";
            richTextBox10.Text = "";
            richTextBox11.Text = "";
            richTextBox12.Text = "";
            richTextBox13.Text = "";
            richTextBox14.Text = "";
            richTextBox16.Text = "";
            button8.Visible = false;
            button9.Visible = false;
            button10.Visible = false;
            maskedTextBox1.Text = " ";
            maskedTextBox2.Text = " ";
            maskedTextBox3.Text = " ";
            maskedTextBox4.Text = " ";
            maskedTextBox5.Text = " ";
            local1.Checked = false;
            local2.Checked = false;
            local3.Checked = false;
            // local4.Checked = false;


            //level_green.Checked = false;
            // level_red.Checked = false;
            //  level_yellow.Checked = false;

        }
        private void FinalFrame_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            eventArgs.Frame.RotateFlip(RotateFlipType.RotateNoneFlipX);
            pictureBox7.Image = (Bitmap)eventArgs.Frame.Clone();
        }

        int vid = 0;
        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text != "")
            {

                if (panel6.Visible == false)
                {

                    if (richTextBox3.Text != "")
                    {
                        comboBox1.Items.Clear();
                        StreamReader sr = new StreamReader(@"C:\compartilhamento\vessels.txt");
                        string x = sr.ReadToEnd();
                        string[] y = x.Split('\n');
                        foreach (string s in y)
                        {
                            comboBox1.Items.Add(s);
                        }
                        sr.Close();
                    }
                }

                label54.Visible = false;
                label44.Visible = false;
                label45.Visible = false;
                label46.Visible = false;
                label47.Visible = false;
                label48.Visible = false;
                label49.Visible = false;
                label50.Visible = false;


                dataGridView1.Visible = false;
                local1.Enabled = false;
                local2.Enabled = false;
                local4.Enabled = false;
                plant = 0;

                if (band == 0)
                {
                    button3.Text = Label_Show_data[0];
                }
                else
                {
                    button3.Text = Label_Show_data[1];
                }

                // button3.Text = "Show DataBase";


                button1.Enabled = true;


                tempo = 0;
                panel11.Visible = true;
                pictureBox7.Image = Properties.Resources.barcode1;
                panel11.BackColor = Color.White;
                label8.Visible = false;
                maskedTextBox1.ReadOnly = true;
                maskedTextBox2.ReadOnly = true;
                maskedTextBox3.ReadOnly = true;
                maskedTextBox4.ReadOnly = true;
                maskedTextBox5.ReadOnly = true;
                maskedTextBox1.Text = " ";
                maskedTextBox2.Text = " ";
                maskedTextBox3.Text = " ";
                maskedTextBox4.Text = " ";
                maskedTextBox5.Text = " ";
                vid = 1;


                //  if (vid == 1)
                // {

                if (textBox7.SelectionLength >= 0)
                {
                    textBox7.Focus();
                    textBox7.Text = "";
                }
                richTextBox1.ReadOnly = true;
                richTextBox2.ReadOnly = true;
                richTextBox3.ReadOnly = true;
                richTextBox4.ReadOnly = true;
                //richTextBox5.ReadOnly = true;
                richTextBox6.ReadOnly = true;
                richTextBox7.ReadOnly = true;
                richTextBox8.ReadOnly = true;
                //richTextBox9.ReadOnly = true;
                richTextBox10.ReadOnly = true;
                richTextBox11.ReadOnly = true;
                richTextBox12.ReadOnly = true;
                richTextBox13.ReadOnly = true;
                richTextBox14.ReadOnly = true;
                richTextBox16.ReadOnly = true;
                // button2.Enabled = false;
                if (band == 0)
                {
                    button1.Text = Label_Read_QRcode_On[0];
                    label7.Text = button1.Text;
                }
                else
                {
                    button1.Text = Label_Read_QRcode_On[1];
                    label7.Text = button1.Text;
                }

                panel10.Visible = false;

                videoCaptureDevice = new VideoCaptureDevice(filterInfoCollection[cboCamera.SelectedIndex].MonikerString);
                videoCaptureDevice.NewFrame += FinalFrame_NewFrame;
                dateTimePicker1.Visible = false;
                dateTimePicker2.Visible = false;
                pictureBox7.Image = Properties.Resources.barcode1;
                //  pictureBox1.BackgroundImage = Properties.Resources.frame;

                richTextBox1.Text = "";
                richTextBox2.Text = "";
                richTextBox3.Text = "";
                richTextBox4.Text = "";
                // richTextBox5.Text = "";
                richTextBox6.Text = "";
                richTextBox7.Text = "";
                richTextBox8.Text = "";
                // richTextBox9.Text = "";
                richTextBox10.Text = "";
                richTextBox11.Text = "";
                richTextBox12.Text = "";
                richTextBox13.Text = "";
                richTextBox14.Text = "";
                richTextBox16.Text = "";

                richTextBox6.Visible = true;
                richTextBox7.Visible = true;

                ckecked_false();
                //  if (timer4.Enabled == false)
                //  {
                //  button2.Enabled = false;
                // timer4.Start();
                //  }


                //}

                if (vid == 2)
                {

                }
            }
            else
            {
                MessageBox.Show(_read);
            }
        }

        private void ler_linha()
        {
            //   String locked = File.ReadLines(@"C:\compartilhamento\lock.txt").ElementAtOrDefault(0);


            //  timer8.Stop();

            // textBox16.Text = "";
            string[] linhas22 = System.IO.File.ReadAllLines(@"C:\compartilhamento\data_txt\data.txt");


            try
            {
                if (textBox16.Text == "") {

                    if (comboBox2.Text.Trim() == "ALL")
                    {
                        listBox1.Items.Clear();
                        lista_ = 0;
                        foreach (string linha22 in linhas22)
                        {
                            if (linha22.Contains("Vessel"))
                            {
                                listBox1.Items.Add(linha22);
                                lista_++;
                                label27.Text = lista_.ToString();
                                label67.Text = lista_.ToString();






                            }
                        }
                        if (lista_ == 0)
                        {
                            label27.Text = lista_.ToString();
                            //label67.Text = label27.Text;
                        }
                    }
                    else
                    {
                        listBox1.Items.Clear();
                        lista_ = 0;
                        foreach (string linha222 in linhas22)
                        {

                            if (linha222.Contains(comboBox2.Text.Trim()))
                            {
                                //  if (linha22.Contains("\n"))
                                // {
                                listBox1.Items.Add(linha222);
                                lista_++;
                                label27.Text = lista_.ToString();
                                label67.Text = lista_.ToString();
                                // }



                            }
                        }
                        if (lista_ == 0)
                        {
                            label27.Text = lista_.ToString();
                            label67.Text = lista_.ToString();
                            //label67.Text = label27.Text;
                        }
                    }
                }



                // string FileToRead = @"C:\data_txt\data.txt";
                string FileToRead = @"C:\compartilhamento\data_txt\data.txt";
                TextReader Leitor = new StreamReader(@"C:\compartilhamento\data_txt\data.txt", true);//Inicializa o Leitor
                int Linhas = 0;

                while (Leitor.Peek() != -1)
                {//Enquanto o arquivo não acabar, o Peek não retorna -1 sendo adequando para o loop while...
                    Linhas++;//Incrementa 1 na contagem
                    Leitor.ReadLine();//Avança uma linha no arquivo
                }
                Leitor.Close(); //Fecha o Leitor, dando acesso ao arquivo para outros programas....
                int count = File.ReadAllLines(@"C:\compartilhamento\data_txt\data2.txt").Length;
                string number = count.ToString().Trim();

                //   MessageBox.Show(rich5);


                //string criterio = comboBox1.Text.Trim();// id

                // string[] linhas2 = File.ReadAllLines(@"c:\compartilhamento\data_txt\data.txt");

                //  foreach (string linha in linhas2)
                //  {
                // if (linha.Contains(criterio))
                //  lbResultado.Items.Add(linha);
                // label27.Text = Linhas.ToString();
                //  }
                label27.Text = Linhas.ToString();
              //  label67.Text = label27.Text;

                //  string number = System.IO.File.ReadAllText(@"C:\compartilhamento\data_txt\count.txt");
                label3.Text = number;
                //  number2 = Int32.Parse(label3.Text);
                // MessageBox.Show("linha ok");
            }
            catch
            {
                //   MessageBox.Show("não consegui acessar o caminho  " + @"C:\compartilhamento\data_txt\data.txt");
            }

        }

        private void wifi_level()
        {
            try
            {
                System.Diagnostics.Process p = new System.Diagnostics.Process();
                ///  System.Diagnostics.h
                p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                p.StartInfo.CreateNoWindow = true;
                p.StartInfo.FileName = "netsh.exe";
                p.StartInfo.Arguments = "wlan show interfaces";
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.RedirectStandardOutput = true;

                p.Start();

                string s = p.StandardOutput.ReadToEnd();
                string s1 = s.Substring(s.IndexOf("Perfil"));
                s1 = s1.Substring(2, s1.IndexOf(":"));
                s1 = s1.Substring(2, s1.IndexOf("\n")).Trim();

                string s2 = s.Substring(s.IndexOf("Sinal"));
                s2 = s2.Substring(s2.IndexOf(":"));
                s2 = s2.Substring(2, s2.IndexOf("\n")).Trim();
                string s3 = s2.Replace("%", "");

                s4 = Int32.Parse(s3);
                label20.Text = s3;

                p.WaitForExit();

                if (s4 <= 10 || s4 == null)
                {
                    button6.Image = global::criptoqrcode.Properties.Resources.wifi2_level0_1;
                }
                if (s4 <= 20)
                {
                    button6.Image = global::criptoqrcode.Properties.Resources.wifi2_level1_1;
                }
                if (s4 <= 50)
                {
                    button6.Image = global::criptoqrcode.Properties.Resources.wifi2_level2_1;
                }
                if (s4 >= 100)
                {
                    button6.Image = global::criptoqrcode.Properties.Resources.wifi2_level31;
                }
            }
            catch (ArgumentException ex)
            {

                button6.Image = global::criptoqrcode.Properties.Resources.wifi2_level0_1;
            }

        }
        private void ler_linha_projeto()
        {

            // string FileToRead = @"C:\data_txt\data.txt";
            string[] arquivo = File.ReadAllLines(@"C:\compartilhamento\data_txt\PROJETO.txt");
            // TextReader Leitor = new StreamReader(@"C:\data_txt\ROJETO.txt", true);//Inicializa o Leitor


            // comboBox1.Text = arquivo[0].ToString();
            richTextBox9.Text = arquivo[1].ToString();
            // Leitor.Close(); //Fecha o Leitor, dando acesso ao arquivo para outros programas....


        }
        private void ler_data_txt()
        {
            using (StreamReader sr = new StreamReader("data_time.txt"))
            {
                string line;

                while ((line = sr.ReadLine()) != null)
                {
                    z = line;
                }
            }
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            client.Disconnect();
        }
        void client_MqttMsgPublishReceived(object sender, MqttMsgPublishEventArgs e)
        {
            string ReceivedMessage = Encoding.UTF8.GetString(e.Message);

            // we need this construction because the receiving code in the library and the UI with textbox run on different threads
            SetText(ReceivedMessage);

        }
        private void SetText(string text)
        {
            // we need this construction because the receiving code in the library and the UI with textbox run on different threads
            if (this.RecText.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(SetText);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                this.RecText.Text = text;
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {

            ProcessStartInfo startInfo2 = new ProcessStartInfo("cmd", "/K netsh interface set interface \"Ethernet\" ENABLED");
            startInfo2.RedirectStandardOutput = true;
            startInfo2.UseShellExecute = false;
            // Do not create the black window.
            startInfo2.CreateNoWindow = true;
            startInfo2.WindowStyle = ProcessWindowStyle.Hidden;
            Process.Start(startInfo2);

            //  client = new MqttClient("stellar-veterinary.cloudmqtt.com", 1883, false, null, null, MqttSslProtocols.TLSv1_2);

            //  client.MqttMsgPublishReceived += client_MqttMsgPublishReceived;
            //  clientId = Guid.NewGuid().ToString();
            //  client.Connect(clientId, "suport1", "12345");

            //   client.Subscribe(new string[] { "suport1" }, new byte[] { 0 });   // we need arrays as parameters because we can subscribe to different topics with one call
            //   SetText("");


            StreamReader sr = new StreamReader(@"C:\compartilhamento\IP_NEW.txt");
            string x = sr.ReadToEnd();
            sr.Close();

            IP_START.Text = x.Split(',')[0].Trim();
            IP_STOP.Text = x.Split(',')[1].Trim();

            _ipstart = x.Split(',')[0].Trim();
            _ipstop = x.Split(',')[1].Trim();

            //  MessageBox.Show(IP_START.Text);


            label44.Visible = false;
            label45.Visible = false;
            label46.Visible = false;
            label47.Visible = false;
            label48.Visible = false;
            label49.Visible = false;
            label50.Visible = false;
    
            passall = File.ReadAllText(@"C:\compartilhamento\pass\pass.txt");
            //  rede = System.IO.File.ReadAllText(@"C:\compartilhamento\rede.txt");
            // rede1 = File.ReadLines(@"C:\compartilhamento\rede.txt").ElementAtOrDefault(0);
            // rede2 = File.ReadLines(@"C:\compartilhamento\rede.txt").ElementAtOrDefault(1);
            // rede3 = File.ReadLines(@"C:\compartilhamento\rede.txt").ElementAtOrDefault(2);
            // rede = System.IO.File.ReadAllText(@"C:\compartilhamento\rede.txt");



            fproj = File.GetLastWriteTime(@"C:\compartilhamento\data_txt\PROJETO.txt");
            FileIOPermission f2 = new FileIOPermission(FileIOPermissionAccess.AllAccess, "C:\\compartilhamento\\");






            int count = File.ReadAllLines(@"C:\compartilhamento\data_txt\data2.txt").Length;
            string number = count.ToString().Trim();//System.IO.File.ReadAllText(@"C:\compartilhamento\data_txt\count.txt");
            label3.Text = number;
            number2 = Int32.Parse(label3.Text);
            //string constr = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=YES;IMEX=1;""", "C:\\data_base\\" + label18.Text);
            try
            {
                wifi_level();


                textBox6.Select();
                wifi = new Wifi();
                List<AccessPoint> aps = wifi.GetAccessPoints();
                foreach (AccessPoint ap in aps)
                {
                    ListViewItem lvItem = new ListViewItem(ap.Name);
                    lvItem.SubItems.Add(ap.SignalStrength + "%");
                    lvItem.Tag = ap;
                    lview_AP.Items.Add(lvItem);
                }
            }
            catch
            {

            }
            ler_linha_projeto();
            ler_linha();
            ler_data_txt();
            filterInfoCollection = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            foreach (FilterInfo Device in filterInfoCollection)
                cboCamera.Items.Add(Device.Name);
            cboCamera.SelectedIndex = 0;
            videoCaptureDevice = new VideoCaptureDevice();
            richTextBox1.Text = "";
            richTextBox2.Text = "";
            richTextBox3.Text = "";
            richTextBox4.Text = "";
            //richTextBox5.Text = "";
            richTextBox6.Text = "";
            richTextBox7.Text = "";
            richTextBox8.Text = "";

            lcompany.Text = nome;
            lname.Text = emp;
            lfunc.Text = function;
            lid.Text = id;
            lvessel.Text = vessel;
            lcheckin.Text = Label_initial[1];
            lcheckout.Text = Label_final[1];
            // button7.Visible = false;

            dateTimePicker1.Visible = true;
            dateTimePicker2.Visible = true;
            // textBox2.Text = dateTimePicker1.Value.ToString();
            richTextBox6.Visible = false;
            richTextBox7.Visible = false;
            filterInfoCollection = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            foreach (FilterInfo Device in filterInfoCollection)
                cboCamera.Items.Add(Device.Name);
            cboCamera.SelectedIndex = 0;
            videoCaptureDevice = new VideoCaptureDevice();



            label18.Text = GetFiles(@"C:\compartilhamento\data_base");

            inicialize_qrreader();

            lname.Text = label_nome[0];
            lcompany.Text = label_emp[0];
            lfunc.Text = label_Function[0];
            lid.Text = label_Id[0];
            lvessel.Text = label_vessel[0];
            lproject.Text = label_porj[0];
            lv1.Text = label_vaccine1[0];
            lv2.Text = label_vaccine2[0];
            lbustter.Text = label_reforco[0];
            local1.Text = place1[0];
            local2.Text = place2[0];
            local3.Text = place3[0];
            local4.Text = place4[0];
            laccess.Text = label_acc[0];
            level_yellow.Text = label_yellow[0];
            level_green.Text = label_green[0];
            level_red.Text = label_red[0];
            lcheckin.Text = Label_initial[0];
            lcheckout.Text = Label_final[0];
            button1.Text = Label_Read_QRcode_On[0];
            button28.Text = Label_Read_QRcode_Off[0];
            button2.Text = Label_Create_QRcode[0];
            button3.Text = Label_Show_data[0];
            button4.Text = Label_Save_data[0];
            // button5.Text = Label_Config[0];
            button6.Text = Label_wifi[0];
            button17.Text = Label_email[0];
            button19.Text = Label_Mostrar_checkin[0];
            button21.Text = Label_reset[0];
            button22.Text = Label_fechar[0];
            button8.Text = Label_entrada[0];
            button9.Text = Label_saida[0];
            button10.Text = Label_cancel[0];
            label23.Text = label_onboard[0];
            button27.Text = Label_reset_project[0];
            button29.Text = label_reg[0];
            label7.Text = button1.Text;
            label6.Text = onboard[0];
            label_cad.Text = label_cad1[0];

            label53.Text = Label_53[0];
            button41.Text = bt_41[0];
            button42.Text = bt_42[0];
            button43.Text = bt_43[0];
            button44.Text = bt_44[0];
            button45.Text = bt_45[0];
            regs.Text = bt_regis[0];
            _cad = cad_mode[0];
            _read = read_mode[0];

            /*
             Label_reset
                        String[] Label_initial = { "Inicio", "Check-in" };
                        String[] Label_final = { "Fim", "Check-out" };
                        String[] Label_Read_QRcode_On = { "Ler Qrcode Ligado", "Read QRcode On" };
                        String[] Label_Read_QRcode_Off = { "Ler Qrcode Desligado", "Read QRcode Off" };
                        String[] Label_Create_QRcode = { "Imprimir Qrcode:", "Print QRcode" };
                        String[] Label_Show_data = { "Mostrar banco de dados:", "Show DataBase" };
                        String[] Label_Save_data = { "Salvar banco de dados:", "Save Database Backup" };
                        String[] Label_Config = { "Configurações:", "Settings" };
                        String[] Label_wifi = { "Conexão Wi-Fi:", "Wi-Fi connection" };
                        String[] Label_email = { "Enviar Qrcode por E-mail:", "Send Qr Code  by E-mail" };
                        String[] Label_Mostrar_checkin = { "Mostrar Check-in:", "Show Check-in" };
                        String[] Label_fechar = { "Desligar:", "Turn Off" };
            */
            panel11.Size = new Size(460, 455);
            //  panel11.Location = new System.Drawing.Point(95, 150);
            pictureBox7.Size = new Size(430, 420);
            pictureBox7.Location = new System.Drawing.Point(15, 18);
            ip = System.IO.File.ReadAllText(@"C:\compartilhamento\IP.txt");
            ip1 = File.ReadLines(@"C:\compartilhamento\IP.txt").ElementAtOrDefault(0);
            ip2 = File.ReadLines(@"C:\compartilhamento\IP.txt").ElementAtOrDefault(1);
            ip3 = File.ReadLines(@"C:\compartilhamento\IP.txt").ElementAtOrDefault(2);

            textBox7.Focus();
            textBox7.Text = "";

            GetBiosInformation();
            // CloseExcel();
           // myThread.Abort();
          //  checa_host();
            // atualiza_compartilhamento();
            comp = 1;


            try
            {
                // open file dialog
                OpenFileDialog open = new OpenFileDialog();
                // image filters
                //  open.Filter = "Image Files(*.jpg; *.jpeg; *.gif; *.bmp)|*.jpg; *.jpeg; *.gif; *.bmp";
                String foto = @"C:\compartilhamento\logo_criptoqrcode\logo.png";
                //  if (open.ShowDialog() == DialogResult.OK)
                // {
                // display image in picture box
                pictureBox2.Image = new Bitmap(foto);
                // image file path
                // textBox1.Text = open.FileName;
                // }
                StreamReader sr2 = new StreamReader(@"C:\compartilhamento\vessels.txt");
                string x2 = sr2.ReadToEnd();
                string[] y = x2.Split('\n');
                foreach (string s in y)
                {
                    comboBox1.Items.Add(s);
                    comboBox2.Items.Add(s);
                }
                sr.Close();
            }
            catch
            {
                MessageBox.Show("O arquivo com a logo deve estar no formato PNG. Corriga o formato e inicie o programa novamente");
                // throw new ApplicationException("Image loading error....");
            }

    
            int milliseconds = 5000;
            Thread.Sleep(milliseconds);

            MyhostName = System.Net.Dns.GetHostName();
            string hostName = System.Net.Dns.GetHostName();
            string myIP = Dns.GetHostByName(hostName).AddressList[0].ToString();
            label66.Text = "IP: " + myIP;
            timer10.Enabled = true;
        }


        Thread myThread = null;
        int plant = 0;


        public void checa_host()
        {

            if (IP_START.Text == string.Empty)
            {
                //MessageBox.Show("No IP address entered.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // lblStatus.ForeColor = System.Drawing.Color.Red;
                // lblStatus.Text = "No IP address entered.";
            }
            else
            {

                //Create new thread for pinging
                //myThread = new Thread(() => scan(txtIP.Text));
                myThread = new Thread(() => scan2(IP_START.Text, IP_STOP.Text));
                myThread.Start();

                if (myThread.IsAlive == true)
                {
                    // cmdStop.Enabled = true;
                    //  cmdScan.Enabled = false;
                    // txtIP.Enabled = false;
                    // txtIP2.Enabled = false;
                    //  MessageBox.Show("checa_host ok");
                }
            }

        }
        static void SetDoubleBuffer(System.Windows.Forms.Control dgv, bool DoubleBuffered)
        {
            typeof(System.Windows.Forms.Control).InvokeMember("DoubleBuffered",
                BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty,
                null, null, new object[] { DoubleBuffered });
        }
        private async void button3_Click(object sender, EventArgs e)
        {
           
            try
            {
                String[] Label_initial = { "Inicio", "Check-in" };
                String[] Label_final = { "Fim", "Check-out" };
                String[] Label_Read_QRcode_On = { "Ler Qrcode Ligado", "Read QRcode On" };
                String[] Label_Read_QRcode_Off = { "Ler Qrcode Desligado", "Read QRcode Off" };
                String[] Label_Create_QRcode = { "Imprimir Qrcode:", "Print QRcode" };
                String[] Label_Show_data = { "Mostrar banco de dados:", "Show DataBase" };
                String[] Label_Save_data = { "Salvar banco de dados:", "Save Database Backup" };
                String[] Label_Config = { "Configurações:", "Settings" };
                String[] Label_wifi = { "Conexão Wi-Fi:", "Wi-Fi connection" };
                String[] Label_email = { "Enviar Qrcode por E-mail:", "Send Qr Code  by E-mail" };
                String[] Label_Mostrar_checkin = { "Mostrar Check-in:", "Show Check-in" };
                String[] Label_fechar = { "Desligar:", "Turn Off" };



                criar_excel();

                if (textBox7.SelectionLength >= 0)
                {
                    textBox7.Focus();
                    textBox7.Text = "";
                }
                plant++;
                if (plant == 1)
                {
                    criar_excel();
                    //if (libera == true)
                    // {
                    button34.Visible = true;
                    // }
                    // else
                    // {
                    //  button34.Visible = false;
                    // }

                    if (band == 0)
                    {
                        button3.Text = Label_close_data[0];
                    }
                    else
                    {
                        button3.Text = Label_close_data[1];
                    }
                    //  button3.Text = "Close DataBase";
                    //  button1.Enabled = false;
                    //button2.Enabled = false;
                    String name = "Planilha1";



                    string constr = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=YES;IMEX=1;""", "C:\\compartilhamento\\data_base\\" + label18.Text);

                    OleDbConnection con = new OleDbConnection(constr);
                    OleDbCommand oconn = new OleDbCommand("Select * From [" + name + "$]", con);
                    con.Open();

                    OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                    System.Data.DataTable data = new System.Data.DataTable();
                    sda.Fill(data);
                    con.Close();


                    dataGridView1.Size = new System.Drawing.Size(1693, 800);
                    dataGridView1.DataSource = data;
                    this.dataGridView1.DefaultCellStyle.ForeColor = Color.White;
                    this.dataGridView1.DefaultCellStyle.BackColor = Color.FromArgb(51, 51, 76);
                    dataGridView1.Visible = true;
                    dataGridView1.FilterAndSortEnabled = true;
                    bindingSource1.DataSource = data;
                    panel5.Visible = false;

                }
                else
                {
                    dataGridView1.Visible = false;
                    dataGridView1.Visible = false;
                    button34.Visible = false;
                    panel5.Visible = true;

                    plant = 0;

                    if (band == 0)
                    {
                        button3.Text = Label_Show_data[0];
                    }
                    else
                    {
                        button3.Text = Label_Show_data[1];
                    }




                    button1.Enabled = true;

                }
            }
            catch
            {
                MessageBox.Show("data base corrompida ou não encontrada. feche a planilha se estiver aberta e inicie o programa novamente!");
            }

        }
        private void acha_palavra_txt()
        {
            if (l == 0)
            {

                lbResultado.Items.Clear();
                Refresh();

                string criterio = richTextBox4.Text;// id

                string[] linhas = File.ReadAllLines(@"c:\compartilhamento\data_txt\data.txt");

                foreach (string linha in linhas)
                {
                    if (linha.Contains(criterio))
                        lbResultado.Items.Add(linha);
                }



                if (lbResultado.Items.Count == 1)
                {
                    button8.Enabled = false;
                    button9.Enabled = true;

                }
                if (lbResultado.Items.Count == 0)
                {
                    button8.Enabled = true;
                    button9.Enabled = false;

                }

            }
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            
            BarcodeReader Reader = new BarcodeReader();
            Result result = Reader.Decode((Bitmap)pictureBox1.Image);
            if (result != null)
            {

                try
                {


                    // timer2.Enabled = false;





                    txtResult.Text = result.ToString();

                    if (txtResult.Text.Length > 10)
                    {


                        string text1 = txtResult.Lines[0];
                        string mystring = text1.Remove(0, 6);
                        richTextBox1.Text = mystring;

                        string text2 = txtResult.Lines[1];
                        mystring = text2.Remove(0, 9);
                        richTextBox2.Text = mystring;

                        string text3 = txtResult.Lines[2];
                        mystring = text3.Remove(0, 10);
                        richTextBox3.Text = mystring;

                        string text4 = txtResult.Lines[3];
                        mystring = text4.Remove(0, 4);
                        richTextBox4.Text = mystring;

                        string text5 = txtResult.Lines[4];
                        mystring = text5.Remove(0, 8);
                        comboBox1.Text = mystring;

                        string text8 = txtResult.Lines[5];
                        mystring = text8.Remove(0, 1);
                        richTextBox8.Text = mystring;

                        string text6 = txtResult.Lines[6];
                        mystring = text6.Remove(0, 9);
                        richTextBox6.Text = mystring;

                        string text7 = txtResult.Lines[7];
                        mystring = text7.Remove(0, 7);
                        richTextBox7.Text = mystring;

                        videoCaptureDevice.Stop();
                        //  pictureBox1.BackgroundImage= Image.
                        pictureBox7.BackgroundImage = Properties.Resources.barcode1;
                        button8.Visible = true;
                        button9.Visible = true;
                        button10.Visible = true;






                        compare_data();

                        acha_palavra_txt();

                        timer1.Stop();


                    }
                }
                catch
                {

                }
            }
        }
        int input_ok = 0;
        private void zera()
        {
            if (input_ok == 0)
            {
                videoCaptureDevice = new VideoCaptureDevice(filterInfoCollection[cboCamera.SelectedIndex].MonikerString);
                videoCaptureDevice.NewFrame += FinalFrame_NewFrame;
                videoCaptureDevice.Start();
                timer1.Start();

                richTextBox1.Text = "";
                richTextBox2.Text = "";
                richTextBox3.Text = "";
                richTextBox4.Text = "";
                // richTextBox5.Text = "";
                richTextBox6.Text = "";
                richTextBox7.Text = "";
                richTextBox8.Text = "";
                button8.Visible = false;
                button9.Visible = false;
                button10.Visible = false;
            }
        }
        private void cancelar()
        {
            // panel11.Visible = false;
            // panel4.Visible = true;

            /// maskedTextBox1.Text = " ";
            /// maskedTextBox2.Text = " ";
            ///maskedTextBox3.Text = " ";
            ///maskedTextBox4.Text = " ";
            ///maskedTextBox5.Text = " ";

            panel11.Visible = true;
            pictureBox7.Image = Properties.Resources.barcode1;
            panel11.BackColor = Color.White;
            label8.Visible = false;
            //pictureBox1.Image = Properties.Resources.frame;
            input_ok = 0;

            //  pictureBox1.Image = null;
            /// pictureBox1.Image = Modern_Qr_code.Properties.Resources.

            // panel4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(76)))));
            videoCaptureDevice = new VideoCaptureDevice(filterInfoCollection[cboCamera.SelectedIndex].MonikerString);
            //videoCaptureDevice.NewFrame += FinalFrame_NewFrame;
            //videoCaptureDevice.Start();
            // timer1.Start();
            if (textBox7.SelectionLength >= 0)
            {
                textBox7.Focus();
                textBox7.Text = "";
            }
            richTextBox1.Text = "";
            richTextBox2.Text = "";
            richTextBox3.Text = "";
            richTextBox4.Text = "";
            //richTextBox5.Text = "";
            richTextBox6.Text = "";
            richTextBox7.Text = "";
            richTextBox8.Text = "";

            // richTextBox9.Text = "";
            richTextBox10.Text = "";
            richTextBox11.Text = "";
            richTextBox12.Text = "";
            richTextBox13.Text = "";
            richTextBox14.Text = "";
            richTextBox16.Text = "";
            button8.Visible = false;
            button9.Visible = false;
            button10.Visible = false;
            pictureBox1.Image = criptoqrcode.Properties.Resources.barcode1;
        }
        private void timer2_Tick(object sender, EventArgs e)
        {
            cancelar();
            ckecked_false();
            timer2.Stop();
        }
        int lista_ = 0;
        private void mostra_conteudo_txt()
        {
            listBox1.Items.Clear();
            lista_ = 0;

            if (checkBox1.Checked == true)
            {
                panel6.Size = new Size(1695, 996); //1695; 996
                panel6.Visible = true;
                string[] linhas22 = System.IO.File.ReadAllLines(@"C:\compartilhamento\data_txt\data.txt");

                foreach (string linha22 in linhas22)
                {
                    if (linha22.Contains(comboBox1.Text.Trim()))
                    {
                        listBox1.Items.Add(linha22);
                        lista_++;
                        label27.Text = lista_.ToString();
                      //  label67.Text = label27.Text;


                    }
                }
              //  ler_linha();
            }
            else
            {

                if (new FileInfo(@"C:\compartilhamento\data_txt\data.txt").Length >= 0)
                {

                    panel6.Size = new Size(1695, 996);
                    panel6.Visible = true;
                    string text = System.IO.File.ReadAllText(@"C:\compartilhamento\data_txt\data.txt");

                    foreach (string s in Regex.Split(text, "\n"))
                    {

                        listBox1.Items.Add(s);



                    }

                  //  ler_linha();
                }
            }


        }
        private void beep()
        {

            int frequency = 800;

            // Set the Duration
            int duration = 250;

            // Play beep sound once
            Console.Beep(frequency, duration);
        }
        private void timer3_Tick(object sender, EventArgs e)
        {
            try
            {
                string caminhoArquivo = @"C:\compartilhamento\data_txt\data4.txt";
                string primeiraLinha = File.ReadLines(caminhoArquivo).First().Trim();
                if (primeiraLinha == "1")
                {
                    //  panel12.BackColor = Color.GreenYellow;
                   
                    
                    //ler_linha();
                    Console.WriteLine($"A primeira linha do arquivo é: {primeiraLinha}");

                    using (StreamWriter sw = new StreamWriter(caminhoArquivo))
                    {
                        sw.WriteLine("0"); // Escreve o número 1 na primeira linha do arquivo
                    }
                }
            }

            catch
            {

            }



            if (online == true)
            {
                wifi_level();
            }

            string datetime = DateTime.Now.ToString();
            label19.Text = DateTime.Now.ToString();
            label62.Text = DateTime.Now.ToString();
            DateTime now = DateTime.Now;
            int h = now.Hour;
            int m = now.Minute;
            int s = now.Second;

            label22.Text = h + ":" + m + ":" + s.ToString().Trim();
            //label22.Text = "11:27:00";

            if (label22.Text == z)
            {
                //
                //
                //
                //
                // ();
                //listBox1.Text = varPalavra;
                mostra_conteudo_txt();
                beep();
                beep();
                beep();
                beep();
                beep();
            }
        }



        private void escrever_palavra()

        {
            string nomeArquivo = @"C:\compartilhamento\data_txt\data.txt";


            // Name: Henrique Kaique Costa Compay: Local 9  E - Mail: cristiano.engenharia.ac @gmail.com ID: 111099444
            string textoInserir = "Number: " + richTextBox16.Text + ":  Name: " + richTextBox2.Text + ":  Vessel: " + comboBox1.Text.Trim() + ":  Compay: " + richTextBox1.Text + "  Id: " + richTextBox4.Text + "  :E-Mail: " + richTextBox8.Text+" :  "+DateTime.Now;//
            int numeroLinha = Convert.ToInt32(Linhas);
            ArrayList linhas = new ArrayList();

            if (File.Exists(nomeArquivo))
            {
                try
                {
                    rdr = new StreamReader(nomeArquivo);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao acessar o arquivo : " + ex.Message);
                    return;
                }
            }
            else
            {
                MessageBox.Show("O arquivo : " + nomeArquivo + " não existe...");
                return;
            }

            string linha;

            while ((linha = rdr.ReadLine()) != null)
            {
                try
                {
                    linhas.Add(linha);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao acessar o arquivo : " + ex.Message);
                    return;
                }
            }
            rdr.Close();

            if (linhas.Count > numeroLinha)
                linhas.Insert(numeroLinha, textoInserir);
            else
                linhas.Add(textoInserir);

            StreamWriter wrtr = new StreamWriter(nomeArquivo);

            foreach (string strNewLine in linhas)
            {
                wrtr.WriteLine(strNewLine);
            }
            wrtr.Close();
            textoInserir = "";
            // txtArquivo.Text = AbreArquivoTexto(nomeArquivo);

        }

        string ll;
        private void apaga_palavra_txt()
        {
            if (l == 0)
            {
                textBox7.Focus();
                textBox7.Text = "";
                string tempFile = Path.GetTempFileName();

                using (var sr = new StreamReader(@"C:\compartilhamento\data_txt\data.txt"))
                {
                    using (var sw = new StreamWriter(tempFile))
                    {
                        string line;

                        while ((line = sr.ReadLine()) != null)
                        {

                            //Name: Rodrigo  Compay: Googlemarine  Id: 111222333
                          
                            ll = RemoveEnd(line,23);
                           //  MessageBox.Show(ll);

                            // Number: 10  Name: Cristiano de Araujo Calheiros  Compay: Googlemarine Id: 111098414  :E - Mail: cristiano.engenharia.ac @gmail.com
                            if (ll != "Number: " + richTextBox16.Text + ":  Name: " + richTextBox2.Text + ":  Vessel: " + comboBox1.Text.Trim() + ":  Compay: " + richTextBox1.Text + "  Id: " + richTextBox4.Text + "  :E-Mail: " + richTextBox8.Text)
                               sw.WriteLine(line);
                        }
                    }
                }
                File.Copy(tempFile, @"C:\compartilhamento\data_txt\data.txt", true);

                // File.Delete(@"C:\compartilhamento\data_txt\data.txt");
                // File.Move(tempFile, @"C:\compartilhamento\data_txt\data.txt");
            }
        }

        private void in_out_alt()
        {


            lbResultado.Items.Clear();
            Refresh();
            string criterio = richTextBox4.Text;// id
            string[] linhas = File.ReadAllLines(@"c:\compartilhamento\data_txt\data.txt");

            foreach (string linha in linhas)
            {
                if (linha.Contains(criterio))
                    lbResultado.Items.Add(linha);
            }



            if (lbResultado.Items.Count == 1)
            {

                // MessageBox.Show("existe");
                try
                {

                    comp = 0;
                  //   ler_linha();
                    // 
                    //timer8.Stop();
                    pictureBox7.Image = Properties.Resources.barcode1;
                    panel11.BackColor = Color.White;
                    label8.Visible = false;



                    if (textBox7.SelectionLength >= 0)
                    {
                        textBox7.Focus();
                        textBox7.Text = "";
                    }
                    listBox1.Items.Clear();
                    input_ok = 0;
                    output_data = DateTime.Now.ToString("MM/dd/yyyy").Trim();
                    output_hora = DateTime.Now.ToString("HH:mm:ss tt").Trim();
                    input_data = "*";
                    input_hora = "*";
                    apaga_palavra_txt();

                    panel4.BackColor = Color.White;
                    videoCaptureDevice = new VideoCaptureDevice(filterInfoCollection[cboCamera.SelectedIndex].MonikerString);
                    videoCaptureDevice.NewFrame += FinalFrame_NewFrame;
                    // videoCaptureDevice.Start();
                    // timer1.Start();
                    pictureBox7.Image = Properties.Resources.barcode1;
                    //  CarregarPlanilha();
                    carrega_planilha_txt();
                    button8.Visible = false;
                    button9.Visible = false;
                    button10.Visible = false;
                     checa_host();
                    // atualiza_compartilhamento();
                   // ler_linha();
                    // ckecked_false();

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                comp = 1;
            }
            if (lbResultado.Items.Count == 0)
            {
                // MessageBox.Show("não existe");
                try
                {
                    if (checado == 1)
                    {
                        comp = 0;
                        // ler_linha();
                        // ckecked_false();
                        // timer8.Stop();

                        pictureBox7.Image = Properties.Resources.barcode1;
                        panel11.BackColor = Color.White;
                        label8.Visible = false;




                        if (textBox7.SelectionLength >= 0)
                        {
                            textBox7.Focus();
                            textBox7.Text = "";
                        }

                        input_data = DateTime.Now.ToString("MM/dd/yyyy").Trim();
                        pictureBox7.Image = Properties.Resources.barcode1;
                        input_ok = 0;

                        input_hora = DateTime.Now.ToString("HH:mm:ss tt").Trim();
                        output_data = "*";
                        output_hora = "*";
                        panel4.BackColor = Color.White;
                        videoCaptureDevice = new VideoCaptureDevice(filterInfoCollection[cboCamera.SelectedIndex].MonikerString);
                        videoCaptureDevice.NewFrame += FinalFrame_NewFrame;

                        acha_palavra_txt();
                        escrever_palavra();




                        button8.Visible = false;
                        button9.Visible = false;
                        button10.Visible = false;

                        ler_linha();
                        //  ckecked_false();
                        // CarregarPlanilha();
                        carrega_planilha_txt();
                        // checa_host();
                        //atualiza_compartilhamento();
                        comp = 1;
                    }
                    else
                    {
                        MessageBox.Show("PLease check local or level");
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }


            }

            /////////////////////////////////////////// in/////////////////////////////////////////

            ///////////////////////////////////////////////////////
            ///
            /// 
            /// 
            /////////////////////////////out///////////////////////

            ////////////////////////////////////////////////////////////////

        }
        private void compare_data()
        {
           
            if (id_onboard2 == true)
            {
               
             
            }
            else
            {

                var parameterDate_ASo = DateTime.ParseExact(richTextBox10.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                var parameterDate_initial = DateTime.ParseExact(richTextBox6.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                var parameterDate_final = DateTime.ParseExact(richTextBox7.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                var todaysDate = DateTime.Today;



                if (rich5 != comboBox1.Text.Trim() || rich9 != richTextBox9.Text.Trim())
                {
                    //  richTextBox5.Text = text10[7].Remove(0, 8);
                    //  richTextBox9.Text = text10[8].Remove(0, 0);

                    panel11.BackColor = Color.Red;
                    beep();
                    beep();
                    beep();
                    beep();

                    button8.Visible = false;
                    button9.Visible = false;
                    button10.Visible = true;
                    input_ok = 0;
                    // label8.ForeColor = Color.Red;
                    label8.Visible = true;
                    label8.Text = "Embarcação " + rich5 + " ou Projeto divergente";
                    try
                    {

                        pictureBox7.LoadAsync(@"C:\compartilhamento\data_picture\" + richTextBox4.Text.ToString() + ".jpg");

                    }
                    catch
                    {
                        pictureBox7.LoadAsync(@"C:\compartilhamento\data_picture\face.jpg");
                    }




                    timer2.Start();
                    //  DateTime now = DateTime.Now;
                    //  while (DateTime.Now.Subtract(now).Seconds < 3)
                    // {
                    // wait for 60 seconds.
                    //  }
                    // cancelar_teste();
                    //  Thread.Sleep(3000);
                }



                if (company_loc == true)
                {

                    panel11.BackColor = Color.Red;
                    beep();
                    beep();
                    beep();
                    beep();

                    button8.Visible = false;
                    button9.Visible = false;
                    button10.Visible = true;
                    input_ok = 0;
                    label8.Visible = true;
                    label8.Text = "BLOQUEADO: CARTÃO INUTILIZADO !";

                    try
                    {
                        pictureBox7.LoadAsync(@"C:\compartilhamento\data_picture\" + richTextBox4.Text.ToString() + ".jpg");

                    }
                    catch
                    {
                        pictureBox7.LoadAsync(@"C:\compartilhamento\data_picture\face.jpg");

                    }


                    timer2.Start();
                }

                if (bb == "Bloqueado")
                {
                    // pictureBox7.Image 
                    //pictureBox7.Image 

                    panel11.BackColor = Color.Red;
                    beep();
                    beep();
                    beep();
                    beep();
                    button8.Visible = false;
                    button9.Visible = false;
                    button10.Visible = true;
                    input_ok = 0;
                    //label8.ForeColor = Color.Red;
                    label8.Visible = true;
                    button7.Visible = false;
                    btloc.Visible = true;
                    label8.Text = "BLOQUEADO : " + label37.Text;

                    try
                    {
                        pictureBox7.LoadAsync(@"C:\compartilhamento\data_picture\" + richTextBox4.Text.ToString() + ".jpg");

                    }
                    catch
                    {
                        pictureBox7.LoadAsync(@"C:\compartilhamento\data_picture\face.jpg");

                    }
                    DateTime now = DateTime.Now;
                    while (DateTime.Now.Subtract(now).Seconds < 5)
                    {
                        // wait for 60 seconds.
                    }
                    timer2.Start();
                }

                if (parameterDate_initial > parameterDate_ASo)
                {

                    panel11.BackColor = Color.Red;
                    beep();
                    beep();
                    beep();
                    beep();

                    button8.Visible = false;
                    button9.Visible = false;
                    button10.Visible = true;
                    input_ok = 0;
                    //label8.ForeColor = Color.Red;
                    label8.Visible = true;
                    label8.Text = "Aso vencido";

                    try
                    {
                        pictureBox7.LoadAsync(@"C:\compartilhamento\data_picture\" + richTextBox4.Text.ToString() + ".jpg");

                    }
                    catch
                    {
                        pictureBox7.LoadAsync(@"C:\compartilhamento\data_picture\face.jpg");

                    }
                    timer2.Start();
                }






                if ((parameterDate_initial < todaysDate) && (parameterDate_final < todaysDate))
                {

                    label8.Visible = true;
                    label8.Text = "Data limite expirou";
                    panel11.BackColor = Color.Red;
                    beep();
                    beep();
                    beep();
                    beep();

                    button8.Visible = false;
                    button9.Visible = false;
                    button10.Visible = true;
                    input_ok = 0;

                    try
                    {
                        pictureBox7.LoadAsync(@"C:\compartilhamento\data_picture\" + richTextBox4.Text.ToString() + ".jpg");

                    }
                    catch
                    {
                        pictureBox7.Load(@"C:\compartilhamento\data_picture\face.jpg");
                    }
                    timer2.Start();
                }


                if ((parameterDate_initial > todaysDate) && (parameterDate_final > todaysDate))
                {

                    label8.Visible = true;
                    label8.Text = "Data limite expirou";
                    panel11.BackColor = Color.Red;
                    beep();
                    beep();
                    beep();
                    beep();

                    button8.Visible = false;
                    button9.Visible = false;
                    button10.Visible = true;
                    input_ok = 0;

                    try
                    {
                        pictureBox7.LoadAsync(@"C:\compartilhamento\data_picture\" + richTextBox4.Text.ToString() + ".jpg");

                    }
                    catch
                    {
                        pictureBox7.Load(@"C:\compartilhamento\data_picture\face.jpg");
                    }
                    timer2.Start();
                }
                if (parameterDate_initial > parameterDate_final)
                {

                    label8.Visible = true;
                    label8.Text = "Data limite expirou";
                    panel11.BackColor = Color.Red;
                    beep();
                    beep();
                    beep();
                    beep();

                    button8.Visible = false;
                    button9.Visible = false;
                    button10.Visible = true;
                    input_ok = 0;

                    try
                    {
                        pictureBox7.LoadAsync(@"C:\compartilhamento\data_picture\" + richTextBox4.Text.ToString() + ".jpg");

                    }
                    catch
                    {
                        pictureBox7.Load(@"C:\compartilhamento\data_picture\face.jpg");
                    }
                    timer2.Start();
                }


                if (parameterDate_final == todaysDate && parameterDate_initial < parameterDate_ASo)
                {





                    if (rich5 == comboBox1.Text.Trim() && rich9 == richTextBox9.Text.Trim() && company_loc == false && bb != "Bloqueado")
                    {

                        entrou = true;
                        label8.Visible = true;
                        label8.Text = "Liberado";
                        input_ok = 1;
                        panel11.BackColor = Color.GreenYellow;
                        beep();
                        button8.Visible = true;
                        button9.Visible = true;
                        button10.Visible = true;
                        button7.Visible = true;
                        btloc.Visible = false;
                        try
                        {
                            pictureBox7.LoadAsync(@"C:\compartilhamento\data_picture\" + richTextBox4.Text.ToString() + ".jpg");

                            button7.Visible = true;

                        }
                        catch
                        {
                            button7.Visible = true;
                            pictureBox7.LoadAsync(@"C:\compartilhamento\data_picture\face.jpg");

                        }
                        if (panel11.BackColor == Color.GreenYellow)
                        {
                            in_out_alt();
                        }


                    }






                }


                if ((parameterDate_final > todaysDate) && (parameterDate_initial <= todaysDate) && (parameterDate_initial < parameterDate_ASo))
                {

                    if ((rich5 == comboBox1.Text.Trim()) && (rich9 == richTextBox9.Text.Trim()) && company_loc == false && bb != "Bloqueado")
                    {

                        entrou=true;
                        label8.Visible = true;
                        label8.Text = "Liberado";
                        input_ok = 1;
                        panel11.BackColor = Color.GreenYellow;
                        beep();
                        button8.Visible = true;
                        button9.Visible = true;
                        button10.Visible = true;
                        button7.Visible = true;
                        btloc.Visible = false;
                        input_ok = 1;
                        try
                        {
                            pictureBox7.LoadAsync(@"C:\compartilhamento\data_picture\" + richTextBox4.Text.ToString() + ".jpg");
                            button7.Visible = true;
                        }
                        catch
                        {
                            pictureBox7.Load(@"C:\compartilhamento\data_picture\face.jpg");
                            button7.Visible = true;
                        }
                        if (panel11.BackColor == Color.GreenYellow)
                        {
                            in_out_alt();
                        }


                    }


                }

            }
        }

        private void alterado()
        {
            //  richTextBox6.Text = secondLine.Split(':')[31];
            // richTextBox7.Text = secondLine.Split(':')[32];
            try
            {
                int rich1 = Int16.Parse(richTextBox16.Text.Trim()) - 1;
                string secondLine2 = File.ReadLines(@"C:\compartilhamento\data_txt\data2.txt").ElementAtOrDefault(rich1);
                string sec = secondLine2.Split(':')[26].Trim();
                string pass11 = secondLine2.Split(':')[5].Trim();
                string pass12 = secondLine2.Split(':')[9].Trim();
                string pass13 = secondLine2.Split(':')[3].Trim();
                string pass14 = secondLine2.Split(':')[11].Trim();
                string pass15 = secondLine2.Split(':')[17].Trim();
                string pass16 = secondLine2.Split(':')[19].Trim(); // NR34
                string pass17 = secondLine2.Split(':')[21].Trim(); // VA1
                string pass18 = secondLine2.Split(':')[23].Trim(); // V2                                                                                                                                                                   //17
                string pass19 = secondLine2.Split(':')[25].Trim();// V BOOSTER
                string pass20 = secondLine2.Split(':')[31].Trim();// V BOOSTER
                string pass21 = secondLine2.Split(':')[32].Trim();// V BOOSTER
                string pass22 = secondLine2.Split(':')[33].Trim();// V BOOSTER
                                                                  //  0   1      2                       3        4              5         6              7  8           9         10                                  11      12             13        14      15    16           17     18           19          20            21          22           23               24            25 26     27   
                                                                  //Number : 1 : Name : CRISTIANO CALHEIROS 3 : Compay : GOOGLEMARINE :Funcition:  ENGENHEIRO  :Id: 1110988400 : E-mail : cristiano.engenharia.ac@gmail.com : Vessel : Googlemarine : Project : 2001 : ASO : 02/02/2024 : NR-34 : 02/02/2025 : Vaccine-1 : 02/02/2026 : Vaccine-2 : 02/02/2027 : Booster vaccine : 02/02/2028 :  : COMUM :
                if (pass11.Trim() != richTextBox1.Text.Trim())  //COMPANY
                {
                    alter = true;
                }



                if (pass12.Trim() != richTextBox4.Text.Trim())  // IDENTIDADE
                {
                    alter = true;
                }

                if (pass13.Trim() != richTextBox2.Text.Trim()) // NOME
                {
                    alter = true;
                }

                if (pass14.Trim() != richTextBox8.Text.Trim()) // email
                {
                    alter = true;
                }
                if (pass15.Trim() != maskedTextBox1.Text.Trim()) // aso
                {
                    alter = true;
                }
                if (pass16.Trim() != maskedTextBox2.Text.Trim()) // nr34
                {
                    alter = true;
                }
                if (pass17.Trim() != maskedTextBox3.Text.Trim()) // vacina 1
                {
                    alter = true;
                }
                if (pass18.Trim() != maskedTextBox4.Text.Trim()) // vacina 2
                {
                    alter = true;
                }
                if (pass19.Trim() != maskedTextBox5.Text.Trim()) // vacina reforço
                {
                    alter = true;
                }

                if (pass20.Trim() != richTextBox6.Text.Trim()) // vacina reforço
                {
                    alter = true;
                }
                if (pass21.Trim() != richTextBox7.Text.Trim()) // vacina reforço
                {
                    alter = true;
                }
                if (pass22.Trim() != richTextBox15.Text.Trim()) // vacina reforço
                {
                    alter = true;
                }

                if (pass11.Trim() == richTextBox1.Text.Trim() && pass12.Trim() == richTextBox4.Text.Trim() && pass13.Trim() == richTextBox2.Text.Trim())
                {
                    if (pass14.Trim() == richTextBox8.Text.Trim() && pass15.Trim() == maskedTextBox1.Text.Trim() && pass16.Trim() == maskedTextBox2.Text.Trim())
                    {
                        if (pass17.Trim() == maskedTextBox3.Text.Trim() && pass18.Trim() == maskedTextBox4.Text.Trim() && pass19.Trim() == maskedTextBox5.Text.Trim() && pass20.Trim() == richTextBox6.Text.Trim() && pass21.Trim() == richTextBox7.Text.Trim() && pass21.Trim() == richTextBox15.Text.Trim())
                        {
                            alter = false; // 
                        }
                    }
                }
            }
            catch
            {

            }
        }
        int test2 = 0;
        private void timer4_Tick(object sender, EventArgs e)
        { 
        
            if (vid == 1)
            {
                  



                ckecked_false();
                // string cripto = "";
                textBox7.Text.Replace(';', '/');
                string zzz = textBox7.Text.Replace(';', '/');
                textBox8.Text = zzz.Trim();

                try
                {
                    Criptografia criptografia = new Criptografia(CryptProvider.RC2);
                    criptografia.Key = "Etec2017";
                    textBox8.Text = criptografia.Decrypt(zzz);
                    textBox10.Text = textBox8.Text;
                    test2 = 1;
                    string textoqr = zzz;
                }
                catch (Exception ex)
                {
                    textBox7.Text = " ";
                    textBox8.Text = " ";
                    timer4.Stop();
                    MessageBox.Show("Invalid card", ex.Message.Substring(2, 0));
                    test2 = 0;
                    tempo = 0;
                }

                if (test2 == 1)
                {
                    int count = textBox8.Lines.Length;
                    TextReader read = new System.IO.StringReader(textBox8.Text);
                    int rows = 150;

                    string[] text10 = new string[rows];
                    for (int r = 1; r < rows; r++)
                    {
                        text10[r] = read.ReadLine();
                    }



                    if (count >= 8)
                    {

                        try
                        {








                            richTextBox16.Text = text10[1].Remove(0, 6);
                            richTextBox1.Text = text10[3].Remove(0, 9);
                            richTextBox2.Text = text10[2].Remove(0, 6);
                            richTextBox3.Text = text10[4].Remove(0, 10);
                            richTextBox4.Text = text10[5].Remove(0, 4);
                            richTextBox8.Text = text10[6].Remove(0, 1);
                            //  richTextBox5.Text = text10[7].Remove(0, 8);
                            //  richTextBox9.Text = text10[8].Remove(0, 0);


                            rich5 = text10[7].Remove(0, 8).Trim();
                            //    MessageBox.Show(rich5);
                            rich9 = text10[8].Remove(0, 0).Trim();
                            // label33.Text = rich5;
                            //  label34.Text = rich9;
                            richTextBox10.Text = text10[9].Remove(0, 0);
                            richTextBox11.Text = text10[10].Remove(0, 0);
                            richTextBox12.Text = text10[11].Remove(0, 0);
                            richTextBox13.Text = text10[12].Remove(0, 0);
                            richTextBox14.Text = text10[13].Remove(0, 0);
                            richTextBox6.Text = text10[14].Remove(0, 6).Trim();
                            richTextBox7.Text = text10[15].Remove(0, 4).Trim();
                            richTextBox15.Text = "https://drive.google.com/file/d/" + text10[16] + "/view?usp=sharing";
                            maskedTextBox1.Text = richTextBox10.Text;
                            maskedTextBox2.Text = richTextBox11.Text;
                            maskedTextBox3.Text = richTextBox12.Text;
                            maskedTextBox4.Text = richTextBox13.Text;
                            maskedTextBox5.Text = richTextBox14.Text;
                          ///  MessageBox.Show(richTextBox16.Text);

                            check_id_onboard2();


                         
                            if(id_onboard2== true)
                            {

                                timer4.Stop();
                                MessageBox.Show("ESTA PESSOA ESTÁ EM OUTRA EMBARCAÇÃO! SERÁ NECESSÁRIO DAR A SAÍDA PARA A LIBERAÇÃO DE ENTRADA A ESTA EMBARCAÇÃO");
                                   
                          
                             
                            }


                            if (id_onboard2 == false)
                            {



                                if (text10[17].Trim() == "1")
                            {
                                local1.Checked = true;

                            }
                            if (text10[18].Trim() == "1")
                            {
                                local2.Checked = true;
                            }
                            if (text10[19].Trim() == "1")
                            {
                                local3.Checked = true;
                            }
                            if (text10[20].Trim() == "1")
                            {
                                local4.Checked = true;
                            }
                            if (text10[21].Trim() == "1")
                            {
                                level_yellow.Checked = true;
                            }
                            if (text10[22].Trim() == "1")
                            {
                                level_green.Checked = true;
                            }
                            if (text10[23].Trim() == "1")
                            {
                                level_red.Checked = true;
                            }
                         

                   
                        
                        
                                // path3 = "https://drive.google.com/file/d/" + subs2[5] + "/view?usp=sharing";
                                // panel11.Visible = true;
                                //  panel4.Visible = false;

                                /////////////////////

                                // string ve = textBox10.Lines[6];
                                textBox15.Text = text10[7];

                            if (text10[7] == "VESSEL: " + comboBox1.Text.Trim())
                            {
                                try
                                {
                                    path2 = richTextBox15.Text;
                                    subs2 = path2.Split('/');
                                    path3 = subs2[5];
                                }
                                catch
                                {

                                }
                                string teste2 = "Number : " + richTextBox16.Text.Trim() + " : Name : " + richTextBox2.Text + " : Compay : " + richTextBox1.Text + " :Funcition:  " + richTextBox3.Text + "  :Id: " + richTextBox4.Text + " : E-mail : " + richTextBox8.Text + " : Vessel : " + comboBox1.Text.Trim() + " : Project : " + richTextBox9.Text + " : ASO : " + richTextBox10.Text + " : NR-34 : " + richTextBox11.Text + " : NR-10 : " + richTextBox12.Text + " : NR-33 : " + richTextBox13.Text + " : NR-35 : " + richTextBox14.Text + " : " + bb + " : " + comuser.Text + " :" + richTextBox17.Text + " :" + " :" + local22 + " :" + richTextBox6.Text + " :" + richTextBox7.Text + " :" + path3;
                                string filePath = @"C:\compartilhamento\data_txt\data2.txt";
                                string[] lines = File.ReadAllLines(filePath);

                                for (int i = 0; i < lines.Length; i++)
                                {
                                    if (lines[i].Contains(richTextBox4.Text.Trim()))
                                    {
                                        ver = 1;
                                        lines[i] = teste2.Trim();
                                        /// MessageBox.Show("Achei: " + richTextBox4.Text.Trim());
                                    }

                                }
                                //and save it:

                                File.WriteAllLines(filePath, lines);
                                ver = 0;

                            }
                            //////////////////


                            //  pictureBox1.Visible = false;
                            //  panel4.Visible = false;
                            //  panel4.Size = new Size(360, 355);
                            // panel4.Location = new System.Drawing.Point(95, 150);
                            // pictureBox1.Size = new Size(330, 320);
                            // pictureBox1.Location = new System.Drawing.Point(15, 18);

                            // panel4.Visible = true;
                            //  pictureBox1.Visible = true;
                            // panel10.Visible = false;



                            // button8.Visible = true;
                            // button9.Visible = true;
                            // button10.Visible = true;
                            // pictureBox1.BackgroundImage = Properties.Resources.barcode1;
                            int rich1 = Int16.Parse(richTextBox16.Text.Trim()) - 1;
                            string secondLine2 = File.ReadLines(@"C:\compartilhamento\data_txt\data2.txt").ElementAtOrDefault(rich1);
                            string sec = secondLine2.Split(':')[26].Trim();
                            string pass11 = secondLine2.Split(':')[5].Trim();
                            string pass12 = secondLine2.Split(':')[9].Trim();
                            string pass13 = secondLine2.Split(':')[3].Trim();
                            string pass14 = secondLine2.Split(':')[11].Trim();
                            string pass15 = secondLine2.Split(':')[17].Trim();
                            string pass16 = secondLine2.Split(':')[19].Trim(); // NR34
                            string pass17 = secondLine2.Split(':')[21].Trim(); // VA1
                            string pass18 = secondLine2.Split(':')[23].Trim(); // V2                                                                                                                                                                   //17
                            string pass19 = secondLine2.Split(':')[25].Trim();// V BOOSTER

                            //   string pass19 = secondLine2.Split(':')[25].Trim();// V BOOSTER





                            //  0   1      2                       3        4              5         6              7  8           9         10                                  11      12             13        14      15    16           17     18           19          20            21          22           23               24            25 26     27   
                            //Number : 1 : Name : CRISTIANO CALHEIROS 3 : Compay : GOOGLEMARINE :Funcition:  ENGENHEIRO  :Id: 1110988400 : E-mail : cristiano.engenharia.ac@gmail.com : Vessel : Googlemarine : Project : 2001 : ASO : 02/02/2024 : NR-34 : 02/02/2025 : Vaccine-1 : 02/02/2026 : Vaccine-2 : 02/02/2027 : Booster vaccine : 02/02/2028 :  : COMUM :
                            if (pass11.Trim() != richTextBox1.Text.Trim())  //COMPANY
                            {
                                company_loc = true;
                            }



                            if (pass12.Trim() != richTextBox4.Text.Trim())  // IDENTIDADE
                            {
                                company_loc = true;
                            }

                            if (pass13.Trim() != richTextBox2.Text.Trim()) // NOME
                            {
                                company_loc = true;
                            }

                            if (pass14.Trim() != richTextBox8.Text.Trim()) // email
                            {
                                company_loc = true;
                            }
                            if (pass15.Trim() != maskedTextBox1.Text.Trim()) // aso
                            {
                                company_loc = true;
                            }
                            if (pass16.Trim() != maskedTextBox2.Text.Trim()) // nr34
                            {
                                company_loc = true;
                            }
                            if (pass17.Trim() != maskedTextBox3.Text.Trim()) // vacina 1
                            {
                                company_loc = true;
                            }
                            if (pass18.Trim() != maskedTextBox4.Text.Trim()) // vacina 2
                            {
                                company_loc = true;
                            }
                            if (pass19.Trim() != maskedTextBox5.Text.Trim()) // vacina reforço
                            {
                                company_loc = true;
                            }


                            if (pass11.Trim() == richTextBox1.Text.Trim() && pass12.Trim() == richTextBox4.Text.Trim() && pass13.Trim() == richTextBox2.Text.Trim())
                            {
                                if (pass14.Trim() == richTextBox8.Text.Trim() && pass15.Trim() == maskedTextBox1.Text.Trim() && pass16.Trim() == maskedTextBox2.Text.Trim())
                                {
                                    if (pass17.Trim() == maskedTextBox3.Text.Trim() && pass18.Trim() == maskedTextBox4.Text.Trim() && pass19.Trim() == maskedTextBox5.Text.Trim())
                                    {
                                        company_loc = false; // 
                                    }
                                }
                            }



                            if (sec == "Bloqueado")
                            {
                                bb = "Bloqueado";
                                label37.Text = secondLine2.Split(':')[28];
                                // company_loc = true;
                            }
                            else
                            {
                                bb = "";
                                label37.Text = "";
                                // company_loc = false;
                            }





                                acha_palavra_txt();
                                compare_data();

                                //  timer2.Enabled = true;
                                textBox7.Text = "";
                                textBox8.Text = " ";
                                label37.Text = "";
                                tempo = 0;
                            }

                            timer4.Stop();
                        }
                        catch (Exception ex)
                        {
                            //  MessageBox.Show("",ex.Message);
                        }
                    }

                }
            }
            
        }
        public async Task TestaPing(string url)
        {
            try
            {
                Ping pinger = new Ping();
                PingReply resposta = await pinger.SendPingAsync(url);
                ExibeInfoRespostaPing(resposta);
                pinger.PingCompleted += pinger_PingCompleted;


            }
            catch
            {

            }
        }
        private static void pinger_PingCompleted(object sender, PingCompletedEventArgs e)
        {
            try
            {
                PingReply resposta = e.Reply;
                if (e.Cancelled)
                {
                    MessageBox.Show($"Ping para {e.UserState.ToString()} foi cancelado");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Excepition lançada durante o ping:{ex.Message}");
            }
        }

        private void ExibeInfoRespostaPing(PingReply resposta)
        {

            st = resposta.Status.ToString().Trim();
            /// MessageBox.Show(st);
            if (st == "Success")
            {
                // panel12.BackColor = Color.YellowGreen;
                // label40.Visible = false;

                if (timer9.Enabled == false)
                {

                    // timer9.Start();
                }
                p = 1;
                ping_local = 1;
            }
            if (st == "TimedOut")
            {
                // panel12.BackColor = Color.Red;
                // label40.Visible = true;
                zzz = 0;
                p = 0;
                ping_local = 0;
                // MessageBox.Show("sem comunicação");
            }
        }

        private void timer8_Tick(object sender, EventArgs e)
        {
            //TestaPing(ip1);
            // TestaPing(ip2);
            //  TestaPing(ip3);
            // checa_host();
            //  MessageBox.Show(comp.ToString());
            label51.Text = count2.ToString();
            //  label51.Text = hostName;
            String block = File.ReadLines(@"C:\compartilhamento\lock.txt").ElementAtOrDefault(0);

            if (comp == 1)
            {
                //checa_host();
                //  atualiza_compartilhamento();
                /// ler_linha();
            }

            if (count >= 1)
            {
                // label51.Text = count.ToString();
                // online_ = true;
            }

            DateTime fproj2 = File.GetLastWriteTime(@"C:\compartilhamento\data_txt\PROJETO.txt");
            if (fproj2 > fproj)
            {

                System.Windows.Forms.Application.Restart();

            }


            DateTime flogo2 = File.GetLastWriteTime(@"C:\compartilhamento\logo_criptoqrcode\logo.png");
            if (flogo2 > flogo3)
            {


                System.Windows.Forms.Application.Restart();


            }
            //  MessageBox.Show(flogo3.ToString() + "    " + flogo2.ToString());
            //    DateTime flogo;
            //   DateTime flogo1;
        }
        private void ler_linha3()
        {
            //   String locked = File.ReadLines(@"C:\compartilhamento\lock.txt").ElementAtOrDefault(0);


            //  timer8.Stop();
            try
            {
                // string FileToRead = @"C:\data_txt\data.txt";
                string FileToRead = @"C:\compartilhamento\data_txt\data.txt";
                TextReader Leitor = new StreamReader(@"C:\compartilhamento\data_txt\data.txt", true);//Inicializa o Leitor
                int Linhas = 0;
                //  MessageBox.Show("acesso ok");
                while (Leitor.Peek() != -1)
                {//Enquanto o arquivo não acabar, o Peek não retorna -1 sendo adequando para o loop while...
                    Linhas++;//Incrementa 1 na contagem
                    Leitor.ReadLine();//Avança uma linha no arquivo
                }



                label27.Text = Linhas.ToString();
             //   label67.Text = label27.Text;
                if (rich5 == comboBox1.Text.Trim())
                {

                }
                Leitor.Close();
            }
            catch
            {
                //  MessageBox.Show("não consegui acessar o caminho  " + @"C:\compartilhamento\data_txt\data.txt");
            }

        }
      //  int comp = 1;
        private void timer9_Tick(object sender, EventArgs e)
        {
           //myThread.Start();
           // myThread.Abort();
        //    panel12.BackColor = Color.Black;

            try
            {
               if (compr == 0)
                {
                    // atualiza_compartilhamento();
                    checa_host();
                    //  timer9.Stop();
                    // ler_linha3();
                  compr = 1;

                }
            }
            catch
            {

            }
        }
        private void userloc()
        {
            panel5.Visible = true;
            button1.Enabled = true;
            button28.Enabled = true;
           // button2.Enabled = true;
          //  button17.Enabled = true;
          //  button29.Enabled = true;
            button3.Enabled = true;
            button4.Enabled = true;
            button19.Enabled = true;
        }
        private void userlocon()
        {
            button1.Enabled = false;
            button28.Enabled = false;
           // button2.Enabled = false;
          //  button17.Enabled = false;
          //  button29.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            button19.Enabled = false;
            panel5.Visible = false;
        }
        bool verz = false;

        public void loc_crew()
        {
            button21.Visible = false;
            button6.Visible = false;
            button4.Visible = false;
            button3.Visible = false;
           // button29.Visible = false;
            button28.Visible = false;
           // button17.Visible = false;
          //  button2.Visible = false;
            button1.Visible = false;
            button19.Visible = false;
            monitor();
        }

        public void unloc_crew()
        {
            button21.Visible = true;
            button6.Visible = true;
            button4.Visible = true;
            button3.Visible = true;
            //button29.Visible = true;
            button28.Visible = true;
            //button17.Visible = true;
           // button2.Visible = true;
            button1.Visible = true;
            button19.Visible = true;
        }
        private void button35_Click(object sender, EventArgs e)
        {
          
            verz = true;
            String pass0 = passall.Split(',')[0];
            String pass1 = passall.Split(',')[1];
            String pass2 = passall.Split(',')[2];
            String pass3 = passall.Split(',')[3];
            String pass4 = passall.Split(',')[4];
            String pass5 = passall.Split(',')[5];
            String pass6 = passall.Split(',')[6];
            String pass7 = passall.Split(',')[7];
            String pass8 = passall.Split(',')[8];
            String pass9 = passall.Split(',')[9];
            String pass10 = passall.Split(',')[10];
            String pass11 = passall.Split(',')[11];
            //   if (verz == true)
            // {


            //  }
            try
            {
                if (comuser.SelectedItem.ToString() == pass0.Trim() && textpass.Text.Trim() == pass1.Trim())
                {


                    MessageBox.Show("Usuário Admin " + comuser.Text + " com acesso liberado");
                    dataGridView1.ReadOnly = false;
                    // button27.Visible = true;
                    regs.Visible = true;
                    panel17.Visible = true;
                    panel19.Visible = true;

                    libera = true;
                    userloc();
                    textBox7.Focus();
                    textBox7.Text = "";
                    unloc_crew();

                }
                if (comuser.SelectedItem.ToString() == pass0.Trim() && textpass.Text.Trim() != pass1.Trim())
                {
                    MessageBox.Show("A SENHA DO USUARIO " + comuser.Text + " ESTÁ INCORRETA!");
                    userlocon();
                    libera = false;
                    button27.Visible = false;
                    regs.Visible = false;
                    panel17.Visible = false;
                    panel19.Visible = false;
                }





                if (comuser.SelectedItem.ToString() == pass2.Trim() && textpass.Text.Trim() == pass3.Trim())
                {
                    MessageBox.Show("Usuário Admin nivel 2 " + comuser.Text + " com acesso liberado");
                    dataGridView1.ReadOnly = false;
                    libera = true;
                    userloc();
                    verz = false;
                    // button27.Visible = true;
                    panel17.Visible = false;
                    panel19.Visible = false;
                    regs.Visible = true;
                    textBox7.Focus();
                    textBox7.Text = "";
                    unloc_crew();

                }
                if (comuser.SelectedItem.ToString() == pass2.Trim() && textpass.Text.Trim() != pass3.Trim())
                {
                    MessageBox.Show("A SENHA DO USUARIO " + comuser.Text + " ESTÁ INCORRETA!");
                    userlocon();
                    libera = false;
                    button27.Visible = false;
                    regs.Visible = false;
                    panel17.Visible = false;
                    panel19.Visible = false;


                }




                if (comuser.SelectedItem.ToString() == pass4.Trim() && textpass.Text.Trim() == pass5.Trim())
                {
                    MessageBox.Show("Usuário " + comuser.Text + " com acesso liberado");
                    dataGridView1.ReadOnly = true;
                    libera = false;
                    userloc();
                    verz = false;
                    button27.Visible = false;
                    regs.Visible = false;
                    textBox7.Focus();
                    unloc_crew();
                    textBox7.Text = "";
                }
                if (comuser.SelectedItem.ToString() == pass4.Trim() && textpass.Text.Trim() != pass5.Trim())
                {
                    MessageBox.Show("A SENHA DO USUARIO " + comuser.Text + " ESTÁ INCORRETA!");
                    userlocon();
                    libera = false;
                    button27.Visible = false;
                    regs.Visible = false;

                }

                if (comuser.SelectedItem.ToString() == pass6.Trim() && textpass.Text.Trim() == pass7.Trim())
                {
                    MessageBox.Show("Usuário " + comuser.Text + " com acesso liberado");
                    dataGridView1.ReadOnly = true;
                    libera = false;
                    userloc();
                    verz = false;
                    button27.Visible = false;
                    regs.Visible = false;
                    textBox7.Focus();
                    textBox7.Text = "";
                    unloc_crew();
                }
                if (comuser.SelectedItem.ToString() == pass6.Trim() && textpass.Text.Trim() != pass7.Trim())
                {
                    MessageBox.Show("A SENHA DO USUARIO " + comuser.Text + " ESTÁ INCORRETA!");
                    userlocon();
                    libera = false;
                    regs.Visible = false;

                }


                if (comuser.SelectedItem.ToString() == pass8.Trim() && textpass.Text.Trim() == pass9.Trim())
                {
                    MessageBox.Show("Usuário " + comuser.Text + " com acesso liberado");
                    dataGridView1.ReadOnly = true;
                    libera = false;
                    verz = false;
                    userloc();
                    button27.Visible = false;
                    textBox7.Focus();
                    textBox7.Text = "";
                    regs.Visible = false;
                    unloc_crew();
                }

                if (comuser.SelectedItem.ToString() == pass8.Trim() && textpass.Text.Trim() != pass9.Trim())
                {
                    MessageBox.Show("A SENHA DO USUARIO " + comuser.Text + " ESTÁ INCORRETA!");
                    userlocon();
                    libera = false;
                    regs.Visible = false;

                }

                if (comuser.SelectedItem.ToString() == pass10.Trim() && textpass.Text.Trim() == pass11.Trim())
                {
                    MessageBox.Show("Usuário " + comuser.Text + " com acesso liberado");
                    dataGridView1.ReadOnly = true;
                    libera = false;
                    userloc();
                    verz = false;


                    textBox7.Focus();
                    textBox7.Text = "";
                    regs.Visible = false;
                    loc_crew();
                }

                if (comuser.SelectedItem.ToString() == pass10.Trim() && textpass.Text.Trim() != pass11.Trim())
                {
                    MessageBox.Show("A SENHA DO USUARIO " + comuser.Text + " ESTÁ INCORRETA!");
                    userlocon();
                    libera = false;


                }


                if (comuser.SelectedItem.ToString() == "" && textpass.Text.Trim() == "")
                {

                    MessageBox.Show("ESCOLHA O USUÁRIO E DIGITE A SENHA!");
                    userlocon();
                    libera = false;
                    regs.Visible = false;

                }
            }
            catch
            {

            }

            // MessageBox.Show(libera.ToString());
        }

        private void button28_Click(object sender, EventArgs e)
        {

            if (comboBox1.Text != "")
            {

                label44.Visible = true;
                label54.Visible = true;
                label45.Visible = true;
                label46.Visible = true;
                label47.Visible = true;
                //  label48.Visible = true;
                label49.Visible = true;
                label50.Visible = true;

                local1.Checked = false;
                local2.Checked = false;
                local3.Checked = false;
                local1.Enabled = true;
                local2.Enabled = true;
                local4.Enabled = true;
                label37.Visible = false;

                dataGridView1.Visible = false;
                button2.Enabled = false;
                button17.Enabled = false;
              //  button29.Enabled = true;
                btloc.Visible = false;
                button7.Visible = false;
                // CHLocked.Visible = false;
                richTextBox4.ReadOnly = false;
                richTextBox16.ReadOnly = false;
                plant = 0;

                if (band == 0)
                {
                    button3.Text = Label_Show_data[0];
                }
                else
                {
                    button3.Text = Label_Show_data[1];
                }

                // button3.Text = "Show DataBase";


                button1.Enabled = true;

                panel11.Visible = true;
                pictureBox7.Image = Properties.Resources.barcode1;
                panel11.BackColor = Color.White;
                label8.Visible = false;
                //  button7.Visible = true;
                maskedTextBox1.ReadOnly = false;
                maskedTextBox2.ReadOnly = false;
                maskedTextBox3.ReadOnly = false;
                maskedTextBox4.ReadOnly = false;
                maskedTextBox5.ReadOnly = false;
                maskedTextBox1.Visible = true;
                maskedTextBox2.Visible = true;
                maskedTextBox3.Visible = true;
                maskedTextBox4.Visible = true;
                maskedTextBox5.Visible = true;
                maskedTextBox1.Text = " ";
                //maskedTextBox2.Text = " ";
                //maskedTextBox3.Text = " ";
                // maskedTextBox4.Text = " ";
                //  maskedTextBox5.Text = " ";
                // MaskedTextBox m = new MaskedTextBox();
                // m.Text = "00000000";
                maskedTextBox2.Text = "00000000";
                maskedTextBox3.Text = "00000000";
                maskedTextBox4.Text = "00000000";
                maskedTextBox5.Text = "00000000";




                // panel4.Size = new Size(296, 215);
                // panel4.Location = new System.Drawing.Point(80, 277);
                //  pictureBox1.Location = new System.Drawing.Point(5, 25);
                //  pictureBox1.Size = new Size(178, 176);
                // panel10.Visible = true;

                if (textBox7.SelectionLength >= 0)
                {
                    textBox7.Focus();
                    textBox7.Text = "";
                }
                richTextBox1.ReadOnly = false;
                richTextBox2.ReadOnly = false;
                richTextBox3.ReadOnly = false;
                richTextBox4.ReadOnly = false;
                //richTextBox5.ReadOnly = false;
                // richTextBox6.ReadOnly = false;
                // richTextBox7.ReadOnly = false;
                richTextBox8.ReadOnly = false;

                richTextBox10.ReadOnly = true;
                richTextBox11.ReadOnly = true;
                richTextBox12.ReadOnly = true;
                richTextBox13.ReadOnly = true;
                richTextBox14.ReadOnly = true;

                // if (timer4.Enabled)
                /// {

                cancelar();
                //button2.Enabled = true;
                if (band == 0)
                {
                    button28.Text = Label_Read_QRcode_Off[0];
                    label7.Text = button28.Text;
                }
                else
                {
                    button28.Text = Label_Read_QRcode_Off[1];
                    label7.Text = button28.Text;
                }

                //richTextBox9.Text = "";
                richTextBox10.Text = "";
                richTextBox11.Text = "";
                richTextBox12.Text = "";
                richTextBox13.Text = "";
                richTextBox14.Text = "";
                //button1.Text = "Read QRcode Off";
                richTextBox1.Enabled = true;
                richTextBox2.Enabled = true;
                richTextBox3.Enabled = true;
                richTextBox4.Enabled = true;
                comboBox1.Enabled = true;
                richTextBox6.Enabled = true;
                richTextBox7.Enabled = true;
                richTextBox8.Enabled = true;
                richTextBox9.Enabled = true;
                richTextBox10.Enabled = true;
                richTextBox11.Enabled = true;
                richTextBox12.Enabled = true;
                richTextBox13.Enabled = true;
                richTextBox14.Enabled = true;


                button4.Enabled = true;
                richTextBox1.Text = "";
                richTextBox2.Text = "";
                richTextBox3.Text = "";
                richTextBox4.Text = "";
                // richTextBox5.Text = "";
                richTextBox6.Text = "";
                richTextBox7.Text = "";

                //richTextBox9.Text = "";
                richTextBox10.Text = "";
                richTextBox11.Text = "";
                richTextBox12.Text = "";
                richTextBox13.Text = "";
                richTextBox14.Text = "";
                richTextBox16.Text = "";
                // richTextBox8.Text = "cristiano.engenharia.ac@gmail.com";
                dateTimePicker1.Visible = true;
                dateTimePicker2.Visible = true;
                richTextBox6.Visible = false;
                richTextBox7.Visible = false;
                //    timer4.Stop();
                // videoCaptureDevice.Stop();
                //  pictureBox1.Image = Properties.Resources.frame;
                //button2.Enabled = true;
                // }
                vid = 0;
            }
            else
            {
                MessageBox.Show(_cad);
            }
        }
        private void compare_aso()
        {
            try
            {

                string texto = "Aqui está um exemplo de texto com a palavra 'dado'.";

                string palavraProcurada = "dado";

                if (texto.Contains(palavraProcurada))
                {
                    Console.WriteLine($"A palavra '{palavraProcurada}' foi encontrada no texto.");
                }
                var parameterDate_ASo = DateTime.ParseExact(maskedTextBox1.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    //var parameterDate_initial = DateTime.ParseExact(richTextBox6.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    /// var parameterDate_final = DateTime.ParseExact(richTextBox7.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    var todaysDate = DateTime.Today;

                    if (todaysDate > parameterDate_ASo)
                    {


                        beep();
                        beep();
                        beep();
                        beep();
                        aso_1 = 1;
                        MessageBox.Show("Aso vencido, Favor verificar");
                    }
                    else
                    {

                        aso_1 = 0;

                    }


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            // compare_id();
        }
        int ver = 0;
        String local22 = "";
        private void read_write()
        {
            String data_new;
            String data2_new;
            if (dateTimePicker1.Visible == true)
            {
                data_new = dateTimePicker1.Value.ToString("dd/MM/yyyy").ToString().Trim();
                richTextBox6.Text = data_new;
            }
            else
            {
                data_new = richTextBox6.Text.Trim();
            }
            if (dateTimePicker2.Visible == true)
            {
                data2_new = dateTimePicker2.Value.ToString("dd/MM/yyyy").ToString().Trim();
                richTextBox7.Text = data2_new;
            }
            else
            {
                data2_new = richTextBox7.Text.Trim();
            }
            if (richTextBox16.Text.Trim() == "")
            {


                int count = File.ReadAllLines(@"C:\compartilhamento\data_txt\data2.txt").Length;

                string number = count.ToString().Trim(); //   teste count System.IO.File.ReadAllText(@"C:\compartilhamento\data_txt\count.txt");
                label3.Text = number;
                number2 = count;//Int32.Parse(number);
                lb4.Visible = true;
                number2 = number2 + 1;
                // File.WriteAllText(@"C:\compartilhamento\data_txt\count.txt", number2.ToString());
                if (zzz == 1)
                {
                    // string text = System.IO.File.ReadAllText(@"C:\compartilhamento\rede.txt");
                    // \\DOF_ACCESS\\compartilhamento\\
                    //  File.WriteAllText(text + @"data_txt\count.txt", number2.ToString());
                }

                label3.Text = number2.ToString();
                lb4.Text = label3.Text;
                richTextBox16.Text = number2.ToString();

                if (local1val == 1)
                {
                    local22 = place1[band];

                }
                if (local2val == 1)
                {
                    local22 = place2[band];

                }
                if (local4val == 1)
                {
                    local22 = place4[band];

                }
                using (StreamWriter writer = new StreamWriter(@"C:\compartilhamento\data_txt\data2.txt", true)) //@"\data_full\vessel\local_list.txt("C:\\data_full\\local.txt", true))
                {
                    try
                    {
                        path2 = richTextBox15.Text;
                        subs2 = path2.Split('/');
                        path3 = subs2[5];
                    }
                    catch
                    {

                    }

                    //  richTextBox16.Text = label3.Text.Trim();
                    string teste2 = "Number : " + richTextBox16.Text.Trim() + " : Name : " + richTextBox2.Text + " : Compay : " + richTextBox1.Text + " :Funcition:  " + richTextBox3.Text + "  :Id: " + richTextBox4.Text + " : E-mail : " + richTextBox8.Text + " : Vessel : " + comboBox1.Text.Trim() + " : Project : " + richTextBox9.Text + " : ASO : " + richTextBox10.Text + " : NR-34 : " + richTextBox11.Text + " : NR-10 : " + richTextBox12.Text + " : NR-33 : " + richTextBox13.Text + " : NR-35 : " + richTextBox14.Text + " : " + bb + " : " + comuser.Text + " :" + richTextBox17.Text + " :" + " :" + local22 + " :" + data_new + " :" + data2_new + " :" + path3;

                    writer.WriteLine(teste2);
                    writer.Close();
                }
                // MessageBox.Show("Não Achei ");
            }

            else
            {
                //  richTextBox16.Text = number2.ToString();
                if (local1val == 1)
                {
                    local22 = place1[band];

                }
                if (local2val == 1)
                {
                    local22 = place2[band];

                }
                if (local4val == 1)
                {
                    local22 = place4[band];

                }
                if (richTextBox16.Text.Trim() != "")
                {
                    try
                    {
                        path2 = richTextBox15.Text;
                        subs2 = path2.Split('/');

                        path3 = subs2[5];
                    }
                    catch
                    {

                    }
                    string teste2 = "Number : " + richTextBox16.Text.Trim() + " : Name : " + richTextBox2.Text + " : Compay : " + richTextBox1.Text + " :Funcition:  " + richTextBox3.Text + "  :Id: " + richTextBox4.Text + " : E-mail : " + richTextBox8.Text + " : Vessel : " + comboBox1.Text.Trim() + " : Project : " + richTextBox9.Text + " : ASO : " + richTextBox10.Text + " : NR-34 : " + richTextBox11.Text + " : NR-10 : " + richTextBox12.Text + " : NR-33 : " + richTextBox13.Text + " : NR-35 : " + richTextBox14.Text + " : " + bb + " : " + comuser.Text + " :" + richTextBox17.Text + " :" + " :" + local22 + " :" + data_new + " :" + data2_new + " :" + path3;
                    string filePath = @"C:\compartilhamento\data_txt\data2.txt";
                    string[] lines = File.ReadAllLines(filePath);

                    for (int i = 0; i < lines.Length; i++)
                    {
                        if (lines[i].Contains(richTextBox4.Text.Trim()))
                        {
                            ver = 1;
                            lines[i] = teste2.Trim();
                            /// MessageBox.Show("Achei: " + richTextBox4.Text.Trim());
                        }

                    }
                    //and save it:

                    File.WriteAllLines(filePath, lines);
                    ver = 0;

                }
            }




        }
        public Bitmap GerarQRCode(int width, int height, string text)
        {
            try
            {
                var bw = new ZXing.BarcodeWriter();
                var encOptions = new ZXing.Common.EncodingOptions() { Width = width, Height = height, Margin = 0 };
                bw.Options = encOptions;
                bw.Format = ZXing.BarcodeFormat.QR_CODE;
                var resultado = new Bitmap(bw.Write(text));
                return resultado;

            }
            catch
            {
                throw;
            }
        }

        private void criar_excel()
        {

            string txtFilePath = @"C:\compartilhamento\data_base\novo.txt";
            string excelFilePath = @"C:\compartilhamento\data_base\novo.xls";

            IWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Planilha1");

            using (StreamReader reader = new StreamReader(txtFilePath))
            {
                string headerLine = reader.ReadLine();
                string[] headerParts = { "NUMBER", "NAME", "COMPANY FUNCTION ", "FUNCTION", "ID", "EMAIL", "VESSEL", "CHECK-IN VALIDATION", "CHECK-OUT VALIDATION", "CHECK-IN  DATA", "CHECK-IN  HORA", "CHECK-OUT DATA", "CHECK-OUT HORA", "PROJECT", "ASO" , "NR-34", "NR-10", "NR-33", "NR-35", "LOCAL", "LEVEL", "ESTADO", "MOTIVO", "USUARIO" };//headerLine.Split(':');

                IRow headerRow = sheet.CreateRow(0);
                for (int colIndex = 0; colIndex < headerParts.Length; colIndex++)
                {
                    ICell cell = headerRow.CreateCell(colIndex);
                    cell.SetCellValue(headerParts[colIndex]);
                }

                int rowIndex = 1;

                while (!reader.EndOfStream)
                {
                    string line = reader.ReadLine();
                    string[] parts = line.Split(',');

                    IRow row = sheet.CreateRow(rowIndex);

                    for (int colIndex = 0; colIndex < parts.Length; colIndex++)
                    {
                        ICell cell = row.CreateCell(colIndex);
                        cell.SetCellValue(parts[colIndex]);
                    }

                    rowIndex++;
                }
            }

            using (FileStream stream = new FileStream(excelFilePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(stream);
            }
        }
        private void cadastrar_invited()
        {

            rec = false;
            Boolean tempo = false;
            if (dateTimePicker1.Visible == true)
            {
                if (dateTimePicker2.Value.Date >= dateTimePicker1.Value.Date)
                {
                    tempo = true;
                }
                else
                {
                    tempo = false;
                    MessageBox.Show("A DATA FINAL ESTÁ MENOR QUE A DATA INICIAL, CORRIJA POR FAVOR!");

                }
            }

            if (dateTimePicker1.Visible == false)
            {
                tempo = true;
            }
            if (tempo == true)
            {
                timer10.Enabled = false;
                timer12.Enabled = false;
                ProcessStartInfo startInfo = new ProcessStartInfo("cmd", "/K netsh interface set interface \"Ethernet\" DISABLED");
                startInfo.RedirectStandardOutput = true;
                startInfo.UseShellExecute = false;
                // Do not create the black window.
                startInfo.CreateNoWindow = true;
                startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                Process.Start(startInfo);

                check_id_onboard();

                if (id_onboard == false)
                {
                    if (dateTimePicker1.Visible == true)
                    {
                        richTextBox6.Text = dateTimePicker1.Value.ToString("dd/MM/yyyy").ToString().Trim();
                        richTextBox7.Text = dateTimePicker2.Value.ToString("dd/MM/yyyy").ToString().Trim();
                    }


                   // alterado();
                    System.Threading.Thread.Sleep(2000);
                    compare_id();
                    check_if_exist_id();

                    if (id_exist == true)
                    {

                        //  MessageBox.Show("id existe");
                        richTextBox10.Text = "VISITANTE";//maskedTextBox1.Text;
                        richTextBox11.Text = "N/A";
                        richTextBox12.Text = "N/A";
                        richTextBox13.Text = "N/A";
                        richTextBox14.Text = "N/A";
                        textBox13.Text = richTextBox1.Text;
                        textBox7.Focus();
                        textBox7.Text = "";

                        if (band == 0)
                        {
                            button2.Text = Label_Create_QRcode[0];
                        }
                        else
                        {
                            button2.Text = Label_Create_QRcode[1];
                        }





                        // var parameterDate2_initial = DateTime.ParseExact(dateTimePicker1.Value.Date.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                        // var parameterDate2_final = DateTime.ParseExact(dateTimePicker2.Value.Date.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);

                        if (resultado == 0)
                        {
                            if (textBox7.SelectionLength >= 0)
                            {
                                // textBox7.Focus();
                                // textBox7.Text = "";
                            }
                            // dataGridView1.Visible = false;
                            if (richTextBox1.Text != "" && richTextBox2.Text != "" && richTextBox3.Text != "" && richTextBox4.Text != "" && richTextBox6.Text != "" && richTextBox7.Text != "" && comboBox1.Text != "" && richTextBox8.Text != "" && checado == 1 && richTextBox6.Text != " " && richTextBox7.Text != " ")  //  /  /
                            {



                                // compare_aso();

                                // if (aso_1 == 0)
                                //  {
                              //  richTextBox1.Text = richTextBox1.Text.Trim()+" VISITANTE ";

                                    if (richTextBox8.Text == "")
                                    {
                                        richTextBox8.Text = "N/A";
                                    }
                                if (richTextBox10.Text == "")
                                {
                                    richTextBox10.Text = "N/A";
                                }
                                if (richTextBox11.Text == "")
                                {
                                    richTextBox11.Text = "N/A";
                                }
                                    if (richTextBox12.Text == "")
                                    {
                                        richTextBox12.Text = "N/A";
                                    }
                                if (richTextBox13.Text == "")
                                {
                                    richTextBox13.Text = "N/A";
                                }
                                if (richTextBox14.Text == "")
                                {
                                    richTextBox14.Text = "N/A";
                                }

                                    maskedTextBox1.Visible = false;
                                    maskedTextBox2.Visible = false;
                                    maskedTextBox3.Visible = false;
                                    maskedTextBox4.Visible = false;
                                    maskedTextBox5.Visible = false;




                                    read_write();
                                    confere = 1;
                                    lb4.Visible = true;
                                    label5.Visible = true;
                                    panel10.Visible = true;
                                    label5.Text = richTextBox2.Text;




                                    qr_generate = "Printed Qrcode";

                                    //
                                   // CarregarPlanilha2();
                                carrega_planilha2_txt();
                                //  create_qrcode();
                                // create_qrcode_new();
                                create_qrcode_invited_new();
                                    print_qrcode();

                                    ProcessStartInfo startInfo2 = new ProcessStartInfo("cmd", "/K netsh interface set interface \"Ethernet\" ENABLED");
                                    startInfo2.RedirectStandardOutput = true;
                                    startInfo2.UseShellExecute = false;
                                    // Do not create the black window.
                                    startInfo2.CreateNoWindow = true;
                                    startInfo2.WindowStyle = ProcessWindowStyle.Hidden;
                                    Process.Start(startInfo2);

                               /// }





                                //
                                //string teste = "Name : " + richTextBox2.Text + " : Compay : " + richTextBox1.Text + " :Funcition: " + richTextBox3.Text + "  :Id: " + richTextBox4.Text + " : E-mail : " + richTextBox8.Text + " : Vessel :" + richTextBox5.Text + " : Project : " + richTextBox9.Text + ": ASO : " + richTextBox10.Text + " : NR-34 : " + richTextBox11.Text + " : Vaccine-1 : " + richTextBox12.Text + " : Vaccine-2 : " + richTextBox13.Text + " : Booster vaccine : " + richTextBox14.Text;
                                //  atualiza_compartilhamento();

                                // 

                                // else
                                //  {
                                //  MessageBox.Show(id_check[band]);
                                //  }
                                checado = 0;

                            }
                            else
                            {
                                if (band == 0)
                                {
                                    MessageBox.Show("Favor preencher todos os campos");
                                }

                                if (band == 1)
                                {
                                    MessageBox.Show("Please complete all informations places");
                                }
                            }

                        }


                        if (resultado == 1)
                        {
                            MessageBox.Show("ID duplicated");
                        }
                        textBox7.Focus();
                        textBox7.Text = " ";
                    }
                    ok_but2 = false;

                }
                else
                {
                    MessageBox.Show("ESTA PESSOA ESTÁ A BORDO! SÓ É PERMITIDO IMPRIMIR OU CADASTRAR SE A PESSOA ESTIVER FORA DA EMBARCAÇÃO");
                    id_onboard = false;
                }
            }
            rec = true;
            timer10.Enabled = true;
            timer12.Enabled = true;
        }

        private void create_qrcode_invited_new()
        {
            String data_new;
            String data2_new;
            if (dateTimePicker1.Visible == true)
            {
                data_new = dateTimePicker1.Value.ToString("dd/MM/yyyy").ToString().Trim();
            }
            else
            {
                data_new = richTextBox6.Text.Trim();
            }
            if (dateTimePicker2.Visible == true)
            {
                data2_new = dateTimePicker2.Value.ToString("dd/MM/yyyy").ToString().Trim();
            }
            else
            {
                data2_new = richTextBox7.Text.Trim();
            }

            label5.Text = " ";
            label28.Text = " ";
            label30.Text = " ";
            label31.Text = " ";
            panel4.BackColor = Color.White;
            panel11.Visible = false;
            panel4.Visible = true;

            label5.Text = richTextBox2.Text;
            label28.Text = "Vessel: " + comboBox1.Text;
            label30.Text = richTextBox9.Text;
            lb4.Text = richTextBox16.Text; ;
            label31.Text = "De: " + data_new;
            label32.Text = "A:    " + data2_new;


            if (richTextBox15.Text != "")
            {
                try
                {
                    path2 = richTextBox15.Text;
                    subs2 = path2.Split('/');

                    path3 = subs2[5];
                }
                catch
                {

                }
            }
            else
            {
                path3 = ".";
            }



         //   richTextBox1.Text = " VISITANTE  " + richTextBox1.Text;
            data2 = number + " " + richTextBox16.Text + "\r\n" + nome + " " + richTextBox2.Text + "\r\n" + emp + " " + richTextBox1.Text + " \r\n" + function + " " + richTextBox3.Text + "\r\n" + id + " " +
            this.richTextBox4.Text + "\r\n" + email + " " + this.richTextBox8.Text + "\r\n" + vessel + " " + this.comboBox1.Text.Trim() + "\r\n" + this.richTextBox9.Text + "\r\n" + this.richTextBox10.Text + "\r\n" + this.richTextBox11.Text + "\r\n" + this.richTextBox12.Text + "\r\n" + this.richTextBox13.Text + "\r\n" + this.richTextBox14.Text + "\r\n" +
            initial + " " + data_new + "\r\n" +
            final + " " + data2_new + "\r\n" + path3 + "\r\n" + local1val + "\r\n" + local2val + "\r\n" + local3val + "\r\n" + local4val + "\r\n" + levelyellow + "\r\n" + levelgreen + "\r\n" + levelred;

            Criptografia criptografia = new Criptografia(CryptProvider.RC2);
            criptografia.Key = "Etec2017"; // chave
            data2 = criptografia.Encrypt(data2.ToString());
            Bitmap bmp = new Bitmap(GerarQRCode(300, 300, data2));
            pictureBox1.Image = bmp;
        }
        private void create_qrcode_new()
        {

            String data_new;
            String data2_new;
            if (dateTimePicker1.Visible == true)
            {
                data_new = dateTimePicker1.Value.ToString("dd/MM/yyyy").ToString().Trim();
            }
            else
            {
                data_new = richTextBox6.Text.Trim();
            }
            if (dateTimePicker2.Visible == true)
            {
                data2_new = dateTimePicker2.Value.ToString("dd/MM/yyyy").ToString().Trim();
            }
            else
            {
                data2_new = richTextBox7.Text.Trim();
            }

            label5.Text = " ";
            label28.Text = " ";
            label30.Text = " ";
            label31.Text = " ";
            panel4.BackColor = Color.White;
            panel11.Visible = false;
            panel4.Visible = true;

            label5.Text = richTextBox2.Text;
            label28.Text = "Vessel: "+comboBox1.Text;
            label30.Text = richTextBox9.Text;
            lb4.Text = richTextBox16.Text; ;
            label31.Text = "De: " + data_new;
            label32.Text = "A:    " + data2_new;


            if (richTextBox15.Text != "")
            {
                try
                {
                    path2 = richTextBox15.Text;
                    subs2 = path2.Split('/');

                    path3 = subs2[5];
                }
                catch
                {

                }
            }
            else
            {
                path3 = ".";
            }




            data2 = number + " " + richTextBox16.Text + "\r\n" + nome + " " + richTextBox2.Text + "\r\n" + emp + " " + richTextBox1.Text + "\r\n" + function + " " + richTextBox3.Text + "\r\n" + id + " " +
            this.richTextBox4.Text + "\r\n" + email + " " + this.richTextBox8.Text + "\r\n" + vessel + " " + this.comboBox1.Text.Trim() + "\r\n" + this.richTextBox9.Text + "\r\n" + this.richTextBox10.Text + "\r\n" + this.richTextBox11.Text + "\r\n" + this.richTextBox12.Text + "\r\n" + this.richTextBox13.Text + "\r\n" + this.richTextBox14.Text + "\r\n" +
            initial + " " + data_new + "\r\n" +
            final + " " + data2_new + "\r\n" + path3 + "\r\n" + local1val + "\r\n" + local2val + "\r\n" + local3val + "\r\n" + local4val + "\r\n" + levelyellow + "\r\n" + levelgreen + "\r\n" + levelred;

            Criptografia criptografia = new Criptografia(CryptProvider.RC2);
            criptografia.Key = "Etec2017"; // chave
            data2 = criptografia.Encrypt(data2.ToString());
            Bitmap bmp = new Bitmap(GerarQRCode(300, 300, data2));
            pictureBox1.Image = bmp;
        }
        //bool a = false;
        private void button2_Click(object sender, EventArgs e)
        {
           
            rec = false;
            Boolean tempo = false;
            if (dateTimePicker1.Visible == true)
            {
                if (dateTimePicker2.Value.Date >= dateTimePicker1.Value.Date)
                {
                    tempo = true;
                }
                else
                {
                    tempo = false;
                    MessageBox.Show("A DATA FINAL ESTÁ MENOR QUE A DATA INICIAL, CORRIJA POR FAVOR!");

                }
            }

            if (dateTimePicker1.Visible == false)
            {
                tempo = true;
            }
            if (tempo == true)
            {
                timer10.Enabled = false;
                timer12.Enabled = false;
                ProcessStartInfo startInfo = new ProcessStartInfo("cmd", "/K netsh interface set interface \"Ethernet\" DISABLED");
                startInfo.RedirectStandardOutput = true;
                startInfo.UseShellExecute = false;
                // Do not create the black window.
                startInfo.CreateNoWindow = true;
                startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                Process.Start(startInfo);

                check_id_onboard();

                if (id_onboard == false)
                {
                    if (dateTimePicker1.Visible == true)
                    {
                        richTextBox6.Text = dateTimePicker1.Value.ToString("dd/MM/yyyy").ToString().Trim();
                        richTextBox7.Text = dateTimePicker2.Value.ToString("dd/MM/yyyy").ToString().Trim();
                    }


                    alterado();
                    System.Threading.Thread.Sleep(2000);
                    compare_id();
                    check_if_exist_id();

                    if (id_exist == true)
                    {

                        //  MessageBox.Show("id existe");
                        richTextBox10.Text = maskedTextBox1.Text;
                        richTextBox11.Text = maskedTextBox2.Text;
                        richTextBox12.Text = maskedTextBox3.Text;
                        richTextBox13.Text = maskedTextBox4.Text;
                        richTextBox14.Text = maskedTextBox5.Text;
                        textBox13.Text = richTextBox1.Text;
                        textBox7.Focus();
                        textBox7.Text = "";

                        if (band == 0)
                        {
                            button2.Text = Label_Create_QRcode[0];
                        }
                        else
                        {
                            button2.Text = Label_Create_QRcode[1];
                        }





                        // var parameterDate2_initial = DateTime.ParseExact(dateTimePicker1.Value.Date.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                        // var parameterDate2_final = DateTime.ParseExact(dateTimePicker2.Value.Date.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);

                        if (resultado == 0)
                        {
                            if (textBox7.SelectionLength >= 0)
                            {
                                // textBox7.Focus();
                                // textBox7.Text = "";
                            }
                            // dataGridView1.Visible = false;
                            if (richTextBox1.Text != "" && richTextBox2.Text != "" && richTextBox3.Text != "" && richTextBox4.Text != "" && richTextBox6.Text != "" && richTextBox7.Text != "" && comboBox1.Text != "" && richTextBox8.Text != "" && checado == 1 && maskedTextBox1.Text != "  /  /"
                                && maskedTextBox2.Text != "  /  /" && maskedTextBox3.Text != "  /  /" && maskedTextBox4.Text != "  /  /" && maskedTextBox5.Text != "  /  /" && richTextBox6.Text != " " && richTextBox7.Text != " ")  //  /  /
                            {



                                compare_aso();

                                if (aso_1 == 0)
                                {

                                    if (richTextBox8.Text == "")
                                    {
                                        richTextBox8.Text = "N/A";
                                    }

                                    maskedTextBox1.Visible = false;
                                    maskedTextBox2.Visible = false;
                                    maskedTextBox3.Visible = false;
                                    maskedTextBox4.Visible = false;
                                    maskedTextBox5.Visible = false;




                                    read_write();
                                    confere = 1;
                                    lb4.Visible = true;
                                    label5.Visible = true;
                                    panel10.Visible = true;
                                    label5.Text = richTextBox2.Text;




                                    qr_generate = "Printed Qrcode";

                                    //
                                    //  CarregarPlanilha2();
                                    carrega_planilha2_txt();
                                    //  create_qrcode();
                                    create_qrcode_new();
                                    print_qrcode();

                                    ProcessStartInfo startInfo2 = new ProcessStartInfo("cmd", "/K netsh interface set interface \"Ethernet\" ENABLED");
                                    startInfo2.RedirectStandardOutput = true;
                                    startInfo2.UseShellExecute = false;
                                    // Do not create the black window.
                                    startInfo2.CreateNoWindow = true;
                                    startInfo2.WindowStyle = ProcessWindowStyle.Hidden;
                                    Process.Start(startInfo2);
                                   
                                }





                                //
                                //string teste = "Name : " + richTextBox2.Text + " : Compay : " + richTextBox1.Text + " :Funcition: " + richTextBox3.Text + "  :Id: " + richTextBox4.Text + " : E-mail : " + richTextBox8.Text + " : Vessel :" + richTextBox5.Text + " : Project : " + richTextBox9.Text + ": ASO : " + richTextBox10.Text + " : NR-34 : " + richTextBox11.Text + " : Vaccine-1 : " + richTextBox12.Text + " : Vaccine-2 : " + richTextBox13.Text + " : Booster vaccine : " + richTextBox14.Text;
                                //  atualiza_compartilhamento();

                                // 

                                // else
                                //  {
                                //  MessageBox.Show(id_check[band]);
                                //  }
                                checado = 0;

                            }
                            else
                            {
                                if (band == 0)
                                {
                                    MessageBox.Show("Favor preencher todos os campos");
                                }

                                if (band == 1)
                                {
                                    MessageBox.Show("Please complete all informations places");
                                }
                            }

                        }


                        if (resultado == 1)
                        {
                            MessageBox.Show("ID duplicated");
                        }
                        textBox7.Focus();
                        textBox7.Text = " ";
                    }
                    ok_but2 = false;

                }
                else
                {
                    MessageBox.Show("ESTA PESSOA ESTÁ A BORDO! SÓ É PERMITIDO IMPRIMIR OU CADASTRAR SE A PESSOA ESTIVER FORA DA EMBARCAÇÃO");
                    id_onboard = false;
                }
            }
            rec = true;
            timer10.Enabled = true;
            timer12.Enabled = true;
        }
        private void compare_id()
        {
            String l = "";
            bool ESIM = false;
            bool dois = false;
            bool tres = false;
            string nume = "";
            string[] lines = File.ReadAllLines(@"C:\compartilhamento\data_txt\data2.txt");
            id_1 = 0;
            for (int i = 0; i < lines.Length; i++)
            {

                if (lines[i].Split(':')[9].Trim() == richTextBox4.Text.Trim())
                {

                    l = lines[i].Split(':')[9].Trim();
                    if (lines[i].Split(':')[1].Trim() != richTextBox16.Text.Trim())
                    {
                        MessageBox.Show("O NÚMERO DA IDENTIDADE  * " + lines[i].Split(':')[9].Trim() + " *  JÁ ESTÁ CADASTRADO NO ACESSO DE NÚMERO " + lines[i].Split(':')[1].Trim());
                    }
                    ESIM = true;
                    richTextBox4.Text = lines[int.Parse(richTextBox16.Text) - 1].Split(':')[9].Trim();
                    string text4 = "Number : " + richTextBox16.Text + " : Name : " + richTextBox2.Text + " : Compay : " + richTextBox1.Text + " :Funcition:  " + richTextBox3.Text + "  :Id: " + richTextBox4.Text + " : E-mail : " + richTextBox8.Text + " : Vessel : " + comboBox1.Text.Trim() + " : Project : " + richTextBox9.Text + " : ASO : " + maskedTextBox1.Text + " : NR-34 : " + maskedTextBox2.Text + " : NR-10 : " + maskedTextBox3.Text + " : NR-33 : " + maskedTextBox4.Text + " : NR-35 : " + maskedTextBox5.Text + " : " + bb + " : " + comuser.Text + " :" + richTextBox17.Text + " :" + local22.Trim();

                    string text = File.ReadAllText(@"C:\compartilhamento\data_txt\data2.txt");
                    text = text.Replace(lines[Int16.Parse(richTextBox16.Text) - 1], text4);
                    File.WriteAllText(@"C:\compartilhamento\data_txt\data2.txt", text);
                    //  MessageBox.Show(" CADASTRO REALIZADO COM SUCESSO, A IDENTIDADE NÃO FOI ALTERADA POIS JÁ EXISTE UMA IDENTIDADE COM ESTA NÚMERO");
                    break;

                }
                else
                {
                    ESIM = false;

                }



            }

            if (ESIM == false)
            {
                if (local1val == 1)
                {
                    local22 = place1[band];

                }
                if (local2val == 1)
                {
                    local22 = place2[band];

                }
                if (local4val == 1)
                {
                    local22 = place4[band];

                }
                id_1 = 1;
                string[] lines2 = File.ReadAllLines(@"C:\compartilhamento\data_txt\data2.txt");
                if (richTextBox4.Text.Trim() != l)
                {
                    string text4 = "Number : " + richTextBox16.Text.Trim() + " : Name : " + richTextBox2.Text + " : Compay : " + richTextBox1.Text + " :Funcition:  " + richTextBox3.Text + "  :Id: " + richTextBox4.Text.Trim() + " : E-mail : " + richTextBox8.Text + " : Vessel : " + comboBox1.Text.Trim() + " : Project : " + richTextBox9.Text + " : ASO : " + maskedTextBox1.Text + " : NR-34 : " + maskedTextBox2.Text + " : NR-10 : " + maskedTextBox3.Text + " : NR-33 : " + maskedTextBox4.Text + " : NR-35 : " + maskedTextBox5.Text + " : " + bb + " : " + comuser.Text + " :" + richTextBox17.Text + " :" + local22 + " :" + richTextBox15.Text + "\r\n";

                    string text = File.ReadAllText(@"C:\compartilhamento\data_txt\data2.txt");
                    text = text.Replace(lines2[Int16.Parse(richTextBox16.Text) - 1], text4);
                    File.WriteAllText(@"C:\compartilhamento\data_txt\data2.txt", text);
                    //  MessageBox.Show("CADASTRO REALIZADO COM SUCESSO");
                }
                else
                {

                }
            }

            timer10.Enabled = true;
        }
        private void button29_Click(object sender, EventArgs e)
        {
            
            rec = false;
            Boolean tempo2 = false;
            if (dateTimePicker1.Visible == true)
            {
                if (dateTimePicker2.Value.Date >= dateTimePicker1.Value.Date)
                {
                    tempo2 = true;
                }
                else
                {
                    tempo2 = false;
                    MessageBox.Show("A DATA FINAL ESTÁ MENOR QUE A DATA INICIAL, CORRIJA POR FAVOR!");

                }
            }

            if (dateTimePicker1.Visible == false)
            {
                tempo2 = true;
            }
            if (tempo2 == true)
            {
                timer10.Enabled = false;
                timer12.Enabled = false;

                ProcessStartInfo startInfo = new ProcessStartInfo("cmd", "/K netsh interface set interface \"Ethernet\" DISABLED");
                startInfo.RedirectStandardOutput = true;
                startInfo.UseShellExecute = false;
                // Do not create the black window.
                startInfo.CreateNoWindow = true;
                startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                Process.Start(startInfo);

                string[] lines = File.ReadAllLines(@"C:\compartilhamento\data_txt\data2.txt");
                bool id_ok = false;
                for (int i = 0; i < lines.Length; i++)
                {

                    if (lines[i].Split(':')[9].Trim() == richTextBox4.Text.Trim())
                    {
                        id_ok = true;
                        MessageBox.Show("JÁ EXISTE UM CADASTRO COM ESTA IDENTIDADE");
                        break;
                    }

                }
                if (id_ok == false)
                {
                    check_id_onboard();
                    if (id_onboard == false)
                    {





                        tempo = 0;

                        //timer4.Stop();
                        richTextBox10.Text = maskedTextBox1.Text;
                        richTextBox11.Text = maskedTextBox2.Text;
                        richTextBox12.Text = maskedTextBox3.Text;
                        richTextBox13.Text = maskedTextBox4.Text;
                        richTextBox14.Text = maskedTextBox5.Text;
                        textBox13.Text = richTextBox1.Text;
                        // 
                        //  if (textBox7.SelectionLength >= 0)
                        // {
                        textBox7.Focus();
                        textBox7.Text = "";

                        if (band == 0)
                        {
                            button2.Text = Label_Create_QRcode[0];
                        }
                        else
                        {
                            button2.Text = Label_Create_QRcode[1];
                        }




                        // acha_palavra_txt2();



                        if (resultado == 0)
                        {
                            if (textBox7.SelectionLength >= 0)
                            {
                                // textBox7.Focus();
                                // textBox7.Text = "";
                            }
                            // dataGridView1.Visible = false;
                            if (richTextBox1.Text != " " && richTextBox2.Text != " " && richTextBox3.Text != " " && richTextBox4.Text != " " && comboBox1.Text != " " && richTextBox8.Text != " " && checado == 1 && maskedTextBox1.Text != "  /  /"
                                && maskedTextBox2.Text != "  /  /" && maskedTextBox3.Text != "  /  /" && maskedTextBox4.Text != "  /  /" && maskedTextBox5.Text != "  /  /")  //  /  /
                            {
                                //compare_id();






                                compare_aso();

                                if (aso_1 == 0)
                                {
                                    if (richTextBox8.Text == "")
                                    {
                                        richTextBox8.Text = "N/A";
                                    }

                                    read_write();

                                    maskedTextBox1.Visible = false;
                                    maskedTextBox2.Visible = false;
                                    maskedTextBox3.Visible = false;
                                    maskedTextBox4.Visible = false;
                                    maskedTextBox5.Visible = false;


                                    // f


                                    confere = 1;
                                    lb4.Visible = true;
                                    label5.Visible = true;
                                    panel10.Visible = true;
                                    label5.Text = richTextBox2.Text;


                                    qr_generate = "registered";
                                    cad = true;
                                  //  CarregarPlanilha2();
                                    carrega_planilha2_txt();
                                    create_qrcode_new();

                                    ProcessStartInfo startInfo2 = new ProcessStartInfo("cmd", "/K netsh interface set interface \"Ethernet\" ENABLED");
                                    startInfo2.RedirectStandardOutput = true;
                                    startInfo2.UseShellExecute = false;
                                    // Do not create the black window.
                                    startInfo2.CreateNoWindow = true;
                                    startInfo2.WindowStyle = ProcessWindowStyle.Hidden;
                                    Process.Start(startInfo2);


                                    richTextBox1.Text = "";
                                    richTextBox2.Text = "";
                                    richTextBox3.Text = "";
                                    richTextBox4.Text = "";
                                    richTextBox6.Text = "";
                                    richTextBox7.Text = "";
                                    richTextBox8.Text = "";
                                    richTextBox10.Text = "";
                                    richTextBox11.Text = "";
                                    richTextBox12.Text = "";
                                    richTextBox13.Text = "";
                                    richTextBox14.Text = "";
                                    richTextBox15.Text = "";
                                    richTextBox16.Text = "";
                                    maskedTextBox1.Text = "";
                                    maskedTextBox2.Text = "";
                                    maskedTextBox3.Text = "";
                                    maskedTextBox4.Text = "";
                                    maskedTextBox5.Text = "";
                                    if (band == 0)
                                    {
                                        MessageBox.Show(Label_cadastro[0]);
                                    }
                                    if (band == 1)
                                    {
                                        MessageBox.Show(Label_cadastro[1]);
                                    }

                                    local1.Checked = false;
                                    local2.Checked = false;
                                    local4.Checked = false;
                                    checado = 0;
                                }
                                //  MessageBox.Show(id_check[band]);
                            }


                            // atualiza_compartilhamento();

                            else
                            {
                                if (band == 0)
                                {
                                    MessageBox.Show("Favor preencher todos os campos");
                                }

                                if (band == 1)
                                {
                                    MessageBox.Show("Please complete all informations places");
                                }
                            }
                        }
                        if (resultado == 1)
                        {
                            MessageBox.Show("ID duplicated");
                        }
                        textBox7.Focus();
                        textBox7.Text = " ";

                    }
                    else
                    {
                        MessageBox.Show("ESTA PESSOA ESTÁ A BORDO!, SÓ É PERMITIDO IMPRIMIR OU CADASTRAR SE A PESSOA ESTIVER FORA DA EMBARCAÇÃO");
                    }
                }
            }
            ///  timer4.Start();
            rec = true;
      
            timer10.Enabled = true;
            timer12.Enabled = true;
        }

        private void button4_Click(object sender, EventArgs e)
        {

            panel15.Visible = true;



        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            // panel8.Location= new Point(Location.X,Location.Y);//Point(521,155);
            panel8.Size = new Size(490, 353); //new Size(1030, 600);
            wi++;
            if (wi == 1)
            {
                //online = true;
                lview_AP.Items.Clear();
                wifi = new Wifi();
                List<AccessPoint> aps = wifi.GetAccessPoints();
                foreach (AccessPoint ap in aps)
                {
                    ListViewItem lvItem = new ListViewItem(ap.Name);
                    lvItem.SubItems.Add(ap.SignalStrength + "%");
                    lvItem.Tag = ap;
                    lview_AP.Items.Add(lvItem);
                }


                panel8.Visible = true;
            }
            if (wi == 2)
            {
                // online = false;
                if (textBox7.SelectionLength >= 0)
                {
                    textBox7.Focus();
                    textBox7.Text = "";
                }
                panel8.Visible = false;

                wi = 0;
            }

            // System.Windows.Forms.Application.Exit();
            //   this.Close();
        }
        int inp = 0;
        public void monitor()
        {
            // MessageBox.Show(valores.ToString());

            pictureBox4.Visible = false;


            if (new FileInfo(@"C:\compartilhamento\data_txt\data.txt").Length >= 0)
            {
                inp++;
                if (inp == 1)
                {
                    // panel6.Visible = true;
                    mostra_conteudo_txt();
                    comboBox2.SelectedIndex = 0;

                    //  comboBox2.Text = "ok";

                    comboBox2.Visible = true;
                    beep();
                    beep();
                    textBox7.Focus();
                    textBox7.Text = "";
                }
                if (inp == 2)
                {

                    //if (textBox7.SelectionLength >= 0)
                    // {
                    textBox7.Focus();
                    textBox7.Text = "";
                    // }
                    panel6.Visible = false;
                    comboBox2.Visible = false;
                    inp = 0;
                }
            }
            textBox7.Focus();
            textBox7.Text = "";

            //this.Close();
        }

        private void button19_Click(object sender, EventArgs e)
        {
            monitor();

        }

        private void button21_Click(object sender, EventArgs e)
        {
            //panel16.Visible = true;
            // ler_linha();
            //CloseExcel();
            ProcessStartInfo startInfo2 = new ProcessStartInfo("cmd", "/K netsh interface set interface \"Ethernet\" DISABLED");
            startInfo2.RedirectStandardOutput = true;
            startInfo2.UseShellExecute = false;
            // Do not create the black window.
            startInfo2.CreateNoWindow = true;
            startInfo2.WindowStyle = ProcessWindowStyle.Hidden;
            Process.Start(startInfo2);
            System.Windows.Forms.Application.Restart();
        }

        private void button22_Click(object sender, EventArgs e)
        {
            CloseExcel();
            this.Close();
        }
        private void check_if_exist_number()
        {

            // richTextBox17.Text = "";
            //    label37.Text = "Bloquear";
            int rich1 = Int16.Parse(richTextBox16.Text.Trim()) - 1;
            string secondLine2 = File.ReadLines(@"C:\compartilhamento\data_txt\data2.txt").ElementAtOrDefault(rich1);
            //  MessageBox.Show(secondLine);
            //string checline = secondLine.Trim();
            if (secondLine2 != null)
            {

                //  Number: 1 : Name: Cristiano: Compay: Googlemarine: Funcition: Engenheiro: Id: 111098414 : E - mail : 1 : Vessel: Googlemarine: Project: 190603 : ASO: 22 / 02 / 2023 : NR - 34 : 22 / 02 / 2023 : Vaccine - 1 : 22 / 02 / 2023 : Vaccine - 2 : 22 / 02 / 2023 : Booster vaccine : 22 / 02 / 2023 : Bloqueado: GUSTAVO: Falta da quarta dose da vacina
                int rich16 = Int16.Parse(richTextBox16.Text.Trim()) - 1;
                int lab3 = Int16.Parse(label3.Text);
                if (rich16 <= lab3)
                {
                    string secondLine = File.ReadLines(@"C:\compartilhamento\data_txt\data2.txt").ElementAtOrDefault(rich16);

                    //  MessageBox.Show(secondLine);
                    //string checline = secondLine.Trim();

                    try
                    {


                        if (ok_but2 == false)
                        {
                            button2.Enabled = true;
                            // richTextBox15.Text = "";
                            /*
                             richTextBox2.Text = secondLine.Split(':')[3].Trim();
                             richTextBox3.Text = secondLine.Split(':')[7].Trim();
                             richTextBox4.Text = secondLine.Split(':')[9].Trim();
                             richTextBox1.Text = secondLine.Split(':')[5].Trim();
                             richTextBox8.Text = secondLine.Split(':')[11].Trim();
                            */


                            //Number : 1 : Name : GUSTAVO MAGALHAES : Compay : DOF :Funcition:  GERENTE DE PROJETO  :Id: 8866640719 : E-mail : cristiano.engenharia.ac@gmail.com : Vessel : Skandi Rio : Project : Docagem : ASO : 07/01/2023 : NR-34 : 00/00/0000 : Vaccine-1 : 00/00/0000 : Vaccine-2 : 00/00/0000 : Booster vaccine : 00/00/0000 :  : COMUM : : :Convés :14/10/2022 :28/10/2022 :.

                            richTextBox2.Text = secondLine.Split(':')[3].Trim();
                            richTextBox3.Text = secondLine.Split(':')[7].Trim();
                            richTextBox4.Text = secondLine.Split(':')[9].Trim();
                            richTextBox1.Text = secondLine.Split(':')[5].Trim();

                            comboBox1.Items.Clear();
                            comboBox1.Items.Insert(0, secondLine.Split(':')[13].Trim());
                            comboBox1.SelectedIndex = 0;

                            richTextBox8.Text = secondLine.Split(':')[11].Trim();
                            if (secondLine.Split(':')[11].Trim() == "")
                            {
                                richTextBox8.Text = "N/A";
                            }

                            maskedTextBox1.Visible = true;
                            maskedTextBox2.Visible = true;
                            maskedTextBox3.Visible = true;
                            maskedTextBox4.Visible = true;
                            maskedTextBox5.Visible = true;
                            // maskedTextBox1.Text = "00/00/0000";
                            // maskedTextBox2.Text = "00/00/0000";
                            try
                            {

                                string ok = secondLine.Split(':')[17].Trim();
                                //  MessageBox.Show(ok);
                                string teste = DateTime.ParseExact(ok, "M/d/yyyy", CultureInfo.InvariantCulture).ToString("dd/MM/yyyy");

                                maskedTextBox1.Text = teste;

                                if (ok == "")
                                {
                                    maskedTextBox1.Text = "00/00/0000";
                                }

                                ok = secondLine.Split(':')[19].Trim();
                                //  MessageBox.Show(ok);
                                teste = DateTime.ParseExact(ok, "M/d/yyyy", CultureInfo.InvariantCulture).ToString("dd/MM/yyyy");
                                // MessageBox.Show(teste);
                                maskedTextBox2.Text = teste;

                                if (ok == "")
                                {
                                    maskedTextBox2.Text = "00/00/0000";
                                }


                            }
                            catch
                            {

                            }

                            comboBox1.Items.Clear();
                            comboBox1.Items.Insert(0, secondLine.Split(':')[13].Trim());
                            comboBox1.SelectedIndex = 0;


                            if (richTextBox3.Text != "")
                            {
                                //  comboBox1.Items.Clear();
                                StreamReader sr = new StreamReader(@"C:\compartilhamento\vessels.txt");
                                string x = sr.ReadToEnd();
                                string[] y = x.Split('\n');
                                foreach (string s in y)
                                {
                                    comboBox1.Items.Add(s);
                                }
                                sr.Close();
                            }




                            maskedTextBox1.Text = secondLine.Split(':')[17];
                            maskedTextBox2.Text = secondLine.Split(':')[19];
                            maskedTextBox3.Text = secondLine.Split(':')[21];
                            maskedTextBox4.Text = secondLine.Split(':')[23];
                            maskedTextBox5.Text = secondLine.Split(':')[25];

                            if (secondLine.Split(':')[17].Trim() == "")
                            {
                                // maskedTextBox1.Text = "00/00/0000";
                            }

                            if (secondLine.Split(':')[19].Trim() == "")
                            {
                                maskedTextBox2.Text = "00/00/0000";
                            }

                            if (secondLine.Split(':')[21].Trim() == "")
                            {
                                maskedTextBox3.Text = "00/00/0000";
                            }
                            if (secondLine.Split(':')[23].Trim() == "")
                            {
                                maskedTextBox4.Text = "00/00/0000";
                            }
                            if (secondLine.Split(':')[25].Trim() == "")
                            {
                                maskedTextBox5.Text = "00/00/0000";
                            }

                            try
                            {
                                richTextBox15.Text = secondLine.Split(':')[33].Trim();
                            }
                            catch
                            {
                                // richTextBox15.Text = "";
                            }
                            richTextBox6.Visible = true;
                            richTextBox7.Visible = true;
                            dateTimePicker1.Visible = false;
                            dateTimePicker2.Visible = false;

                            try
                            {

                                richTextBox6.Text = secondLine.Split(':')[31];
                                richTextBox7.Text = secondLine.Split(':')[32];
                                String local222 = secondLine.Split(':')[30].Trim();

                                label34.Text = local222;
                                if (local222 == place1[band])
                                {
                                    local1.Checked = true;
                                    local2.Checked = false;
                                    local4.Checked = false;
                                }
                                if (local222 == place2[band])
                                {
                                    local1.Checked = false;
                                    local2.Checked = true;
                                    local4.Checked = false;
                                }
                                if (local222 == place4[band])
                                {
                                    local1.Checked = false;
                                    local2.Checked = false;
                                    local4.Checked = true;
                                }
                            }
                            catch
                            {

                            }
                        }

                        string sec = secondLine.Split(':')[26].Trim();
                        if (sec == "Bloqueado")
                        {
                            button2.Enabled = false;
                            button17.Enabled = false;
                            //  button5.Image = ((System.Drawing.Image)(resources.GetObject("_bButton.Image")));
                            label37.Text = secondLine.Split(':')[28];
                            textbloc = richTextBox17.Text.Trim();
                            textbloc2 = richTextBox17.Text.Trim();
                            //label37.Text = sec;
                            btloc.Visible = true;
                            button7.Visible = false;
                            label37.Visible = true;
                            btloc.Image = Properties.Resources.lock0;
                            ////  btloc.Text = "Bloqueado";
                            CHLocked.Text = "Desbloquear";
                            //  textbloc = "";
                            //  textbloc2 = "";
                        }
                        else
                        {
                            richTextBox17.Text = "";
                            //  label37.Text = "Desbloqueado";
                            CHLocked.Text = "Bloquear";
                            btloc.Visible = false;
                            button7.Visible = true;
                            button2.Enabled = true;
                            button17.Enabled = true;
                            //  btloc.Text = " ";
                            // btloc.Image = Properties.Resources.lock2;
                        }
                        String pass11 = secondLine.Split(':')[11];
                    }

                    catch
                    {
                        // MessageBox.Show("Não há cadastro com este Número");
                    }
                }




                //  button2.Enabled = true;
                //   button17.Enabled = true;
              //  button29.Enabled = false;
                // richTextBox4.ReadOnly = true;
                // richTextBox4.
                //   pictureBox1.Visible = false;
                //   panel4.Visible = false;



                //panel11.Size = new Size(360, 355);
                //panel11.Location = new System.Drawing.Point(95, 150);
                //pictureBox7.Size = new Size(330, 320);
                // pictureBox7.Location = new System.Drawing.Point(15, 18);


                try
                {

                    pictureBox7.LoadAsync(@"C:\compartilhamento\data_picture\" + richTextBox4.Text.ToString() + ".jpg");

                }
                catch
                {
                    pictureBox7.LoadAsync(@"C:\compartilhamento\data_picture\face.jpg");
                }

            }
            else
            {
                id_exist = false;
                richTextBox16.Text = "";
                richTextBox2.Text = "";
                richTextBox3.Text = "";
                richTextBox4.Text = "";
                richTextBox1.Text = "";
                richTextBox8.Text = "";
                richTextBox10.Text = "";
                richTextBox11.Text = "";
                richTextBox12.Text = "";
                richTextBox13.Text = "";
                richTextBox14.Text = "";
                richTextBox15.Text = "";
                maskedTextBox1.Text = "";
                maskedTextBox2.Text = "";
                maskedTextBox3.Text = "";
                maskedTextBox4.Text = "";
                maskedTextBox5.Text = "";
                maskedTextBox1.Visible = false;
                maskedTextBox2.Visible = false;
                maskedTextBox3.Visible = false;
                maskedTextBox4.Visible = false;
                maskedTextBox5.Visible = false;
                button2.Enabled = false;
                button17.Enabled = false;


                MessageBox.Show("NÃO EXISTE CADASTRO COM ESTE NÚMERO!");
            }
        }

        private void richTextBox16_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (richTextBox16.Text != "")
            {
                richTextBox6.Text = "";
                richTextBox7.Text = "";
                richTextBox15.Text = "";
                label37.Text = "";
                local1.Checked = false;
                local2.Checked = false;
                local3.Checked = false;
                dateTimePicker1.Visible = false;
                dateTimePicker2.Visible = false;
                check_if_exist_number();
            }



        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (richTextBox17.Text != "")
            {
                DialogResult dialogResult = MessageBox.Show("DESEJA REALMENTE BLOQUEAR O ACESSO?", "BLOQUEIO DE ACESSO", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    bb = "Bloqueado";
                    bloqueio();
                    //  richTextBox17.Text = "";
                    btloc.Image = Properties.Resources.lock0;
                    btloc.Visible = true;
                    button7.Visible = false;
                    richTextBox17.Text = "";
                    button2.Enabled = false;
                    button17.Enabled = false;
                   
                    //myThread.Abort();
                    checa_host();
                    // MessageBox.Show("button7_Click");
                    // atualiza_compartilhamento();
                }
                else if (dialogResult == DialogResult.No)
                {

                }

            }
            else
            {
                MessageBox.Show("FAVOR JUSTIFICAR O MOTIVO DO BLOQUEIO");
            }

        }

        private void btloc_Click(object sender, EventArgs e)
        {

            if (richTextBox17.Text != "")
            {

                if (r7 == true)
                {
                    DialogResult dialogResult = MessageBox.Show("DESEJA REALMENTE DESBLOQUEAR O ACESSO ? ", "DESBLOQUEIO DE ACESSO", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        bb = "Desbloqueado";
                        bloqueio();
                        //  richTextBox17.Text = "";
                        btloc.Visible = false;
                        button7.Visible = true;
                        button2.Enabled = true;
                        button17.Enabled = true;
                        richTextBox17.Text = "";
                        //myThread.Abort();
                        checa_host();
                        MessageBox.Show("btloc_Click");
                        // atualiza_compartilhamento();
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        //   btloc.Image = Properties.Resources.lock0;
                        //  btloc.Image = Properties.Resources.lock2;

                    }
                    r7 = false;
                }
            }
            else
            {
                MessageBox.Show("FAVOR JUSTIFICAR O MOTIVO DO DESBLOQUEIO");
            }

            /*

            textbloc = richTextBox17.Text.Trim();
            bt1++;
            if(bt1== 1)
            {


              //  btloc.Image = Properties.Resources.lock0;
              //  btloc.Text = "Bloqueado";
            }
            if (bt1 >= 2)
            {
                if (textbloc!= textbloc2 && richTextBox17.Text != "")
                {

                    btloc.Image = Properties.Resources.lock2;

                    
                    DialogResult dialogResult = MessageBox.Show("Sure", "Some Title", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        bloqueio();
                        richTextBox17.Text = "";
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        btloc.Image = Properties.Resources.lock0;

                        textbloc = textbloc2;
                        richTextBox17.Text = textbloc2;
                    }

                    //  r7 = false;
                    bt1 = 0;
                }
             
            }
            */
        }
        int get_pic = 0;
        private void button15_Click(object sender, EventArgs e)
        {
            try
            {

                get_pic++;
                if (get_pic == 1)
                {
                    if (richTextBox1.Text != "" && richTextBox2.Text != "" && richTextBox3.Text != "" && richTextBox4.Text != "" && comboBox1.Text != "")
                    {
                        try
                        {

                            button15.Text = "Press again to take a picture";
                            // if(cboCamera.)
                            //  if(cboCamera.Items(1)=="teste")

                            int iten = cboCamera.Items.Count;
                            if (iten >= 4)
                            {
                                iten = 1;
                            }
                            else
                            {
                                iten = 0;

                            }
                            //  MessageBox.Show(iten.ToString());
                            videoCaptureDevice = new VideoCaptureDevice(filterInfoCollection[0].MonikerString);
                            videoCaptureDevice.NewFrame += FinalFrame_NewFrame;
                            videoCaptureDevice.Start();
                            pictureBox7.Image = Properties.Resources.barcode1;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Erro " + ex.Message);
                        }
                    }
                }



                if (get_pic == 2)
                {

                    if (richTextBox1.Text != "" && richTextBox2.Text != "" && richTextBox3.Text != "" && richTextBox4.Text != "" && comboBox1.Text != "")
                    {
                        try
                        {
                            if (videoCaptureDevice.IsRunning)
                            {

                                caminhoImagemSalva2 = @"c:\compartilhamento\data_new_picture\" + richTextBox4.Text.Trim() + ".jpg";
                                //caminhoImagemSalva2.
                                pictureBox7.Image.Save(caminhoImagemSalva2, ImageFormat.Jpeg);

                                copyAll(@"C:\compartilhamento\data_new_picture\", @"C:\compartilhamento\data_picture\");
                                //  clearFolder(@"C:\compartilhamento\data_new_picture\");
                                button15.Text = "Face picture Include";
                                button15.Visible = false;



                            }

                            if (videoCaptureDevice.IsRunning)
                            {
                                videoCaptureDevice.Stop();
                                pictureBox7.Image = Properties.Resources.barcode1;

                            }
                        }
                        catch
                        {

                        }
                        //myThread.Abort();
                        checa_host();
                        //  MessageBox.Show("button15_Click");
                        // atualiza_compartilhamento();
                    }

                    get_pic = 0;
                }

            }
            catch
            {

            }


        }
        private void check_locked()
        {
            string[] lines = File.ReadAllLines(@"C:\compartilhamento\data_txt\data2.txt");

            // Start at line number 2 because there is a header
            for (int i = 1; i < lines.Length; i++)
            {
                // 2 ways to do this:
                if (lines[i].Contains(richTextBox4.Text))
                {

                    String pass0 = lines[1].Split(':')[26];

                    label37.Text = pass0;

                    if (label37.Text.Trim() == "Bloqueado")
                    {
                        String pass1 = lines[1].Split(':')[28];
                        richTextBox17.Text = pass1;
                    }
                    //   MessageBox.Show(pass0);
                    // string text3 = lines[i];
                    // string text4 = "Number : " + richTextBox16.Text + " : Name : " + richTextBox2.Text + " : Compay : " + richTextBox1.Text + " :Funcition:  " + richTextBox3.Text + "  :Id: " + richTextBox4.Text + " : E-mail : " + richTextBox16.Text + " : Vessel : " + richTextBox5.Text + " : Project : " + richTextBox9.Text + " : ASO : " + maskedTextBox1.Text + " : NR-34 : " + maskedTextBox2.Text + " : Vaccine-1 : " + maskedTextBox3.Text + " : Vaccine-2 : " + maskedTextBox4.Text + " : Booster vaccine : " + maskedTextBox5.Text + " : " + label37.Text + " : " + comuser.Text + " :" + richTextBox17.Text;
                    // string text = File.ReadAllText(@"C:\compartilhamento\data_txt\data2.txt");
                    // text = text.Replace(text3, text4);
                    // File.WriteAllText(@"C:\compartilhamento\data_txt\data2.txt", text);

                }


                // or a more structured way:
                //  if (lines[i].Split('|')[2].Contains("ABC"))
                // {
                // Copy it where you want
                // }
            }
        }
        private void ler_linha2()
        {
            criterio = " ";
            lbResultado.Items.Clear();
            criterio = richTextBox4.Text;
            string[] linhas = File.ReadAllLines(@"C:\compartilhamento\data_txt\data2.txt");




            if (richTextBox4.Text.Length >= 7)
            {
                foreach (string linha in linhas)
                {
                    // cri++;
                    if (linha.Contains(criterio) && criterio != " ")
                    {
                        cri = 1;
                        lbResultado.Items.Add(linha);
                        //  Number: 0 : Name: 2 : Compay: 2 :Funcition: 2  :Id: 22222222222 : E - mail : 2 : Vessel: 2 : Project: 2 : ASO: 2 : NR - 34 : 2 : Vaccine - 1 : 2 : Vaccine - 2 : 2 : Booster vaccine : 2
                        //     Number:  : Name: Cristiano de Araujo Calheiros : Compay: Googlemarine: Funcition: 
                        //Engenheiro: Id: 111098414 : E - mail : cristiano.engengaria.ac @gmail.com: Vessel:
                        //Skandi Salvador : Project: Docagem: ASO: 11 / 12 / 23 : NR - 34 : 11 / 12 / 23 :
                        //Vaccine - 1 : 11 / 12 / 23 : Vaccine - 2 : 11 / 12 / 23 : Booster vaccine : 11 / 12 / 23           // Number: 14 : Name: 1 : Compay: 4 :Funcition: 2  :Id: 366666666 : E - mail : 7 : Vessel: 5 : Project: 6 : ASO: 8 : NR - 34 : 9 : Vaccine - 1 : 10 : Vaccine - 2 : 11 : Booster vaccine : 12
                        string[] parts = linha.Split(':');
                        string whatIWant0 = parts[1].Trim();// + " ";// + parts[1];
                        string whatIWant1 = parts[1].Trim();// + " ";// + parts[1];
                        string whatIWant2 = parts[3].Trim();// + " ";// + parts[1];
                        string whatIWant3 = parts[5].Trim();// + " ";// + parts[1];
                        string whatIWant4 = parts[7].Trim();// + " ";// + parts[1];
                        string whatIWant5 = parts[9].Trim();// + " ";// + parts[1];
                        string whatIWant6 = parts[11].Trim();// + " ";// + parts[1];
                        string whatIWant7 = parts[13].Trim();// + " ";// + parts[1];
                        string whatIWant8 = parts[15].Trim();// + " ";// + parts[1];
                        string whatIWant9 = parts[17].Trim();// + " ";// + parts[1];
                        string whatIWant10 = parts[19].Trim();// + " ";// + parts[1];
                        string whatIWant11 = parts[21].Trim();// + " ";// + parts[1];
                        string whatIWant12 = parts[23].Trim();// + " ";// + parts[1];
                        string whatIWant13 = parts[25].Trim();// + " ";// + parts[1];
                                                              // string whatIWant14 = parts[23].Trim();// + " ";// + parts[1];

                        // Name: Googlemarine: Compay: Cristiano calheiros  :Funcition: Engenheiro: Id: 111098888 : E - mail : cristiano.engenharia.ac @gmail.com: Vessel: Skandi Salvador : Project: reparo: ASO: 11 / 02 / 2023 : NR - 34 : 11 / 02 / 2023 : Vaccine - 1 : 11 / 02 / 2023 : Vaccine - 2 : 11 / 02 / 2023 : Booster vaccine : 11 / 02 / 2023
                        richTextBox16.Text = whatIWant0;//.Remove(0, 8);   // number
                        richTextBox2.Text = whatIWant2;//.Remove(0, 6);    //name
                        richTextBox1.Text = whatIWant3; //.Remove(0, 9);   // company
                        richTextBox3.Text = whatIWant4;//.Remove(0, 10);   // function
                        richTextBox4.Text = whatIWant5;//.Remove(0, 4);    //  id
                        richTextBox8.Text = whatIWant6;//.Remove(0, 4);   // e-mail
                                                       //  richTextBox5.Text = whatIWant7;//.Remove(0, 8);    // vessel
                                                       //  richTextBox9.Text = whatIWant8;//.Remove(0, 8);    // projeto
                        richTextBox10.Text = whatIWant9;//.Remove(0, 8);   // aso
                        richTextBox11.Text = whatIWant10;//.Remove(0, 8);   // nr34
                        richTextBox12.Text = whatIWant11;//.Remove(0, 8);  // vaccine 1
                        richTextBox13.Text = whatIWant12;//.Remove(0, 8);  // vaccine 2
                        richTextBox14.Text = whatIWant13;//.Remove(0, 8);  // refoço


                        maskedTextBox1.Text = richTextBox10.Text;
                        maskedTextBox2.Text = richTextBox11.Text;
                        maskedTextBox3.Text = richTextBox12.Text;
                        maskedTextBox4.Text = richTextBox13.Text;
                        maskedTextBox5.Text = richTextBox14.Text;
                        // richTextBox15.Text = whatIWant13;//.Remove(0, 8);  // refoço

                        check_locked();
                    }
                    else
                    {
                        cri = 0;
                    }



                }




            }

            if (lbResultado.Items.Count == 0)
            {
                if (confere == 1)
                {
                    escrever_palavra2();
                }

                richTextBox1.Text = "";
                richTextBox2.Text = "";
                richTextBox3.Text = "";
                // richTextBox4.Text = "";
                richTextBox8.Text = "";
                // richTextBox5.Text = "";



            }

            confere = 0;

        }
        private void escrever_palavra2()
        {


            string nomeArquivo = @"C:\compartilhamento\data_txt\data2.txt";
            string textoInserir = "Number : " + richTextBox16.Text + " : Name : " + richTextBox2.Text + " : Compay : " + richTextBox1.Text + "  :Funcition: " + richTextBox3.Text + "  :Id: " + richTextBox4.Text + " : E-mail : " + richTextBox8.Text + " : Vessel : " + comboBox1.Text.Trim() + " : Project : " + richTextBox9.Text + " : ASO : " + richTextBox10.Text + " : NR-34 : " + richTextBox11.Text + " : NR-10 : " + richTextBox12.Text + " : NR-33 : " + richTextBox13.Text + " : NR-35 : " + richTextBox14.Text;
            int numeroLinha = Convert.ToInt32(Linhas);

            ArrayList linhas = new ArrayList();

            if (File.Exists(nomeArquivo))
            {
                try
                {
                    rdr = new StreamReader(nomeArquivo);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao acessar o arquivo : " + ex.Message);
                    return;
                }
            }
            else
            {
                MessageBox.Show("O arquivo : " + nomeArquivo + " não existe...");
                return;
            }
            string linha;

            while ((linha = rdr.ReadLine()) != null)
            {
                try
                {
                    linhas.Add(linha);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao acessar o arquivo : " + ex.Message);
                    return;
                }
            }
            rdr.Close();

            if (linhas.Count > numeroLinha)
                linhas.Insert(numeroLinha, textoInserir);
            else
                linhas.Add(textoInserir);

            StreamWriter wrtr = new StreamWriter(nomeArquivo);

            foreach (string strNewLine in linhas)
            {
                wrtr.WriteLine(strNewLine);
            }
            wrtr.Close();
            textoInserir = "";
            // txtArquivo.Text = AbreArquivoTexto(nomeArquivo);

        }
        private void check_id_onboard()
        {
            try
            {
                //  Number: 1 : Name: Cristiano: Compay: Googlemarine: Funcition: Engenheiro: Id: 111098414 : E - mail : 1 : Vessel: Googlemarine: Project: 190603 : ASO: 22 / 02 / 2023 : NR - 34 : 22 / 02 / 2023 : Vaccine - 1 : 22 / 02 / 2023 : Vaccine - 2 : 22 / 02 / 2023 : Booster vaccine : 22 / 02 / 2023 : Bloqueado: GUSTAVO: Falta da quarta dose da vacina
                // int rich2 = Int32.Parse(label3.Text);
                // int rich4;
                //  int lab3 = Int16.Parse(label3.Text);
                for (int i = 0; i < Int32.Parse(label27.Text); i++)
                {

                    string secondLine = File.ReadLines(@"C:\compartilhamento\data_txt\data.txt").ElementAtOrDefault(i);
                    
                    //  MessageBox.Show(secondLine.Split(':')[5].Trim() + "//" + comboBox1.SelectedItem);
                //    string selected = this.comboBox1.SelectedItem.ToString();
                //    label63.Text = selected;
                 //   MessageBox.Show(selected);


                    if (secondLine != null)
                    {
                        if (secondLine.Split(':')[1].Trim() == richTextBox16.Text.Trim())
                        {
                            id_onboard = true;
                            richTextBox16.Text = "";
                            richTextBox2.Text = "";
                            richTextBox3.Text = "";
                            richTextBox4.Text = "";
                            richTextBox1.Text = "";
                            richTextBox8.Text = "";
                            richTextBox10.Text = "";
                            richTextBox11.Text = "";
                            richTextBox12.Text = "";
                            richTextBox13.Text = "";
                            richTextBox14.Text = "";
                            richTextBox15.Text = "";
                            maskedTextBox1.Text = "";
                            maskedTextBox2.Text = "";
                            maskedTextBox3.Text = "";
                            maskedTextBox4.Text = "";
                            maskedTextBox5.Text = "";
                            button2.Enabled = false;
                        }

                    }


                }
                if (label27.Text == "0")
                {
                    id_onboard = false;
                }
            }
            catch
            {
                //  id_onboard = false;
            }
        }

        private void check_id_onboard2()
        {
            try
            {
                id_onboard2 = false;
                //  Number: 1 : Name: Cristiano: Compay: Googlemarine: Funcition: Engenheiro: Id: 111098414 : E - mail : 1 : Vessel: Googlemarine: Project: 190603 : ASO: 22 / 02 / 2023 : NR - 34 : 22 / 02 / 2023 : Vaccine - 1 : 22 / 02 / 2023 : Vaccine - 2 : 22 / 02 / 2023 : Booster vaccine : 22 / 02 / 2023 : Bloqueado: GUSTAVO: Falta da quarta dose da vacina
                // int rich2 = Int32.Parse(label3.Text);
                // int rich4;
                //  int lab3 = Int16.Parse(label3.Text);

                int qtdLinhas = File.ReadLines(@"C:\compartilhamento\data_txt\data.txt").Count();
                    for (int i = 0; i < qtdLinhas; i++)
                    {

                        string secondLine = File.ReadLines(@"C:\compartilhamento\data_txt\data.txt").ElementAtOrDefault(i);
                        //  MessageBox.Show(secondLine.Split(':')[5].Trim());
                        if (secondLine != null)
                        {

                            if (secondLine.Split(':')[1].Trim() == richTextBox16.Text.Trim() && secondLine.Split(':')[5].Trim() != label63.Text.Trim())
                            {
                                
                                //  MessageBox.Show("diferente");
                         
                                id_onboard2 = true;

                            }

                        }


                    }
               

            }
            catch
            {
                //  id_onboard = false;
            }
        }


        private void check_if_exist_id()
        {
            try
            {
                //  Number: 1 : Name: Cristiano: Compay: Googlemarine: Funcition: Engenheiro: Id: 111098414 : E - mail : 1 : Vessel: Googlemarine: Project: 190603 : ASO: 22 / 02 / 2023 : NR - 34 : 22 / 02 / 2023 : Vaccine - 1 : 22 / 02 / 2023 : Vaccine - 2 : 22 / 02 / 2023 : Booster vaccine : 22 / 02 / 2023 : Bloqueado: GUSTAVO: Falta da quarta dose da vacina
                int rich2 = Int32.Parse(label3.Text);
                int rich4;
                //  int lab3 = Int16.Parse(label3.Text);
                for (int i = 0; i < Int32.Parse(label3.Text); i++)
                {

                    string secondLine = File.ReadLines(@"C:\compartilhamento\data_txt\data2.txt").ElementAtOrDefault(i);
                    if (secondLine != null)
                    {
                        if (secondLine.Split(':')[9].Trim() == richTextBox4.Text)
                        {

                            try
                            {

                                pictureBox7.LoadAsync(@"C:\compartilhamento\data_picture\" + richTextBox4.Text.ToString() + ".jpg");

                            }
                            catch
                            {
                                pictureBox7.LoadAsync(@"C:\compartilhamento\data_picture\face.jpg");
                            }

                            if (ok_but2 == false)
                            {

                                try
                                {
                                    String local222 = secondLine.Split(':')[30].Trim();
                                    label34.Text = local222;
                                    if (local222 == place1[band])
                                    {
                                        local1.Checked = true;
                                        local2.Checked = false;
                                        local4.Checked = false;
                                    }
                                    if (local222 == place2[band])
                                    {
                                        local1.Checked = false;
                                        local2.Checked = true;
                                        local4.Checked = false;
                                    }
                                    if (local222 == place4[band])
                                    {
                                        local1.Checked = false;
                                        local2.Checked = false;
                                        local4.Checked = true;
                                    }
                                    richTextBox6.Text = secondLine.Split(':')[31];
                                    richTextBox7.Text = secondLine.Split(':')[32];

                                }
                                catch
                                {

                                }

                                /*
                                richTextBox16.Text = secondLine.Split(':')[1].Trim();
                                richTextBox2.Text = secondLine.Split(':')[3].Trim();
                                richTextBox3.Text = secondLine.Split(':')[7].Trim();
                                richTextBox4.Text = secondLine.Split(':')[9].Trim();
                                richTextBox1.Text = secondLine.Split(':')[5].Trim();
                                richTextBox8.Text = secondLine.Split(':')[11].Trim();
                                */


                                richTextBox2.Text = secondLine.Split(':')[3].Trim();
                                richTextBox3.Text = secondLine.Split(':')[7].Trim();
                                richTextBox4.Text = secondLine.Split(':')[9].Trim();
                                richTextBox1.Text = secondLine.Split(':')[5].Trim();
                                richTextBox16.Text = secondLine.Split(':')[1].Trim();
                                richTextBox8.Text = secondLine.Split(':')[11].Trim();

                                comboBox1.Items.Clear();



                                comboBox1.Items.Insert(0, secondLine.Split(':')[13].Trim());
                                comboBox1.SelectedIndex = 0;

                                if (secondLine.Split(':')[11].Trim() == "")
                                {
                                    richTextBox8.Text = "N/A";
                                }

                                maskedTextBox1.Visible = true;
                                maskedTextBox2.Visible = true;
                                maskedTextBox3.Visible = true;
                                maskedTextBox4.Visible = true;
                                maskedTextBox5.Visible = true;
                                /*
                                try
                                {
                                    string ok = secondLine.Split(':')[17].Trim();
                                    //  MessageBox.Show(ok);
                                    string teste = DateTime.ParseExact(ok, "M/d/yyyy", CultureInfo.InvariantCulture).ToString("dd/MM/yyyy");
                                    // MessageBox.Show(teste);
                                    maskedTextBox1.Text = teste;


                                    ok = secondLine.Split(':')[19].Trim();
                                    MessageBox.Show(ok);
                                    //  MessageBox.Show(ok);
                                    teste = DateTime.ParseExact(ok, "M/d/yyyy", CultureInfo.InvariantCulture).ToString("dd/MM/yyyy");
                                    // MessageBox.Show(teste);
                                    maskedTextBox2.Text = teste;
                                }
                                catch
                                {

                                }
                                */
                                // MessageBox.Show(secondLine);
                                // MessageBox.Show(secondLine.Split(':')[9].Trim());
                                maskedTextBox1.Text = secondLine.Split(':')[17];
                                maskedTextBox2.Text = secondLine.Split(':')[19];
                                maskedTextBox3.Text = secondLine.Split(':')[21];
                                maskedTextBox4.Text = secondLine.Split(':')[23];
                                maskedTextBox5.Text = secondLine.Split(':')[25];


                                if (secondLine.Split(':')[17].Trim() == "")
                                {
                                    maskedTextBox1.Text = "00/00/0000";
                                }

                                if (secondLine.Split(':')[19].Trim() == "")
                                {
                                    maskedTextBox2.Text = "00/00/0000";
                                }

                                if (secondLine.Split(':')[21].Trim() == "")
                                {
                                    maskedTextBox3.Text = "00/00/0000";
                                }
                                if (secondLine.Split(':')[23].Trim() == "")
                                {
                                    maskedTextBox4.Text = "00/00/0000";
                                }
                                if (secondLine.Split(':')[25].Trim() == "")
                                {
                                    maskedTextBox5.Text = "00/00/0000";
                                }


                                // richTextBox15.Text = "";

                                try
                                {
                                    richTextBox15.Text = secondLine.Split(':')[33];
                                }
                                catch
                                {
                                    // richTextBox15.Text = "";
                                }
                                // richTextBox6.Visible = true;
                                // richTextBox7.Visible = true;
                                // dateTimePicker1.Visible = false;
                                // dateTimePicker2.Visible = false;
                                //dateTimePicker1.Value.ToString("dd/MM/yyyy").ToString();
                                // dateTimePicker2.Value.ToString("dd/MM/yyyy").ToString();


                            }

                            if (secondLine.Split(':')[26].Trim() == "Bloqueado")
                            {
                                button2.Enabled = false;
                                button17.Enabled = false;
                                button7.Visible = false;
                                btloc.Visible = true;
                                label37.Text = secondLine.Split(':')[28];
                                label37.Visible = true;
                            }
                            else
                            {
                                button2.Enabled = true;
                                button17.Enabled = true;
                                button7.Visible = true;
                                btloc.Visible = false;
                                label37.Visible = false;
                            }
                            id_exist = true;

                            break;

                        }

                        label39.Text = i.ToString();
                        //  MessageBox.Show((i).ToString());

                        if (ok_but2 == false)
                        {
                            if (i == rich2 - 1 && secondLine.Split(':')[9].Trim() != richTextBox4.Text)
                            {
                                richTextBox16.Text = "";
                                richTextBox2.Text = "";
                                richTextBox3.Text = "";
                                richTextBox4.Text = "";
                                richTextBox1.Text = "";
                                richTextBox8.Text = "";
                                richTextBox10.Text = "";
                                richTextBox11.Text = "";
                                richTextBox12.Text = "";
                                richTextBox13.Text = "";
                                richTextBox14.Text = "";
                                richTextBox15.Text = "";
                                maskedTextBox1.Text = "";
                                maskedTextBox2.Text = "";
                                maskedTextBox3.Text = "";
                                maskedTextBox4.Text = "";
                                maskedTextBox5.Text = "";
                                maskedTextBox1.Visible = false;
                                maskedTextBox2.Visible = false;
                                maskedTextBox3.Visible = false;
                                maskedTextBox4.Visible = false;
                                maskedTextBox5.Visible = false;
                                button2.Enabled = false;
                                button17.Enabled = false;
                                id_exist = false;
                                MessageBox.Show("NÃO EXISTE CADASTRO COM ESTE NÚMERO!");

                            }

                        }
                    }

                }
            }
            catch
            {
                /// MessageBox.Show("NÃO A DADOS CADASTRADOS!");
            }

            //  try
            // {

            //  string sec = secondLine.Split(':')[26].Trim();

            // }
            //  }




            /*


            if (richTextBox4.Text.Length >= 5)
            {
                button2.Enabled = true;
                button17.Enabled = true;
                button29.Enabled = false;
                richTextBox4.ReadOnly = true;
                // richTextBox4.
                //   pictureBox1.Visible = false;
                //   panel4.Visible = false;



                panel11.Size = new Size(360, 355);
                panel11.Location = new System.Drawing.Point(95, 150);
                pictureBox7.Size = new Size(330, 320);
                pictureBox7.Location = new System.Drawing.Point(15, 18);



                richTextBox17.Text = "";
                //label37.Text = "Bloquear";





                try
                {
                    string path2 = richTextBox4.Text.ToString();
                    Stream stream = File.Open(@"C:\compartilhamento\data_picture\" + richTextBox4.Text + ".jpg", FileMode.Open,
                    FileAccess.Read, FileShare.Delete);
                    pictureBox7.Image = Image.FromStream(stream);
                    Bitmap bmp = new Bitmap(pictureBox7.Image);
                    bmp.RotateFlip(RotateFlipType.Rotate180FlipY);
                    pictureBox7.Image = bmp;
                    stream.Close();

                    //pictureBox7.Load(@"C:\compartilhamento\data_picture\" + path + ".jpg";
                    //  pictureBox7.Image = new Bitmap(@"C:\compartilhamento\data_picture\" + path + ".jpg");

                    // pictureBox7.Dispose();

                }
                catch (Exception e)
                {
                    MessageBox.Show("Não há foto cadastrada");
                    string path = richTextBox4.Text.ToString();
                    Stream stream = File.Open(@"C:\compartilhamento\data_picture\face.jpg", FileMode.Open,
                    FileAccess.Read, FileShare.Delete);
                    pictureBox7.Image = Image.FromStream(stream);

                    Bitmap bmp = new Bitmap(pictureBox7.Image);
                    bmp.RotateFlip(RotateFlipType.Rotate180FlipY);
                    pictureBox7.Image = bmp;
                    stream.Close();
                    richTextBox4.Text = "";
                    richTextBox4.ReadOnly = false;
                }
                catch
                {





                }

                ler_linha2();





            }
            */
        }
        private void richTextBox4_TextChanged(object sender, EventArgs e)
        {
            if (richTextBox4.Text.Length >= 5)
            {
                button15.Visible = true;

                if (label7.Text == button28.Text)
                {
                    button2.Enabled = true;
                    button17.Enabled = true;
                    button29.Enabled = true;
                }
                else
                {
                    button2.Enabled = false;
                    button17.Enabled = false;
                    button29.Enabled = false;
                }
            }
            else
            {
                button15.Visible = false;
                button2.Enabled = false;
                button17.Enabled = false;
                button29.Enabled = false;
            }

            if (richTextBox4.Text == "")
            {
                label46.Visible = true;
                button2.Enabled = false;
                button17.Enabled = false;
                button29.Enabled = false;
            }
            else
            {
                label46.Visible = false;
            }
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            band = 0;

            lname.Text = label_nome[0];
            lcompany.Text = label_emp[0];
            lfunc.Text = label_Function[0];
            lid.Text = label_Id[0];
            lvessel.Text = label_vessel[0];
            lproject.Text = label_porj[0];
            lv1.Text = label_vaccine1[0];
            lv2.Text = label_vaccine2[0];
            lbustter.Text = label_reforco[0];
            local1.Text = place1[0];
            local2.Text = place2[0];
            local3.Text = place3[0];
            local4.Text = place4[0];
            laccess.Text = label_acc[0];
            level_yellow.Text = label_yellow[0];
            level_green.Text = label_green[0];
            level_red.Text = label_red[0];
            lcheckin.Text = Label_initial[0];
            lcheckout.Text = Label_final[0];
            button1.Text = Label_Read_QRcode_On[0];
            button28.Text = Label_Read_QRcode_Off[0];
            button2.Text = Label_Create_QRcode[0];
            button3.Text = Label_Show_data[0];
            button4.Text = Label_Save_data[0];
            //button5.Text = Label_Config[0];
            button6.Text = Label_wifi[0];
            button17.Text = Label_email[0];
            button19.Text = Label_Mostrar_checkin[0];
            button21.Text = Label_reset[0];
            button22.Text = Label_fechar[0];
            button8.Text = Label_entrada[0];
            button9.Text = Label_saida[0];
            button10.Text = Label_cancel[0];
            label23.Text = label_onboard[0];
            button27.Text = Label_reset_project[0];
            button29.Text = label_reg[0];
            label7.Text = Label_Read_QRcode_On[0];
            label6.Text = onboard[0];
            label_cad.Text = label_cad1[0];
            label53.Text = Label_53[0];
            button41.Text = bt_41[0];
            button42.Text = bt_42[0];
            button43.Text = bt_43[0];
            button44.Text = bt_44[0];
            button45.Text = bt_45[0];
            regs.Text = bt_regis[0];
            _cad = cad_mode[0];
            _read = read_mode[0];
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            band = 1;
            lname.Text = label_nome[1];
            lcompany.Text = label_emp[1];
            lfunc.Text = label_Function[1];
            lid.Text = label_Id[1];
            lvessel.Text = label_vessel[1];
            lproject.Text = label_porj[1];
            lv1.Text = label_vaccine1[1];
            lv2.Text = label_vaccine2[1];
            lbustter.Text = label_reforco[1];
            local1.Text = place1[1];
            local2.Text = place2[1];
            local3.Text = place3[1];
            local4.Text = place4[1];
            laccess.Text = label_acc[1];
            level_yellow.Text = label_yellow[1];
            level_green.Text = label_green[1];
            level_red.Text = label_red[1];
            lcheckin.Text = Label_initial[1];
            lcheckout.Text = Label_final[1];
            button1.Text = Label_Read_QRcode_On[1];
            button28.Text = Label_Read_QRcode_Off[1];
            button2.Text = Label_Create_QRcode[1];
            button3.Text = Label_Show_data[1];
            button4.Text = Label_Save_data[1];
            // button5.Text = Label_Config[1];
            button6.Text = Label_wifi[1];
            button17.Text = Label_email[1];
            button19.Text = Label_Mostrar_checkin[1];
            button21.Text = Label_reset[1];
            button22.Text = Label_fechar[1];
            button8.Text = Label_entrada[1];
            button9.Text = Label_saida[1];
            button10.Text = Label_cancel[1];
            label23.Text = label_onboard[1];
            button27.Text = Label_reset_project[1];
            button29.Text = label_reg[1];
            label7.Text = Label_Read_QRcode_On[1];
            label6.Text = onboard[1];
            label_cad.Text = label_cad1[1];
            label53.Text = Label_53[1];
            button41.Text = bt_41[1];
            button42.Text = bt_42[1];
            button43.Text = bt_43[1];
            button44.Text = bt_44[1];
            button45.Text = bt_45[1];
            regs.Text = bt_regis[1];
            _cad = cad_mode[1];
            _read = read_mode[1];


        }

        private void richTextBox4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                richTextBox1.Focus();
                // MessageBox.Show("You pressed enter! Good job!");
            }
        }

        private void richTextBox4_MouseDoubleClick(object sender, MouseEventArgs e)
        {

            local1.Checked = false;
            local2.Checked = false;
            local3.Checked = false;
            richTextBox6.Text = " ";
            richTextBox7.Text = " ";
            richTextBox15.Text = "";
            check_if_exist_id();
            if (richTextBox16.Text != "")
            {
                richTextBox6.Visible = true;
                richTextBox7.Visible = true;
                dateTimePicker1.Visible = false;
                dateTimePicker2.Visible = false;

               // button29.Enabled = false;
            }
            int lista_2 = 0;
            for (int i = 0; i < Int32.Parse(label3.Text); i++)
            {





                //
                lista_2++;


            }



        }


        private void button11_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click_2(object sender, EventArgs e)
        {
            // panel8.Location= new Point(Location.X,Location.Y);//Point(521,155);
            panel8.Size = new Size(490, 353); //new Size(1030, 600);
            wi++;
            if (wi == 1)
            {
                //online = true;
                lview_AP.Items.Clear();
                wifi = new Wifi();
                List<AccessPoint> aps = wifi.GetAccessPoints();
                foreach (AccessPoint ap in aps)
                {
                    ListViewItem lvItem = new ListViewItem(ap.Name);
                    lvItem.SubItems.Add(ap.SignalStrength + "%");
                    lvItem.Tag = ap;
                    lview_AP.Items.Add(lvItem);
                }


                panel8.Visible = true;
            }
            if (wi == 2)
            {
                // online = false;
                if (textBox7.SelectionLength >= 0)
                {
                    textBox7.Focus();
                    textBox7.Text = "";
                }
                panel8.Visible = false;

                wi = 0;
            }

            // System.Windows.Forms.Application.Exit();
            //   this.Close();
        }
        private void testa_local()
        {

            if (local1.Checked == false)
            {
                // checado = 1;
                local1val = 0;
            }

            if (local2.Checked == false)
            {
                // checado = 1;
                local2val = 0;
            }

            if (local3.Checked == false)
            {
                // checado = 1;
                local3val = 0;
            }

            if (local4.Checked == false)
            {
                // checado = 1;
                local4val = 0;
            }
            if (level_yellow.Checked == false)
            {
                levelyellow = 0;
            }
            if (level_green.Checked == false)
            {
                levelgreen = 0;
            }
            if (level_red.Checked == false)
            {
                levelred = 0;
            }
        }
        private void local1_CheckedChanged(object sender, EventArgs e)
        {
            if (richTextBox4.Text != "")
            {



                if (local1.Checked == true)
                {
                    checado = 1;
                    local1val = 1;
                    local2.Checked = false;
                    local4.Checked = false;
                    label50.Visible = false;
                }
                if (local1.Checked == false && local2.Checked == false && local3.Checked == false && local4.Checked == false)
                {
                    checado = 0;
                    local1val = 0;
                    label50.Visible = true;
                }

                testa_local();
            }
            else
            {
                local1.Checked = false;
                local2.Checked = false;
                local4.Checked = false;
            }
        }

        private void local2_CheckedChanged(object sender, EventArgs e)
        {
            if (richTextBox4.Text != "")
            {
                if (local2.Checked == true)
                {
                    checado = 1;
                    local2val = 1;
                    local1.Checked = false;
                    local4.Checked = false;
                    label50.Visible = false;
                }
                if (local1.Checked == false && local2.Checked == false && local3.Checked == false && local4.Checked == false)
                {
                    checado = 0;
                    local2val = 0;
                    label50.Visible = true;
                }

                testa_local();
            }
            else
            {
                local1.Checked = false;
                local2.Checked = false;
                local4.Checked = false;
            }
        }

        private void local4_CheckedChanged(object sender, EventArgs e)
        {
            if (richTextBox4.Text != "")
            {
                if (local4.Checked == true)
                {
                    checado = 1;
                    local4val = 1;
                    local1.Checked = false;
                    local2.Checked = false;
                    label50.Visible = false;
                }
                if (local1.Checked == false && local2.Checked == false && local3.Checked == false && local4.Checked == false)
                {
                    checado = 0;
                    local4val = 0;
                    label50.Visible = true;
                }

                testa_local();
            }
            else
            {
                local1.Checked = false;
                local2.Checked = false;
                local4.Checked = false;
            }
        }

        private void button27_Click(object sender, EventArgs e)
        {
            password frm = new password();
            frm.Show();
        }

        private void richTextBox17_Click(object sender, EventArgs e)
        {
            label37.Visible = false;
        }

        private void label37_Click(object sender, EventArgs e)
        {
            label37.Visible = false;
        }

        private void richTextBox17_TextChanged(object sender, EventArgs e)
        {
            if (richTextBox17.Text != "")
            {
                r7 = true;
            }
        }

        private void richTextBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = "%&'çÇ´`~^áéíãť@#$¨*".Contains(e.KeyChar.ToString());
            e.KeyChar = Char.ToUpper(e.KeyChar);
        }

        private void richTextBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = "%&'çÇ´`~^áéíãť@#$¨*".Contains(e.KeyChar.ToString());
            e.KeyChar = Char.ToUpper(e.KeyChar);
        }

        private void richTextBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = "%&'çÇ´`~^áéíãť@#$¨*".Contains(e.KeyChar.ToString());
            e.KeyChar = Char.ToUpper(e.KeyChar);
        }

        private void richTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = "%&'çÇ´`~^áéíãť@#$¨*".Contains(e.KeyChar.ToString());
            e.KeyChar = Char.ToUpper(e.KeyChar);
        }

        private void richTextBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = "%&'çÇ´`~^áéíãť#$¨*".Contains(e.KeyChar.ToString());
            // e.KeyChar = Char.ToUpper(e.KeyChar);
        }

        private void richTextBox15_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = "%&'çÇ´`~^áéíãť@#$¨*".Contains(e.KeyChar.ToString());
            e.KeyChar = Char.ToUpper(e.KeyChar);
        }

        private void richTextBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                richTextBox3.Focus();
                // MessageBox.Show("You pressed enter! Good job!");
            }
        }

        private void richTextBox3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                richTextBox4.Focus();
                // MessageBox.Show("You pressed enter! Good job!");
            }
        }

        private void richTextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                richTextBox8.Focus();
                // MessageBox.Show("You pressed enter! Good job!");
            }
        }

        private void richTextBox8_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                maskedTextBox1.Focus();
                // MessageBox.Show("You pressed enter! Good job!");
            }
        }

        private void maskedTextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                maskedTextBox2.Focus();
                // MessageBox.Show("You pressed enter! Good job!");
            }
        }

        private void maskedTextBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                maskedTextBox3.Focus();
                // MessageBox.Show("You pressed enter! Good job!");
            }
        }

        private void maskedTextBox3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                maskedTextBox4.Focus();
                // MessageBox.Show("You pressed enter! Good job!");
            }
        }

        private void maskedTextBox4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                maskedTextBox5.Focus();
                // MessageBox.Show("You pressed enter! Good job!");
            }
        }

        private void maskedTextBox5_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                richTextBox15.Focus();
                // MessageBox.Show("You pressed enter! Good job!");
            }
        }

        private void richTextBox15_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                richTextBox2.Focus();
                // MessageBox.Show("You pressed enter! Good job!");
            }
        }


        public void sendGmail()
        {
            text1 = richTextBox2.Text;
            text2 = richTextBox1.Text;
            text5 = comboBox1.Text;
            text6 = dateTimePicker1.Value.ToString("ddd, dd MMM yyyy");
            text7 = dateTimePicker2.Value.ToString("ddd, dd MMM yyyy");
            text8 = richTextBox8.Text;

            // foi = true;
            // panel1.BackColor = Color.Yellow;
            // panel14.BackColor = Color.Yellow;
            //  timer1.Stop();









            int width = panel4.Size.Width;
            int height = panel4.Size.Height;
            Bitmap bm = new Bitmap(width, height);
            panel4.DrawToBitmap(bm, new System.Drawing.Rectangle(0, 0, width, height));
            bm.Save(@"C:\compartilhamento\data_picture\qr\Qrcode10.png", ImageFormat.Png);
            //label4.Visible = false;


            //pictureBox1.Image.Save(@"C:\data_picture\qr\Qrcode10.png", ImageFormat.Png);
            System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage();
            mail.From = new System.Net.Mail.MailAddress("dof.qrcode@gmail.com");
            try
            {
                mail.To.Add(text8); // para
            }
            catch (Exception ex)
            {

                MessageBox.Show("erro send email", ex.Message);
            }
            mail.Subject = "Liberação " + text5 + " - " + text6 + " até " + text7 + " -" + text2;//id.Trim(); // assunto
            mail.Body = "<br>" + text1 + ",<br> " + "<br>Seja bem vindo a DOF<br> <br>Segue o QR Code:<br>Apresente este voucher de liberação, para acessar a embarcação <b>" + text5 + "</b> <font color=#FF0000>, juntamente com documento de identificação.";//id.Trim(); // mensagem
            mail.IsBodyHtml = true;
            System.Net.Mail.Attachment attachment;
            mail.Attachments.Add(new System.Net.Mail.Attachment(@"C:\compartilhamento\data_picture\qr\Qrcode10.png"));




            try
            {
                using (var smtp = new SmtpClient("smtp.gmail.com"))
                {
                    smtp.EnableSsl = true; // GMail requer SSL
                    smtp.Port = 587;       // porta para SSL
                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network; // modo de envio
                    smtp.UseDefaultCredentials = false; // vamos utilizar credencias especificas

                    // seu usuário e senha para autenticação
                    smtp.Credentials = new NetworkCredential("dof.qrcode@gmail.com", "wammvtijopzdghps");
                    //    smtp.Credentials = new NetworkCredential("alarm.boat@gmail.com", "damhkxldmyegacvi");
                    //  MessageBox.Show("acesso liberado");
                    smtp.Send(mail);
                    MessageBox.Show("E-mail sent with success");
                    //  soma2 = 0;
                    //   soma = 0;


                }

            }
            catch (Exception ex)
            {

                MessageBox.Show("erro send email", ex.Message);
            }

            mail.Attachments.Dispose();


            // timer1.Start();

        }
        private void button17_Click(object sender, EventArgs e)
        {
            
            rec = false;
            Boolean tempo = false;
            if (dateTimePicker1.Visible == true)
            {
                if (dateTimePicker2.Value.Date >= dateTimePicker1.Value.Date)
                {
                    tempo = true;
                }
                else
                {
                    tempo = false;
                    MessageBox.Show("A DATA FINAL ESTÁ MENOR QUE A DATA INICIAL, CORRIJA POR FAVOR!");

                }
            }

            if (dateTimePicker1.Visible == false)
            {
                tempo = true;
            }
            if (tempo == true)
            {
                timer10.Enabled = false;
                timer12.Enabled = false;
                ProcessStartInfo startInfo = new ProcessStartInfo("cmd", "/K netsh interface set interface \"Ethernet\" DISABLED");
                startInfo.RedirectStandardOutput = true;
                startInfo.UseShellExecute = false;
                // Do not create the black window.
                startInfo.CreateNoWindow = true;
                startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                Process.Start(startInfo);

                check_id_onboard();
                if (id_onboard == false)
                {
                    if (dateTimePicker1.Visible == true)
                    {
                        richTextBox6.Text = dateTimePicker1.Value.ToString("dd/MM/yyyy").ToString().Trim();
                        richTextBox7.Text = dateTimePicker2.Value.ToString("dd/MM/yyyy").ToString().Trim();
                    }


                    alterado();
                    compare_id();
                    check_if_exist_id();

                    if (id_exist == true)
                    {
                        richTextBox10.Text = maskedTextBox1.Text;
                        richTextBox11.Text = maskedTextBox2.Text;
                        richTextBox12.Text = maskedTextBox3.Text;
                        richTextBox13.Text = maskedTextBox4.Text;
                        richTextBox14.Text = maskedTextBox5.Text;
                        textBox13.Text = richTextBox1.Text;





                        // 
                        //  if (textBox7.SelectionLength >= 0)
                        // {
                        textBox7.Focus();
                        textBox7.Text = "";
                        // }
                        /*
                        String[] Label_initial = { "Inicio", "Check-in" };
                        String[] Label_final = { "Fim", "Check-out" };
                        String[] Label_Read_QRcode_On = { "Ler Qrcode Ligado", "Read QRcode On" };
                        String[] Label_Read_QRcode_Off = { "Ler Qrcode Desligado", "Read QRcode Off" };
                        String[] Label_Create_QRcode = { "Imprimir Qrcode:", "Print QRcode" };
                        String[] Label_Show_data = { "Mostrar banco de dados:", "Show DataBase" };
                        String[] Label_Save_data = { "Salvar banco de dados:", "Save Database Backup" };
                        String[] Label_Config = { "Configurações:", "Settings" };
                        String[] Label_wifi = { "Conexão Wi-Fi:", "Wi-Fi connection" };
                        String[] Label_email = { "Enviar Qrcode por E-mail:", "Send Qr Code  by E-mail" };
                        String[] Label_Mostrar_checkin = { "Mostrar Check-in:", "Show Check-in" };
                        String[] Label_fechar = { "Desligar:", "Turn Off" };
            */
                        if (band == 0)
                        {
                            button2.Text = Label_Create_QRcode[0];
                        }
                        else
                        {
                            button2.Text = Label_Create_QRcode[1];
                        }





                        // var parameterDate2_initial = DateTime.ParseExact(dateTimePicker1.Value.Date.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                        // var parameterDate2_final = DateTime.ParseExact(dateTimePicker2.Value.Date.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);

                        if (resultado == 0)
                        {
                            if (textBox7.SelectionLength >= 0)
                            {
                                // textBox7.Focus();
                                // textBox7.Text = "";
                            }
                            // dataGridView1.Visible = false;
                            if (richTextBox1.Text != " " && richTextBox2.Text != " " && richTextBox3.Text != " " && richTextBox4.Text != " " && comboBox1.Text != " " && richTextBox8.Text != "" && checado == 1 && maskedTextBox1.Text != "  /  /"
                                && maskedTextBox2.Text != "  /  /" && maskedTextBox3.Text != "  /  /" && maskedTextBox4.Text != "  /  /" && maskedTextBox5.Text != "  /  /")  //  /  /
                            {



                                compare_aso();

                                if (aso_1 == 0)
                                {



                                    maskedTextBox1.Visible = false;
                                    maskedTextBox2.Visible = false;
                                    maskedTextBox3.Visible = false;
                                    maskedTextBox4.Visible = false;
                                    maskedTextBox5.Visible = false;




                                    read_write();
                                    confere = 1;
                                    lb4.Visible = true;
                                    label5.Visible = true;
                                    panel10.Visible = true;
                                    label5.Text = richTextBox2.Text;




                                    qr_generate = "Qrcode Sent by E-mail";

                                    //
                                    // CarregarPlanilha2();
                                    carrega_planilha2_txt();
                                    //  create_qrcode();
                                    ProcessStartInfo startInfo2 = new ProcessStartInfo("cmd", "/K netsh interface set interface \"Ethernet\" ENABLED");
                                    startInfo2.RedirectStandardOutput = true;
                                    startInfo2.UseShellExecute = false;
                                    // Do not create the black window.
                                    startInfo2.CreateNoWindow = true;
                                    startInfo2.WindowStyle = ProcessWindowStyle.Hidden;
                                    Process.Start(startInfo2);
                                    create_qrcode_new();
                                    sendGmail();
                                    
                                    //  print_qrcode();
                                }





                                //
                                //string teste = "Name : " + richTextBox2.Text + " : Compay : " + richTextBox1.Text + " :Funcition: " + richTextBox3.Text + "  :Id: " + richTextBox4.Text + " : E-mail : " + richTextBox8.Text + " : Vessel :" + richTextBox5.Text + " : Project : " + richTextBox9.Text + ": ASO : " + richTextBox10.Text + " : NR-34 : " + richTextBox11.Text + " : Vaccine-1 : " + richTextBox12.Text + " : Vaccine-2 : " + richTextBox13.Text + " : Booster vaccine : " + richTextBox14.Text;
                                //  atualiza_compartilhamento();

                                // 

                                // else
                                //  {
                                //  MessageBox.Show(id_check[band]);
                                //  }
                                checado = 0;

                            }
                            else
                            {
                                if (band == 0)
                                {
                                    MessageBox.Show("Favor preencher todos os campos");
                                }

                                if (band == 1)
                                {
                                    MessageBox.Show("Please complete all informations places");
                                }
                            }

                        }


                        if (resultado == 1)
                        {
                            MessageBox.Show("ID duplicated");
                        }
                        textBox7.Focus();
                        textBox7.Text = " ";
                    }
                    ok_but2 = false;

                }
                else
                {
                    MessageBox.Show("ESTA PESSOA ESTÁ A BORDO! SÓ É PERMITIDO ENVIAR EMAIL SE A PESSOA ESTIVER FORA DA EMBARCAÇÃO");
                }
            }
            rec = true;
            timer10.Enabled = true;
            timer12.Enabled = true;
        }

        private void timer6_Tick(object sender, EventArgs e)
        {
            //  label52.Text = count2.ToString();




            if (online_ == true)
            {
                //  panel12.BackColor = Color.GreenYellow;

                // label40.Visible = false;

            }

            if (online_ == false)
            {
                //  panel12.BackColor = Color.Red;
                // label40.Visible = false;
            }

            online_ = false;
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged_1(object sender, EventArgs e)
        {
            if (tempo == 0)
            {
                tempo = 1;
                pictureBox7.Image = Properties.Resources.barcode1;
                timer4.Start();

            }
        }
        int fil = 0;
        private void button34_Click(object sender, EventArgs e)
        {
            //  MessageBox.Show(libera.ToString());
            // if (libera == true)
            // {
            var drives = DriveInfo.GetDrives().Where(drive => drive.IsReady && drive.DriveType == DriveType.Removable);
            if (drives.FirstOrDefault() != null)
            {

                object misValue = System.Reflection.Missing.Value;
                String line;

                //   foreach (DataGridViewRow item in this.dataGridView1.SelectedRows)
                //    {
                //dataGridView1.Rows.RemoveAt(item.Index);
                //   }

                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);

                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

                // app.Visible = false;

                worksheet = workbook.Sheets["Planilha1"];
                //  worksheet = workbook.ActiveSheet;

                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }
                // storing Each row and column value to excel sheet  
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }




                string folder = @"D:\CRIPTOQRCODE_backup\"; //nome do diretorio a ser criado

                //Se o diretório não existir...

                if (!Directory.Exists(folder))
                {

                    //Criamos um com o nome folder
                    Directory.CreateDirectory(folder);

                }

                if (Directory.Exists(folder))
                {
                    fil++;

                    worksheet.SaveAs(@"D:\CRIPTOQRCODE_backup\" + fil.ToString() + "FILTRO.xlsx", Type.Missing);
                    //  app.

                    app.Quit();

                    System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel");
                    foreach (System.Diagnostics.Process p in process)
                    {
                        if (!string.IsNullOrEmpty(p.ProcessName))
                        {
                            try
                            {
                                p.Kill();
                            }
                            catch { }
                        }
                    }
                }

                MessageBox.Show("Backup concluido com sucesso!");
            }
            else
            {
                MessageBox.Show("No Pendrive found..");
                if (textBox7.SelectionLength >= 0)
                {
                    // textBox7.Focus();
                    //textBox7.Text = "";
                }
            }
            ///  }
        }

        private void richTextBox3_TextChanged(object sender, EventArgs e)
        {
            if (richTextBox3.Text == "")
            {
                label45.Visible = true;
            }
            else
            {
                label45.Visible = false;
            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            if (richTextBox1.Text == "")
            {
                label47.Visible = true;
            }
            else
            {
                label47.Visible = false;
            }
        }

        private void richTextBox8_TextChanged(object sender, EventArgs e)
        {
            if (richTextBox8.Text == "")
            {
                // label48.Visible = true;
            }
            else
            {
                //  label48.Visible = false;
            }
        }

        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            timer10.Enabled = false;
            timer12.Enabled = false;
            //Criar um MessageBox com os botões Sim e Não e deixar o botão 2(Não) selecionado por padrão e comparar o botão apertado
            if (DialogResult.Yes == MessageBox.Show("TEM CERTEZA QUE DESEJA DAR SAIDA MANUAL A ESTA PESSOA?", "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
            {
               
                string bio = listBox1.SelectedItem.ToString().Trim();
                // MessageBox.Show(bio);
                string teste = "";//"Name : " + richTextBox2.Text + " : Compay : " + richTextBox1.Text + " :Funcition:  " + richTextBox3.Text + "  :Id: " + richTextBox4.Text + " : E-mail : " + richTextBox8.Text + " : Vessel : " + richTextBox5.Text + " : Project : " + richTextBox9.Text + " : ASO : " + richTextBox10.Text + " : NR-34 : " + richTextBox11.Text + " : Vaccine-1 : " + richTextBox12.Text + " : Vaccine-2 : " + richTextBox13.Text + " : Booster vaccine : " + richTextBox14.Text;
                int ver = 0;
                string filePath = @"C:\compartilhamento\data_txt\data.txt";
               

                string tempFile = Path.GetTempFileName();

                using (var sr = new StreamReader(filePath))
                {
                    using (var sw = new StreamWriter(tempFile))
                    {
                        string line;

                        while ((line = sr.ReadLine()) != null)
                        {
                            if (line != listBox1.SelectedItem.ToString().Trim())
                                sw.WriteLine(line);
                        }
                        sr.Close();
                    }

                }


                string line_to_delete = bio;
                var oldLines = System.IO.File.ReadAllLines(@"C:\compartilhamento\data_txt\data.txt");
                var newLines = oldLines.Where(line => !line.Contains(line_to_delete));
                System.IO.File.WriteAllLines(@"C:\compartilhamento\data_txt\data.txt", newLines);

                out_by_user();




                // File.Delete(filePath);
                // File.Move(tempFile, filePath);
                // File.
               // ler_linha();



                while (listBox1.SelectedItems.Count > 0)
                {



                    listBox1.Items.Remove(listBox1.SelectedItems[0]);
                }

                label67.Text = listBox1.SelectedItems.Count.ToString();

                // Get first 12 characters substring from a string    
                //   string authorName = bio.Substring(0,5);
                //Sua rotina de exclusão
                //Confirmando exclusão para o usuário
                //   MessageBox.Show("Registro apagado com sucesso", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //  MessageBox.Show(authorName, "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                // myThread.Abort();
                checa_host();
                //   MessageBox.Show("listBox1_MouseDoubleClick");
                //atualiza_compartilhamento();

                if (label27.Text == "0")
                {
                    panel6.Visible = false;

                }
            }
            else
            {
                // MessageBox.Show("Registro cancelado");
            }

            // ler_linha();

            //  }

            timer10.Enabled = true;
            timer12.Enabled = true;
        }

        private void listBox1_MouseClick(object sender, MouseEventArgs e)
        {
            pictureBox4.Visible = true;

            try
            {


                if (listBox1.SelectedItem != null) {

                    string path = listBox1.SelectedItem.ToString();

                    string s = path;
                    string id = s.Split(":".ToCharArray())[8];
                    // MessageBox.Show(id);
                    pictureBox4.Visible=true; 
                    pictureBox4.Load(@"C:\compartilhamento\data_picture\" + id.Trim() + ".jpg");
                }
            }
            catch
            {
                pictureBox4.Visible = true;
                pictureBox4.Load(@"C:\compartilhamento\data_picture\face.jpg");
            }
        }
        int click_ = 0;
        private void pictureBox4_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            click_++;
            if (click_ == 1)
            {
                pictureBox4.Size = new Size(500, 500);
                pictureBox4.Location = new System.Drawing.Point(600, 300);
            }
            if (click_ == 2)
            {
                pictureBox4.Size = new Size(170, 140);
                pictureBox4.Location = new System.Drawing.Point(1544, 6);
                click_ = 0;//847; 56
            }
        }

        private void pictureBox4_MouseDoubleClick_1(object sender, MouseEventArgs e)
        {
            click_++;
            if (click_ == 1)
            {
                pictureBox4.Size = new Size(500, 500);
                pictureBox4.Location = new System.Drawing.Point(600, 300);
            }
            if (click_ == 2)
            {
                pictureBox4.Size = new Size(170, 140);
                pictureBox4.Location = new System.Drawing.Point(1544, 6);
                click_ = 0;//847; 56
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            if (textBox7.SelectionLength >= 0)
            {
                textBox7.Focus();
                textBox7.Text = "";
            }
            panel6.Visible = false;
        }

        private void richTextBox16_TextChanged(object sender, EventArgs e)
        {
            if (richTextBox16.Text == "")
            {
                button7.Visible = false;
                btloc.Visible = false;
                button15.Visible = false;

            }



        }

        private void richTextBox4_Click(object sender, EventArgs e)
        {
            if (richTextBox4.Text == "")
            {
                button15.Visible = false;
            }
        }
        private void GetBiosInformation()
        {
            string relDt = "";
            try
            {
                string login = System.IO.File.ReadAllText(@"C:\compartilhamento\login.txt");
                ManagementObjectSearcher mSearcher = new ManagementObjectSearcher("SELECT SerialNumber, SMBIOSBIOSVersion, ReleaseDate FROM Win32_BIOS");
                ManagementObjectCollection collection = mSearcher.Get();
                foreach (ManagementObject obj in collection)
                {

                    // MessageBox.Show((string)obj["SerialNumber"]);
                    // textBox13.Text = (string)obj["SerialNumber"];
                    // lblBiosSerial.Text = (string)obj["SerialNumber"];
                    // lblBiosVersion.Text = (string)obj["SMBIOSBIOSVersion"];
                    relDt = (string)obj["ReleaseDate"];
                    DateTime dt = ManagementDateTimeConverter.ToDateTime(relDt);
                    //  lblBiosDate.Text = dt.ToString("dd-MMM-yyyy");//date format
                    if (login.Trim() != (string)obj["SerialNumber"])
                    {
                        panel9.Visible = true;
                        panel9.Size = new Size(1693, 997);
                        //36; 632
                    }
                    else
                    {

                        panel9.Visible = false;
                    }
                }
            }
            catch (Exception ex)
            {
                // Console.WriteLine(ex.ToString());
            }


        }
        private void button5_Click(object sender, EventArgs e)
        {
            check_id_onboard2();
          //  check_id_onboard();
            //File.WriteAllText(@"C:\teste\PROJETO.txt", "teste1" + "\r\n" + "teste1");


            /*
            String l = "";
            bool ESIM = false;
            bool dois = false;
            bool tres = false;
            string nume = "";
            string[] lines = File.ReadAllLines(@"C:\compartilhamento\data_txt\
            
           ");
            id_1 = 0;
            for (int i = 0; i < lines.Length; i++)
            {

                if (lines[i].Split(':')[9].Trim() == richTextBox4.Text.Trim())
                {




                    l = lines[i].Split(':')[9].Trim();
                    if (lines[i].Split(':')[1].Trim() != richTextBox16.Text.Trim()) {
                        MessageBox.Show("O NÚMERO DA IDENTIDADE" + lines[i].Split(':')[9].Trim() + " JÁ ESTÁ CADASTRADO NO ACESSO DE NÚMERO " + lines[i].Split(':')[1].Trim());
                    }
                    ESIM = true;
                    richTextBox4.Text = lines[int.Parse(richTextBox16.Text) - 1].Split(':')[9].Trim();
                    string text4 = "Number : " + richTextBox16.Text + " : Name : " + richTextBox2.Text + " : Compay : " + richTextBox1.Text + " :Funcition:  " + richTextBox3.Text + "  :Id: " + richTextBox4.Text + " : E-mail : " + richTextBox8.Text + " : Vessel : " + richTextBox5.Text + " : Project : " + richTextBox9.Text + " : ASO : " + maskedTextBox1.Text + " : NR-34 : " + maskedTextBox2.Text + " : Vaccine-1 : " + maskedTextBox3.Text + " : Vaccine-2 : " + maskedTextBox4.Text + " : Booster vaccine : " + maskedTextBox5.Text + " : " + bb + " : " + comuser.Text + " :" + richTextBox17.Text;
                    string text = File.ReadAllText(@"C:\compartilhamento\data_txt\data2.txt");
                    text = text.Replace(lines[Int16.Parse(richTextBox16.Text) - 1], text4);
                    File.WriteAllText(@"C:\compartilhamento\data_txt\data2.txt", text);
                    //  MessageBox.Show(" CADASTRO REALIZADO COM SUCESSO, A IDENTIDADE NÃO FOI ALTERADA POIS JÁ EXISTE UMA IDENTIDADE COM ESTA NÚMERO");
                    break;

                }
                else
                {
                    ESIM = false;

                }



            }

            if (ESIM == false)
            {

                id_1 = 1;
                string[] lines2 = File.ReadAllLines(@"C:\compartilhamento\data_txt\data2.txt");
                if (richTextBox4.Text.Trim() != l) {
                    string text4 = "Number : " + richTextBox16.Text.Trim() + " : Name : " + richTextBox2.Text + " : Compay : " + richTextBox1.Text + " :Funcition:  " + richTextBox3.Text + "  :Id: " + richTextBox4.Text.Trim() + " : E-mail : " + richTextBox8.Text + " : Vessel : " + richTextBox5.Text + " : Project : " + richTextBox9.Text + " : ASO : " + maskedTextBox1.Text + " : NR-34 : " + maskedTextBox2.Text + " : Vaccine-1 : " + maskedTextBox3.Text + " : Vaccine-2 : " + maskedTextBox4.Text + " : Booster vaccine : " + maskedTextBox5.Text + " : " + bb + " : " + comuser.Text + " :" + richTextBox17.Text;
                    string text = File.ReadAllText(@"C:\compartilhamento\data_txt\data2.txt");
                    text = text.Replace(lines2[Int16.Parse(richTextBox16.Text) - 1], text4);
                    File.WriteAllText(@"C:\compartilhamento\data_txt\data2.txt", text);
                    //  MessageBox.Show("CADASTRO REALIZADO COM SUCESSO");
                }
                else
                {

                }
            }





            /*
                    

                        if (id_1 == 0)
                        {
                            string id_3;
                            int exs = 0;
                                   string[] lines3 = File.ReadAllLines(@"C:\compartilhamento\data_txt\data2.txt");
                                   string id_11 = lines[Int16.Parse(richTextBox16.Text)-1].Split(':')[1].Trim();

                            for (int i = 0; i < lines3.Length; i++)
                            {

                                if (richTextBox4.Text.Trim() == lines[i].Split(':')[9].Trim())
                                {
                                    MessageBox.Show("JÁ EXISTE UMA IDENTIDADE COM ESTA NÚMERO");
                                    exs = 1;
                                }

                                else
                                {
                                    exs = 1; 

                                }

                            }

                            if (exs == 1)
                            {
                                id_3 = richTextBox4.Text.Trim();//lines[Int16.Parse(richTextBox16.Text) - 1].Split(':')[9].Trim();
                            }
                            string text4 = "Number : " + id_11 + " : Name : " + richTextBox2.Text + " : Compay : " + richTextBox1.Text + " :Funcition:  " + richTextBox3.Text + "  :Id: " + id_3 + " : E-mail : " + richTextBox8.Text + " : Vessel : " + richTextBox5.Text + " : Project : " + richTextBox9.Text + " : ASO : " + maskedTextBox1.Text + " : NR-34 : " + maskedTextBox2.Text + " : Vaccine-1 : " + maskedTextBox3.Text + " : Vaccine-2 : " + maskedTextBox4.Text + " : Booster vaccine : " + maskedTextBox5.Text + " : " + bb + " : " + comuser.Text + " :" + richTextBox17.Text;
                            string text = File.ReadAllText(@"C:\compartilhamento\data_txt\data2.txt");
                            text = text.Replace(lines[Int16.Parse(richTextBox16.Text) - 1], text4);
                            File.WriteAllText(@"C:\compartilhamento\data_txt\data2.txt", text);
                            MessageBox.Show("NOVO CADASTRO REALIZADO COM SUCESSO!");

                        }

                        //  GetBiosInformation();
                        /*
                        if (dataGridView1.Rows.Count > 0)
                        {
                            try
                            {
                                XcelApp.Application.Workbooks.Add(Type.Missing);
                                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                                {
                                    XcelApp.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                                }
                                //
                                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                                {
                                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                                    {
                                        XcelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                                    }
                                }
                                //
                                XcelApp.Columns.AutoFit();
                                //
                                XcelApp.Visible = true;
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Erro : " + ex.Message);
                                XcelApp.Quit();
                            }
                        }
                        */
            //out_by_user();
            //string line2 = null;
            // string line_to_delete = "the line i want to delete";


            // var oldLines = System.IO.File.ReadAllLines(@"C:\compartilhamento\data_txt\data.txt");
            // var newLines = oldLines.Where(line => !line.Contains(line_to_delete));
            // System.IO.File.WriteAllLines(@"C:\compartilhamento\data_txt\data.txt", newLines);

        }

        private void button9_Click(object sender, EventArgs e)
        {

        }

        private void button36_Click(object sender, EventArgs e)
        {
            String ma = "teste";
            string str2 = string.Empty;
            int val2 = 0;
            //  label5.Text = ma.ToString();
            for (int i = 0; i < ma.Length; i++)
            {
                if (Char.IsDigit(ma[i]))
                    str2 += ma[i];
            }
            //label5.Text = str2;

            while (str2.Length > 5)
            {
                str2 = str2.Substring(0, str2.Length - 1);
            }

            if (str2.Length > 0)
            {
                val2 = int.Parse(str2);
            }


            double str3 = ((Math.Sqrt(val2)) * val2) / 2;
            String str5 = str3.ToString();
            str5.Substring(0, 5);
            StreamReader srb = new StreamReader("code.txt"); //C:\Users\win 10\Documents   Users\\win 10\\Documents\\local.txt
            string xb = srb.ReadToEnd().Trim();
            srb.Close();
            id = val2.ToString();
            label2.Text = "ID: " + val2.ToString();
            // label9.Text = "ID: " + val.ToString();
            if (xb != str5.Substring(0, 5).Trim())
            {

            }





            using (StreamWriter writer = new StreamWriter("code.txt", false)) //@"\data_full\vessel\local_list.txt("C:\\data_full\\local.txt", true))
            {
                // writer.WriteLine(textBox1.Text + "," + textBox2.Text + "," + textBox3.Text + "," + comboBox2.Text + ",");
                writer.WriteLine(textBox11.Text);
                writer.Close();
            }
            //Task.Delay(1000).Wait();
            System.Windows.Forms.Application.Restart();


            int val = 0;
            val = int.Parse(textBox11.Text);
            double str33 = ((Math.Sqrt(val)) * val) / 2;
            textBox12.Text = str3.ToString();
        }

        private void richTextBox6_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (richTextBox16.Text != "")
            {
                DialogResult dialogResult = MessageBox.Show("DESEJA REALMENTE ALTERAR A DATA?", "ALTERAÇÃO DE DATA", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    dateTimePicker1.Visible = true;
                    dateTimePicker2.Visible = true;
                }
                else if (dialogResult == DialogResult.No)
                {
                    dateTimePicker1.Visible = false;
                    dateTimePicker2.Visible = false;
                }




            }
        }

        private void maskedTextBox2_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void richTextBox7_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (richTextBox16.Text != "")
            {
                DialogResult dialogResult = MessageBox.Show("DESEJA REALMENTE ALTERAR A DATA?", "ALTERAÇÃO DE DATA", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    dateTimePicker1.Visible = true;
                    dateTimePicker2.Visible = true;
                }
                else if (dialogResult == DialogResult.No)
                {
                    dateTimePicker1.Visible = false;
                    dateTimePicker2.Visible = false;
                }




            }
        }

        private void dataGridView1_SortStringChanged(object sender, Zuby.ADGV.AdvancedDataGridView.SortEventArgs e)
        {
            bindingSource1.Sort = dataGridView1.SortString;
        }

        private void button8_Click(object sender, EventArgs e)
        {

        }

        private void button10_Click(object sender, EventArgs e)
        {

        }

        private void button30_Click(object sender, EventArgs e)
        {
            label34.Text = DateTime.Now.ToString("dd/MM/yyyy").Trim();
        }

        private void button33_Click(object sender, EventArgs e)
        {

        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void richTextBox15_TextChanged(object sender, EventArgs e)
        {

        }
        private void get_link_txt()
        {
            try
            {
                if (richTextBox15.Text != "")
                {
                    // https://drive.google.com/file/d/12wpEOD6oDNL4eWgJXRFEzkPlmJ5fwdsT/view?usp=sharing

                    // string path = richTextBox15.Text;
                    //  String path2;
                    //  string[] subs = path.Split('/');
                    panel9.Size = new Size(1050, 43);
                    path2 = "https://drive.google.com/file/d/" + richTextBox15.Text.Trim() + "/view?usp=sharing";
                    // path2 = "https://drive.google.com/file/d/12wpEOD6oDNL4eWgJXRFEzkPlmJ5fwdsT/view?usp=sharing";
                    panel9.Visible = true;
                    panel14.Visible = true;
                    panel14.Size = new Size(1145, 44);

                    webBrowser1.Visible = true;
                    // webBrowser1.DocumentText = "<html>< head >< title > HTML Backgorund Color</ title ></ head >< body style = 'background-color:grey;' >< h1 > Products </ h1 ></ body ></ html >";
                    webBrowser1.Size = new Size(1145, 680);


                    this.webBrowser1.Navigate(path2);
                }
                // MessageBox.Show(subs[5]);
            }
            catch
            {

            }
        }
        private void richTextBox15_DoubleClick(object sender, EventArgs e)
        {
            get_link_txt();
        }

        private void button37_Click(object sender, EventArgs e)
        {
            panel9.Visible = false;
            panel14.Visible = false;
            webBrowser1.Visible = false;
        }

        private void lcompany_Click(object sender, EventArgs e)
        {
            richTextBox1.Text = comboBox1.Text + " CREW";
        }

        int ip_find = 0;
        int count2 = 0;
        int soma = 0;
        public void scan2(string start, string end)
        {
            comp = 0;
            //   timer9.Enabled = false;
            if (soma == 0)
            {
           
                try
                {

                    //Split IP string into a 4 part array
                    string[] startIPString = start.Split('.');
                    int[] startIP = Array.ConvertAll<string, int>(startIPString, int.Parse); //Change string array to int array


                    string[] endIPString = end.Split('.');
                    int[] endIP = Array.ConvertAll<string, int>(endIPString, int.Parse);
                    //Count the number of successful pings
                    // count2 = 0;
                    Ping myPing;
                    PingReply reply;
                    IPAddress addr;
                    IPHostEntry host;
                  //  

                    //Progress bar

                    //   listVAddr.Items.Clear();

                    //Loops through the IP range, maxing out at 255
                    for (int i = startIP[2]; i <= endIP[2]; i++)
                    { //3rd octet loop
                        for (int y = startIP[3]; y <= 255; y++)
                        { //4th octet loop
                            string ipAddress = startIP[0] + "." + startIP[1] + "." + i + "." + y; //Convert IP array back into a string
                            string endIPAddress = endIP[0] + "." + endIP[1] + "." + endIP[2] + "." + (endIP[3] + 1); // +1 is so that the scanning stops at the correct range

                            //If current IP matches final IP in range, break
                            if (ipAddress == endIPAddress)
                            {
                                break;
                            }

                            myPing = new Ping();
                            try
                            {
                                reply = myPing.Send(ipAddress, 10); //Ping IP address with 500ms timeout
                                myping2 = reply.RoundtripTime.ToString();
                            }
                            catch (Exception ex)
                            {
                                break;
                            }



                            //Log pinged IP address in listview
                            //Grabs DNS information to obtain system info
                            if (reply != null && reply.Status == IPStatus.Success)
                            {


                                try
                                {

                                    addr = IPAddress.Parse(ipAddress);
                                    host = Dns.GetHostEntry(addr);
                                    // addr.Address.Trim();

                                    //  listVAddr.Items.Add(new ListViewItem(new String[] { host.HostName, "Up" })); //Log successful pings
                                    host.HostName.Trim();

                                    String input = host.HostName.Trim();//host.HostName.Substring(0, host.HostName.LastIndexOf("."));
                                                                        // label21.Text = input;
                                                                        //  MessageBox.Show(input.ToString());
                                                                        //input = "";
                                                                        //  MessageBox.Show(ipAddress);
                                    if (host.HostName.Trim() != hostName.Trim() && hostName != "")
                                    {
                                        //  
                                        //  ip_find++;
                                        //  count2=2;

                                        count2++;

                                    }
                                    if (host.HostName.Trim() == hostName)
                                    {

                                        // count2=1;

                                        //  ip_find--;
                                        //  MessageBox.Show("ok" + host.HostName.Trim());
                                        // online_ = false;
                                    }


                                    //  if (host.HostName!= "LAPTOP-DRSLFUQS.local")
                                    /// {

                                    //  MessageBox.Show(input.Trim());
                                    // MessageBox.Show("teste ip " + input + "  " + count);
                                    // 
                                    //  \\CRIPTOQRCODE2\\compartilhamento\\
                                   // rede1 = "\\\\" + input.Trim() + "\\compartilhamento\\";
                                    // MessageBox.Show(rede1);
                                    // host.HostName = "";

                                    online_ = true;
                                    p = 1;
                                    int at = 0;
                                    //timer7.Stop();
                                    //escrever_lock();
                                    string hostName2 = System.Net.Dns.GetHostName();
                                    string str = input.Substring(0, 6).Trim();
                                    //  MessageBox.Show(input);

                                    if (input != " " && input != hostName2+".lan" && input != hostName2)
                                    {

                                        if (str == "QRCODE")
                                        {
                                            ip_find++;
                                            
                                            rede1 = "\\\\" + input.Trim() + "\\compartilhamento\\";
                                            rede10= input.Trim();
                                            //  panel12.BackColor = Color.Black;
                                            //   MessageBox.Show(input);
                                            atualiza_compartilhamento();

                                            // timer11.Start();
                                            str = "";


                                        }
                                        else
                                        {

                                        }
                                    }
                                    else
                                    {

                                    }


                                    //System.Threading.Thread.Sleep(1000);
                                    // atualiza_compartilhamento();
                                   // input = "";

                                    //  }

                                    //  MessageBox.Show(count.ToString());

                                    //  
                                    //    MessageBox.Show(host.HostName.Trim());




                                }

                                catch
                                {

                                    // listVAddr.Items.Add(new ListViewItem(new String[] { ipAddress, "Could not retrieve", "Up" })); //Logs pings that are successful, but are most likely not windows machines


                                    //  online_ = false;
                                    //  count2--;




                                }
                                // MessageBox.Show("teste ip " + count);

                                // count2 = 0;



                            }
                            else
                            {



                            }

                            //  count2--;




                            if (count2 == 0)
                            {
                                //  online_ = false;
                                // count2 = 0;
                            }
                            if (count2 == 2)
                            {

                                // online_ = true;
                            }



                            if (count2 == 0)
                            {
                                // MessageBox.Show("sem rede");
                            }
                        }

                        startIP[3] = 1; //If 4th octet reaches 255, reset back to 1


                    }

                  //  panel12.BackColor = Color.Black;
                    //   MessageBox.Show("Scanning done!\nFound " + ip_find + " hosts.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    compr = 0;
                 //   label62.Text = ip_find.ToString();
                  //  ip_find = 0;


                    //comp = 1;

                    ///System.Threading.Thread.Sleep(2000);
                    //  MessageBox.Show("ok");
                    //  timer7.Start();

                    //Re-enable buttons

                    //    MessageBox.Show("Scanning done!\nFound " + ip_find + " hosts.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //Catch exception that throws when stopping thread, caused by ping waiting to be acknowledged
                   // ip_find = 0;

                }
                catch (ThreadAbortException tex)
                {
                    Console.WriteLine(tex.StackTrace);
                    // txtIP.Enabled = true;
                    // txtIP2.Enabled = true;

                }


                //Catch invalid IP types
                catch (Exception ex)
                {
                    Console.WriteLine(ex.StackTrace);

                    //txtIP.Enabled = true;
                    // txtIP2.Enabled = true;

                }


                if (count >= 2)
                {
                    // online_ = true;
                    // if (online_ == true)
                    //  {
                    // label40.Visible = false;
                    // panel12.BackColor = Color.GreenYellow;


                    //  }
                    //  if (online_ == false)
                    //  {
                    //  panel12.BackColor = Color.Red;
                    //  label40.Visible = true;

                    // }
                    // online_ = false;
                    //MessageBox.Show(count.ToString());
                }
            }
           /// ip_find = 0;
            //  timer9.Enabled = true;
        }

        private void timer7_Tick(object sender, EventArgs e)
        {


            // String block = File.ReadLines(@"C:\compartilhamento\lock.txt").ElementAtOrDefault(0);

            if (comp == 1)
            {
                // ler_linha();
                //   checa_host();
                // atualiza_compartilhamento();

                //    MessageBox.Show(comp.ToString());
            }
        }

        private void label42_Click(object sender, EventArgs e)
        {

        }
        int ok = 1;
        private void cria_CarregarPlanilha()
        {
            var drives = DriveInfo.GetDrives().Where(drive => drive.IsReady && drive.DriveType == DriveType.Removable);
            try
            {
                File.Delete(drives.FirstOrDefault().Name.ToString() + "\\Controle de Acesso_backup.xlsx");
            }
            catch
            {

            }
            try
            {

                int count = File.ReadAllLines(@"C:\compartilhamento\data_txt\data2.txt").Length;
                // MessageBox.Show(count.ToString());

                var pasta = app.Workbooks.Open(drives.FirstOrDefault().Name.ToString() + "\\Controle de Acesso_backup.xlsx");

                var plan = pasta.Worksheets["Planilha1"];
                // plan.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                plan.Columns.AutoFit();
                int lastRow = plan.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                plan.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                plan.Cells[1, 1] = "NUMBER";
                plan.Cells[1, 2] = "NAME";
                plan.Cells[1, 3] = "COMPANY";
                plan.Cells[1, 4] = "FUNCTION";
                plan.Cells[1, 5] = "ID";
                plan.Cells[1, 6] = "EMAIL";
                plan.Cells[1, 7] = "VESSEL";
                plan.Cells[1, 8] = "CHECK-IN VALIDATION";
                plan.Cells[1, 9] = "CHECK-OUT VALIDATION";


                for (int a = 0; a < count; a++)
                {
                    ok++;
                    string secondLine = File.ReadLines(@"C:\compartilhamento\data_txt\data2.txt").ElementAtOrDefault(a);
                    //  ss++;
                    //  String local = secondLine.Split(':')[0].Trim();

                    //   txtCodigoFunci.Text = lastRow.ToString();
                    lastRow++;
                    plan.Cells[ok, 1] = secondLine.Split(':')[0].Trim();
                    // plan.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    plan.Cells[ok, 2] = secondLine.Split(':')[1].Trim();
                    plan.Cells[ok, 3] = secondLine.Split(':')[2].Trim();
                    plan.Cells[ok, 4] = secondLine.Split(':')[3].Trim();
                    plan.Cells[ok, 5] = secondLine.Split(':')[4].Trim();
                    plan.Cells[ok, 6] = secondLine.Split(':')[5].Trim();
                    plan.Cells[ok, 7] = secondLine.Split(':')[6].Trim();
                    plan.Cells[ok, 8] = secondLine.Split(':')[7].Trim();
                    plan.Cells[ok, 9] = secondLine.Split(':')[8].Trim();
                    // plan.Cells[lastRow, 10] = secondLine.Split(':')[9].Trim();
                    // plan.Cells[lastRow, 11] = secondLine.Split(':')[10].Trim();
                    // MessageBox.Show("ok");
                }
                //  plan.quit();

                pasta.SaveAs(drives.FirstOrDefault().Name.ToString() + "\\Controle de Acesso_backup.xlsx", Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                // pasta.Save();
                pasta.Close();
                // plan.Close();
                MessageBox.Show("Backup concluido com sucesso!");

                //  atualiza_compartilhamento();



            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            //CarregarPlanilha();

        }
        void PausaComThreadSleep()
        {
            Thread.Sleep(5000);
        }
        async Task PausaComTaskDelay()
        {
            await Task.Delay(5000);
        }
        string secondLine2;
        string secondLine5;
        private void update()
        {

            var drives = DriveInfo.GetDrives().Where(drive => drive.IsReady && drive.DriveType == DriveType.Removable);

            if (drives.FirstOrDefault() != null)
            {


                File.Delete(drives.FirstOrDefault().Name.ToString() + "//data2.txt");






                int ss = 7;

                try
                {
                    var wbook = new XLWorkbook(drives.FirstOrDefault().Name.ToString() + "//Controle de Acesso.xlsx");
                    var ws1 = wbook.Worksheet(1);






                    //ID Capacete  Identificação	Empresa	Nome Completo		Identidade	CPF	Função	E-mail	Data  ASO	Data dose 1	Data dose 2	Data Reforço 1	Data Reforço 2


                    //Number : 4 : Name : ARTHUR LOPES : Compay : ALTN :Funcition:  COORDENADOR DE LOGISTICA  :Id: 6376702601 : E-mail : altn.comercial@gmail.com : Vessel : Skandi Rio : Project : Docagem : ASO : 31/01/2023 : NR-34 : 00/00/0000 : Vaccine-1 : 00/00/0000 : Vaccine-2 : 00/00/0000 : Booster vaccine : 00/00/0000 :  : COMUM : :Convés
                    //  int count = File.ReadAllLines(@"D:/Controle de Acesso.xlsx").Length;

                    var columnCount = ws1.LastRowUsed().RowNumber();
                    //  MessageBox.Show(columnCount.ToString());
                    for (int i = 0; i < (columnCount) - 7; i++)
                    {
                        ss++;
                        String muda = ss.ToString();
                        var data1 = ws1.Cell("B" + muda).GetValue<string>();// numero
                        var data2 = ws1.Cell("D" + muda).GetValue<string>(); // Nome
                        var data3 = ws1.Cell("C" + muda).GetValue<string>();//  empresa 
                        var data4 = ws1.Cell("G" + muda).GetValue<string>();// CPF
                        var data5 = ws1.Cell("H" + muda).GetValue<string>();// função
                        var data6 = ws1.Cell("I" + muda).GetValue<string>(); // E-mail
                        var data7 = ws1.Cell("J" + muda).GetValue<string>(); // ASO
                        var data8 = ws1.Cell("K" + muda).GetValue<string>(); //  dose 1
                        var data9 = ws1.Cell("L" + muda).GetValue<string>(); //  dose 2
                        var data10 = ws1.Cell("M" + muda).GetValue<string>();//  data reforço1
                        var data11 = ws1.Cell("N" + muda).GetValue<string>();// data1 reforço 2
                                                                             // var data12 = ws1.Cell("K" + muda).GetValue<string>(); // NR34
                                                                             //   var data13 = ws1.Cell("E" + muda).GetValue<string>(); // Nome
                        var data14 = ws1.Cell("F" + muda).GetValue<string>(); // Nome
                                                                              //   lineChanger("Number:" + data1 + ":Name:" + data2 + ":Company:" + data3 + ":Function:" + data4, "D:\\data2.txt",ss);


                        using (StreamWriter writer = new StreamWriter(drives.FirstOrDefault().Name.ToString() + "//data2.txt", true)) //@"\data_full\vessel\local_list.txt("C:\\data_full\\local.txt", true))
                        {
                            string teste2 = "Number: " + data1 + " : Name :" + data2 + " : Company : " + data3 + " : Function : " + data5+ ": Id : " + data14 + ": E-mail : " + data6 + ": Vessel : " + comboBox1.Text + ": Project : " + richTextBox9.Text+ ": ASO : " + data7 + ": NR34 : " + data8 + ": NR-10 : " + data9 + ": NR-33 : " + data10 + ": NR-35 :" + data11;
                            writer.WriteLine(teste2);
                            writer.Close();
                        }


                    }

                    try
                    {
                        System.IO.File.Copy(@"C:\compartilhamento\data_txt\data2.txt", @"C:\compartilhamento\data_txt\data2_backup.txt", true);
                    }
                    catch
                    {

                    }
                    File.Delete(@"C:\compartilhamento\data_txt\data2.txt");
                    string text = File.ReadAllText(drives.FirstOrDefault().Name.ToString() + "//data2.txt");
                    text = text.Replace("00:00:00", "");
                    File.WriteAllText(@"C:\compartilhamento\data_txt\data2.txt", text);


                }
                catch
                {

                }
                label43.Text = "UPDATE CONDLUIDO COM SUCESSO!";
            }
            else
            {
                label43.Text = "";
                MessageBox.Show("PENDRIVE OU ARQUIVO NÃO ENCONTRADO");
            }
            //  MessageBox.Show("PENDRIVE NÃO ENCONTRADA");


        }
        private void button39_Click(object sender, EventArgs e)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo("cmd", "/K netsh interface set interface \"Ethernet\" DISABLED");
            startInfo.RedirectStandardOutput = true;
            startInfo.UseShellExecute = false;
            // Do not create the black window.
            startInfo.CreateNoWindow = true;
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            Process.Start(startInfo);
            //https://zetcode.com/csharp/excel/ dicas excel
            label43.Text = "AGUARDE, FAZENDO UPDATE DA PLANILHA!";
            update();

          

            //  PausaComThreadSleep();

            //  panel15.Visible = true;
            /// label43.Text = "UPDATE CONDLUIDO COM SUCESSO!";
        }
        public bool RemoveFirstLinesFromFile(string filePath, int skip)
        {
            if (!File.Exists(filePath))
                return false;
            try
            {
                var filePathOld = Path.Combine(filePath, ".old");
                File.Move(filePath, filePathOld);
                File.WriteAllLines(filePath, File.ReadAllLines(filePathOld).Skip(skip));
                return true;
            }
            catch (System.Exception)
            {
                return false;
            }
        }
        private void button38_Click(object sender, EventArgs e)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo("cmd", "/K netsh interface set interface \"Ethernet\" DISABLED");
            startInfo.RedirectStandardOutput = true;
            startInfo.UseShellExecute = false;
            // Do not create the black window.
            startInfo.CreateNoWindow = true;
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            Process.Start(startInfo);
            var drives = DriveInfo.GetDrives().Where(drive => drive.IsReady && drive.DriveType == DriveType.Removable);
            if (drives.FirstOrDefault() != null)
            {
                label43.Text = "AGUARDE, CRIANDO BACKUP!";

                lblname.Text = drives.FirstOrDefault().Name.ToString();
                string fileName = "novo.xlsx";
                string sourcePath = @"C:\compartilhamento\data_base";
                string targetPath = drives.FirstOrDefault().Name.ToString() + "\\CRIPTOQRCODE_AllBackup\\";

                string destFile = System.IO.Path.Combine(targetPath, fileName) ;
                System.IO.Directory.CreateDirectory(targetPath);




  

                //Se o diretório não existir...

                if (!Directory.Exists(targetPath))
                {

                    //Criamos um com o nome folder
                    Directory.CreateDirectory(targetPath);

                }



                if (System.IO.Directory.Exists(sourcePath))
                {
                    string[] files = System.IO.Directory.GetFiles(sourcePath);

                    foreach (string s in files)
                    {

                        fileName = System.IO.Path.GetFileName(s);
                        destFile = System.IO.Path.Combine(targetPath, fileName);
                        System.IO.File.Copy(s, destFile, true);
                    }

                    // cria_CarregarPlanilha();
                   ;


                    try
                    {
                        File.Delete(@"D:\Result.txt");
                        File.Delete(@"D:\teste.txt");
                        File.Delete(@"D:\teste3.txt");
                    }
                    catch
                    {

                    }
                    string tempo = DateTime.Now.ToString("dd_MM_yyyy_hh_mm_ss");
                    /// MessageBox.Show("Backup creaded with success");
                    File.Copy(lblname.Text + "\\CRIPTOQRCODE_AllBackup\\novo.xlsx", lblname.Text + "\\CRIPTOQRCODE_AllBackup\\" + tempo + ".xlsx");
                    label43.Text = "BACKUP CRIADO COM SUCESSO!";
                    if (textBox7.SelectionLength >= 0)
                    {
                        textBox7.Focus();
                        textBox7.Text = "";
                    }
                }


                else
                {
                    Console.WriteLine("Source path does not exist!");
                    if (textBox7.SelectionLength >= 0)
                    {
                        textBox7.Focus();
                        textBox7.Text = "";
                    }
                }
            }
            else
            {
                MessageBox.Show("No Pendrive found..");
                if (textBox7.SelectionLength >= 0)
                {
                    textBox7.Focus();
                    textBox7.Text = "";
                }
            }


        }




        private string ReplaceFirst(string thise, string oldValue, string newValue)
        {
            int startindex = thise.IndexOf(oldValue);

            if (startindex == -1)
            {
                return thise;
            }

            return thise.Remove(startindex, oldValue.Length).Insert(startindex, newValue);
        }

        private void back_up()
        {
            try
            {
                int ss = 7;
                /// ler excel
                /*
                using ClosedXML.Excel;

                using var wbook = new XLWorkbook("simple.xlsx");

                var ws1 = wbook.Worksheet(1);
                var data = ws1.Cell("A1").GetValue<string>();
                */
                //ID Capacete  Identificação	Empresa	Nome Completo		Identidade	CPF	Email	Função	Data  ASO	Data dose 1	Data dose 2	Data Reforço 1	Data Reforço 2

                // Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                // if (xlApp == null)

                /// {
                //  MessageBox.Show("Excel is not properly installed!!");
                // return;
                //   }




                int count = File.ReadAllLines(@"C:\compartilhamento\data_txt\data2.txt").Length;
                var wb = new XLWorkbook();

                //  wb.Application.OnKey("^v", "");


                var ws = wb.Worksheets.Add("Planilha1");
                ws.Range("B7", "V7").Style.Fill.BackgroundColor = XLColor.FromArgb(91, 155, 213); //Color.FromArgb(91, 155, 213);
                ws.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                ws.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                ws.Column(7).Width = 35;
                ws.Row(7).Height = 50;
                ws.Cell("B7").Value = "ID Capacete  Identificação";//1
                ws.Cell("C7").Value = "Empresa";//5
                ws.Cell("D7").Value = "Nome Completo";//3
                ws.Cell("F7").Value = "Identidade";
                ws.Cell("G7").Value = "CPF";//34
                ws.Cell("I7").Value = "Email";//11
                ws.Cell("H7").Value = "Função";//7
                ws.Cell("J7").Value = "Data  ASO";//13
                ws.Cell("V7").Value = "Data  NR34";//15
                ws.Cell("K7").Value = "Data dose 1";//17
                ws.Cell("L7").Value = "Data dose 2";//19
                ws.Cell("M7").Value = "Data Reforço 1";//21
                ws.Cell("N7").Value = "Data Reforço 2";//23


                /// ClosedXML.Excel.CutCopyMode = XlCutCopyMode.xlCopy;
                // Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)ws.Worksheets(1);
                // wks.Application.CutCopyMode = (Microsoft.Office.Interop.Excel.XlCutCopyMode)0;
                // wb.Worksheets.Application.CutCopyMode = 0;
                //XlCutCopyMode.xlCopy = false;
                ///  Application.CutCopyMode = (Microsoft.Office.Interop.Excel.XlCutCopyMode)0;
                //ws.CutCopyMode = false;
                for (int a = 0; a < count; a++)
                {

                    string secondLine = File.ReadLines(@"C:\compartilhamento\data_txt\data2.txt").ElementAtOrDefault(a);

                    ss++;

                    ws.Cell("B" + ss).Value = secondLine.Split(':')[1].Trim();
                    ws.Cell("C" + ss).Value = secondLine.Split(':')[5].Trim();
                    ws.Cell("D" + ss).Value = secondLine.Split(':')[3].Trim();
                    ws.Cell("F" + ss).Value = secondLine.Split(':')[9].Trim();
                    ws.Cell("I" + ss).Value = secondLine.Split(':')[11].Trim();
                    ws.Cell("H" + ss).Value = secondLine.Split(':')[7].Trim();
                    ws.Cell("G" + ss).Value = secondLine.Split(':')[15].Trim();
                    ws.Cell("J" + ss).Value = secondLine.Split(':')[17].Trim();
                    ws.Cell("K" + ss).Value = secondLine.Split(':')[19].Trim();
                    ws.Cell("L" + ss).Value = secondLine.Split(':')[21].Trim();
                    ws.Cell("M" + ss).Value = secondLine.Split(':')[23].Trim();
                    ws.Cell("N" + ss).Value = secondLine.Split(':')[25].Trim();
                    ws.Cell("V" + ss).Value = secondLine.Split(':')[19].Trim();
                }
                ws.Columns().AdjustToContents();

                try
                {
                    var drives = DriveInfo.GetDrives().Where(drive => drive.IsReady && drive.DriveType == DriveType.Removable);
                    wb.SaveAs(drives.FirstOrDefault().Name.ToString() + "\\Controle de Acesso_backup.xlsx");

                }
                catch
                {
                    // 
                }
            }catch
            {

                MessageBox.Show("PENDRIVE NÃO ENCONTRADA");
            }

        }


        private void button40_Click(object sender, EventArgs e)
        {
       
            panel15.Visible = false;
            System.Windows.Forms.Application.Restart();

        }

        private void panel15_Paint(object sender, PaintEventArgs e)
        {

        }

        private void richTextBox7_TextChanged(object sender, EventArgs e)
        {

        }
    
        private void timer10_Tick(object sender, EventArgs e)
        {
            /// DirectoryCanListFiles(rede1.Trim() + "data_txt\\data2.txt");
            /// "\\\\" + input.Trim() + "\\compartilhamento\\"
            // MessageBox.Show(MyhostName); 
            //   MessageBox.Show(rede1);

            string nome3;
            Boolean rec2 = false;
            // string myIP = Dns.GetHostByName("QRCODE-50").AddressList[1].ToString();
            //label64.Text = myIP;
            textBox22.Text = " ";
            nome3 = "";

            try
            {

                if (rede1 != null) {



                    if (rede1 != "\\\\" + MyhostName.Trim() + "\\compartilhamento\\")
                    {

                        // MessageBox.Show(rede1);
                        nome3 = rede1;
                        if (rede1 != "")
                        {
                            string nome = rede1.Split('\\')[2].Trim();
                            string nome2 = nome.Remove(nome.Length - 6);
                            textBox22.Text = nome + " Online";
                            //myString.Remove(myString.Length-3)
                            //  panel12.BackColor = Color.YellowGreen;


                            // 

                            if (rec == true)
                            {
                                DateTime fdata0 = File.GetLastWriteTime(@"C:\compartilhamento\data_txt\data2.txt");
                                DateTime fdata1 = File.GetLastWriteTime(rede1.Trim() + "data_txt\\data2.txt");
                                if (fdata1 == fdata0)
                                {
                                    int count = File.ReadAllLines(@"C:\compartilhamento\data_txt\data2.txt").Length;
                                    label3.Text = count.ToString().Trim();
                                    
                                    //   saida_manual();
                                    rec = false;

                                }
                              //  if (fdata1 != fdata0)
                                //{
                                 //   saida_manual();

                               // }
                            }
                            //  }
                            else
                            {

                            }
                        }
                      //  rede1 = "";


                    }
                    else
                    {
                        //  MessageBox.Show("nulo");
                    }
                }
                if (entrou == true)
                {
                   
                }
                
            }
            catch
            {
               // textBox22.Text = "";
            }
            // rede1 = "";



        }
        

        
        private void textBox14_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox14_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            string text = System.IO.File.ReadAllText(@"C:\compartilhamento\data_txt\data.txt");
            listBox1.Items.Clear();
            foreach (string s in Regex.Split(text, textBox14.Text))
            {

                listBox1.Items.Add(s);



            }
        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {
            if (richTextBox2.Text == "") {
                label44.Visible = true;
                label50.Visible = true;
            }
            else
            {
                label44.Visible = false;
            }
        }

        private void maskedTextBox1_TextChanged(object sender, EventArgs e)
        {
            if (maskedTextBox1.Text == "  /  /")
            {
                label49.Visible = true;
            }
            else
            {
                label49.Visible = false;
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void txtResult_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }

        private void label53_Paint(object sender, PaintEventArgs e)
        {
            Font myfont = new Font("Arial", 14);
            Brush mybrush = new System.Drawing.SolidBrush(System.Drawing.Color.Black);
            e.Graphics.TranslateTransform(90, 90);
            e.Graphics.RotateTransform(90);
            e.Graphics.DrawString("teste12345678", myfont, mybrush, 0, 0);
        }

        private void button41_Click(object sender, EventArgs e)
        {
            label34.Text = comboBox1.Text;
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            label63.Text = comboBox1.Text;
            if (comboBox1.Text.Trim() == "")
            {
                label54.Visible = true;
         
                // label50.Visible = true;
            }
            else
            {
                label54.Visible = false;
           

            }



            listBox1.Items.Clear();
            textBox16.Text = "";
            string[] linhas22 = System.IO.File.ReadAllLines(@"C:\compartilhamento\data_txt\data.txt");
            if (comboBox1.Text.Trim() != "ALL")
            {
                comboname = comboBox1.Text.Trim();
                lista_ = 0;
                foreach (string linha22 in linhas22)
                {
                    if (linha22.Contains(comboBox1.Text.Trim()))
                    {
                        listBox1.Items.Add(linha22);
                        lista_++;
                        label27.Text = lista_.ToString();
                       // label67.Text = label27.Text;

          
                    }
                }
                if (lista_ == 0)
                {
                    label27.Text = lista_.ToString();
                    //label67.Text = label27.Text;
                }
            }
            else
            {
                lista_ = 0;
                foreach (string linha22 in linhas22)
                {
                    //  if (linha22.Contains("\n"))
                    // {
                    listBox1.Items.Add(linha22);
                    lista_++;
                    label27.Text = lista_.ToString();
                    //label67.Text = label27.Text;


                    // }
                }
                if (lista_ == 0)
                {
                    label27.Text = lista_.ToString();
                    //label67.Text = label27.Text;
                }
            }
            if (richTextBox16.Text == "")
            {
                // button1.PerformClick();
            }








        }
        Boolean existe = false;
        private void button42_Click(object sender, EventArgs e)
        {
            existe = false;
            try
            {

                string result = string.Empty;
                var lines = File.ReadAllLines(@"C:\compartilhamento\vessels.txt");
                foreach (var line in lines)
                {
                    if (line.Contains(textBox17.Text))
                    {
                        //var text = line.Replace("Customer ID :", "");
                        //result = text.Trim();
                        existe = true;
                    }
                }



                if (existe == true)
                {
                    MessageBox.Show(" Nome ja existente");

                }

                else
                {


                    comboBox1.Items.Clear();
                    using (StreamWriter writer = new StreamWriter(@"C:\compartilhamento\vessels.txt", true)) //@"\data_full\vessel\local_list.txt("C:\\data_full\\local.txt", true))
                    {
                        // writer.WriteLine(textBox1.Text + "," + textBox2.Text + "," + textBox3.Text + "," + comboBox2.Text + ",");
                        writer.WriteLine(textBox17.Text);
                        writer.Close();
                    }

                    MessageBox.Show("nome cadastrado com sucesso!");
                    StreamReader sr = new StreamReader(@"C:\compartilhamento\vessels.txt");
                    string x = sr.ReadToEnd();
                    string[] y = x.Split('\n');
                    foreach (string s in y)
                    {
                        comboBox1.Items.Add(s);
                    }
                    sr.Close();


                    //  comboBox1.Items.Add(textBox17.Text.Trim());




                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //  ListBox1.Items.Add(TextBox1.Text)
            //  Else
            //     MessageBox.Show("Item ja existente", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            //  End If
        }

        private void button41_Click_1(object sender, EventArgs e)
        {
            //panel16.Visible = false;
            pcad.Visible = false;
        }

        private void checkBox1_CheckStateChanged(object sender, EventArgs e)
        {
            //if(checkBox1.CheckState == )
        }

        private void textBox16_TextChanged(object sender, EventArgs e)
        {
            // button1.PerformClick();
            listBox1.Items.Clear();

            // lista2_ = 0;
            string[] linhas22 = System.IO.File.ReadAllLines(@"C:\compartilhamento\data_txt\data.txt");

            foreach (string linha22 in linhas22)
            {
                if (linha22.Contains(textBox16.Text.Trim()))
                {
                    //   listBox1.Items.Add(linha22);
                    //lista2_++;
                    //label27.Text = lista_.ToString();
                }
            }

            if (comboBox2.Text.Trim() != "ALL")
            {
                listBox1.Items.Clear();
                lista_ = 0;
                foreach (string linha22 in linhas22)
                {
                    if (linha22.Contains(textBox16.Text.Trim()) && linha22.Contains(comboBox2.Text.Trim()))
                    {
                        listBox1.Items.Add(linha22);

                        //  lista_++;
                        //  label27.Text = lista_.ToString();
                    }
                }
            }
            else
            {
                listBox1.Items.Clear();
                lista_ = 0;
                foreach (string linha22 in linhas22)
                {
                    if (linha22.Contains(textBox16.Text.Trim()))
                        // {
                        listBox1.Items.Add(linha22);

                    // lista_++;
                    // label27.Text = lista_.ToString();
                }
            }







        }

        private void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            pictureBox4.Visible = false;
            textBox16.Text = "";
            string[] linhas22 = System.IO.File.ReadAllLines(@"C:\compartilhamento\data_txt\data.txt");
            if (comboBox2.Text.Trim() != "ALL")
            {
                lista_ = 0;
               // label67.Text = lista_.ToString();
                foreach (string linha22 in linhas22)
                {
                    if (linha22.Contains(comboBox2.Text.Trim()))
                    {
                        listBox1.Items.Add(linha22);
                        lista_++;
                        label27.Text = lista_.ToString();
                       /// label67.Text = lista_.ToString();

                    }
                }
            }
            else
            {
                lista_ = 0;
               // label67.Text = lista_.ToString();
                foreach (string linha22 in linhas22)
                {
                    //  if (linha22.Contains("\n"))
                    // {
                    listBox1.Items.Add(linha22);
                    lista_++;
                    label27.Text = lista_.ToString();
                  //  label67.Text = lista_.ToString();


                    // }
                }

                if (lista_ == 0)
                {
                    // label27.Text = lista_.ToString();
                }

            }
            button1.PerformClick();
            if (richTextBox16.Text == " ")
            {

            }
        }

        private void comboBox1_MouseClick(object sender, MouseEventArgs e)
        {
            if (richTextBox3.Text != "")
            {
                comboBox1.Items.Clear();
                StreamReader sr = new StreamReader(@"C:\compartilhamento\vessels.txt");
                string x = sr.ReadToEnd();
                string[] y = x.Split('\n');
                foreach (string s in y)
                {
                    comboBox1.Items.Add(s);
                }
                sr.Close();
            }
        }

        private void textBox19_TextChanged(object sender, EventArgs e)
        {

        }

        private void regs_MouseClick(object sender, MouseEventArgs e)
        {
            if (pcad.Visible == false)
            {
                pcad.Visible = true;
                StreamReader sr = new StreamReader(@"C:\compartilhamento\IP_NEW.txt");
                string x = sr.ReadToEnd();
                sr.Close();

                IP_START.Text = x.Split(',')[0].Trim();
                IP_STOP.Text = x.Split(',')[1].Trim();

            }
            else
            {
                pcad.Visible = false;
            }
        }
        static string ReplaceLastLetter(string text, string newLetter)
        {
            string substring = text.Substring(0, text.Length - 1); // ABC -> AB
            return substring + newLetter; // ABD
        }
        private void button43_Click(object sender, EventArgs e)
        {
            // string text = "Guatavo,1234567890,";
            // string newLetter = ",Rosana,123456,";


            //  string replaced = ReplaceLastLetter(text, newLetter);
            //  MessageBox.Show(replaced);



            if (comboBox3.Text != "" && textBox19.Text != "" && textBox21.Text != "" && textBox20.Text != "")
            {
                StreamReader sr = new StreamReader(@"C:\compartilhamento\pass\pass.txt");
                string x = sr.ReadToEnd();
                sr.Close();

                string replaced = x.Replace(comboBox3.Text, textBox21.Text).Replace(textBox19.Text, textBox20.Text);

                using (StreamWriter writer = new StreamWriter(@"C:\compartilhamento\pass\pass.txt", false)) //@"\data_full\vessel\local_list.txt("C:\\data_full\\local.txt", true))
                {
                    // writer.WriteLine(textBox1.Text + "," + textBox2.Text + "," + textBox3.Text + "," + comboBox2.Text + ",");
                    writer.Write(replaced);
                    writer.Close();
                }
                MessageBox.Show("CONCLUIDO COM SUCESSO!");
            }
            else
            {
                MessageBox.Show("INFORME OS DADOS PARA CADASTRO OU ALTERAÇÃO");
            }
            System.Windows.Forms.Application.Restart();
        }

        private void button44_Click(object sender, EventArgs e)
        {


            StreamReader sr = new StreamReader(@"C:\compartilhamento\pass\pass.txt");
            string x = sr.ReadToEnd();
            sr.Close();
            int numVirgulas = x.Split(',').Length;

            if (numVirgulas >= 16)
            {
                MessageBox.Show("O NÚMRO DE USUÁRIOS JÁ ESTÁ PREEENCHIDO!");

            }
            else
            {
                string text = x;
                string newLetter = "," + comboBox3.Text + "," + textBox19.Text + ",";
                string replaced = ReplaceLastLetter(text, newLetter);

                using (StreamWriter writer = new StreamWriter(@"C:\compartilhamento\pass\pass.txt", false)) //@"\data_full\vessel\local_list.txt("C:\\data_full\\local.txt", true))
                {
                    // writer.WriteLine(textBox1.Text + "," + textBox2.Text + "," + textBox3.Text + "," + comboBox2.Text + ",");
                    writer.Write(replaced);
                    writer.Close();
                }
            }
            System.Windows.Forms.Application.Restart();
        }


        private void button45_Click(object sender, EventArgs e)
        {
            pcad.Visible = false;
           // System.Windows.Forms.Application.Restart();
        }

        private void button46_Click(object sender, EventArgs e)
        {

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

            StreamReader sr = new StreamReader(@"C:\compartilhamento\pass\pass.txt");
            string x = sr.ReadToEnd();
            sr.Close();


            String pass0 = x.Split(',')[1];
            String pass1 = x.Split(',')[3];
            String pass2 = x.Split(',')[5];
            String pass3 = x.Split(',')[7];
            String pass4 = x.Split(',')[9];
            String pass5 = x.Split(',')[11];

            if (comboBox3.SelectedIndex == 0)
            {
                textBox19.Text = pass0;
            }
            if (comboBox3.SelectedIndex == 1)
            {
                textBox19.Text = pass1;
            }
            if (comboBox3.SelectedIndex == 2)
            {
                textBox19.Text = pass2;
            }
            if (comboBox3.SelectedIndex == 3)
            {
                textBox19.Text = pass3;
            }
            if (comboBox3.SelectedIndex == 4)
            {
                textBox19.Text = pass4;
            }
            if (comboBox3.SelectedIndex == 5)
            {
                textBox19.Text = pass5;
            }



        }

        private void comboBox3_SelectedValueChanged(object sender, EventArgs e)
        {

        }

        private void comboBox3_MouseClick(object sender, MouseEventArgs e)
        {



            StreamReader sr = new StreamReader(@"C:\compartilhamento\pass\pass.txt");
            string x = sr.ReadToEnd();
            sr.Close();


            String pass0 = x.Split(',')[0];
            String pass1 = x.Split(',')[2];
            String pass2 = x.Split(',')[4];
            String pass3 = x.Split(',')[6];
            String pass4 = x.Split(',')[8];
            String pass5 = x.Split(',')[10];

            comboBox3.Items.Clear();
            comboBox3.Items.Insert(0, pass0);
            comboBox3.Items.Insert(1, pass1);
            comboBox3.Items.Insert(2, pass2);
            comboBox3.Items.Insert(3, pass3);
            comboBox3.Items.Insert(4, pass4);
            comboBox3.Items.Insert(5, pass5);


        }

        private void comuser_MouseClick(object sender, MouseEventArgs e)
        {
            StreamReader sr = new StreamReader(@"C:\compartilhamento\pass\pass.txt");
            string x = sr.ReadToEnd();
            sr.Close();


            String pass0 = x.Split(',')[0];
            String pass1 = x.Split(',')[2];
            String pass2 = x.Split(',')[4];
            String pass3 = x.Split(',')[6];
            String pass4 = x.Split(',')[8];
            String pass5 = x.Split(',')[10];

            comuser.Items.Clear();
            comuser.Items.Insert(0, pass0);
            comuser.Items.Insert(1, pass1);
            comuser.Items.Insert(2, pass2);
            comuser.Items.Insert(3, pass3);
            comuser.Items.Insert(4, pass4);
            comuser.Items.Insert(5, pass5);
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {

        }

        private void button46_Click_1(object sender, EventArgs e)
        {
            using (StreamWriter writer = new StreamWriter(@"C:\compartilhamento\IP_NEW.txt", false)) //@"\data_full\vessel\local_list.txt("C:\\data_full\\local.txt", true))
            {
                // writer.WriteLine(textBox1.Text + "," + textBox2.Text + "," + textBox3.Text + "," + comboBox2.Text + ",");
                writer.Write(IP_START.Text + "," + IP_STOP.Text);
                writer.Close();
            }
            System.Windows.Forms.Application.Restart();
        }

        private void regs_Click(object sender, EventArgs e)
        {

        }
        public static bool HasWriteAccessToFolder(string folderPath)
        {
            try
            {
                // Attempt to get a list of security permissions from the folder. 
                // This will raise an exception if the path is read only or do not have access to view the permissions. 
                System.Security.AccessControl.DirectorySecurity ds = Directory.GetAccessControl(folderPath);
                MessageBox.Show("sim");
                return true;
            }
            catch (UnauthorizedAccessException)
            {
                MessageBox.Show("não");
                return false;
            }
        }

        public static String RemoveEnd(String str, int len)
        {
            if (str.Length < len)
            {
                return string.Empty;
            }

            return str.Remove(str.Length - len);
        }


        private void button47_Click(object sender, EventArgs e)


        {
            apaga_palavra_txt();

        }

        private void button50_Click(object sender, EventArgs e)
        {
            var drives = DriveInfo.GetDrives().Where(drive => drive.IsReady && drive.DriveType == DriveType.Removable);
            if (drives.FirstOrDefault() != null)
            {

                lblname.Text = drives.FirstOrDefault().Name.ToString();
                string fileName = "2022_02_22.xls";
                string sourcePath = @"C:\compartilhamento\data_base";
                string targetPath = drives.FirstOrDefault().Name.ToString() + "\\CRIPTOQRCODE_AllBackup\\";

                string destFile = System.IO.Path.Combine(targetPath, fileName);
                System.IO.Directory.CreateDirectory(targetPath);

                // string folder = @"D:\CRIPTOQRCODE_AllBackup\"; //nome do diretorio a ser criado

                //Se o diretório não existir...

                if (!Directory.Exists(targetPath))
                {

                    //Criamos um com o nome folder
                    Directory.CreateDirectory(targetPath);

                }



                if (System.IO.Directory.Exists(sourcePath))
                {
                    string[] files = System.IO.Directory.GetFiles(sourcePath);

                    foreach (string s in files)
                    {

                        fileName = System.IO.Path.GetFileName(s);
                        destFile = System.IO.Path.Combine(targetPath, fileName);
                        System.IO.File.Copy(s, destFile, true);
                    }

                    cria_CarregarPlanilha();
                    MessageBox.Show("Backup creaded with success");

                    if (textBox7.SelectionLength >= 0)
                    {
                        textBox7.Focus();
                        textBox7.Text = "";
                    }
                }


                else
                {
                    Console.WriteLine("Source path does not exist!");
                    if (textBox7.SelectionLength >= 0)
                    {
                        textBox7.Focus();
                        textBox7.Text = "";
                    }
                }
            }
            else
            {
                MessageBox.Show("No Pendrive found..");
                if (textBox7.SelectionLength >= 0)
                {
                    textBox7.Focus();
                    textBox7.Text = "";
                }
            }

        }

        private void button49_Click(object sender, EventArgs e)
        {
            //https://zetcode.com/csharp/excel/ dicas excel
            SautinSoft.UseOffice u = new SautinSoft.UseOffice();

            string inpFile = Path.GetFullPath(@"D:\Controle de Acesso.xlsx");
            string outFile = Path.GetFullPath(@"D:\Result.txt");
            int ret = u.InitExcel();
            if (ret == 1)
            {
                Console.WriteLine("Error! Can't load MS Excel library in memory");
                return;
            }

            ret = u.ConvertFile(inpFile, outFile, SautinSoft.UseOffice.eDirection.XLS_to_TEXT);

            u.CloseExcel();

            string text = File.ReadAllText(@"D:\Result.txt");
            text = text.Replace("	", ":");
            text = text.Replace(".", "");
            text = text.Replace("-", "");
            //var lines2 = File.ReadAllLines(@"D:\teste3.txt");
            //File.WriteAllLines(@"D:\teste3.txt", lines2.Skip(1).ToArray());

            File.WriteAllText(@"D:\teste.txt", text);

            string text2 = File.ReadAllText(@"D:\teste.txt");
            text2 = text2.Replace("::", ":");
            text2 = text2.Replace("::::", ":"); //:::::::::::::::
            text2 = text2.Replace(":::::::::::::::", ":"); //:::::::::::::::
            text2 = text2.Replace(":::", ":"); //:::::::::::::::
            text2 = text2.Replace("::::::", ":"); //:::::::::::::::
            text2 = text2.Replace("::", ":");
            text2 = text2.Replace("-", "");
            text2 = text2.ToUpper(new CultureInfo("en-US", false));
            File.WriteAllText(@"D:\teste3.txt", text2);
            //var lines = File.ReadAllLines(@"D:\teste3.txt");
            //File.WriteAllLines(@"D:\teste3.txt", lines.Skip(1).ToArray());


            var lines = System.IO.File.ReadAllLines(@"D:\teste3.txt");
            System.IO.File.WriteAllLines(@"D:\teste3.txt", lines.Skip(1));
            lines = System.IO.File.ReadAllLines(@"D:\teste3.txt");
            System.IO.File.WriteAllLines(@"D:\teste3.txt", lines.Skip(1));
            //  RemoveFirstLinesFromFile(@"D:\teste3.txt",1);
            String palavra = File.ReadAllText(@"D:\teste3.txt");
            File.WriteAllText(@"C:\compartilhamento\data_txt\data2.txt", palavra);
            MessageBox.Show("Update concluido com sucesso!");
        }

        private void button48_Click(object sender, EventArgs e)
        {
            /*
           using ClosedXML.Excel;

           using var wbook = new XLWorkbook("simple.xlsx");

           var ws1 = wbook.Worksheet(1);
           var data = ws1.Cell("A1").GetValue<string>();
           */
            //ID Capacete  Identificação	Empresa	Nome Completo		Identidade	CPF	Email	Função	Data  ASO	Data dose 1	Data dose 2	Data Reforço 1	Data Reforço 2

            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)

            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;


            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);

            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            var plan = xlWorkBook.Worksheets["Planilha1"];
            plan.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            plan.Cells.Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            //ws.Ranges("C5, F5:G8").Style.Fill.BackgroundColor = XLColor.Gray;
            // xlWorkSheet.Rows.AutoFit();
            xlWorkSheet.Cells[1, 1] = "ID Capacete  Identificação";
            xlWorkSheet.Cells[1, 1].EntireColumn.ColumnWidth = 35;
            xlWorkSheet.Cells[1, 1].RowHeight = 50;
            xlWorkSheet.Range["B1:M1"].Interior.Color = Color.FromArgb(91, 155, 213);// System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                                                                                     // xlWorkSheet.Range["A1:M1"].Rows;
            xlWorkSheet.Cells[1, 2] = "Empresa";
            xlWorkSheet.Cells[1, 3] = "Nome Completo";
            xlWorkSheet.Cells[1, 4] = "Identidade";
            xlWorkSheet.Cells[1, 5] = "CPF";
            xlWorkSheet.Cells[1, 6] = "Email";
            xlWorkSheet.Cells[1, 7] = "Função";
            xlWorkSheet.Cells[1, 8] = "Data  ASO";
            xlWorkSheet.Cells[1, 9] = "Data  NR34";
            xlWorkSheet.Cells[1, 10] = "Data dose 1";
            xlWorkSheet.Cells[1, 11] = "Data dose 2";
            xlWorkSheet.Cells[1, 12] = "Data Reforço 1";
            xlWorkSheet.Cells[1, 13] = "Data Reforço 2";
            Excel.Range _range;

            _range = xlWorkSheet.get_Range("B1", "M2000");

            //Get the borders collection.

            Excel.Borders borders = _range.Borders;

            //Set the hair lines style.

            borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            borders.Weight = 2d;

            xlWorkSheet.Columns.AutoFit();

            xlWorkBook.SaveAs("d:\\Controle de Acesso_backup.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            MessageBox.Show("Excel file created , you can find the file d:\\Controle de Acesso_backup.xls");

            //panel15.Visible = false;
        }

        private void panel11_Paint(object sender, PaintEventArgs e)
        {

        }

        private void timer5_Tick(object sender, EventArgs e)
        {

        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {

        }

        private void timer11_Tick(object sender, EventArgs e)
        {
           
           // timer11.Stop();
        }

        private void timer11_Tick_1(object sender, EventArgs e)
        {
            
        }

        private void timer12_Tick(object sender, EventArgs e)
        {

        }

        private void timer11_Tick_2(object sender, EventArgs e)
        {
            Ping pingClass = new Ping();
            PingReply pingReply = pingClass.Send(textBox18.Text.Trim());
            label65.Text = (pingReply.RoundtripTime.ToString())+ " ms";
            //+ "ms");
           // label64.Text = (pingReply.Status.ToString());
        }

        private void label61_Click(object sender, EventArgs e)
        {

        }

        private void button24_Click(object sender, EventArgs e)
        {
            if (textBox18.Text != "") {
                if (timer11.Enabled == false)
                {
                    button24.Text = "PING ON";
                    timer11.Enabled = true;
                }
                else
                {
                    button24.Text = "PING OFF";
                    timer11.Enabled = false;
                    label65.Text = "";
                }
            }
        }

        private void label65_Click(object sender, EventArgs e)
        {

        }

        private void timer12_Tick_1(object sender, EventArgs e)
        {
            DateTime fdataa = File.GetLastWriteTime(@"C:\compartilhamento\data_txt\data.txt");
            if (pega == false)
            {
                DateTime fdatab = fdataa;
                pega = true;
            }



            if (fdatab != fdataa)
            {
                ler_linha();
                pega = false;
            }
        }

        private void button48_Click_1(object sender, EventArgs e)
        {
            String data_new;
            String data2_new;
            if (dateTimePicker1.Visible == true)
            {
                data_new = dateTimePicker1.Value.ToString("dd/MM/yyyy").ToString().Trim();
            }
            else
            {
                data_new = richTextBox6.Text.Trim();
            }
            if (dateTimePicker2.Visible == true)
            {
                data2_new = dateTimePicker2.Value.ToString("dd/MM/yyyy").ToString().Trim();
            }
            else
            {
                data2_new = richTextBox7.Text.Trim();
            }
            richTextBox1.Text = " VISITANTE  "+richTextBox1.Text ;
            data2 = number + " " + richTextBox16.Text + "\r\n" + nome + " " + richTextBox2.Text + "\r\n" + emp + " " + "VISITANTE " + richTextBox1.Text + " \r\n" + function + " " + richTextBox3.Text + "\r\n" + id + " " +
         this.richTextBox4.Text + "\r\n" + email + " " + this.richTextBox8.Text + "\r\n" + vessel + " " + this.comboBox1.Text.Trim() + "\r\n" + this.richTextBox9.Text + "\r\n" + this.richTextBox10.Text + "\r\n" + this.richTextBox11.Text + "\r\n" + this.richTextBox12.Text + "\r\n" + this.richTextBox13.Text + "\r\n" + this.richTextBox14.Text + "\r\n" +
         initial + " " + data_new + "\r\n" +
         final + " " + data2_new + "\r\n" + path3 + "\r\n" + local1val + "\r\n" + local2val + "\r\n" + local3val + "\r\n" + local4val + "\r\n" + levelyellow + "\r\n" + levelgreen + "\r\n" + levelred;
            MessageBox.Show(data2);
        }

        private void bindingSource1_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void bindingSource2_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void button49_Click_1(object sender, EventArgs e)
        {
            //create_qrcode_invited_new(); ;// cadastrar_invited()
            cadastrar_invited();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button31_Click(object sender, EventArgs e)
        {

        }
    }
}
