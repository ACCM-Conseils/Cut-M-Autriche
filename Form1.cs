using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Reflection;
using System.Resources;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Advantech.Adam;
using Microsoft.VisualBasic.FileIO;

namespace CUT_M
{
    public partial class Form1 : Form
    {
        private static readonly log4net.ILog log =
            log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private bool m_bStart;
        private AdamSocket adamModbus1;
        private AdamSocket adamModbus2;
        private int m_iTempo;
        private int m_iTempoOrigine;
        private int m_iTempoImpulsion;
        private string m_szIP1;
        private int m_iPort1;
        private string m_szIP2;
        private int m_iPort2;
        private int m_iDoTotal1, m_iDiTotal1;
        private static bool[] bData1;
        private int m_iDoTotal2, m_iDiTotal2;
        private static bool[] bData2;
        private static bool goodConditions = false;
        private static bool boutonOperateur = false;
        private static bool decoupeencours = false;
        private static bool capot = false;
        private static bool pedale = false;
        private static bool ouverturePorte = false;
        private static bool finProd = false;
        private static bool porte = false;
        private static bool origine = false;
        private static bool laser = false;
        private static bool start = false;
        private static bool demarrage = false;
        private static long timestart = 0;
        private static long dureeWatchdog = 0;

        //Traductions
        private static string ChoixRef = string.Empty;
        private static string ConnectAdam1 = string.Empty;
        private static string ConnectAdam2 = string.Empty;
        private static string InitAdam1 = string.Empty;
        private static string InitAdam2 = string.Empty;
        private static string Timer1Start = string.Empty;
        private static string NoConnectAdam1 = string.Empty;
        private static string Timer2Start = string.Empty;
        private static string NoConnectAdam2 = string.Empty;
        private static string CloseAndStart = string.Empty;
        private static string Adam1NotConnected = string.Empty;
        private static string NoRef = string.Empty;
        private static string Quitter = string.Empty;
        private static string FermerPorte = string.Empty;
        private static string OuvrirShutter = string.Empty;
        private static string DemarrageDecoupe = string.Empty;
        private static string Fermee = string.Empty;
        private static string ChoisirRef = string.Empty;
        private static string OuvrirPorte = string.Empty;
        private static string PlacerPiece = string.Empty;
        private static string DepartCycle = string.Empty;

        //Variables debug
        private static bool forcePorte = false;
        private static bool forceDepart = false;
        private static bool forceShutter = false;

        private static List<Produit> produits = new List<Produit>();
        private static List<Production> production = new List<Production>();
        public Form1()
        {
            CultureInfo ci = new CultureInfo ("en"); Thread.CurrentThread.CurrentCulture = ci; Thread.CurrentThread.CurrentUICulture = ci;

            InitializeComponent();

            timer1.Interval = 300;
            timer1.Tick += new System.EventHandler(this.timer1_Tick);

            timer2.Interval = 300;
            timer2.Tick += new System.EventHandler(this.timer2_Tick);

            Warning warn = new Warning();
            warn.FormBorderStyle = FormBorderStyle.None;
            warn.WindowState = FormWindowState.Maximized;

            if (warn.ShowDialog() == DialogResult.OK)
            {
                log4net.Config.XmlConfigurator.Configure();

                log.Info("Starting Cut-M");

                InitLang();

                try
                {
                    InitCutM();

                    Application.DoEvents();

                    LoadRef();

                    Application.DoEvents();

                    InitLaser();

                }
                catch(Exception e)
                {
                    log.Error(e);
                }
            }
            else
                this.Close();
        }
        private void InitLang()
        {
            ResourceManager res_man = new ResourceManager("CUT_M.Lang_" + System.Globalization.CultureInfo.CurrentUICulture.TwoLetterISOLanguageName, Assembly.GetExecutingAssembly());

            lblRef.Text = res_man.GetString("lblRef");
            lblInfos.Text = res_man.GetString("lblInfos");
            lblProd.Text = res_man.GetString("lblProd");
            btAnnule.Text = res_man.GetString("btAnnule");
            lblTitreDiametre.Text = res_man.GetString("lblTitreDiametre");
            lblTitreQte.Text = res_man.GetString("lblTitreQte");
            lblTitreQte.Text = res_man.GetString("lblTitreQte");
            lblTitreEtat.Text = res_man.GetString("lblTitreEtat");
            lblTitreEtage.Text = res_man.GetString("lblTitreEtage");
            lbltitrePorte.Text = res_man.GetString("lbltitrePorte");
            lbltitreLaser.Text = res_man.GetString("lbltitreLaser");

            ChoixRef = res_man.GetString("ChoixRef");
            lblInfo.Text = ChoixRef;

            ConnectAdam1 = res_man.GetString("ConnectAdam1");
            ConnectAdam2 = res_man.GetString("ConnectAdam2");
            InitAdam1 = res_man.GetString("InitAdam1");
            InitAdam2 = res_man.GetString("InitAdam2");
            Timer1Start = res_man.GetString("Timer1Start");
            NoConnectAdam1 = res_man.GetString("NoConnectAdam1");
            Timer2Start = res_man.GetString("Timer2Start");
            NoConnectAdam2 = res_man.GetString("NoConnectAdam2");
            CloseAndStart = res_man.GetString("CloseAndStart");
            Adam1NotConnected = res_man.GetString("Adam1NotConnected");
            NoRef = res_man.GetString("NoRef");
            Quitter = res_man.GetString("Quitter");
            FermerPorte = res_man.GetString("FermerPorte");
            OuvrirShutter = res_man.GetString("OuvrirShutter");
            Fermee = res_man.GetString("Fermee");
            ChoisirRef = res_man.GetString("ChoisirRef");
            OuvrirPorte = res_man.GetString("OuvrirPorte");
            PlacerPiece = res_man.GetString("PlacerPiece");
            DepartCycle = res_man.GetString("DepartCycle");

            log.Error("Loading translations OK");
        }

        private void InitCutM()
        {
            log.Info("InitCutM");

            if (!this.IsHandleCreated)
            {
                this.CreateHandle();

            }            

            m_bStart = false;			// the action stops at the beginning
            m_szIP1 = ut_xml.ValueXML(@".\CUT-M.xml", "IPAdam1");	// modbus slave IP address for Adam1
            log.Info(string.Format("IPAdam1 : {0}", m_szIP1));
            m_iPort1 = System.Convert.ToInt32(ut_xml.ValueXML(@".\CUT-M.xml", "PortAdam1"));				// modbus TCP port for Adam1
            log.Info(string.Format("PortAdam1 : {0}", m_iPort1));
            m_szIP2 = ut_xml.ValueXML(@".\CUT-M.xml", "IPAdam2");	// modbus slave IP address for Adam1
            log.Info(string.Format("IPAdam2 : {0}", m_szIP2));
            m_iPort2 = System.Convert.ToInt32(ut_xml.ValueXML(@".\CUT-M.xml", "PortAdam2"));				// modbus TCP port for Adam1
            log.Info(string.Format("PortAdam2 : {0}", m_iPort2));
            dureeWatchdog = System.Convert.ToInt32(ut_xml.ValueXML(@".\CUT-M.xml", "dureeWatchdog"));
            log.Info(string.Format("dureeWatchdog : {0}", dureeWatchdog));
            m_iTempo = System.Convert.ToInt32(ut_xml.ValueXML(@".\CUT-M.xml", "TempoMoteur"));
            log.Info(string.Format("TempoMoteur : {0}", m_iTempo));
            m_iTempoOrigine = System.Convert.ToInt32(ut_xml.ValueXML(@".\CUT-M.xml", "TempoOrigine"));
            log.Info(string.Format("TempoOrigine : {0}", m_iTempoOrigine));
            m_iTempoImpulsion = System.Convert.ToInt32(ut_xml.ValueXML(@".\CUT-M.xml", "TempoImpulsion"));
            log.Info(string.Format("TempoImpulsion : {0}", m_iTempoImpulsion));

            log.Error("Loading params OK");

            adamModbus1 = new AdamSocket();
            adamModbus1.SetTimeout(1000, 1000, 1000);
            adamModbus2 = new AdamSocket();
            adamModbus2.SetTimeout(1000, 1000, 1000);
            label42.Invoke(new EventHandler(delegate { label42.Text = finProd.ToString(); }));
            label39.Invoke(new EventHandler(delegate { label39.Text = start.ToString(); }));

            if (adamModbus1.Connect(m_szIP1, ProtocolType.Tcp, m_iPort1))
            {
                log.Info(ConnectAdam1);

                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, ConnectAdam1); }));

                lblEtat.Text = ConnectAdam1;
                lblEtat.ForeColor = Color.Green;

                int iDI = 0, iDO = 0;

                InitChannelItems(true, true, ref iDI, ref iDO);
                InitChannelItems(true, true, ref iDI, ref iDO);
                InitChannelItems(true, true, ref iDI, ref iDO);
                InitChannelItems(true, true, ref iDI, ref iDO);
                InitChannelItems(true, true, ref iDI, ref iDO);
                InitChannelItems(true, true, ref iDI, ref iDO);
                InitChannelItems(true, true, ref iDI, ref iDO);
                InitChannelItems(true, true, ref iDI, ref iDO);
                InitChannelItems(true, false, ref iDI, ref iDO);
                InitChannelItems(true, false, ref iDI, ref iDO);
                InitChannelItems(true, false, ref iDI, ref iDO);
                InitChannelItems(true, false, ref iDI, ref iDO);
                InitChannelItems(true, false, ref iDI, ref iDO);
                InitChannelItems(true, false, ref iDI, ref iDO);
                InitChannelItems(true, false, ref iDI, ref iDO);
                InitChannelItems(true, false, ref iDI, ref iDO);

                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, InitAdam1); }));

                m_iDoTotal1 = iDO;
                m_iDiTotal1 = iDI;

                int iCnt;
                bool bCommFSV;
                bool bPtoPFSV;
                bool[] bWDT;

                //if (m_Adam6000Type == Adam6000Type.Adam6055) // no DO for 6055
                if (m_iDoTotal1 == 0)
                {
                    return;
                }

                if (adamModbus1.DigitalOutput().GetWDTMask(out bCommFSV, out bPtoPFSV, out bWDT) && bWDT.Length == 8)
                {
                    iCnt = 0;
                }
                else
                    MessageBox.Show("GetWDTMask() failed;");                

                timer1.Start();

                log.Info(Timer1Start);

                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, Timer1Start); }));

                ChangeOID1(8, 0);
                ChangeOID1(9, 0);
                ChangeOID1(10, 0);
                ChangeOID1(11, 0);
                ChangeOID1(12, 0);
                ChangeOID1(13, 0);
                ChangeOID1(14, 0);
                ChangeOID1(15, 0);

                RefreshDIO1();
            }
            else
            {
                lblEtat.Text = NoConnectAdam1;
                lblEtat.ForeColor = Color.Red;

                log.Error("Adam 1 module initialisation impossible");

                lblPorte.Text = Adam1NotConnected;
                lblLaser.Text = Adam1NotConnected;

                ListViewItem lv = new ListViewItem(NoConnectAdam1);
                lv.ForeColor = Color.Red;
                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));

                timer1.Start();
            }

            if (adamModbus2.Connect(m_szIP2, ProtocolType.Tcp, m_iPort2))
            {
                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, ConnectAdam2); }));

                lblEtat2.Text = ConnectAdam2;
                lblEtat2.ForeColor = Color.Green;

                int iDI = 0, iDO = 0;

                InitChannelItems(true, false, ref iDI, ref iDO);
                InitChannelItems(true, false, ref iDI, ref iDO);
                InitChannelItems(true, false, ref iDI, ref iDO);
                InitChannelItems(true, false, ref iDI, ref iDO);
                InitChannelItems(true, false, ref iDI, ref iDO);
                InitChannelItems(true, false, ref iDI, ref iDO);
                InitChannelItems(true, false, ref iDI, ref iDO);
                InitChannelItems(true, false, ref iDI, ref iDO);
                InitChannelItems(true, false, ref iDI, ref iDO);
                InitChannelItems(true, false, ref iDI, ref iDO);
                InitChannelItems(true, false, ref iDI, ref iDO);
                InitChannelItems(true, false, ref iDI, ref iDO);
                InitChannelItems(true, false, ref iDI, ref iDO);
                InitChannelItems(true, false, ref iDI, ref iDO);
                InitChannelItems(true, false, ref iDI, ref iDO);
                InitChannelItems(true, false, ref iDI, ref iDO);

                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, InitAdam2); }));

                m_iDoTotal2 = iDO;
                m_iDiTotal2 = iDI;

                timer2.Start();

                log.Info(Timer2Start);

                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, Timer2Start); }));

                ChangeOID2(10, 1);

                RefreshDIO2();

            }
            else
            {
                log.Error("Adam 2 module initialisation impossible");

                lblEtat2.Text = NoConnectAdam2;
                lblEtat2.ForeColor = Color.Red;

                ListViewItem lv = new ListViewItem(NoConnectAdam2);
                lv.ForeColor = Color.Red;
                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));
            }

            this.timer3.Interval = 300;
            this.timer3.Tick += new System.EventHandler(this.timer3_Tick);
        }

        private void InitLaser()
        {
            try
            {
                log.Info("InitLaser");

                Application.DoEvents();

                if (!porte || !laser)
                {
                    while (!porte || !laser)
                    {
                        Thread.Sleep(1000);

                        Application.DoEvents();
                        this.TopMost = false;
                        MessageBox.Show(new Form { TopMost = true }, CloseAndStart, "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = CloseAndStart; }));

                        Application.DoEvents();
                    }
                }

                ChangeOID1(15, 1);

                Thread.Sleep(m_iTempoImpulsion);

                ChangeOID1(15, 0);

                Process[] workers = Process.GetProcessesByName("CutMaker.exe");
                foreach (Process worker in workers)
                {
                    log.Info("Closing CutMaker.exe");
                    worker.Kill();
                    worker.WaitForExit();
                    worker.Dispose();
                }

                Thread.Sleep(2000);

                String exePath = ut_xml.ValueXML(@".\CUT-M.xml", "LaserExe");
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = exePath;
                //startInfo.Arguments = "header.h";
                Process.Start(startInfo);

                ChangeOID1(15, 1);

                Thread.Sleep(1500);

                while (!origine)
                {
                    Thread.Sleep(1000);
                    Application.DoEvents();
                }

                log.Info("Origin OK");

                ChangeOID1(15, 0);

                Application.DoEvents();
            }
            catch(Exception ex)
            {
                log.Error(ex);
            }
        }

        protected void InitChannelItems(bool i_bVisable, bool i_bIsDI, ref int i_iDI, ref int i_iDO)
        {
            int iCh;
            if (i_bVisable)
            {
                iCh = i_iDI + i_iDO;
                if (i_bIsDI) // DI
                {
                    i_iDI++;
                }
                else // DO
                {
                    i_iDO++;
                }
            }
        }

        protected void RefreshDIO1()
        {
            try
            {
                int iDiStart = 1, iDoStart = 17;
                int iChTotal;
                bool[] bDiData, bDoData;

                if (adamModbus1.Modbus().ReadCoilStatus(iDiStart, m_iDiTotal1, out bDiData) &&
                    adamModbus1.Modbus().ReadCoilStatus(iDoStart, m_iDoTotal1, out bDoData))
                {
                    iChTotal = m_iDiTotal1 + m_iDoTotal1;
                    bData1 = new bool[iChTotal];
                    Array.Copy(bDiData, 0, bData1, 0, m_iDiTotal1);
                    Array.Copy(bDoData, 0, bData1, m_iDiTotal1, m_iDoTotal1);

                    bool good = true;
                    String Message = string.Empty;

                    if (!bData1[0])
                    {
                        Message += FermerPorte + Environment.NewLine;
                        porte = false;
                    }
                    else
                        porte = true;

                    //label37.Invoke(new EventHandler(delegate { label37.Text = porte.ToString(); }));

                    if (!bData1[1] && !forceShutter)
                    {
                        Message += OuvrirShutter + Environment.NewLine;
                        laser = false;
                    }
                    else
                        laser = true;
                    //label38.Invoke(new EventHandler(delegate { label38.Text = laser.ToString(); }));
                    if (bData1[2] || forceDepart)
                    {
                        if (laser && porte)
                        {                     
                            log.Info("Départ cycle reçu");
                            lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, "Départ cycle reçu"); }));

                            boutonOperateur = true;
                            checkBox6.Checked = false;
                            forceDepart = false;

                            lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = DemarrageDecoupe; }));
                        }
                        else
                        {
                            lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = CloseAndStart; }));
                        }
                    }
                    if (bData1[3])
                    {
                        if (!decoupeencours)
                        {
                            ouverturePorte = true;

                            ChangeOID2(10, 0);
                            log.Info("Door open signal");
                           lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, "Ouverture porte reçu"); }));

                            Application.DoEvents();
                        }
                    }
                    else
                    {
                        ouverturePorte = false;

                        ChangeOID2(10, 1);
                    }

                    if (!bData1[4])
                    {
                        pedale = false;
                    }
                    else
                        pedale = true;

                    if (!bData1[6])
                    {
                        origine = false;
                    }
                    else
                        origine = true;

                    if (bData1[5])
                    {
                        finProd = true;
                    }
                    else
                        finProd = false;

                    //label42.Invoke(new EventHandler(delegate { label42.Text = finProd.ToString(); }));
                    //label39.Invoke(new EventHandler(delegate { label39.Text = start.ToString(); }));

                    txtDO0.Text = bData1[8].ToString();
                    txtDO1.Text = bData1[9].ToString();
                    txtDO2.Text = bData1[10].ToString();
                    txtDO3.Text = bData1[11].ToString();
                    txtDO4.Text = bData1[12].ToString();
                    txtDO5.Text = bData1[13].ToString();
                    txtDO6.Text = bData1[14].ToString();
                    txtDO7.Text = bData1[15].ToString();

                    txtDI0.Text = bData1[0].ToString();
                    txtDI1.Text = bData1[1].ToString();
                    txtDI2.Text = bData1[2].ToString();
                    txtDI3.Text = bData1[3].ToString();
                    txtDI4.Text = bData1[4].ToString();
                    txtDI5.Text = bData1[5].ToString();
                    txtDI6.Text = bData1[6].ToString();

                    if (porte)
                    {
                        lblPorte.Invoke(new EventHandler(delegate { lblPorte.ForeColor = Color.Green; lblPorte.Text = Fermee; }));
                    }
                    else
                    {
                        lblPorte.Invoke(new EventHandler(delegate { lblPorte.ForeColor = Color.Red; lblPorte.Text = "Ouverte"; }));
                    }

                    if (laser)
                    {
                        lblLaser.Invoke(new EventHandler(delegate { lblLaser.ForeColor = Color.Green; lblLaser.Text = "Shutter ouvert"; }));
                    }
                    else
                    {
                        lblLaser.Invoke(new EventHandler(delegate { lblLaser.ForeColor = Color.Red; lblLaser.Text = "Shutter fermé"; }));
                    }

                    if (porte && laser)
                    {
                        goodConditions = true;
                    }
                    else if ((!porte || !laser) && start)
                    {
                        goodConditions = false;
                        MessageBox.Show("Production interrompue");
                        RazProd();
                    }

                    //label41.Invoke(new EventHandler(delegate { label41.Text = goodConditions.ToString(); }));

                    /*if (goodConditions && bData1[16] && comboBox1.SelectedIndex > 0)
                    {
                        ChangeOID1(16, 0);
                    }*/
                    /*if (goodConditions && comboBox1.SelectedIndex > 0 && !start && !demarrage)
                    {
                        Message = "Positionner la pièce";
                    }
                    else if (goodConditions && comboBox1.SelectedIndex > 0 && start && !boutonOperateur)
                    {
                        Message = "En attente de départ de cycle";
                    }
                    else if (goodConditions && comboBox1.SelectedIndex == 0)
                    {
                        Message = "Choisir ou saisir une référence";
                    }
                    else if (goodConditions && comboBox1.SelectedIndex > 0 && start && boutonOperateur)
                    {
                        Message = "Découpe en cours";
                    }*/



                    lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, "goodConditions : "+ goodConditions); }));
                    lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, "start : "+ start); }));

                    if (!goodConditions && start)
                    {
                        start = false;
                        finProd = true;
                        RazProd();
                    }
                }
            }
            catch(Exception ex)
            {
                log.Error(ex);
            }
        }

        protected void RefreshDIO2()
        {
            int iDiStart = 1, iDoStart = 17;
            int iChTotal;
            bool[] bDiData, bDoData, bData;

            if (adamModbus2.Modbus().ReadCoilStatus(iDoStart, m_iDoTotal2, out bDoData))
            {
                iChTotal = m_iDoTotal2;
                bData = new bool[iChTotal];
                Array.Copy(bDoData, 0, bData, 0, m_iDoTotal2);
                if (iChTotal > 0)
                    txt1DO0.Text = bData[0].ToString();
                if (iChTotal > 1)
                    txt1DO1.Text = bData[1].ToString();
                if (iChTotal > 2)
                    txt1DO2.Text = bData[2].ToString();
                if (iChTotal > 3)
                    txt1DO3.Text = bData[3].ToString();
                if (iChTotal > 4)
                    txt1DO4.Text = bData[4].ToString();
                if (iChTotal > 5)
                    txt1DO5.Text = bData[5].ToString();
                if (iChTotal > 6)
                    txt1DO6.Text = bData[6].ToString();
                if (iChTotal > 7)
                    txt1DO7.Text = bData[7].ToString();
                if (iChTotal > 8)
                    txt1DO8.Text = bData[8].ToString();
                if (iChTotal > 9)
                    txt1DO9.Text = bData[9].ToString();
                if (iChTotal > 10)
                    txt1DO10.Text = bData[10].ToString();
            }
            else
            {
                txt1DO0.Text = "Fail";
                txt1DO1.Text = "Fail";
                txt1DO2.Text = "Fail";
                txt1DO3.Text = "Fail";
                txt1DO4.Text = "Fail";
                txt1DO5.Text = "Fail";
                txt1DO6.Text = "Fail";
                txt1DO7.Text = "Fail";
                txt1DO8.Text = "Fail";
                txt1DO9.Text = "Fail";
                txt1DO10.Text = "Fail";
            }
        }

        protected bool LoadRef()
        {
            try
            {
                DataTable table1 = new DataTable("references");
                table1.Columns.Add("reference");
                table1.Columns.Add("quantite");

                string pathClient = ut_xml.ValueXML(@".\CUT-M.xml", "DossierClient");
                string fileClient = ut_xml.ValueXML(@".\CUT-M.xml", "FichierClient");

                if (File.Exists(Path.Combine(pathClient, fileClient)))
                {
                    production = new List<Production>();
                    using (TextFieldParser parser = new TextFieldParser(Path.Combine(pathClient, fileClient)))
                    {
                        parser.TextFieldType = FieldType.Delimited;
                        parser.SetDelimiters(";");
                        while (!parser.EndOfData)
                        {
                            //Processing row
                            string[] fields = parser.ReadFields();
                            table1.Rows.Add(fields[0], System.Convert.ToInt32(fields[1]));
                            production.Add(new Production() { reference = fields[0], restant = System.Convert.ToInt32(fields[1]) });
                        }

                        DataRow row = table1.NewRow();
                        row[0] = ChoisirRef;
                        row[1] = 0;
                        table1.Rows.InsertAt(row, 0);

                        comboBox1.DataSource = table1;
                        comboBox1.DisplayMember = "reference";
                    }

                    string pathProduit = ut_xml.ValueXML(@".\CUT-M.xml", "DossierRef");
                    string fileProduit = ut_xml.ValueXML(@".\CUT-M.xml", "FichierRef");

                    using (TextFieldParser parser = new TextFieldParser(Path.Combine(pathProduit, fileProduit)))
                    {
                        parser.TextFieldType = FieldType.Delimited;
                        parser.SetDelimiters(";");
                        while (!parser.EndOfData)
                        {
                            try
                            {
                                //Processing row
                                string[] fields = parser.ReadFields();
                                produits.Add(new Produit() { reference = fields[0], diametre = System.Convert.ToInt32(fields[1]), positionangle = fields[2], masque = fields[3], etage = System.Convert.ToInt32(fields[4]) });
                            }
                            catch
                            { }
                        }
                    }
                    return true;
                }
                else
                {
                    table1 = new DataTable("references");
                    table1.Columns.Add("reference");
                    table1.Columns.Add("quantite");

                    DataRow row = table1.NewRow();
                    row[0] = NoRef;
                    row[1] = 0;
                    table1.Rows.InsertAt(row, 0);

                    comboBox1.DataSource = table1;
                    comboBox1.DisplayMember = "reference";

                    return false;
                }

            }
            catch (Exception e)
            {
                return false;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex > 0)
            {
                start = true;
                ChangeOID2(0, 0);
                ChangeOID2(1, 0);
                ChangeOID2(2, 0);
                ChangeOID2(3, 0);
                ChangeOID2(4, 0);
                ChangeOID2(5, 0);
                ChangeOID2(6, 0);
                ChangeOID2(7, 0);
                ChangeOID2(8, 0);
                ChangeOID2(9, 0);

                String _ref = comboBox1.Text;

                Produit produit = produits.FirstOrDefault(m => m.reference == _ref);

                if (produit != null)
                {
                    log.Info(string.Format("Selected reference : {0}", comboBox1.Text));
                    lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("Selection référénce : {0}", comboBox1.Text)); }));

                    Production prod = production.FirstOrDefault(m => m.reference == _ref);

                    lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("Diametre : {0}", produit.diametre.ToString())); }));
                    log.Info(string.Format("Diameter : {0}", produit.diametre.ToString()));
                    lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("Etage : {0}", produit.etage.ToString())); }));
                    log.Info(string.Format("Floor : {0}", produit.etage.ToString()));
                    lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("Qte : {0}", prod.restant.ToString())); }));
                    log.Info(string.Format("Quantity : {0}", prod.restant.ToString()));

                    lblDiametre.Text = produit.diametre.ToString();
                    lblEtage.Text = produit.etage.ToString();
                    lblQte.Text = prod.restant.ToString();

                    try
                    {
                        log.Info(string.Format("Mask : {0}", produit.masque.ToString()));
                        lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("Masque : {0}", produit.masque.ToString())); }));

                        log.Info("Loading mask");
                        lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = "Chargement du masque"; }));

                        Application.DoEvents();

                        int DO0_1 = int.Parse(produit.masque[0].ToString());
                        int DO0_2 = int.Parse(produit.masque[1].ToString());
                        int DO0_3 = int.Parse(produit.masque[2].ToString());
                        int DO0_4 = int.Parse(produit.masque[3].ToString());
                        int DO0_5 = int.Parse(produit.masque[4].ToString());
                        int DO0_6 = int.Parse(produit.masque[5].ToString());

                        lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("DO0_1 : {0}", DO0_1.ToString())); }));
                        log.Info(String.Format("DO0_1 : {0}", DO0_1.ToString()));
                        lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("DO0_2 : {0}", DO0_2.ToString())); }));
                        log.Info(String.Format("DO0_2 : {0}", DO0_2.ToString()));
                        lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("DO0_3 : {0}", DO0_3.ToString())); }));
                        log.Info(String.Format("DO0_3 : {0}", DO0_3.ToString()));
                        lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("DO0_4 : {0}", DO0_4.ToString())); }));
                        log.Info(String.Format("DO0_4 : {0}", DO0_4.ToString()));
                        lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("DO0_5 : {0}", DO0_5.ToString())); }));
                        log.Info(String.Format("DO0_5 : {0}", DO0_5.ToString()));
                        lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("DO0_6 : {0}", DO0_6.ToString())); }));
                        log.Info(String.Format("DO0_6 : {0}", DO0_6.ToString()));

                        Application.DoEvents();

                        ChangeOID2(1, DO0_6);
                        ChangeOID2(2, DO0_5);
                        ChangeOID2(3, DO0_4);
                        ChangeOID2(4, DO0_3);
                        ChangeOID2(5, DO0_2);
                        ChangeOID2(6, DO0_1);

                        Thread.Sleep(1000);
                        ChangeOID2(0, 1);
                        Thread.Sleep(m_iTempoImpulsion);
                        ChangeOID2(0, 0);

                        log.Info("Loading mask OK");
                        lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, "Chargement masque OK"); }));

                        Application.DoEvents();
                    }
                    catch
                    {
                        log.Info("Unable to load mask");
                        ListViewItem lv = new ListViewItem("Chargement masque impossible");
                        lv.ForeColor = Color.Red;
                        lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));
                    }

                    var obj = production.FirstOrDefault(x => x.reference == comboBox1.Text);
                    int qte = obj.restant;


                    comboBox1.Invoke(new EventHandler(delegate { comboBox1.Enabled = false; }));
                    label39.Invoke(new EventHandler(delegate { label39.Text = start.ToString(); }));

                    Application.DoEvents();

                    while (qte > 0)
                    {

                        demarrage = true;
                        decoupeencours = false;

                        if (porte)
                        {
                            lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = OuvrirPorte; }));

                            Application.DoEvents();

                            while (porte && !forcePorte)
                            {
                                Thread.Sleep(1000);

                                Application.DoEvents();
                            }
                        }
                        lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = PlacerPiece; }));

                        Application.DoEvents();

                        while (!porte || forcePorte)
                        {
                            Thread.Sleep(1000);

                            Application.DoEvents();
                        }

                        ListViewItem lv = new ListViewItem(DepartCycle);
                        lv.ForeColor = Color.Green;
                        lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));

                        log.Info(DepartCycle);

                        lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = DepartCycle; }));

                        Application.DoEvents();

                        lv = new ListViewItem("boutonOperateur : " + boutonOperateur);
                        if(boutonOperateur)
                            lv.ForeColor = Color.Green;
                        else
                            lv.ForeColor = Color.Red;
                        lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));

                        lv = new ListViewItem("demarrage : " + demarrage);
                        if (demarrage)
                            lv.ForeColor = Color.Green;
                        else
                            lv.ForeColor = Color.Red;
                        lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));

                        while (!boutonOperateur && demarrage)
                        {
                            lv = new ListViewItem("boutonOperateur : " + boutonOperateur);
                            if (boutonOperateur)
                                lv.ForeColor = Color.Green;
                            else
                                lv.ForeColor = Color.Red;
                            lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));

                            lv = new ListViewItem("demarrage : " + demarrage);
                            if (demarrage)
                                lv.ForeColor = Color.Green;
                            else
                                lv.ForeColor = Color.Red;
                            lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));

                            lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = DepartCycle; }));

                            Thread.Sleep(1000);

                            Application.DoEvents();
                        }

                        ChangeOID1(15, 1);

                        lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = "Démarrage découpe"; }));

                        lv = new ListViewItem("Démarrage découpe");
                        lv.ForeColor = Color.Green;
                        lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));

                        lv = new ListViewItem("Goodcondition : "+goodConditions);
                        lv.ForeColor = Color.Green;
                        lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));

                        if (goodConditions)
                        {
                            decoupeencours = true;

                            boutonOperateur = false;

                            int Angle_1 = int.Parse(produit.positionangle[0].ToString());
                            int Angle_2 = int.Parse(produit.positionangle[1].ToString());
                            int Angle_3 = int.Parse(produit.positionangle[2].ToString());
                            int Angle_4 = int.Parse(produit.positionangle[3].ToString());
                            int Angle_5 = int.Parse(produit.positionangle[4].ToString());
                            int Angle_6 = int.Parse(produit.positionangle[5].ToString());

                            lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("Angle_1 : {0}", Angle_1.ToString())); }));
                            log.Info(String.Format("Angle_1 : {0}", Angle_1.ToString()));
                            lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("Angle_2 : {0}", Angle_2.ToString())); }));
                            log.Info(String.Format("Angle_2 : {0}", Angle_2.ToString()));
                            lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("Angle_3 : {0}", Angle_3.ToString())); }));
                            log.Info(String.Format("Angle_3 : {0}", Angle_3.ToString()));
                            lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("Angle_4 : {0}", Angle_4.ToString())); }));
                            log.Info(String.Format("Angle_4 : {0}", Angle_4.ToString()));
                            lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("Angle_5 : {0}", Angle_5.ToString())); }));
                            log.Info(String.Format("Angle_5 : {0}", Angle_5.ToString()));
                            lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("Angle_6 : {0}", Angle_6.ToString())); }));
                            log.Info(String.Format("Angle_6 : {0}", Angle_6.ToString()));

                            Application.DoEvents();

                            ChangeOID1(9, Angle_6);
                            ChangeOID1(10, Angle_5);
                            ChangeOID1(11, Angle_4);
                            ChangeOID1(12, Angle_3);
                            ChangeOID1(13, Angle_2);
                            ChangeOID1(14, Angle_1);

                            Thread.Sleep(1000);

                            ChangeOID1(8, 1);
                            Thread.Sleep(m_iTempoImpulsion);
                            ChangeOID1(8, 0);

                            Thread.Sleep(m_iTempo);

                            ChangeOID2(7, 1);
                            Thread.Sleep(m_iTempoImpulsion);
                            ChangeOID2(7, 0);

                            while (!finProd)
                            {
                                Thread.Sleep(1000);
                                Application.DoEvents();

                                lv = new ListViewItem("finProd");
                                if(finProd)
                                    lv.ForeColor = Color.Green;
                                else
                                    lv.ForeColor = Color.Red;
                                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));

                                lvOpe.Refresh();
                            }

                            finProd = false;

                            lv = new ListViewItem("Qte : " + qte);
                            lv.ForeColor = Color.Green;

                            lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));

                            ChangeOID1(15, 0);

                            Thread.Sleep(2000);

                            demarrage = false;
                            decoupeencours = false;
                        }
                    }

                    start = false;

                    /*if (goodConditions)
                        lblInfo.Text = "Placer le capot";
                    else
                    {
                        string Message = "";
                        if (!porte)
                            Message += "Fermer la porte" + Environment.NewLine;
                        if (!laser)
                            Message += "Ouvrir le shutter" + Environment.NewLine;

                        lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = Message; }));
                    }

                    button1.Invoke(new EventHandler(delegate { button1.Enabled = true; }));

                    try
                    {

                        demarrage = true;

                    }
                    catch
                    {
                        ListViewItem lv = new ListViewItem("Initialisation module Adam 1 impossible");
                        lv.ForeColor = Color.Red;
                        lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));
                    }                    

                    if (goodConditions)
                    {
                        var obj = production.FirstOrDefault(x => x.reference == comboBox1.Text);
                        int qte = obj.restant;


                        comboBox1.Invoke(new EventHandler(delegate { comboBox1.Enabled = false; }));
                        label39.Invoke(new EventHandler(delegate { label39.Text = start.ToString(); }));

                        Application.DoEvents();

                        while (qte > 0)
                        {
                            //ListViewItem lv = new ListViewItem("En attente de départ cycle");
                            //lv.ForeColor = Color.Green;
                            //lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));

                            //Application.DoEvents();

                            //while (!boutonOperateur && demarrage)
                            //{
                            //    Thread.Sleep(1000);

                            //    Application.DoEvents();

                            //    lvOpe.Refresh();
                            //}

                            if (demarrage)
                            {
                                Thread.Sleep(1000);

                                lblCapot.Invoke(new EventHandler(delegate { lblCapot.Visible = true; }));

                                capot = true;

                                Application.DoEvents();

                                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, "Début origine"); }));
                                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, "Début impulsion"); }));

                                ChangeOID1(13, 1);
                                Thread.Sleep(m_iTempoImpulsion);
                                ChangeOID1(13, 0);

                                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, "Fin impulsion"); }));

                                Thread.Sleep(m_iTempoOrigine);

                                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("Angle : {0}", produit.positionangle.ToString())); }));

                                int rotations = System.Convert.ToInt32(produit.positionangle);

                                for (int i = 0; i < rotations; i++)
                                {
                                    ChangeOID1(12, 1);
                                    Thread.Sleep(m_iTempoImpulsion);
                                    ChangeOID1(12, 0);

                                    Thread.Sleep(m_iTempo);

                                    Application.DoEvents();
                                }

                                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, "Initialisation module Adam 1 OK"); }));

                                Application.DoEvents();

                                capot = false;

                                Thread.Sleep(2000);

                                start = true;
                                decoupeencours = true;

                                Application.DoEvents();

                                if (goodConditions)
                                {
                                    Application.DoEvents();

                                    ChangeOID1(16, 1);

                                    lv = new ListViewItem("Impulsion départ");
                                    lv.ForeColor = Color.Green;
                                    lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));

                                    Thread.Sleep(1000);

                                    ChangeOID1(16, 0);

                                    lv = new ListViewItem("Fin impulsion départ");
                                    lv.ForeColor = Color.Green;
                                    lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));

                                    finProd = false;

                                    label42.Invoke(new EventHandler(delegate { label42.Text = finProd.ToString(); }));

                                    lv = new ListViewItem("Découpe en cours");
                                    lv.ForeColor = Color.Green;
                                    lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));

                                    lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = "Découpe en cours"; }));

                                    Application.DoEvents();

                                    timestart = DateTime.Now.Ticks;
                                    timer3.Enabled = true;
                                    timer3.Start();

                                    button4.Invoke(new EventHandler(delegate { button4.Enabled = false; }));

                                    while (!finProd)
                                    {
                                        Thread.Sleep(1000);
                                        Application.DoEvents();

                                        lvOpe.Refresh();
                                    }

                                    if (finProd)
                                    {

                                        lblCapot.Invoke(new EventHandler(delegate { lblCapot.Visible = false; }));

                                        label42.Invoke(new EventHandler(delegate { label42.Text = finProd.ToString(); }));

                                        timer3.Stop();
                                        timer3.Enabled = false;

                                        timer1.Enabled = false;

                                        obj = production.FirstOrDefault(x => x.reference == comboBox1.Text);
                                        qte = obj.restant -= 1;
                                        if (obj != null) obj.restant = qte;

                                        String toReplace = string.Empty;

                                        foreach (Production p in production)
                                        {
                                            toReplace += p.reference + ";" + p.restant + Environment.NewLine;
                                        }

                                        string pathClient = ut_xml.ValueXML(@".\CUT-M.xml", "DossierClient");
                                        string fileClient = ut_xml.ValueXML(@".\CUT-M.xml", "FichierClient");

                                        try
                                        {
                                            File.WriteAllText(Path.Combine(pathClient, fileClient), toReplace);

                                            lblQte.Text = obj.restant.ToString();

                                            boutonOperateur = false;
                                            decoupeencours = false;

                                            Thread.Sleep(1000);

                                            checkBox3.Checked = false;
                                            checkBox4.Checked = false;
                                            finProd = true;

                                            label40.Invoke(new EventHandler(delegate { label40.Text = boutonOperateur.ToString(); }));
                                            label42.Invoke(new EventHandler(delegate { label42.Text = finProd.ToString(); }));

                                            lv = new ListViewItem("Découpe terminée");
                                            lv.ForeColor = Color.Green;
                                            lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));

                                            lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = "Découpe terminée"; }));

                                            Application.DoEvents();

                                            Thread.Sleep(m_iTempoOrigine);

                                            start = false;
                                            timer1.Enabled = true;

                                            button4.Invoke(new EventHandler(delegate { button4.Enabled = true; }));
                                        }
                                        catch(Exception ex)
                                        {
                                            lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, "Erreur : " + ex.Message); }));

                                            bool ok = false;

                                            do
                                            {
                                                if (MessageBox.Show("Impossible de mettre à jour le fichier de production, veuillez fermer le fichier ou redémarrer l'application.", "Information", MessageBoxButtons.RetryCancel) == DialogResult.Retry)
                                                {
                                                    try
                                                    {
                                                        File.WriteAllText(pathClient, toReplace);

                                                        lblQte.Text = obj.restant.ToString();

                                                        ok = true;
                                                    }
                                                    catch
                                                    {
                                                        ok = false;
                                                    }
                                                }
                                            }
                                            while (ok == false);
                                        }
                                    }
                                }
                                else
                                {
                                    boutonOperateur = false;
                                    finProd = true;
                                    start = false;
                                    demarrage = false;

                                    break;
                                }
                            }
                            else
                            {                                
                                MessageBox.Show("Production interrompue.", "Information", MessageBoxButtons.OK);
                                RazProd();

                                RazInfos();

                                LoadRef();

                                comboBox1.Invoke(new EventHandler(delegate { comboBox1.Refresh(); }));

                                Application.DoEvents();
                                break;
                            }
                        }

                        if (qte == 0)
                        {
                            demarrage = false;

                            comboBox1.Invoke(new EventHandler(delegate { comboBox1.SelectedIndex = 0; }));

                            Thread.Sleep(1000);

                            Application.DoEvents();

                            lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = "Production terminée"; }));

                            if (MessageBox.Show("Production terminée.", "Information", MessageBoxButtons.OK) == DialogResult.OK)
                            {
                                production.Remove(obj);

                                string pathClient = ut_xml.ValueXML(@".\CUT-M.xml", "DossierClient");
                                string fileClient = ut_xml.ValueXML(@".\CUT-M.xml", "FichierClient");

                                String toReplace = string.Empty;

                                if (production.Count() > 0)
                                {
                                    foreach (Production p in production)
                                    {
                                        toReplace += p.reference + ";" + p.restant + Environment.NewLine;
                                    }

                                    File.WriteAllText(Path.Combine(pathClient, fileClient), toReplace);
                                }
                                else
                                {
                                    Thread.Sleep(1000);

                                    GC.Collect();
                                    GC.WaitForPendingFinalizers();
                                    FileInfo f = new FileInfo(Path.Combine(pathClient, fileClient));
                                    f.Delete();
                                }

                                RazProd();

                                RazInfos();

                                LoadRef();

                                comboBox1.Invoke(new EventHandler(delegate { comboBox1.Refresh(); }));

                                Application.DoEvents();
                            }
                        }*/
                }
                    else
                        MessageBox.Show("Impossible de démarrer la production, veuillez consulter le(s) message(s) d'erreur ci-dessous", "Information");
                
            }
            else
            {
                btAnnule.Invoke(new EventHandler(delegate { btAnnule.Enabled = false; }));

                lblDiametre.Text = "";
                lblEtage.Text = "";
                lblQte.Text = "";
                lblInfo.Text = ChoixRef;
            }

        }

        private void ChangeOID1(int i_iCh, int etat)
        {
            int iOnOff, iStart = 17 + i_iCh - m_iDiTotal1;

            timer1.Enabled = false;

            iOnOff = etat;

            if (adamModbus1.Modbus().ForceSingleCoil(iStart, iOnOff))
                RefreshDIO1();
            else
            {
                ListViewItem lv = new ListViewItem(String.Format("Mise à jour impossible : {0} {1}", i_iCh.ToString(), etat.ToString()));
                lv.ForeColor = Color.Red;
                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));
            }

            timer1.Enabled = true;
        }

        private void ChangeOID2(int i_iCh, int etat)
        {
            int iOnOff, iStart = 17 + i_iCh - m_iDiTotal2;

            timer2.Enabled = false;

            iOnOff = etat;

            if (adamModbus2.Modbus().ForceSingleCoil(iStart, iOnOff))
                RefreshDIO2();
            else
            {
                ListViewItem lv = new ListViewItem(String.Format("Mise à jour impossible : {0} {1}", i_iCh.ToString(), etat.ToString()));
                lv.ForeColor = Color.Red;
                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));
            }

            timer2.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Etes vous sûr ?", "Information", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                demarrage = false;

                Application.DoEvents();

                RazProd();

                RazInfos();

                LoadRef();

                lblCapot.Invoke(new EventHandler(delegate { lblCapot.Visible = false; }));

                comboBox1.Invoke(new EventHandler(delegate { comboBox1.Refresh(); }));

                Application.DoEvents();

                button4.Invoke(new EventHandler(delegate { button4.Enabled = true; }));
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                timer1.Enabled = false;

                RefreshDIO1();

                String Message = string.Empty;

                if (!porte)
                {
                    porte = false;
                }
                else
                    porte = true;

                label37.Invoke(new EventHandler(delegate { label37.Text = porte.ToString(); }));

                if (!laser)
                {
                    laser = false;
                }
                else
                    laser = true;

                label38.Invoke(new EventHandler(delegate { label38.Text = laser.ToString(); }));
                label44.Invoke(new EventHandler(delegate { label44.Text = capot.ToString(); }));
                label45.Invoke(new EventHandler(delegate { label45.Text = decoupeencours.ToString(); }));

                if (!start && !demarrage && !finProd && comboBox1.SelectedIndex == 0 && laser)
                {
                    Message = "Choisir ou saisir un référence";
                    goodConditions = true;
                }
                /*else if (!start && !demarrage && !finProd && comboBox1.SelectedIndex > 0)
                {
                    Message = "En attente chargement capot et départ cycle";
                    goodConditions = true;
                }
                else if (!start && demarrage && !finProd && !decoupeencours && !capot && !boutonOperateur)
                {
                    Message = "En attente chargement capot et départ cycle";
                    goodConditions = true;
                }*/
                else if (demarrage && !finProd && !decoupeencours && capot)
                {
                    Message = "Mise en référence capot";
                    goodConditions = true;
                }
                else if (demarrage && !finProd && decoupeencours && !capot)
                {
                    Message = "Découpe en cours";
                    goodConditions = true;
                }
                else if (laser && porte && !start)
                {
                    Message = "Découpe terminée";
                    goodConditions = true;
                }
                else if (!laser || !porte)
                {
                    goodConditions = false;
                }

                label41.Invoke(new EventHandler(delegate { label41.Text = goodConditions.ToString(); }));

                if (!string.IsNullOrEmpty(Message))
                    lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = Message; }));

                timer1.Enabled = true;
            }
            catch(Exception ex)
            {
                log.Error(ex);
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            timer2.Enabled = false;

            RefreshDIO2();

            timer2.Enabled = true;
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            long diff = DateTime.Now.Ticks - timestart;

            TimeSpan ts = TimeSpan.FromTicks(diff);
            double secondesFromTs = ts.TotalSeconds;

            if (timer3.Enabled && (secondesFromTs > dureeWatchdog))
            {
                finProd = true;
                start = false;
                demarrage = false;
                boutonOperateur = false;

                checkBox3.Checked = false;
                checkBox4.Checked = false;
                timer3.Stop();

                RazProd();
                MessageBox.Show("Temps de réponse dépassé, production interrompue", "Information");

                lblCapot.Invoke(new EventHandler(delegate { lblCapot.Visible = false; }));

                timestart = DateTime.Now.Ticks;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (!checkBox1.Checked)
            {
                porte = false;
                lblPorte.Invoke(new EventHandler(delegate { lblPorte.ForeColor = Color.Red; lblPorte.Text = "Ouverte"; }));
                ListViewItem lv = new ListViewItem("Porte ouverte");
                lv.ForeColor = Color.Red;
                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));
            }
            else
            {
                porte = true;
                lblPorte.Invoke(new EventHandler(delegate { lblPorte.ForeColor = Color.Green; lblPorte.Text = "Fermée"; }));
                ListViewItem lv = new ListViewItem("Porte fermée");
                lv.ForeColor = Color.Green;
                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));
            }

            label37.Invoke(new EventHandler(delegate { label37.Text = porte.ToString(); }));
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (!checkBox2.Checked)
            {
                laser = false;
                lblLaser.Invoke(new EventHandler(delegate { lblLaser.ForeColor = Color.Red; lblLaser.Text = "Shutter fermé"; }));
                ListViewItem lv = new ListViewItem("Shutter fermé");
                lv.ForeColor = Color.Red;
                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));
            }
            else
            {
                laser = true;
                lblLaser.Invoke(new EventHandler(delegate { lblLaser.ForeColor = Color.Green; lblLaser.Text = "Shutter ouvert"; }));
                ListViewItem lv = new ListViewItem("Shutter ouvert");
                lv.ForeColor = Color.Green;
                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));
            }

            label38.Invoke(new EventHandler(delegate { label38.Text = laser.ToString(); }));
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (!checkBox3.Checked)
            {
                boutonOperateur = false;
            }
            else
            {
                boutonOperateur = true;

                ListViewItem lv = new ListViewItem("Bouton opérateur activé");
                lv.ForeColor = Color.Green;
                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (!checkBox4.Checked)
            {
                finProd = false;
            }
            else
            {
                finProd = true;

                ListViewItem lv = new ListViewItem("Fin de production");
                lv.ForeColor = Color.Green;
                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));
            }

            label42.Invoke(new EventHandler(delegate { label42.Text = finProd.ToString(); }));
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Fermer le programme ?", "Information", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                ChangeOID1(8, 0);
                ChangeOID1(9, 0);
                ChangeOID1(10, 0);
                ChangeOID1(11, 0);
                ChangeOID1(12, 0);
                ChangeOID1(13, 0);
                ChangeOID1(14, 0);
                ChangeOID1(15, 0);

                this.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(LoadRef())
            {
                comboBox1.Invoke(new EventHandler(delegate { comboBox1.Refresh(); }));
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.TopMost = true;
            this.WindowState = FormWindowState.Normal;
            this.StartPosition = FormStartPosition.CenterScreen;
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked)
            {
                forcePorte = true;
            }
            else
                forcePorte = false;
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if(checkBox6.Checked)
            {
                forceDepart = true;
            }
            else
            {
                forceDepart = false;
            }
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked)
            {
                boutonOperateur = true;
            }
            else
            {
                boutonOperateur = false;
            }
        }

        private void RazProd()
        {
            start = false;
            comboBox1.Invoke(new EventHandler(delegate { comboBox1.Enabled = true; }));
            ListViewItem lv = new ListViewItem("Production terminée");
            lv.ForeColor = Color.Green;
            lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));
            timer3.Stop();
            timer1.Stop();

            start = false;
            demarrage = false;
            boutonOperateur = false;

            checkBox3.Checked = false;
            checkBox4.Checked = false;
            finProd = true;

            ChangeOID1(8, 0);
            ChangeOID1(9, 0);
            ChangeOID1(10, 0);
            ChangeOID1(11, 0);
            ChangeOID1(12, 0);
            ChangeOID1(13, 0);
            ChangeOID1(14, 0);
            ChangeOID1(15, 0);
        }

        private void RazInfos()
        {
            comboBox1.Invoke(new EventHandler(delegate { comboBox1.SelectedIndex = 0; }));
        }
    }
}
