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
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.UI.WebControls;
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
        private static bool cycleTermine = false;
        private static bool capot = false;
        private static bool pedale = false;
        private static bool ouverturePorte = false;
        private static bool finProd = false;
        private static bool porte = false;
        private static bool origine = false;
        private static bool mandrin = false;
        private static bool laser = false;
        private static bool start = false;
        private static bool demarrage = false;
        private static bool launch = true;
        private static long timestart = 0;
        private static long dureeWatchdog = 0;
        private static string locale = string.Empty;
        private static string licence = string.Empty;
        private static Process p;

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
        private static string Ouverte = string.Empty;
        private static string ChoisirRef = string.Empty;
        private static string OuvrirPorte = string.Empty;
        private static string PlacerPiece = string.Empty;
        private static string DepartCycle = string.Empty;
        private static string LoadMask = string.Empty;
        private static string Decoupe = string.Empty;
        private static string DecoupeTermine = string.Empty;
        private static string Confirme = string.Empty;
        private static string ShutterOuvert = string.Empty;
        private static string ShutterFerme = string.Empty;
        private static string BoutonOperateur = string.Empty;
        private static string FinProd = string.Empty;
        private static string InterProd = string.Empty;
        private static string TempsReponse = string.Empty;
        private static string FermePrg = string.Empty;
        private static string ErreurFichier = string.Empty;
        private static string NoProd = string.Empty;

        //Variables debug
        /*private static bool forcePorte = true;
        private static bool forceDepart = false;
        private static bool forceShutter = true;
        private static bool forceGoodconditions = false;
        private static bool forceFinprod = false;*/

        private static List<Produit> produits = new List<Produit>();
        private static List<Production> production = new List<Production>();
        public Form1()
        {      
            InitializeComponent();

            InitLang();

            CultureInfo ci = new CultureInfo(locale); Thread.CurrentThread.CurrentCulture = ci; Thread.CurrentThread.CurrentUICulture = ci;

            Application.DoEvents();

            timer1.Interval = 150;
            timer1.Tick += new System.EventHandler(this.timer1_Tick);

            timer2.Interval = 150;
            timer2.Tick += new System.EventHandler(this.timer2_Tick);

            Warning warn = new Warning();
            warn.TopMost = true;
            warn.FormBorderStyle = FormBorderStyle.None;
            warn.WindowState = FormWindowState.Maximized;

            if (warn.ShowDialog() == DialogResult.OK)
            {
                if (DateTime.Today < new DateTime(2024, 1, 1) || licence == "e4067309-5107-43bf-a127-e29ee91ee96e")
                {
                    log4net.Config.XmlConfigurator.Configure();

                    log.Info("Starting Cut-M");

                    try
                    {
                        InitCutM();

                        LoadRef();

                        Application.DoEvents();

                        InitLaser();

                    }
                    catch (Exception e)
                    {
                        log.Error(e);
                    }
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Invalid license", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
                this.Close();
        }
        private void InitLang()
        {
            locale = ut_xml.ValueXML(@".\CUT-M.xml", "Locale");
            log.Info(string.Format("Locale : {0}", locale));
            licence = ut_xml.ValueXML(@".\CUT-M.xml", "Licence");
            log.Info(string.Format("Licence : {0}", licence));

            lblRef.Text = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "lblRef");
            lblProd.Text = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "lblProd");
            btAnnule.Text = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "btAnnule");
            lblTitreDiametre.Text = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "lblTitreDiametre");
            lblTitreQte.Text = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "lblTitreQte");
            lblTitreEtat.Text = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "lblTitreEtat");
            lbltitrePorte.Text = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "lbltitrePorte");
            lbltitreLaser.Text = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "lbltitreLaser");

            ChoixRef = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "ChoixRef");
            lblInfo.Text = ChoixRef;

            ConnectAdam1 = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "ConnectAdam1");
            ConnectAdam2 = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "ConnectAdam2");
            InitAdam1 = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "InitAdam1");
            InitAdam2 = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "InitAdam2");
            Timer1Start = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "Timer1Start");
            NoConnectAdam1 = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "NoConnectAdam1");
            Timer2Start = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "Timer2Start");
            NoConnectAdam2 = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "NoConnectAdam2");
            CloseAndStart = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "CloseAndStart");
            Adam1NotConnected = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "Adam1NotConnected");
            NoRef = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "NoRef");
            Quitter = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "Quitter");
            FermerPorte = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "FermerPorte");
            OuvrirShutter = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "OuvrirShutter");
            Fermee = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "Fermee");
            Ouverte = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "Ouverte");
            ChoisirRef = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "ChoisirRef");
            OuvrirPorte = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "OuvrirPorte");
            PlacerPiece = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "PlacerPiece");
            DepartCycle = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "DepartCycle");
            LoadMask = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "LoadMask");
            Decoupe = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "Decoupe");
            DemarrageDecoupe = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "DemarrageDecoupe");
            DecoupeTermine = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "DecoupeTermine");
            Confirme = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "Confirme");
            ShutterOuvert = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "ShutterOuvert");
            ShutterFerme = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "ShutterFerme");
            BoutonOperateur = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "BoutonOperateur");
            FinProd = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "FinProd");
            InterProd = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "InterProd");
            TempsReponse = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "TempsReponse");
            FermePrg = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "FermePrg");
            ErreurFichier = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "ErreurFichier");
            NoProd = ut_xml.ValueXML(@"C:\CUT-M\Trad\Lang_" + locale + ".xml", "NoProd");

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

                log.Info(InitAdam1);

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

                log.Error(NoConnectAdam1);

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

                log.Info(InitAdam2);

                timer2.Start();

                log.Info(Timer2Start);

                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, Timer2Start); }));

                ChangeOID2(10, 1);

                RefreshDIO2();

            }
            else
            {
                log.Error(NoConnectAdam2);

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

                if ((!porte/* && !forcePorte*/) || (!laser/* && !forceShutter*/))
                {
                    while (!porte || !laser)
                    {
                        Thread.Sleep(1000);

                        Application.DoEvents();
                        this.TopMost = true;
                        MessageBox.Show(new Form { TopMost = true }, CloseAndStart, "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = CloseAndStart; }));

                        Application.DoEvents();
                    }
                }

                ChangeOID1(15, 1);

                Thread.Sleep(m_iTempoImpulsion);

                ChangeOID1(15, 0);

                log.Info("Launching CutMaker");

                Process[] processAlreadyRunning = Process.GetProcesses();
                foreach (Process pr in processAlreadyRunning)
                {
                    if (pr.ProcessName.Contains("CutMaker"))
                    {
                        log.Info("Closing CutMaker.exe");
                        pr.Kill();
                        pr.WaitForExit();
                        pr.Dispose();
                    }
                }

                Thread.Sleep(2000);

                String exePath = ut_xml.ValueXML(@".\CUT-M.xml", "LaserExe");
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = exePath;
                startInfo.UseShellExecute = true;
                startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                //startInfo.Arguments = "header.h";
                p = Process.Start(startInfo);

                log.Info("Launching CutMaker OK");

                ChangeOID1(15, 1);

                Thread.Sleep(500);

                while (!origine)
                {
                    Thread.Sleep(1000);
                    Application.DoEvents();
                }

                origine = false;

                log.Info("Origin OK");

                Thread.Sleep(1000);

                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, "Angle moteur"); }));

                log.Info("Motor angle start");

                Application.DoEvents();

                ChangeOID1(9, 1);
                ChangeOID1(10, 0);
                ChangeOID1(11, 0);
                ChangeOID1(12, 0);
                ChangeOID1(13, 0);
                ChangeOID1(14, 0);

                Thread.Sleep(500);

                Application.DoEvents();

                ChangeOID1(8, 1);
                Thread.Sleep(m_iTempoImpulsion);
                ChangeOID1(8, 0);

                Thread.Sleep(m_iTempo);

                ChangeOID1(15, 0);

                Application.DoEvents();

                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, "Motor angle end"); }));                

                launch = false;

                this.WindowState = FormWindowState.Maximized;
                this.TopMost = true;                

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

                    if (!bData1[0]/* && !forcePorte*/)
                    {
                        Message += FermerPorte + Environment.NewLine;
                        porte = false;
                    }
                    else
                        porte = true;                    

                    if (!bData1[1]/* && !forceShutter*/)
                    {
                        Message += OuvrirShutter + Environment.NewLine;
                        laser = false;
                        pictureBox2.Visible = false;
                    }
                    else
                    {
                        pictureBox2.Visible = true;
                        laser = true;
                    }
                    
                    if (bData1[2] /*|| forceDepart*/)
                    {
                        if (laser && porte)
                        {                     
                            log.Info("Operator signal received");

                            boutonOperateur = true;
                            checkBox6.Checked = false;
                            //forceDepart = false;

                            lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = DemarrageDecoupe; }));
                        }
                        else
                        {
                            lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = CloseAndStart; }));
                        }
                    }
                    if (bData1[3])
                    {
                        if (!decoupeencours && !launch)
                        {
                            ouverturePorte = true;

                            ChangeOID2(10, 0);

                            log.Info("Door open signal");

                            Application.DoEvents();
                        }
                    }
                    else
                    {
                        ouverturePorte = false;

                        ChangeOID2(10, 1);
                    }

                    if (bData1[4])
                    {
                        pedale = true;
                    }
                    else
                        pedale = false;

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

                    if (porte/* || forcePorte*/)
                    {
                        lblPorte.Invoke(new EventHandler(delegate { lblPorte.ForeColor = Color.Green; lblPorte.Text = Fermee; }));
                    }
                    else
                    {
                        lblPorte.Invoke(new EventHandler(delegate { lblPorte.ForeColor = Color.Red; lblPorte.Text = Ouverte; }));
                    }

                    if (laser/* || forceShutter*/)
                    {
                        lblLaser.Invoke(new EventHandler(delegate { lblLaser.ForeColor = Color.Green; lblLaser.Text = ShutterOuvert; }));
                    }
                    else
                    {
                        lblLaser.Invoke(new EventHandler(delegate { lblLaser.ForeColor = Color.Red; lblLaser.Text = ShutterFerme; }));
                    }

                    if (porte && laser)
                    {
                        goodConditions = true;
                    }
                    else if ((!porte || !laser) && decoupeencours)
                    {
                        goodConditions = false;
                        MessageBox.Show(new Form { TopMost = true }, InterProd, "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                        RazProd();
                    }

                    if ((!goodConditions /*|| !forceGoodconditions*/) && start)
                    {
                        start = false;
                        finProd = true;
                        RazProd();
                    }
                }
                else
                {
                    if ((!goodConditions/* || !forceGoodconditions*/) && start)
                    {
                        if (MessageBox.Show("Erreur de communication") == DialogResult.OK)
                        {
                            demarrage = false;
                            finProd = true;

                            Application.DoEvents();

                            RazProd();

                            RazInfos();

                            LoadRef();

                            comboBox1.Invoke(new EventHandler(delegate { comboBox1.Refresh(); }));

                            Application.DoEvents();

                            button4.Invoke(new EventHandler(delegate { button4.Enabled = true; }));

                            timer1.Stop();
                            timer2.Stop();
                            timer3.Stop();
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                if ((!goodConditions /*|| !forceGoodconditions*/) && start)
                {
                    if (MessageBox.Show("Erreur de communication") == DialogResult.OK)
                    {
                        demarrage = false;
                        finProd = true;

                        Application.DoEvents();

                        RazProd();

                        RazInfos();

                        LoadRef();

                        comboBox1.Invoke(new EventHandler(delegate { comboBox1.Refresh(); }));

                        Application.DoEvents();

                        button4.Invoke(new EventHandler(delegate { button4.Enabled = true; }));

                        timer1.Stop();
                        timer2.Stop();
                        timer3.Stop();
                    }
                }
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

                if (!System.Convert.ToBoolean(bData[9]))
                    mandrin = false;
                else
                    mandrin = true;

                    if (laser && porte)
                {
                    if (!System.Convert.ToBoolean(bData[8]))
                    {
                        ChangeOID2(8, 1);
                    }
                }
                else
                {
                    if (System.Convert.ToBoolean(bData[8]))
                    {
                        ChangeOID2(8, 0);
                    }
                }                
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
                                produits.Add(new Produit() { reference = fields[0], diametre = System.Convert.ToInt32(fields[1]), positionangle = fields[2], masque = fields[3] });
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
                demarrage = true;

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

                    Production prod = production.FirstOrDefault(m => m.reference == _ref);

                    log.Info(string.Format("Diameter : {0}", produit.diametre.ToString()));
                    log.Info(string.Format("Floor : {0}", produit.etage.ToString()));
                    log.Info(string.Format("Quantity : {0}", prod.restant.ToString()));

                    lblDiametre.Text = produit.diametre.ToString();
                    lblQte.Text = prod.restant.ToString();

                    try
                    {
                        log.Info(string.Format("Mask : {0}", produit.masque.ToString()));

                        log.Info("Loading mask");

                        Application.DoEvents();

                        ChangeOID1(15, 1);

                        int DO0_1 = int.Parse(produit.masque[0].ToString());
                        int DO0_2 = int.Parse(produit.masque[1].ToString());
                        int DO0_3 = int.Parse(produit.masque[2].ToString());
                        int DO0_4 = int.Parse(produit.masque[3].ToString());
                        int DO0_5 = int.Parse(produit.masque[4].ToString());
                        int DO0_6 = int.Parse(produit.masque[5].ToString());

                        log.Info(String.Format("DO0_1 : {0}", DO0_1.ToString()));
                        log.Info(String.Format("DO0_2 : {0}", DO0_2.ToString()));
                        log.Info(String.Format("DO0_3 : {0}", DO0_3.ToString()));
                        log.Info(String.Format("DO0_4 : {0}", DO0_4.ToString()));
                        log.Info(String.Format("DO0_5 : {0}", DO0_5.ToString()));
                        log.Info(String.Format("DO0_6 : {0}", DO0_6.ToString()));

                        Application.DoEvents();

                        ChangeOID2(1, DO0_6);
                        ChangeOID2(2, DO0_5);
                        ChangeOID2(3, DO0_4);
                        ChangeOID2(4, DO0_3);
                        ChangeOID2(5, DO0_2);
                        ChangeOID2(6, DO0_1);

                        Thread.Sleep(500);
                        ChangeOID2(0, 1);
                        Thread.Sleep(m_iTempoImpulsion);
                        ChangeOID2(0, 0);

                        log.Info("Loading mask OK");

                        ChangeOID1(15, 0);

                        Application.DoEvents();
                    }
                    catch
                    {
                        log.Info("Unable to load mask");
                    }

                    var obj = production.FirstOrDefault(x => x.reference == comboBox1.Text);
                    int qte = obj.restant;


                    comboBox1.Invoke(new EventHandler(delegate { comboBox1.Enabled = false; }));
                    label39.Invoke(new EventHandler(delegate { label39.Text = start.ToString(); }));

                    Application.DoEvents();

                    while (qte > 0)
                    {                 
                        btAnnule.Enabled = true;

                        decoupeencours = false;

                        if (porte/* || forcePorte*/)
                        {
                            lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = OuvrirPorte; }));

                            Application.DoEvents();

                            while (porte && demarrage/* || forcePorte*/)
                            {
                                Thread.Sleep(100);

                                Application.DoEvents();
                            }
                        }
                        lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = PlacerPiece; }));

                        Application.DoEvents();

                        while (!porte && demarrage/* && !forcePorte*/)
                        {
                            Thread.Sleep(100);

                            Application.DoEvents();
                        }

                        ListViewItem lv = new ListViewItem(DepartCycle);
                        lv.ForeColor = Color.Green;
                        lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));

                        log.Info(DepartCycle);

                        lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = DepartCycle; }));

                        Application.DoEvents();

                        while ((!boutonOperateur /*&& !forceDepart*/) && demarrage)
                        {
                            Thread.Sleep(100);

                            Application.DoEvents();
                        }

                        ChangeOID1(15, 1);

                        log.Info("Motor angle start");

                        ChangeOID1(9, 1);
                        ChangeOID1(10, 0);
                        ChangeOID1(11, 0);
                        ChangeOID1(12, 0);
                        ChangeOID1(13, 0);
                        ChangeOID1(14, 0);

                        Thread.Sleep(500);

                        Application.DoEvents();

                        log.Info("Bit validation");

                        ChangeOID1(8, 1);
                        Thread.Sleep(m_iTempoImpulsion);
                        ChangeOID1(8, 0);

                        Thread.Sleep(m_iTempo);

                        log.Info("Motor angle end");

                        lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = DemarrageDecoupe; }));

                        if (goodConditions/* || forceGoodconditions*/)
                        {
                            decoupeencours = true;

                            int Angle_1 = int.Parse(produit.positionangle[0].ToString());
                            int Angle_2 = int.Parse(produit.positionangle[1].ToString());
                            int Angle_3 = int.Parse(produit.positionangle[2].ToString());
                            int Angle_4 = int.Parse(produit.positionangle[3].ToString());
                            int Angle_5 = int.Parse(produit.positionangle[4].ToString());
                            int Angle_6 = int.Parse(produit.positionangle[5].ToString());

                            log.Info(String.Format("Angle_1 : {0}", Angle_1.ToString()));
                            log.Info(String.Format("Angle_2 : {0}", Angle_2.ToString()));
                            log.Info(String.Format("Angle_3 : {0}", Angle_3.ToString()));
                            log.Info(String.Format("Angle_4 : {0}", Angle_4.ToString()));
                            log.Info(String.Format("Angle_5 : {0}", Angle_5.ToString()));
                            log.Info(String.Format("Angle_6 : {0}", Angle_6.ToString()));

                            Application.DoEvents();

                            ChangeOID1(9, Angle_6);
                            ChangeOID1(10, Angle_5);
                            ChangeOID1(11, Angle_4);
                            ChangeOID1(12, Angle_3);
                            ChangeOID1(13, Angle_2);
                            ChangeOID1(14, Angle_1);

                            Thread.Sleep(500);

                            Application.DoEvents();

                            log.Info("Bit validation");

                            ChangeOID1(8, 1);
                            Thread.Sleep(m_iTempoImpulsion);
                            ChangeOID1(8, 0);

                            Thread.Sleep(m_iTempo);

                            ChangeOID2(7, 1);
                            Thread.Sleep(m_iTempoImpulsion);
                            ChangeOID2(7, 0);

                            log.Info("Start cutting");

                            while (!finProd/* && !forceFinprod*/)
                            {
                                Thread.Sleep(100);
                                Application.DoEvents();
                            }

                            log.Info("End signal received");

                            finProd = false;

                            log.Info("Quantity : "+qte);

                            boutonOperateur = false;

                            log.Info("Motor angle start");

                            ChangeOID1(9, 1);
                            ChangeOID1(10, 0);
                            ChangeOID1(11, 0);
                            ChangeOID1(12, 0);
                            ChangeOID1(13, 0);
                            ChangeOID1(14, 0);

                            Thread.Sleep(500);

                            Application.DoEvents();

                            ChangeOID1(8, 1);
                            Thread.Sleep(m_iTempoImpulsion);
                            ChangeOID1(8, 0);

                            Thread.Sleep(m_iTempo);

                            Application.DoEvents();

                            ChangeOID1(15, 0);

                            log.Info("Motor angle end");

                            decoupeencours = false;

                            obj = production.FirstOrDefault(x => x.reference == comboBox1.Text);
                            qte = obj.restant -= 1;
                            if (obj != null) obj.restant = qte;


                            if (qte > 0)
                                decoupeencours = false;

                            String toReplaceFinish = string.Empty;

                            foreach (Production p in production)
                            {
                                toReplaceFinish += p.reference + ";" + p.restant + Environment.NewLine;
                            }

                            string pathClientFinish = ut_xml.ValueXML(@".\CUT-M.xml", "DossierClient");
                            string fileClientFinish = ut_xml.ValueXML(@".\CUT-M.xml", "FichierClient");

                            try
                            {
                                File.WriteAllText(Path.Combine(pathClientFinish, fileClientFinish), toReplaceFinish);

                                log.Info("Write quantity OK");

                                lblQte.Text = obj.restant.ToString();

                            }
                            catch (Exception ex)
                            {
                                log.Error("Write quantity error : " + ex.Message);

                                bool ok = false;

                                do
                                {
                                    if (MessageBox.Show(ErreurFichier, "Information", MessageBoxButtons.RetryCancel) == DialogResult.Retry)
                                    {
                                        try
                                        {
                                            File.WriteAllText(pathClientFinish, fileClientFinish);

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

                    Stopwatch s = new Stopwatch();
                    s.Start();
                    while (s.Elapsed < TimeSpan.FromSeconds(5))
                    {
                        lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = FinProd; }));
                        Application.DoEvents();
                    }

                    s.Stop();

                    log.Info(FinProd);

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

                    demarrage = false;
                    start = false;
                    finProd = true;

                    Application.DoEvents();

                    RazProd();

                    RazInfos();

                    LoadRef();

                    comboBox1.Invoke(new EventHandler(delegate { comboBox1.Refresh(); }));

                    Application.DoEvents();

                    button4.Invoke(new EventHandler(delegate { button4.Enabled = true; }));

                    ChangeOID1(15, 1);

                    log.Info("Motor angle start");

                    ChangeOID1(9, 1);
                    ChangeOID1(10, 0);
                    ChangeOID1(11, 0);
                    ChangeOID1(12, 0);
                    ChangeOID1(13, 0);
                    ChangeOID1(14, 0);

                    Thread.Sleep(1000);

                    Application.DoEvents();

                    ChangeOID1(8, 1);
                    Thread.Sleep(m_iTempoImpulsion);
                    ChangeOID1(8, 0);

                    Thread.Sleep(m_iTempo);

                    Application.DoEvents();

                    log.Info("Motor angle end");

                    ChangeOID1(15, 0);

                    decoupeencours = false;

                    ChangeOID1(8, 0);
                    ChangeOID1(9, 0);
                    ChangeOID1(10, 0);
                    ChangeOID1(11, 0);
                    ChangeOID1(12, 0);
                    ChangeOID1(13, 0);
                    ChangeOID1(14, 0);
                    ChangeOID1(15, 0);

                    timer1.Stop();
                    timer2.Stop();
                    timer3.Stop();
                }
                    else
                        MessageBox.Show(new Form { TopMost = true }, NoProd, "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                
            }
            else
            {
                btAnnule.Invoke(new EventHandler(delegate { btAnnule.Enabled = false; }));

                lblDiametre.Text = "";
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
                log.Error(String.Format("Update error : {0} {1}", i_iCh.ToString(), etat.ToString()));
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
                log.Error(String.Format("Update error : {0} {1}", i_iCh.ToString(), etat.ToString()));
            }

            timer2.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(new Form { TopMost = true }, Confirme, "Information", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                demarrage = false;
                finProd = true;

                Application.DoEvents();

                RazProd();

                RazInfos();

                LoadRef();

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

                try
                {
                    if (pedale && !porte)
                    {
                        if (!mandrin)
                            ChangeOID2(9, 1);
                    }
                    else
                    {
                        if (mandrin)
                            ChangeOID2(9, 0);
                    }
                }
                catch (Exception ex)
                {
                    lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, ex.Message); }));
                }

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

                RefreshDIO1();                

                if (!start && !demarrage && !finProd && comboBox1.SelectedIndex == 0 && laser)
                {
                    Message = ChoixRef;
                    goodConditions = true;
                }
                else if (demarrage && !finProd && decoupeencours && !capot && !cycleTermine)
                {
                    Message = Decoupe;
                    goodConditions = true;
                }
                else if (!laser || !porte)
                {
                    goodConditions = false;
                }       
                /*else if(forcePorte && forceShutter)
                {
                    goodConditions = true;
                }*/

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
                //demarrage = false;
                boutonOperateur = false;

                timer3.Stop();

                RazProd();

                MessageBox.Show(new Form { TopMost = true }, TempsReponse, "Information", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                timestart = DateTime.Now.Ticks;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (!checkBox1.Checked)
            {
                porte = false;
                lblPorte.Invoke(new EventHandler(delegate { lblPorte.ForeColor = Color.Red; lblPorte.Text = Ouverte; }));
            }
            else
            {
                porte = true;
                lblPorte.Invoke(new EventHandler(delegate { lblPorte.ForeColor = Color.Green; lblPorte.Text = Fermee; }));
            }

            label37.Invoke(new EventHandler(delegate { label37.Text = porte.ToString(); }));
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (!checkBox2.Checked)
            {
                laser = false;
                lblLaser.Invoke(new EventHandler(delegate { lblLaser.ForeColor = Color.Red; lblLaser.Text = ShutterFerme; }));
                ListViewItem lv = new ListViewItem(ShutterFerme);
                lv.ForeColor = Color.Red;
                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));
            }
            else
            {
                laser = true;
                lblLaser.Invoke(new EventHandler(delegate { lblLaser.ForeColor = Color.Green; lblLaser.Text = ShutterOuvert; }));
                ListViewItem lv = new ListViewItem(ShutterOuvert);
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

                log.Info(BoutonOperateur);
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

                log.Info(FinProd);
            }

            label42.Invoke(new EventHandler(delegate { label42.Text = finProd.ToString(); }));
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(new Form { TopMost = true }, FermePrg, "Information", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                demarrage = false;
                finProd = true;

                Application.DoEvents();

                RazProd();

                RazInfos();

                LoadRef();

                comboBox1.Invoke(new EventHandler(delegate { comboBox1.Refresh(); }));

                Application.DoEvents();

                button4.Invoke(new EventHandler(delegate { button4.Enabled = true; }));

                ChangeOID1(8, 0);
                ChangeOID1(9, 0);
                ChangeOID1(10, 0);
                ChangeOID1(11, 0);
                ChangeOID1(12, 0);
                ChangeOID1(13, 0);
                ChangeOID1(14, 0);
                ChangeOID1(15, 0);

                timer1.Stop();
                timer2.Stop();
                timer3.Stop();

                this.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.TopMost = true;
            this.WindowState = FormWindowState.Normal;
            this.StartPosition = FormStartPosition.CenterScreen;
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            /*if (checkBox5.Checked)
            {
                forcePorte = true;
            }
            else
                forcePorte = false;*/
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            /*if(checkBox6.Checked)
            {
                forceDepart = true;
            }
            else
            {
                forceDepart = false;
            }*/
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

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            Rectangle rect = Screen.PrimaryScreen.WorkingArea;
            this.Width = rect.Width;
            this.Height = rect.Height;
            this.WindowState = FormWindowState.Maximized;
            this.TopMost = true;

            this.Refresh();
        }

        private void Form1_ResizeEnd(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            this.TopMost = true;
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            /*if (checkBox8.Checked)
            {
                forceGoodconditions = true;
            }
            else
            {
                forceGoodconditions = false;
            }*/
        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            /*if (checkBox8.Checked)
            {
                forceFinprod = true;
            }
            else
            {
                forceFinprod = false;
            }*/
        }

        private void btHide_Click(object sender, EventArgs e)
        {
        }

        private void RazProd()
        {
            log.Info("RazProd");
            start = false;
            comboBox1.Invoke(new EventHandler(delegate { comboBox1.Enabled = true; }));
            log.Info(FinProd);
            timer3.Stop();
            timer1.Stop();

            start = false;
            //demarrage = false;
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
