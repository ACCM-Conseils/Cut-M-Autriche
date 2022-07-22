using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Sockets;
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
        private static bool finProd = true;
        private static bool porte = false;
        private static bool laser = false;
        private static bool start = false;
        private static bool demarrage = false;
        private static long timestart = 0;
        private static long dureeWatchdog = 0;

        private static List<Produit> produits = new List<Produit>();
        private static List<Production> production = new List<Production>();
        public Form1()
        {
            InitializeComponent();

            Warning warn = new Warning();
            warn.FormBorderStyle = FormBorderStyle.None;
            warn.WindowState = FormWindowState.Maximized;

            if (warn.ShowDialog() == DialogResult.OK)
            {
                InitCutM();

                Application.DoEvents();

                LoadRef();

                Application.DoEvents();
            }
            else
                this.Close();
        }

        private void InitCutM()
        {
            if (!this.IsHandleCreated)
            {
                this.CreateHandle();

            }
            String exePath = ut_xml.ValueXML(@".\CUT-M.xml", "LaserExe");
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = exePath;
            //startInfo.Arguments = "header.h";
            Process.Start(startInfo);

            m_bStart = false;			// the action stops at the beginning
            m_szIP1 = ut_xml.ValueXML(@".\CUT-M.xml", "IPAdam1");	// modbus slave IP address for Adam1
            m_iPort1 = System.Convert.ToInt32(ut_xml.ValueXML(@".\CUT-M.xml", "PortAdam1"));				// modbus TCP port for Adam1
            m_szIP2 = ut_xml.ValueXML(@".\CUT-M.xml", "IPAdam2");	// modbus slave IP address for Adam1
            m_iPort2 = System.Convert.ToInt32(ut_xml.ValueXML(@".\CUT-M.xml", "PortAdam2"));				// modbus TCP port for Adam1
            dureeWatchdog = System.Convert.ToInt32(ut_xml.ValueXML(@".\CUT-M.xml", "dureeWatchdog"));
            m_iTempo = System.Convert.ToInt32(ut_xml.ValueXML(@".\CUT-M.xml", "TempoMoteur"));
            m_iTempoOrigine = System.Convert.ToInt32(ut_xml.ValueXML(@".\CUT-M.xml", "TempoOrigine"));
            m_iTempoImpulsion = System.Convert.ToInt32(ut_xml.ValueXML(@".\CUT-M.xml", "TempoImpulsion"));
            adamModbus1 = new AdamSocket();
            adamModbus1.SetTimeout(1000, 1000, 1000);
            adamModbus2 = new AdamSocket();
            adamModbus2.SetTimeout(1000, 1000, 1000);
            label42.Invoke(new EventHandler(delegate { label42.Text = finProd.ToString(); }));
            label39.Invoke(new EventHandler(delegate { label39.Text = start.ToString(); }));
            label40.Invoke(new EventHandler(delegate { label40.Text = boutonOperateur.ToString(); }));

            if (adamModbus1.Connect(m_szIP1, ProtocolType.Tcp, m_iPort1))
            {
                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, "Connecté au module Adam 1"); }));

                lblEtat.Text = "Connecté aux modules ADAM 1";
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

                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, "Initialisation module Adam 1 OK"); }));

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

                this.timer1.Interval = 500;
                this.timer1.Tick += new System.EventHandler(this.timer1_Tick);

                timer1.Start();

                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, "Timer 1 démarré"); }));

                ChangeOID1(12, 0);
                /*ChangeOID1(13, 1);
                Thread.Sleep(m_iTempoImpulsion);*/
                ChangeOID1(13, 0);
                ChangeOID1(14, 0);
                ChangeOID1(15, 0);
                ChangeOID1(16, 0);
                ChangeOID1(17, 0);
                ChangeOID1(18, 0);

                RefreshDIO1();
            }
            else
            {
                lblEtat.Text = "Connexion au module ADAM 1 impossible";
                lblEtat.ForeColor = Color.Red;

                ListViewItem lv = new ListViewItem("Connexion au module ADAM 1 impossible");
                lv.ForeColor = Color.Red;
                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));

                this.timer1.Interval = 500;
                this.timer1.Tick += new System.EventHandler(this.timer1_Tick);

                timer1.Start();
            }

            if (adamModbus2.Connect(m_szIP2, ProtocolType.Tcp, m_iPort2))
            {
                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, "Connecté au module Adam 2"); }));

                lblEtat2.Text = "Connecté aux modules ADAM 2";
                lblEtat2.ForeColor = Color.Green;

                int iDI = 0, iDO = 0;

                InitChannelItems(true, true, ref iDI, ref iDO);
                InitChannelItems(true, true, ref iDI, ref iDO);
                InitChannelItems(true, true, ref iDI, ref iDO);
                InitChannelItems(true, true, ref iDI, ref iDO);
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

                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, "Initialisation module Adam 2 OK"); }));

                m_iDoTotal2 = iDO;
                m_iDiTotal2 = iDI;

                int iCnt;
                bool bCommFSV;
                bool bPtoPFSV;
                bool[] bWDT;

                //if (m_Adam6000Type == Adam6000Type.Adam6055) // no DO for 6055
                if (m_iDoTotal2 == 0)
                {
                    return;
                }

                if (adamModbus2.DigitalOutput().GetWDTMask(out bCommFSV, out bPtoPFSV, out bWDT) && bWDT.Length == 8)
                {
                    iCnt = 0;
                }
                else
                    MessageBox.Show("GetWDTMask() failed;");

                this.timer2.Interval = 500;
                this.timer2.Tick += new System.EventHandler(this.timer2_Tick);
                timer2.Start();

                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, "Timer 2 démarré"); }));

                ChangeOID2(12, 0);
                ChangeOID2(13, 0);
                ChangeOID2(14, 0);
                ChangeOID2(15, 0);
                ChangeOID2(16, 0);
                ChangeOID2(17, 0);
                ChangeOID2(18, 0);

                RefreshDIO2();

            }
            else
            {
                lblEtat2.Text = "Connexion au module ADAM 2 impossible";
                lblEtat2.ForeColor = Color.Red;

                ListViewItem lv = new ListViewItem("Connexion au module ADAM 2 impossible");
                lv.ForeColor = Color.Red;
                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));
            }

            this.timer3.Interval = 500;
            this.timer3.Tick += new System.EventHandler(this.timer3_Tick);
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
                    Message += "Fermer la porte" + Environment.NewLine;
                    porte = false;
                }
                else
                    porte = true;

                label37.Invoke(new EventHandler(delegate { label37.Text = porte.ToString(); }));

                if (!bData1[1])
                {
                    Message += "Ouvrir le shutter" + Environment.NewLine;
                    laser = false;
                }
                else
                    laser = true;
                label38.Invoke(new EventHandler(delegate { label38.Text = laser.ToString(); }));
                if (bData1[2])
                {
                    if (laser && porte)
                    {
                        boutonOperateur = true;
                        start = true;
                    }
                    else
                    {
                        MessageBox.Show("Veuillez fermer la porte et ouvrir le shutter.", "Information", MessageBoxButtons.OK);
                    }
                }
                else
                {
                    boutonOperateur = false;
                    start = false;
                }
                label40.Invoke(new EventHandler(delegate { label40.Text = boutonOperateur.ToString(); }));
                if (bData1[3])
                {
                    finProd = true;

                    lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, "Fin prod reçue"); }));

                    Application.DoEvents();
                }
                else
                {
                    finProd = false;
                }
                label42.Invoke(new EventHandler(delegate { label42.Text = finProd.ToString(); }));
                label39.Invoke(new EventHandler(delegate { label39.Text = start.ToString(); }));                

                txtDO0.Text = bData1[12].ToString();
                txtDO1.Text = bData1[13].ToString();
                txtDO2.Text = bData1[14].ToString();
                txtDO3.Text = bData1[15].ToString();
                txtDO4.Text = bData1[16].ToString();

                txtDI0.Text = bData1[0].ToString();
                txtDI1.Text = bData1[1].ToString();
                txtDI2.Text = bData1[2].ToString();
                txtDI3.Text = bData1[3].ToString();
                txtDI4.Text = bData1[4].ToString();

                if (porte)
                {
                    lblPorte.Invoke(new EventHandler(delegate { lblPorte.ForeColor = Color.Green; lblPorte.Text = "Fermée"; }));
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

                label41.Invoke(new EventHandler(delegate { label41.Text = goodConditions.ToString(); }));

                /*if (goodConditions && bData1[16] && comboBox1.SelectedIndex > 0)
                {
                    ChangeOID1(16, 0);
                }*/
                if (goodConditions && comboBox1.SelectedIndex > 0 && !start && !demarrage)
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
                }

                if (!goodConditions && start)
                {
                    start = false;
                    finProd = true;
                    RazProd();
                }


                lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = Message; }));
            }
        }

        protected void RefreshDIO2()
        {
            int iDiStart = 1, iDoStart = 17;
            int iChTotal;
            bool[] bDiData, bDoData;

            if (adamModbus2.Modbus().ReadCoilStatus(iDiStart, m_iDiTotal2, out bDiData) &&
                adamModbus2.Modbus().ReadCoilStatus(iDoStart, m_iDoTotal2, out bDoData))
            {
                iChTotal = m_iDiTotal2 + m_iDoTotal2;
                bData2 = new bool[iChTotal];
                Array.Copy(bDiData, 0, bData2, 0, m_iDiTotal2);
                Array.Copy(bDoData, 0, bData2, m_iDiTotal2, m_iDoTotal2);

                txt1DO0.Text = bData2[12].ToString();
                txt1DO1.Text = bData2[13].ToString();
                txt1DO2.Text = bData2[14].ToString();
                txt1DO3.Text = bData2[15].ToString();
                txt1DO4.Text = bData2[16].ToString();

                txt1DI0.Text = bData2[0].ToString();
                txt1DI1.Text = bData2[1].ToString();
                txt1DI2.Text = bData2[2].ToString();
                txt1DI3.Text = bData2[3].ToString();
                txt1DI4.Text = bData2[4].ToString();
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
                        row[0] = "Choisissez une référence";
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
                            //Processing row
                            string[] fields = parser.ReadFields();
                            produits.Add(new Produit() { reference = fields[0], diametre = System.Convert.ToInt32(fields[1]), positionangle = fields[2], masque = fields[3], etage = System.Convert.ToInt32(fields[4]) });
                        }
                    }
                    return true;
                }
                else
                {
                    bool ok = false;

                    do
                    {
                        if (MessageBox.Show("Impossible de localiser le fichier Client.csv, veuillez en créer un et recommencer.", "Information", MessageBoxButtons.RetryCancel) == DialogResult.Retry)
                        {
                            if (File.Exists(Path.Combine(pathClient, fileClient)))
                            {
                                using (TextFieldParser parser = new TextFieldParser(Path.Combine(pathClient, fileClient)))
                                {
                                    production = new List<Production>();
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
                                    row[0] = "Choisissez une référence";
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
                                        //Processing row
                                        string[] fields = parser.ReadFields();
                                        produits.Add(new Produit() { reference = fields[0], diametre = System.Convert.ToInt32(fields[1]), positionangle = fields[2], masque = fields[3], etage = System.Convert.ToInt32(fields[4]) });
                                    }
                                }

                                return true;
                            }
                        }
                    }
                    while (ok == false);

                    return false;
                }

            }
            catch (Exception e)
            {
                MessageBox.Show("Impossible de lire le fichier de production");
                return false;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex > 0)
            {
                ChangeOID2(12, 0);
                ChangeOID2(13, 0);
                ChangeOID2(14, 0);
                ChangeOID2(15, 0);
                ChangeOID2(16, 0);
                ChangeOID2(17, 0);

                String _ref = comboBox1.Text;

                Produit produit = produits.FirstOrDefault(m => m.reference == _ref);

                if (produit != null)
                {
                    lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("Selection référénce : {0}", comboBox1.Text)); }));

                    ChangeOID1(16, 0);
                    txtRefManuelle.Enabled = false;
                    button3.Enabled = false;
                    lblInfo.Text = "";

                    Production prod = production.FirstOrDefault(m => m.reference == _ref);

                    lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("Diametre : {0}", produit.diametre.ToString())); }));
                    lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("Etage : {0}", produit.etage.ToString())); }));
                    lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("Qte : {0}", prod.restant.ToString())); }));

                    lblDiametre.Text = produit.diametre.ToString();
                    lblEtage.Text = produit.etage.ToString();
                    lblQte.Text = prod.restant.ToString();

                    if (goodConditions)
                        lblInfo.Text = "Positionner la pièce";
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

                    try
                    {

                        lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("Masque : {0}", produit.masque.ToString())); }));

                        int DO0_1 = int.Parse(produit.masque[0].ToString());
                        int DO0_2 = int.Parse(produit.masque[1].ToString());
                        int DO0_3 = int.Parse(produit.masque[2].ToString());
                        int DO0_4 = int.Parse(produit.masque[3].ToString());
                        int DO0_5 = int.Parse(produit.masque[4].ToString());

                        ChangeOID2(12, DO0_5);
                        ChangeOID2(13, DO0_4);
                        ChangeOID2(14, DO0_3);
                        ChangeOID2(15, DO0_2);
                        ChangeOID2(16, DO0_1);
                        Thread.Sleep(1000);
                        ChangeOID2(17, 1);

                        lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, "Initialisation module Adam 2 OK"); }));
                    }
                    catch
                    {
                        ListViewItem lv = new ListViewItem("Initialisation module Adam 2 impossible");
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
                            ListViewItem lv = new ListViewItem("En attente de départ cycle");
                            lv.ForeColor = Color.Green;
                            lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, lv); }));

                            Application.DoEvents();

                            while (!boutonOperateur && demarrage)
                            {
                                Thread.Sleep(1000);

                                Application.DoEvents();

                                lvOpe.Refresh();
                            }

                            if (demarrage)
                            {
                                Thread.Sleep(1000);

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
                                        label42.Invoke(new EventHandler(delegate { label42.Text = finProd.ToString(); }));

                                        timer3.Stop();

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
                                        catch
                                        {
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

                                foreach (Production p in production)
                                {
                                    toReplace += p.reference + ";" + p.restant + Environment.NewLine;
                                }

                                File.WriteAllText(Path.Combine(pathClient, fileClient), toReplace);

                                RazProd();

                                RazInfos();

                                LoadRef();

                                comboBox1.Invoke(new EventHandler(delegate { comboBox1.Refresh(); }));

                                Application.DoEvents();
                            }
                        }
                    }
                    else
                        MessageBox.Show("Impossible de démarrer la production, veuillez consulter le(s) message(s) d'erreur ci-dessous", "Information");
                }
            }
            else
            {
                button1.Invoke(new EventHandler(delegate { button1.Enabled = false; }));
                txtRefManuelle.Enabled = true;
                button3.Enabled = true;
                lblDiametre.Text = "";
                lblEtage.Text = "";
                lblQte.Text = "";
                lblInfo.Text = "Choisir ou saisir une référence";
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

                comboBox1.Invoke(new EventHandler(delegate { comboBox1.Refresh(); }));

                Application.DoEvents();

                button4.Invoke(new EventHandler(delegate { button4.Enabled = true; }));
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
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

            if (!start && !demarrage && !finProd && comboBox1.SelectedIndex == 0)
            {
                Message = "Choisir ou saisir un référence";
                goodConditions = true;
            }
            else if (!start && !demarrage && !finProd && comboBox1.SelectedIndex > 0)
            {
                Message = "En attente chargement capot et départ cycle";
                goodConditions = true;
            }
            else if (!start && demarrage && !finProd && !decoupeencours && !capot && !boutonOperateur)
            {
                Message = "En attente chargement capot et départ cycle";
                goodConditions = true;
            }
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
            else if (laser && porte && demarrage && finProd && !decoupeencours && !capot)
            {
                Message = "Découpe terminée";
                goodConditions = true;
            }
            else if (!laser || !porte)
            {
                goodConditions = false;
            }

            label41.Invoke(new EventHandler(delegate { label41.Text = goodConditions.ToString(); }));

            lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = Message; }));

            timer1.Enabled = true;
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

            if (start && secondesFromTs > dureeWatchdog)
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
                label40.Invoke(new EventHandler(delegate { label40.Text = boutonOperateur.ToString(); }));
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

        private void button3_Click(object sender, EventArgs e)
        {
            String _ref = txtRefManuelle.Text;

            Produit produit = produits.FirstOrDefault(m => m.reference == _ref);

            if (produit != null)
            {
                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("Selection référénce : {0}", comboBox1.Text)); }));

                ChangeOID1(16, 0);
                txtRefManuelle.Enabled = false;
                lblInfo.Text = "";

                Production prod = production.FirstOrDefault(m => m.reference == _ref);

                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("Diametre : {0}", produit.diametre.ToString())); }));
                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("Etage : {0}", produit.etage.ToString())); }));
                lvOpe.Invoke(new EventHandler(delegate { lvOpe.Items.Insert(0, String.Format("Qte : {0}", prod.restant.ToString())); }));

                lblDiametre.Text = produit.diametre.ToString();
                lblEtage.Text = produit.etage.ToString();
                lblQte.Text = prod.restant.ToString();

                if (goodConditions)
                    lblInfo.Text = "Positionner la pièce";
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
            }
            else
                MessageBox.Show("Référence inconnue");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Fermer le programme ?", "Information", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                ChangeOID1(12, 0);
                ChangeOID1(13, 0);
                ChangeOID1(14, 0);
                ChangeOID1(15, 0);
                ChangeOID1(16, 0);
                ChangeOID1(17, 0);
                ChangeOID1(18, 0);

                ChangeOID2(12, 0);
                ChangeOID2(13, 0);
                ChangeOID2(14, 0);
                ChangeOID2(15, 0);
                ChangeOID2(16, 0);
                ChangeOID2(17, 0);
                ChangeOID2(18, 0);

                this.Close();
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

            ChangeOID1(12, 0);
            ChangeOID1(13, 0);
            ChangeOID1(14, 0);
            ChangeOID1(15, 0);
            ChangeOID1(16, 0);
            ChangeOID1(17, 0);
            ChangeOID1(18, 0);

            ChangeOID2(12, 0);
            ChangeOID2(13, 0);
            ChangeOID2(14, 0);
            ChangeOID2(15, 0);
            ChangeOID2(16, 0);
            ChangeOID2(17, 0);
            ChangeOID2(18, 0);
        }

        private void RazInfos()
        {
            comboBox1.Invoke(new EventHandler(delegate { comboBox1.SelectedIndex = 0; }));
        }
    }
}
