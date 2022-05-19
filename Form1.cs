using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Text;
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
        private string m_szIP1;
        private int m_iPort1;
        private string m_szIP2;
        private int m_iPort2;
        private int m_iDoTotal1, m_iDiTotal1;
        private static bool[] bData1;
        private int m_iDoTotal2, m_iDiTotal2;
        private static bool[] bData2;
        private static bool goodConditions=false;

        private static List<Produit> produits = new List<Produit>();
        private static List<Production> production = new List<Production>();
        public Form1()
        {
            InitializeComponent();

            InitCutM();

            LoadRef();
        }

        private void InitCutM()
        {
            m_bStart = false;			// the action stops at the beginning
            m_szIP1 = ut_xml.ValueXML(@".\CUT-M.xml", "IPAdam1");	// modbus slave IP address for Adam1
            m_iPort1 = System.Convert.ToInt32(ut_xml.ValueXML(@".\CUT-M.xml", "PortAdam1"));				// modbus TCP port for Adam1
            m_szIP2 = ut_xml.ValueXML(@".\CUT-M.xml", "IPAdam2");	// modbus slave IP address for Adam1
            m_iPort2 = System.Convert.ToInt32(ut_xml.ValueXML(@".\CUT-M.xml", "PortAdam2"));				// modbus TCP port for Adam1
            adamModbus1 = new AdamSocket();
            adamModbus1.SetTimeout(1000, 1000, 1000);
            adamModbus2 = new AdamSocket();
            adamModbus2.SetTimeout(1000, 1000, 1000);

            if (adamModbus1.Connect(m_szIP1, ProtocolType.Tcp, m_iPort1))
            {
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
            }
            else
            {
                lblEtat.Text = "Connexion au module ADAM 1 impossible";
                lblEtat.ForeColor = Color.Red;
            }

            if (adamModbus2.Connect(m_szIP2, ProtocolType.Tcp, m_iPort2))
            {
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
            }
            else
            {
                lblEtat2.Text = "Connexion au module ADAM 2 impossible";
                lblEtat2.ForeColor = Color.Red;
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
                    good = false;
                }
                if (!bData1[1])
                {
                    Message += "Activer le laser" + Environment.NewLine;
                    good = false;
                }

                txtDO0.Text = bData1[12].ToString();
                txtDO1.Text = bData1[13].ToString();
                txtDO2.Text = bData1[14].ToString();
                txtDO3.Text = bData1[15].ToString();
                txtDO4.Text = bData1[16].ToString();

                goodConditions = good;

                if(!goodConditions && bData1[16])
                    ChangeOID1(16, 0);

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
            }
        }

        protected bool LoadRef()
        {
            try
            {
                DataTable table1 = new DataTable("references");
                table1.Columns.Add("reference");
                table1.Columns.Add("quantite");

                string pathClient = ut_xml.ValueXML(@".\CUT-M.xml", "FichierClient");

                if (File.Exists(pathClient))
                {
                    using (TextFieldParser parser = new TextFieldParser(pathClient))
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

                    string pathProduit = ut_xml.ValueXML(@".\CUT-M.xml", "FichierRef");

                    using (TextFieldParser parser = new TextFieldParser(pathProduit))
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
                            if (File.Exists(pathClient))
                            {
                                using (TextFieldParser parser = new TextFieldParser(pathClient))
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

                                string pathProduit = ut_xml.ValueXML(@".\CUT-M.xml", "FichierRef");

                                using (TextFieldParser parser = new TextFieldParser(pathProduit))
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
            catch(Exception e)
            {
                MessageBox.Show("Impossible de lire le fichier de production");
                return false;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex > 0)
            {
                timer1.Start();
                timer2.Start();
                ChangeOID1(16, 0);
                txtRefManuelle.Enabled = false;
                txtQteManuelle.Enabled = false;
                lblInfo.Text = "";
                String _ref = comboBox1.Text;

                Produit produit = produits.FirstOrDefault(m => m.reference == _ref);

                if (produit != null)
                {
                    Production prod = production.FirstOrDefault(m => m.reference == _ref);

                    lblDiametre.Text = produit.diametre.ToString();
                    lblEtage.Text = produit.etage.ToString();
                    lblQte.Text = prod.restant.ToString();

                    lblInfo.Text = "Positionner la pièce";

                    ////Vérif DI 0 si niveau bas -> message Fermer la porte

                    ////Vérif DI 1 si niveau bas -> message Activer le laser

                    ////Si DI 0 et DI 1 niveau haut envoi position angulaire passer les 4 bits à DI 0,1,2,3 et mettre DI 4 à niveau haut
                    ///


                }
            }
            else
            {
                txtRefManuelle.Enabled = true;
                txtQteManuelle.Enabled = true;
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
                MessageBox.Show("Set digital output failed!", "Error");

            timer1.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            String _ref = comboBox1.Text;

            Produit produit = produits.FirstOrDefault(m => m.reference == _ref);

            if (produit != null)
            {
                int DO0 = int.Parse(produit.positionangle[0].ToString());
                int DO1 = int.Parse(produit.positionangle[1].ToString());
                int DO2 = int.Parse(produit.positionangle[2].ToString());
                int DO3 = int.Parse(produit.positionangle[3].ToString());
                ChangeOID1(12, DO0);
                ChangeOID1(13, DO1);
                ChangeOID1(14, DO2);
                ChangeOID1(15, DO3);
                ChangeOID1(16, 1);
            }

            if (goodConditions)
            {
            }
            else
                MessageBox.Show("Impossible de démarrer la production, veuillez consulter le(s) message(s) d'erreur ci-dessous", "Information");

                /*var obj = production.FirstOrDefault(x => x.reference == comboBox1.Text);
                int qte = obj.restant-=1;
                if (obj != null) obj.restant = qte;

                String toReplace = string.Empty;

                foreach(Production p in production)
                {
                    toReplace += p.reference + ";" + p.restant + Environment.NewLine;
                }

                string pathClient = ut_xml.ValueXML(@".\CUT-M.xml", "FichierClient");

                try
                {
                    File.WriteAllText(pathClient, toReplace);

                    lblQte.Text = obj.restant.ToString();
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
                }*/
            }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Etes vous sûr ?", "Information", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Enabled = false;

            RefreshDIO1();

            timer1.Enabled = true;
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            timer2.Enabled = false;

            RefreshDIO2();

            timer2.Enabled = true;
        }
    }
}
