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
        private AdamSocket adamModbus;
        private Adam6000Type m_Adam6000Type;
        private string m_szIP;
        private int m_iPort;
        private int m_iDoTotal, m_iDiTotal, m_iCount;
        private static bool[] bData;

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
            m_szIP = "192.168.1.250";	// modbus slave IP address
            m_iPort = 502;				// modbus TCP port is 502
            adamModbus = new AdamSocket();
            adamModbus.SetTimeout(1000, 1000, 1000);

            if (adamModbus.Connect(m_szIP, ProtocolType.Tcp, m_iPort))
            {
                lblEtat.Text = "Connecté au module ADAM";
                lblEtat.ForeColor = Color.Green;

                lblEtat.Text = "Connecté au module ADAM";
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

                m_iDoTotal = iDO;
                m_iDiTotal = iDI;

                int iCnt;
                bool bCommFSV;
                bool bPtoPFSV;
                bool[] bWDT;

                //if (m_Adam6000Type == Adam6000Type.Adam6055) // no DO for 6055
                if (m_iDoTotal == 0)
                {
                    return;
                }

                if (adamModbus.DigitalOutput().GetWDTMask(out bCommFSV, out bPtoPFSV, out bWDT) && bWDT.Length == 8)
                {
                    iCnt = 0;
                }
                else
                    MessageBox.Show("GetWDTMask() failed;");

                this.timer1.Interval = 500;
                this.timer1.Tick += new System.EventHandler(this.timer1_Tick);

                timer1.Start();
            }
            else
            {
                lblEtat.Text = "Connexion au module ADAM impossible";
                lblEtat.ForeColor = Color.Red;
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

        protected void RefreshDIO()
        {
            int iDiStart = 1, iDoStart = 17;
            int iChTotal;
            bool[] bDiData, bDoData;

            if (adamModbus.Modbus().ReadCoilStatus(iDiStart, m_iDiTotal, out bDiData) &&
                adamModbus.Modbus().ReadCoilStatus(iDoStart, m_iDoTotal, out bDoData))
            {
                iChTotal = m_iDiTotal + m_iDoTotal;
                bData = new bool[iChTotal];
                Array.Copy(bDiData, 0, bData, 0, m_iDiTotal);
                Array.Copy(bDoData, 0, bData, m_iDiTotal, m_iDoTotal);

                if (bData[2])
                    lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = "Porte fermée"; }));
                else
                    lblInfo.Invoke(new EventHandler(delegate { lblInfo.Text = "Porte ouverte"; }));
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

        private void button2_Click(object sender, EventArgs e)
        {
            if (adamModbus.Connect(m_szIP, ProtocolType.Tcp, m_iPort))
            {               
                RefreshDIO();   
            }

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

            RefreshDIO();

            timer1.Enabled = true;
        }
    }
}
