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
using Microsoft.VisualBasic.FileIO;

namespace CUT_M
{
    public partial class Form1 : Form
    {
        private static List<Produit> produits = new List<Produit>();
        private static List<Production> production = new List<Production>();
        public Form1()
        {
            InitializeComponent();

            LoadRef();
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
            var obj = production.FirstOrDefault(x => x.reference == comboBox1.Text);
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
            }
        }
    }
}
