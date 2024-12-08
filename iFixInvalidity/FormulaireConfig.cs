using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using TopSolid.Cad.Design.Automating;
using TopSolid.Cad.Drafting.Automating;
using TopSolid.Kernel.Automating;
using TSH = TopSolid.Kernel.Automating.TopSolidHost;
using System.IO;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;




namespace iFixInvalidity
{
    public partial class FormulaireConfig : Form
    {
        public FormulaireConfig()
        {
            InitializeComponent();
            ChargerDico();
        }

        private void ChargerDico()
        {
            try
            {
                // Lire le fichier et afficher son contenu dans la RichTextBox
                if (File.Exists("dictionary.txt"))
                {
                    string[] lines = File.ReadAllLines("dictionary.txt");
                    DicoTextBox.Lines = lines;
                }
                else
                {
                    // Créer le fichier s'il n'existe pas
                    using (StreamWriter sw = File.CreateText("dictionary.txt")) 
                    {
                         
                    }
                    MessageBox.Show("Le fichier dictionnaire n'existait pas, il a été créé.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors du chargement du dictionnaire : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void buttonEnregistrer_Click_1(object sender, EventArgs e)
        {
            SauvegardeDico();
        }

        private void SauvegardeDico()
        {

                // Récupérer les valeurs des TextBox
                string nom = NomtextBox.Text;
                string designation = DesignationtextBox.Text;
                string valeur = ValeurtextBox.Text;

                // Appeler la méthode modifDico avec les valeurs récupérées
                modifDico(nom, designation, valeur);
            

            try
            {
                // Écrire le contenu de la RichTextBox dans le fichier texte
                File.WriteAllLines("dictionary.txt", DicoTextBox.Lines);
                MessageBox.Show("Les données ont été enregistrées avec succès !");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de l'enregistrement du dictionnaire : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void modifDico(string nomtextBox, string DesignationtextBox,string ValeurtextBox)
        {

            String entreeDico = NomtextBox + " , " + DesignationtextBox + " , " + ValeurtextBox;
            DicoTextBox.AppendText(entreeDico + Environment.NewLine);
            
            





        }

        private void quitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
