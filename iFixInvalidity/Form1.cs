using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TopSolid.Cad.Design.Automating;
using TopSolid.Cad.Drafting.Automating;
using TopSolid.Kernel.Automating;
using TSH = TopSolid.Kernel.Automating.TopSolidHost;

namespace iFixInvalidity
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            ConnectToTopSolid(); // Connexion à TopSolid au lancement de l'application
        }

        
        private void ConnectToTopSolid()
        {
            try
            {
                // Utilisation de la méthode statique pour se connecter à TopSolid
                if (!TopSolidHost.IsConnected)
                {
                    TopSolidHost.Connect("iFixInvalidity");
                    MessageBox.Show("Connexion réussie à TopSolid.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de la connexion à TopSolid : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

      
        // Récupère l'identifiant du document courant.
        private DocumentId DocumentCourant()
        {
            try
            {
                return TopSolidHost.Documents.EditedDocument;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de la récupération du document courant : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return DocumentId.Empty;
            }
        }


        // Récupère le nom du document courant.
        private string NomDocumentCourant(DocumentId documentCourantId)
        {
            try
            {
                return TopSolidHost.Documents.GetName(documentCourantId);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de la récupération du nom du document : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return "Nom inconnu";
            }
        }

        private List<ElementId> EntitéesPubliées(DocumentId documentCourantId)
        {

            try
            {
                List<ElementId> EntitéesPubliéesListe = TSH.Entities.GetPublishings(documentCourantId);
                return EntitéesPubliéesListe;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de la récupération des document publié : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return new List<ElementId>();
            }
        }



        private string TxtParametrePublié (SmartText SmartTxtParametrePublié)
        {
            string TxtParametrePubliéString = SmartTxtParametrePublié.ToString();
            return TxtParametrePubliéString;
        }


        // Gestionnaire de clic pour le bouton.
        
        
        private void button1_Click_1(object sender, EventArgs e)
        {
                DocumentId documentCourantId = new DocumentId();

            try
            {
               documentCourantId = DocumentCourant();
                if (documentCourantId != DocumentId.Empty)
                {
                    string nomDocumentCourant = NomDocumentCourant(documentCourantId);
                    MessageBox.Show($"Document courant : {nomDocumentCourant}", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Aucun document courant trouvé.", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Une erreur s'est produite : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
 
            }


            List<ElementId> EntitéesPubliéesId = EntitéesPubliées(documentCourantId);




        }










    }
}
