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
using TSHD = TopSolid.Cad.Design.Automating.TopSolidDesignHost;
using System.IO;
using System.Security;
using System.Security.Cryptography;
using System.Diagnostics;
using static System.Net.Mime.MediaTypeNames;
using TopSolid.Cad.Electrode.Automating;
using TSEH = TopSolid.Cad.Electrode.Automating.TopSolidElectrodeHost;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;
using OutilsTs;

namespace iFixInvalidity
{
    public class Electrode
    {
        // Méthode publique à appeler
        public void TraiterElectrodes()
        {
            try
            {
                MessageBox.Show("Début du traitement des électrodes", "Debug", MessageBoxButtons.OK, MessageBoxIcon.Information);

                Document document = new Document();
                document.DocId = TSH.Documents.EditedDocument;

                if (document.DocId == DocumentId.Empty)
                {
                    MessageBox.Show("Aucun document TopSolid actif détecté !", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                PdmObjectId Projet = TSH.Pdm.GetProject(document.DocPdmObject);

                string projectName = TSH.Pdm.GetName(Projet);

                MessageBox.Show($"Projet détecté : {projectName}", "Debug", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (Projet.Equals(PdmObjectId.Empty))
                {
                    MessageBox.Show("Impossible de récupérer le projet PDM.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                TSH.Pdm.GetConstituents(Projet, out var dossiers, out var documents);

                if (documents == null || documents.Count == 0)
                {
                    MessageBox.Show("Aucun document trouvé dans le projet.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                MessageBox.Show($"Nombre de documents trouvés : {documents.Count}", "Debug", MessageBoxButtons.OK, MessageBoxIcon.Information);

                foreach (var doc in documents)
                {
                    var docId = TSH.Documents.GetDocument(doc);
                    string docName = TSH.Documents.GetName(docId);
                    MessageBox.Show($"Document: {docName}");
                }

                MessageBox.Show("Traitement terminé avec succès.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ERREUR : {ex.Message}\n\nStack:\n{ex.StackTrace}", "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
