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
    /// <summary>
    /// Classe permettant de traiter et d'identifier les ensembles d'électrodes
    /// qui référencent le document actif dans TopSolid.
    /// </summary>
    public class Electrode
    {
        /// <summary>
        /// Extension de fichier pour les documents d'ensemble d'électrodes TopSolid.
        /// </summary>
        private const string ELECTRODE_EXTENSION = ".TopEld";

        #region Méthode publique

        /// <summary>
        /// Traite les électrodes en recherchant tous les ensembles d'électrodes (.TopEld)
        /// qui référencent le document actuellement édité dans TopSolid.
        /// Affiche les résultats dans des boîtes de dialogue.
        /// </summary>
        /// <remarks>
        /// Cette méthode parcourt tous les documents du projet actif,
        /// filtre les ensembles d'électrodes et identifie ceux qui contiennent
        /// une référence vers le document courant.
        /// </remarks>
        public SmartText TraiterElectrodes()
        {
            try
            {
                ShowDebugMessage("Début du traitement des électrodes");

                DocumentId currentDocumentId = TSH.Documents.EditedDocument;

                if (!ValidateCurrentDocument(currentDocumentId))
                {
                    return null;
                }

                List<PdmObjectId> projectDocuments = PDM.GetAllProjectDocuments();
                ShowDebugMessage($"Nombre de documents trouvés : {projectDocuments.Count}");

                DocumentId electrodeAssembly = FindElectrodeAssemblies(currentDocumentId, projectDocuments);

                SmartText nomDocu = NomFormeEroder(electrodeAssembly);
                return nomDocu;
                
            }
            catch (Exception ex)
            {
                ShowErrorMessage($"ERREUR : {ex.Message}\n\nStack:\n{ex.StackTrace}");
                return null;
            }
        }

        #endregion

        #region Méthodes de validation

        /// <summary>
        /// Valide que le document actuel n'est pas vide.
        /// </summary>
        /// <param name="documentId">L'identifiant du document à valider.</param>
        /// <returns>True si le document est valide, False sinon.</returns>
        private bool ValidateCurrentDocument(DocumentId documentId)
        {
            if (documentId == DocumentId.Empty)
            {
                MessageBox.Show("Aucun document TopSolid actif détecté !", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
        }

        #endregion

        #region Méthodes de recherche

        /// <summary>
        /// Recherche tous les ensembles d'électrodes (.TopEld) dans le projet
        /// qui référencent le document actuel.
        /// </summary>
        /// <param name="currentDocumentId">L'identifiant du document actuel.</param>
        /// <param name="projectDocuments">La liste de tous les documents PDM du projet.</param>
        /// <returns>Une liste contenant les noms des ensembles d'électrodes trouvés.</returns>
        private DocumentId FindElectrodeAssemblies(DocumentId currentDocumentId, List<PdmObjectId> projectDocuments)
        {
            DocumentId foundAssembly = DocumentId.Empty;
            
            // Parcourir tous les documents du projet
            foreach (var pdmDoc in projectDocuments)
            {
                // Vérifier si c'est un ensemble d'électrodes qui référence le document actuel
                if (IsElectrodeAssembly(pdmDoc) && ReferencesCurrentDocument(pdmDoc, currentDocumentId))
                {
                    string assemblyName = TryGetDocumentName(pdmDoc);
                    if (!string.IsNullOrEmpty(assemblyName))
                    {
                        foundAssembly = TSH.Documents.GetDocument(pdmDoc);
                        break;
                    }
                }
            }

            return foundAssembly;
        }

        /// <summary>
        /// Détermine si un document PDM est un ensemble d'électrodes (.TopEld).
        /// </summary>
        /// <param name="pdmDocument">L'identifiant PDM du document à vérifier.</param>
        /// <returns>True si le document est un ensemble d'électrodes, False sinon.</returns>
        private bool IsElectrodeAssembly(PdmObjectId pdmDocument)
        {
            TSH.Pdm.GetType(pdmDocument, out string extension);
            return extension == ELECTRODE_EXTENSION;
        }

        /// <summary>
        /// Vérifie si un document PDM contient une référence vers le document spécifié.
        /// </summary>
        /// <param name="pdmDocument">L'identifiant PDM du document à vérifier.</param>
        /// <param name="currentDocumentId">L'identifiant du document recherché dans les références.</param>
        /// <returns>True si le document référence le document actuel, False sinon.</returns>
        /// <remarks>
        /// Cette méthode retourne False en cas d'erreur lors de la récupération des références.
        /// </remarks>
        private bool ReferencesCurrentDocument(PdmObjectId pdmDocument, DocumentId currentDocumentId)
        {
            try
            {
                DocumentId assemblyDocId = TSH.Documents.GetDocument(pdmDocument);
                // Récupérer toutes les références du document sans dépendance
                List<DocumentId> referencedDocs = TSH.Documents.GetReferencedDocuments(assemblyDocId, false);

                // Vérifier si le document actuel est dans la liste des références
                return referencedDocs.Contains(currentDocumentId);
            }
            catch (Exception)
            {
                // Ignorer les documents inaccessibles ou avec révision invalide
                return false;
            }
        }

        /// <summary>
        /// Tente de récupérer le nom d'un document PDM.
        /// </summary>
        /// <param name="pdmDocument">L'identifiant PDM du document.</param>
        /// <returns>Le nom du document, ou null en cas d'erreur.</returns>
        private string TryGetDocumentName(PdmObjectId pdmDocument)
        {
            try
            {
                DocumentId docId = TSH.Documents.GetDocument(pdmDocument);
                return TSH.Documents.GetName(docId);
            }
            catch (Exception)
            {
                // Retourner null si le document n'est pas accessible
                return null;
            }
        }

        private SmartText NomFormeEroder(DocumentId document)
        {
            Document ensembleElec = new Document();
            ensembleElec.DocId = document;
            string commentaire = "";
            string indice3D = "";
            string moule = "";
            foreach (ElementId paremetre in ensembleElec.DocParameters)

            { string parameterType = TSH.Elements.GetTypeFullName(paremetre);
                if (parameterType == "TopSolid.Kernel.DB.Parameters.TextParameterEntity")
                {
                    //string parameterFullName = TSH.Parameters.GetNameParameter(paremetre);
                    ElementId commentaireId = TSH.Elements.SearchByName(ensembleElec.DocId,"Commentaire");
                    commentaire = TSH.Parameters.GetTextValue(commentaireId);
                    ElementId Indice3DId = TSH.Elements.SearchByName(ensembleElec.DocId, "Indice 3D");
                    indice3D = TSH.Parameters.GetTextValue(Indice3DId);
                    ElementId MouleId = TSH.Elements.SearchByName(ensembleElec.DocId, "Moule");
                    moule = TSH.Parameters.GetTextValue(MouleId);
                }
            }
                    string nomDocu = $"{commentaire} Ind {indice3D} {moule}";

                    return new SmartText(nomDocu);
           // return null;
        }

        #endregion

        #region Méthodes d'affichage

        /// <summary>
        /// Affiche les résultats de la recherche d'ensembles d'électrodes.
        /// </summary>
        /// <param name="electrodeAssemblies">La liste des noms d'ensembles d'électrodes trouvés.</param>
        /// <remarks>
        /// Affiche une boîte de dialogue pour chaque ensemble trouvé,
        /// suivie d'un message de confirmation de fin de traitement.
        /// </remarks>
        private void DisplayResults(List<string> electrodeAssemblies)
        {
            if (electrodeAssemblies.Count > 0)
            {
                foreach (string assemblyName in electrodeAssemblies)
                {
                    MessageBox.Show($"Document: {assemblyName}", "Résultat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }

            MessageBox.Show("Traitement terminé avec succès.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Affiche un message de débogage dans une boîte de dialogue.
        /// </summary>
        /// <param name="message">Le message à afficher.</param>
        private void ShowDebugMessage(string message)
        {
            MessageBox.Show(message, "Debug", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Affiche un message d'erreur dans une boîte de dialogue.
        /// </summary>
        /// <param name="message">Le message d'erreur à afficher.</param>
        private void ShowErrorMessage(string message)
        {
            MessageBox.Show(message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        #endregion
    }
}
