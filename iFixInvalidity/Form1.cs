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


namespace iFixInvalidity
{
    public partial class Form1 : Form
    {
        public object TopSolidDesign { get; private set; }

        public Form1()
        {
            InitializeComponent();
            ConnectToTopSolid(); // Connexion à TopSolid au lancement de l'application
            ConnectToTopSolidDesignHost();
            DisplayDocumentName();
            
        }


        private void ConnectToTopSolid()
        {
            try
            {
                // Vérifier si la connexion est déjà établie
                if (!TSH.IsConnected)
                {
                    // Connexion à TopSolid avec un paramètre d'initialisation (si nécessaire)
                    TSH.Connect();

                    // Vérifier à nouveau si la connexion est réussie
                    if (TSH.IsConnected)
                    {
                        LogMessage("Connexion réussie à TopSolid.");
                        //MessageBox.Show("Connexion réussie à TopSolid.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        LogMessage("Connexion échouée à TopSolid.");
                        MessageBox.Show("Connexion échouée à TopSolid.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    LogMessage("TopSolid est déjà connecté.");
                    MessageBox.Show("TopSolid est déjà connecté.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (InvalidOperationException ex)
            {
                // Gérer une exception spécifique si nécessaire
                LogMessage($"Problème opérationnel : {ex.Message}");
                MessageBox.Show($"Problème opérationnel : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                // Gérer d'autres exceptions
                LogMessage($"Erreur lors de la connexion à TopSolid : {ex.Message}");
                MessageBox.Show($"Erreur lors de la connexion à TopSolid : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }




        private void ConnectToTopSolidDesignHost()
        {
            try
            {
                // Vérifier si la connexion est déjà établie
                if (!TopSolidDesignHost.IsConnected)
                {
                    // Connexion à TopSolid avec un paramètre d'initialisation (si nécessaire)
                    TopSolidDesignHost.Connect();

                    // Vérifier à nouveau si la connexion est réussie
                    if (TopSolidDesignHost.IsConnected)
                    {
                        LogMessage("Connexion réussie à TopSolid module design.");
                        //MessageBox.Show("Connexion réussie à TopSolid module design.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        LogMessage("Connexion échouée à TopSolid module design.");
                        MessageBox.Show("Connexion échouée à TopSolid module design.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    LogMessage("TopSolid module design est déjà connecté.");
                    MessageBox.Show("TopSolid module design est déjà connecté.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (InvalidOperationException ex)
            {
                // Gérer une exception spécifique si nécessaire
                LogMessage($"Problème opérationnel : {ex.Message}");
                MessageBox.Show($"Problème opérationnel : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                // Gérer d'autres exceptions
                LogMessage($"Erreur lors de la connexion à TopSolid module design : {ex.Message}");
                MessageBox.Show($"Erreur lors de la connexion à TopSolid module design : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }





        // Récupère l'identifiant du document courant.
        private DocumentId DocumentCourant()
        {
            try
            {
                DocumentId documentId = TopSolidHost.Documents.EditedDocument;
                LogMessage("Document courant récupéré avec succès.");
                return documentId;
            }
            catch (Exception ex)
            {
                LogMessage($"Erreur lors de la récupération du document courant : {ex.Message}");
                MessageBox.Show($"Erreur lors de la récupération du document courant : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return DocumentId.Empty;
            }
        }

        // Récupère le nom du document courant.
        private string NomDocumentCourant(DocumentId documentCourantId)
        {
            try
            {
                string nomDocument = TopSolidHost.Documents.GetName(documentCourantId);
                LogMessage($"Nom du document courant récupéré avec succès : {nomDocument}");
                return nomDocument;
            }
            catch (Exception ex)
            {
                LogMessage($"Erreur lors de la récupération du nom du document : {ex.Message}");
                MessageBox.Show($"Erreur lors de la récupération du nom du document : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return "Nom inconnu";
            }
        }





        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //Fonction recurcive pour remonter au document piece original-------------------------------------------------------------------------------------------------------------------
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        private DocumentId RecupDocuMaster(DocumentId documentCourantId)
        {
            if (documentCourantId != null && documentCourantId != DocumentId.Empty)
            {
                try
                {
                    // Liste les opérations du document
                    List<ElementId> operationsList = OperationsList(documentCourantId);
                    if (operationsList.Count != 0)
                    {
                        // Pour chaque élément de la liste
                        foreach (ElementId operation in operationsList)
                        {
                            // Si c'est une inclusion
                            if (TSHD.Assemblies.IsInclusion(operation))
                            {
                                // Récupération de l'élément de l'inclusion
                                ElementId insertedElementId = TSHD.Assemblies.GetInclusionChildOccurrence(operation);
                                // Récupération de l'instance du document associé
                                DocumentId instanceDocument = TSHD.Assemblies.GetOccurrenceDefinition(insertedElementId);

                                // Récupération du PdmObjectId associé
                                PdmObjectId pdmObjectId = TSH.Documents.GetPdmObject(instanceDocument);

                                // Récupération de l'extension du document
                                TSH.Pdm.GetType(pdmObjectId, out string DocumentExt);

                                // Si c'est un fichier pièce, renvoyer le document ID
                                if (DocumentExt == ".TopPrt")
                                {
                                    DisplayMasterDocumentName(instanceDocument);
                                    LogMessage($"Document pièce trouvé : {instanceDocument}");
                                    if (instanceDocument != DocumentId.Empty)
                                    {
                                        return instanceDocument;
                                    }
                                }
                                // Si c'est un document de prépa, relancer la fonction
                                if (DocumentExt == ".TopNewPrtSet")
                                {
                                    LogMessage($"Document de prépa trouvé : {instanceDocument}");
                                    DocumentId result = RecupDocuMaster(instanceDocument);
                                    if (result != DocumentId.Empty)
                                    {
                                        return result;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        LogMessage("Aucune opération trouvée dans le document courant.");
                    }
                    return DocumentId.Empty;
                }
                catch (Exception ex)
                {
                    LogMessage($"Erreur lors de la récupération du document maître : {ex.Message}");
                    MessageBox.Show("Erreur : " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    // Retourner DocumentId.Empty en cas d'exception
                    return DocumentId.Empty;
                }
            }
            // Retourner DocumentId.Empty si documentCourantId est null ou empty
            LogMessage("Document courant non valide.");
            return DocumentId.Empty;
        }



        // Récupère la liste des opérations d'un document
        private List<ElementId> OperationsList(DocumentId documentCourantId)
        {
            if (documentCourantId != null && documentCourantId != DocumentId.Empty)
            {
                try
                {
                    List<ElementId> operationsList = TSH.Operations.GetOperations(documentCourantId);
                    LogMessage("Liste des opérations récupérée avec succès.");
                    return operationsList;
                }
                catch (Exception ex)
                {
                    LogMessage($"Erreur : lors de la récupération de la liste des opérations : {ex.Message}");
                    MessageBox.Show("Erreur : lors de la récupération de la liste des opérations " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                LogMessage("Document courant non valide.");
            }
            return new List<ElementId>();
        }

        // Déclaration des noms des paramètres texte à rechercher dans les publications
        string Commentaire = "Commentaire"; // Paramètre commentaire
        string Designation = "Designation";
        string Indice_3D = "Indice 3D";

        private void ParametreMaster(in DocumentId docMaster, out ElementId indice3D, out ElementId commentaireOriginal, out ElementId designationOriginal)
        {
            // Initialisation des paramètres out avec des valeurs par défaut
            indice3D = new ElementId();
            commentaireOriginal = new ElementId();
            designationOriginal = new ElementId();

            // Recherche des paramètres publiés dans le document maître
            List<ElementId> ParameterPubliéList = TSH.Entities.GetPublishings(docMaster);

            // Si la liste des paramètres publiés n'est pas vide
            if (ParameterPubliéList.Count > 0)
            {
                try
                {
                    // Pour chaque paramètre publié
                    foreach (ElementId Parameter in ParameterPubliéList)
                    {
                        // Récupération du nom de chaque paramètre publié pour comparaison
                        string ParameterName = TSH.Elements.GetFriendlyName(Parameter);
                        // MessageBox.Show(ParameterName);

                        // Si le nom du paramètre est égal au nom attendu, renvoyer l'Element ID
                        if (ParameterName == Indice_3D)
                        {
                            indice3D = Parameter;
                            LogMessage($"Paramètre '{Indice_3D}' trouvé.");
                        }
                        if (ParameterName == Commentaire)
                        {
                            commentaireOriginal = Parameter;
                            LogMessage($"Paramètre '{Commentaire}' trouvé.");
                        }
                        if (ParameterName == Designation)
                        {
                            designationOriginal = Parameter;
                            LogMessage($"Paramètre '{Designation}' trouvé.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogMessage($"Erreur : lors de la récupération des paramètres maître : {ex.Message}");
                    MessageBox.Show("Erreur : lors de la récupération des paramètres maître " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                LogMessage("Erreur : La liste des paramètres publiés est vide dans le document maître.");
                MessageBox.Show("Erreur : La liste des paramètres publiés est vide dans le document maître");
                indice3D = new ElementId();
                commentaireOriginal = new ElementId();
                designationOriginal = new ElementId();
            }
        }



        //Recuperation des parametres a modifier.
        private void SetSmartTxtParameter(DocumentId documentCourantId, SmartText[] SmartTxtTable)
        {
            if (!TopSolidHost.Application.StartModification("My Action", false)) return;

            bool CommentaireUpdated = false;
            bool DesignationUpdated = false;
            bool Indice_3DUpdated = false;

            try
            {
                TopSolidHost.Documents.EnsureIsDirty(ref documentCourantId);

                // Recherche des paramètres publiés dans le document courant
                List<ElementId> ParameterPubliéList = TSH.Parameters.GetParameters(documentCourantId);

                // Si la liste des paramètres publiés n'est pas vide
                if (ParameterPubliéList.Count > 0)
                {
                    foreach (ElementId ParameterPublié in ParameterPubliéList)
                    {
                        string Parametertype = TSH.Elements.GetTypeFullName(ParameterPublié);
                        // MessageBox.Show(Parametertype);

                        if (Parametertype == "TopSolid.Kernel.DB.Parameters.TextParameterEntity")
                        {
                            string ParameterPubliéName = TSH.Elements.GetFriendlyName(ParameterPublié);

                            if (ParameterPubliéName == Commentaire)
                            {
                                ElementId ParameterPubliéOp = TSH.Elements.GetParent(ParameterPublié);
                                try
                                {
                                    TSH.Parameters.SetSmartTextParameterCreation(ParameterPubliéOp, SmartTxtTable[1]);
                                    CommentaireUpdated = true;
                                    LogMessage($"Paramètre '{Commentaire}' mis à jour.");
                                }
                                catch (Exception ex)
                                {
                                    LogMessage($"Erreur : lors de la mise à jour du paramètre 'Commentaire' : {ex.Message}");
                                    TopSolidHost.Application.EndModification(false, false);
                                    return;
                                }
                            }
                            if (ParameterPubliéName == Designation)
                            {
                                ElementId ParameterPubliéOp = TSH.Elements.GetParent(ParameterPublié);
                                try
                                {
                                    TSH.Parameters.SetSmartTextParameterCreation(ParameterPubliéOp, SmartTxtTable[2]);
                                    DesignationUpdated = true;
                                    LogMessage($"Paramètre '{Designation}' mis à jour.");
                                }
                                catch (Exception ex)
                                {
                                    LogMessage($"Erreur : lors de la mise à jour du paramètre 'Designation' : {ex.Message}");
                                    TopSolidHost.Application.EndModification(false, false);
                                    return;
                                }
                            }
                            if (ParameterPubliéName == Indice_3D)
                            {
                                ElementId ParameterPubliéOp = TSH.Elements.GetParent(ParameterPublié);
                                try
                                {
                                    TSH.Parameters.SetSmartTextParameterCreation(ParameterPubliéOp, SmartTxtTable[3]);
                                    Indice_3DUpdated = true;
                                    LogMessage($"Paramètre '{Indice_3D}' mis à jour.");
                                }
                                catch (Exception ex)
                                {
                                    LogMessage($"Erreur : lors de la mise à jour du paramètre 'Indice_3D' : {ex.Message}");
                                    TopSolidHost.Application.EndModification(false, false);
                                    return;
                                }
                            }
                        }
                    }
                }
                else
                {
                    LogMessage("Erreur : La liste des paramètres courants est vide");
                }

                // Construction du message de confirmation
                string confirmationMessage = $"Paramètres mis à jour :\n" +
                                              $"{Commentaire} : {(CommentaireUpdated ? "Oui" : "Non")}\n" +
                                              $"{Designation} : {(DesignationUpdated ? "Oui" : "Non")}\n" +
                                              $"{Indice_3D} : {(Indice_3DUpdated ? "Oui" : "Non")}";

                LogMessage(confirmationMessage);

                TopSolidHost.Application.EndModification(true, true);
            }
            catch (Exception ex)
            {
                LogMessage($"Erreur : lors de la récupération des paramètres courants : {ex.Message}");
                TopSolidHost.Application.EndModification(false, false);
            }
        }

        private void LogMessage(string message)
        {
            // Ajouter le message au ListBox
            logListBox.Items.Add(message);
            // Faire défiler le ListBox pour afficher le dernier message
            logListBox.TopIndex = logListBox.Items.Count - 1;
        }



        private SmartText CreateSmartTxt(ElementId smartTxt)
        {
            SmartText SmartTxtId = new SmartText(smartTxt);
            LogMessage($"SmartText créé pour l'élément : {smartTxt}");
            return SmartTxtId;
        }


        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // Gestionnaire de clic pour le bouton.-----------------------------------------------------------------------------------------------------------------------------------------
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        private void button1_Click_1(object sender, EventArgs e)
        {
            // Déclaration variable docu courant
            DocumentId documentCourantId = new DocumentId();

            // Récupération docu courant
            try
            {
                documentCourantId = DocumentCourant();
                if (documentCourantId != DocumentId.Empty)
                {
                    string nomDocumentCourant = NomDocumentCourant(documentCourantId);
                    LogMessage($"Document courant : {nomDocumentCourant}");
                    MessageBox.Show($"Document courant : {nomDocumentCourant}", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    LogMessage("Aucun document courant trouvé.");
                    MessageBox.Show("Aucun document courant trouvé.", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Une erreur s'est produite : {ex.Message}");
                MessageBox.Show($"Une erreur s'est produite : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // Fonction recup docu master
            DocumentId docMaster = RecupDocuMaster(documentCourantId);
            if (docMaster != DocumentId.Empty)
            {
                String docMasterName = TSH.Documents.GetName(docMaster);
                LogMessage($"Document maître trouvé : {docMasterName}");
                MessageBox.Show("Document pièce trouvé : " + docMasterName);
            }
            else
            {
                LogMessage("Aucun document maître trouvé.");
                MessageBox.Show("Aucun document maître trouvé.", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            // Récupération des paramètres de base originaux depuis l'opération d'inclusion
            // Déclaration des variables pour recevoir les valeurs de retour
            ElementId indice3D;
            ElementId commentaireOriginal;
            ElementId designationOriginal;

            // Récupération des paramètres maître
            ParametreMaster(in docMaster, out indice3D, out commentaireOriginal, out designationOriginal);

            SmartText SmartTxtCommentaireId = CreateSmartTxt(commentaireOriginal);
            SmartText SmartTxtDesignationId = CreateSmartTxt(designationOriginal);
            SmartText SmartTxtIndice_3DId = CreateSmartTxt(indice3D);

            // Déclarer et initialiser un tableau de SmartText
            SmartText[] SmartTxtTable = new SmartText[4];

            // SmartTxtTable[0] = SmartTxtEmpty;

            SmartTxtTable[1] = SmartTxtCommentaireId; // Index 1
            SmartTxtTable[2] = SmartTxtDesignationId; // Index 2
            SmartTxtTable[3] = SmartTxtIndice_3DId; // Index 3

            SetSmartTxtParameter(documentCourantId, SmartTxtTable);

            //TSH.Disconnect();
            //Environment.Exit(0);
        }




        //Formulaire des options---------------------------------------------------------------------------------------------------
        private void optionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Créez une instance de
            FormulaireConfig form2 = new FormulaireConfig();
            // Affichez le formulaire
            form2.Show();
        }


        //Bouton quiter-----------------------------------------------------------------------------------------------------------
        private void quitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }


        //Bouton build ---------------------------------------------------------------------------------------------------------
        private void button2_Click(object sender, EventArgs e)
        {
            DocumentId documentCourantId = DocumentCourant();

            // Vérification de documentCourantId
            if (documentCourantId == null || documentCourantId == DocumentId.Empty)
            {
                LogMessage("Erreur : Le document courant est invalide ou vide.");
                MessageBox.Show("Erreur : Le document courant est invalide ou vide.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!TopSolidHost.Application.StartModification("My Action", false))
            {
                LogMessage("Erreur : Impossible de démarrer la modification.");
                MessageBox.Show("Erreur : Impossible de démarrer la modification.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                string CommentaireTxt = "Commentaire";
                string DesignationTxt = "Designation";
                string Indice_3DTxt = "Indice 3D";

                bool Indice_3DCreated = false;
                bool DesignationCreated = false;
                bool CommentaireCreated = false;

                ElementId Indice_3DExiste = TSH.Elements.SearchByName(documentCourantId, Indice_3DTxt);
                if (Indice_3DExiste == ElementId.Empty)
                {
                    ElementId Indice_3DParam = TSH.Parameters.CreateSmartTextParameter(documentCourantId, new SmartText(""));
                    TSH.Elements.SetName(Indice_3DParam, Indice_3DTxt);
                    Indice_3DCreated = true;
                    LogMessage($"Paramètre '{Indice_3DTxt}' créé.");
                }
                else
                {
                    LogMessage($"Paramètre '{Indice_3DTxt}' existe déjà.");
                }

                ElementId DesignationExiste = TSH.Elements.SearchByName(documentCourantId, DesignationTxt);
                if (DesignationExiste == ElementId.Empty)
                {
                    ElementId DesignationParam = TSH.Parameters.CreateSmartTextParameter(documentCourantId, new SmartText(""));
                    TSH.Elements.SetName(DesignationParam, DesignationTxt);
                    DesignationCreated = true;
                    LogMessage($"Paramètre '{DesignationTxt}' créé.");
                }
                else
                {
                    LogMessage($"Paramètre '{DesignationTxt}' existe déjà.");
                }

                ElementId CommentaireExiste = TSH.Elements.SearchByName(documentCourantId, CommentaireTxt);
                if (CommentaireExiste == ElementId.Empty)
                {
                    ElementId CommentaireParam = TSH.Parameters.CreateSmartTextParameter(documentCourantId, new SmartText(""));
                    TSH.Elements.SetName(CommentaireParam, CommentaireTxt);
                    CommentaireCreated = true;
                    LogMessage($"Paramètre '{CommentaireTxt}' créé.");
                }
                else
                {
                    LogMessage($"Paramètre '{CommentaireTxt}' existe déjà.");
                }

                // Construction du message de confirmation
                string confirmationMessage = $"Paramètres créés :\n" +
                                              $"{Indice_3DTxt} : {(Indice_3DCreated ? "Oui" : "Non")}\n" +
                                              $"{DesignationTxt} : {(DesignationCreated ? "Oui" : "Non")}\n" +
                                              $"{CommentaireTxt} : {(CommentaireCreated ? "Oui" : "Non")}";

                LogMessage(confirmationMessage);
                MessageBox.Show(confirmationMessage, "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Information);

                TopSolidHost.Application.EndModification(true, true);
            }
            catch (Exception ex)
            {
                LogMessage($"Erreur : {ex.Message}");
                MessageBox.Show("Erreur : " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TopSolidHost.Application.EndModification(false, false);
            }
        }

        private void DisplayDocumentName()
        {
            DocumentId documentCourantId = DocumentCourant();
           string documentCourantName = NomDocumentCourant(documentCourantId);
          
            labelDocumentName.Text = documentCourantName;
        }

        private void DisplayMasterDocumentName(DocumentId documentCourantId)
        {
            
            string documentCourantName = NomDocumentCourant(documentCourantId);

            labelDocumentMasterName.Text = documentCourantName;
        }

        private void buttonRestart_Click(object sender, EventArgs e)
        {
            RestartApplication();
        }

        private void RestartApplication()
        {
            // Obtenir le chemin de l'exécutable actuel
            string applicationPath = Application.ExecutablePath;

            // Démarrer une nouvelle instance de l'application
            Process.Start(applicationPath);

            // Fermer l'application actuelle
            Application.Exit();
        }

        
    }

    internal class Elementid
    {
    }
}
