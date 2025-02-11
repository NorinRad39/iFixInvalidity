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



namespace iFixInvalidity
{

    public partial class Form1 : Form
    {
        public object TopSolidDesign { get; private set; }

            Document currentDoc;
        DocumentId docMaster = DocumentId.Empty;

        public Form1()
        {
            InitializeComponent();
            ConnectToTopSolid(); // Connexion à TopSolid au lancement de l'application
            ConnectToTopSolidDesignHost();
            ConnectToTopSolidElectrodeHost();

            // Initialisation de currentDoc avec l'instance actuelle de Form1
            currentDoc = new Document(this);

            currentDoc.DocId = DocumentCourant();
            DisplayDocumentName(currentDoc.DocNomTxt);
            DisplayMasterDocumentName();
        }

        bool prepaTrouvé = false;

        //Connexion a topsolid
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
                        LogMessage("Connexion réussie à TopSolid.", System.Drawing.Color.Green);
                        // MessageBox.Show("Connexion réussie à TopSolid.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        LogMessage("Connexion échouée à TopSolid.", System.Drawing.Color.Red);
                        MessageBox.Show("Connexion échouée à TopSolid.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    LogMessage("TopSolid est déjà connecté.", System.Drawing.Color.Orange);
                    MessageBox.Show("TopSolid est déjà connecté.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (InvalidOperationException ex)
            {
                // Gérer une exception spécifique si nécessaire
                LogMessage($"Problème opérationnel : {ex.Message}", System.Drawing.Color.Red);
                MessageBox.Show($"Problème opérationnel : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                // Gérer d'autres exceptions
                LogMessage($"Erreur lors de la connexion à TopSolid : {ex.Message}", System.Drawing.Color.Red);
                MessageBox.Show($"Erreur lors de la connexion à TopSolid : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //Connexion a topsolid design
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
                        LogMessage("Connexion réussie à TopSolid module design.", System.Drawing.Color.Green);
                        // MessageBox.Show("Connexion réussie à TopSolid module design.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        LogMessage("Connexion échouée à TopSolid module design.", System.Drawing.Color.Red);
                        MessageBox.Show("Connexion échouée à TopSolid module design.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    LogMessage("TopSolid module design est déjà connecté.", System.Drawing.Color.Orange);
                    MessageBox.Show("TopSolid module design est déjà connecté.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (InvalidOperationException ex)
            {
                // Gérer une exception spécifique si nécessaire
                LogMessage($"Problème opérationnel : {ex.Message}", System.Drawing.Color.Red);
                MessageBox.Show($"Problème opérationnel : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                // Gérer d'autres exceptions
                LogMessage($"Erreur lors de la connexion à TopSolid module design : {ex.Message}", System.Drawing.Color.Red);
                MessageBox.Show($"Erreur lors de la connexion à TopSolid module design : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //Connexion a top electrode
        private void ConnectToTopSolidElectrodeHost()
        {
            try
            {
                // Vérifier si la connexion est déjà établie
                if (!TopSolidElectrodeHost.IsConnected)
                {
                    // Connexion à TopSolid avec un paramètre d'initialisation (si nécessaire)
                    TopSolidElectrodeHost.Connect();

                    // Vérifier à nouveau si la connexion est réussie
                    if (TopSolidElectrodeHost.IsConnected)
                    {
                        LogMessage("Connexion réussie à TopSolid module design.", System.Drawing.Color.Green);
                        // MessageBox.Show("Connexion réussie à TopSolid module design.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        LogMessage("Connexion échouée à TopSolid module design.", System.Drawing.Color.Red);
                        MessageBox.Show("Connexion échouée à TopSolid module design.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    LogMessage("TopSolid module design est déjà connecté.", System.Drawing.Color.Orange);
                    MessageBox.Show("TopSolid module design est déjà connecté.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (InvalidOperationException ex)
            {
                // Gérer une exception spécifique si nécessaire
                LogMessage($"Problème opérationnel : {ex.Message}", System.Drawing.Color.Red);
                MessageBox.Show($"Problème opérationnel : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                // Gérer d'autres exceptions
                LogMessage($"Erreur lors de la connexion à TopSolid module design : {ex.Message}", System.Drawing.Color.Red);
                MessageBox.Show($"Erreur lors de la connexion à TopSolid module design : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Récupère l'identifiant du document courant.
        private DocumentId DocumentCourant()
        {
            try
            {
                // Récupération de l'ID du document courant en cours d'édition
                DocumentId documentId = TopSolidHost.Documents.EditedDocument;
                LogMessage("Document courant récupéré avec succès.", System.Drawing.Color.Green);
                return documentId;
            }
            catch (Exception ex)
            {
                // Log et affichage d'une erreur si la récupération échoue
                LogMessage($"Erreur lors de la récupération du document courant : {ex.Message}", System.Drawing.Color.Red);
                MessageBox.Show($"Erreur lors de la récupération du document courant : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return DocumentId.Empty;
            }
        }

        // Récupère le nom du document courant.
        private string NomDocumentCourant(DocumentId currentDoc)
        {
            try
            {
                // Récupération du nom du document courant
                string nomDocument = TopSolidHost.Documents.GetName(currentDoc);
                LogMessage($"Nom du document courant récupéré avec succès : {nomDocument}", System.Drawing.Color.Green);
                return nomDocument;
            }
            catch (Exception ex)
            {
                // Log et affichage d'une erreur si la récupération échoue
                LogMessage($"Erreur lors de la récupération du nom du document : {ex.Message}", System.Drawing.Color.Red);
                MessageBox.Show($"Erreur lors de la récupération du nom du document : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return "Nom inconnu";
            }
        }

        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //Fonction recurcive pour remonter au document piece original-------------------------------------------------------------------------------------------------------------------
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        private (DocumentId PrepaDocument, DocumentId docMaster) RecupDocuMaster(DocumentId currentDoc, bool firstPrepaDocumentFound = false)
        {
            DocumentId prepaDocument = DocumentId.Empty;
            DocumentId docMaster = DocumentId.Empty;
            bool documentDerivé = TSHD.Tools.IsDerived(currentDoc);

            if (!documentDerivé)
            {
                if (currentDoc != null && currentDoc != DocumentId.Empty)
                {
                    try
                    {
                        // Liste des opérations du document
                        List<ElementId> operationsList = OperationsList(currentDoc);
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
                                        LogMessage($"Document pièce trouvé : {instanceDocument}", System.Drawing.Color.Green);
                                        if (instanceDocument != DocumentId.Empty)
                                        {
                                            docMaster = instanceDocument;
                                            return (prepaDocument, docMaster);
                                        }
                                    }
                                    // Si c'est un document de prépa, relancer la fonction
                                    if (DocumentExt == ".TopNewPrtSet")
                                    {
                                        LogMessage($"Document de prépa trouvé : {instanceDocument}", System.Drawing.Color.Green);
                                        if (!firstPrepaDocumentFound)
                                        {
                                            firstPrepaDocumentFound = true;
                                            prepaDocument = instanceDocument;
                                        }
                                        var result = RecupDocuMaster(instanceDocument, firstPrepaDocumentFound);
                                        if (result.docMaster != DocumentId.Empty)
                                        {
                                            docMaster = result.docMaster;
                                            return (prepaDocument, docMaster);
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            LogMessage("Aucune opération trouvée dans le document courant.", System.Drawing.Color.Red);
                        }
                    }
                    catch (Exception ex)
                    {
                        LogMessage($"Erreur lors de la récupération du document maître : {ex.Message}", System.Drawing.Color.Red);
                        MessageBox.Show("Erreur : " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    LogMessage("Document courant non valide.", System.Drawing.Color.Red);
                }
            }
            else
            {
                // Si le document est dérivé, obtenir le document de base
                DocumentId docBaseDerivation = TSHD.Tools.GetBaseDocument(currentDoc);
                var result = RecupDocuMaster(docBaseDerivation, firstPrepaDocumentFound);
                if (result.docMaster != DocumentId.Empty)
                {
                    docMaster = result.docMaster;
                    return (prepaDocument, docMaster);
                }
            }

            // Retourner les valeurs par défaut si aucun document maître ou de prépa n'a été trouvé
            return (prepaDocument, docMaster);
        }

        // Récupère la liste des opérations d'un document
        private List<ElementId> OperationsList(DocumentId currentDoc)
        {
            // Vérification que le document courant est valide
            if (currentDoc != null && currentDoc != DocumentId.Empty)
            {
                try
                {
                    // Récupération de la liste des opérations du document
                    List<ElementId> operationsList = TSH.Operations.GetOperations(currentDoc);
                    LogMessage("Liste des opérations récupérée avec succès.", System.Drawing.Color.Green);
                    return operationsList;
                }
                catch (Exception ex)
                {
                    // Log et affichage d'une erreur si la récupération échoue
                    LogMessage($"Erreur : lors de la récupération de la liste des opérations : {ex.Message}", System.Drawing.Color.Red);
                    MessageBox.Show("Erreur : lors de la récupération de la liste des opérations " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                // Log si le document courant n'est pas valide
                LogMessage("Document courant non valide.", System.Drawing.Color.Red);
            }

            // Retourne une liste vide si le document n'est pas valide ou en cas d'erreur
            return new List<ElementId>();
        }

        // Déclaration des noms des paramètres texte à rechercher dans les publications
        string Commentaire = "Commentaire"; // Paramètre pour le commentaire
        string Designation = "Designation"; // Paramètre pour la désignation
        string Indice_3D = "Indice 3D";    // Paramètre pour l'indice 3D
        string OP = "OP";                  // Paramètre pour l'opération
        string nomElec = "Nom elec";       // Paramètre pour le nom de l'électrode
        string Nomdocu = "Nom_docu";      // Paramètre pour le nom du document

        //Recuperation des parametre dans le document maitre
        private void ParametreMaster(in DocumentId docMaster, in DocumentId PrepaDocument, out ElementId indice3D, out ElementId commentaireOriginal, out ElementId designationOriginal, out ElementId OPOriginal, out ElementId nomElecOriginal, out ElementId nomDocuOriginal)
        {
            // Affichage du nom du document de préparation s'il est valide
            if (PrepaDocument != DocumentId.Empty)
            {
                string nomDocPrepa = TSH.Documents.GetName(PrepaDocument);
                MessageBox.Show(nomDocPrepa);
                prepaTrouvé = true;
            }

            // Initialisation des paramètres out avec des valeurs par défaut
            indice3D = new ElementId();
            commentaireOriginal = new ElementId();
            designationOriginal = new ElementId();
            OPOriginal = new ElementId();
            nomElecOriginal = new ElementId();
            nomDocuOriginal = new ElementId();

            List<ElementId> ParameterPubliéList = new List<ElementId>();

            // Recherche des paramètres publiés dans le document maître
            if (docMaster != DocumentId.Empty)
            {
                ParameterPubliéList = TSH.Entities.GetPublishings(docMaster);
            }

            List<ElementId> ParameterPubliéListPrepaDocument = new List<ElementId>();
            if (prepaTrouvé)
            {
                ParameterPubliéListPrepaDocument = TSH.Entities.GetPublishings(PrepaDocument);
            }

            bool opOriginalFound = false; // Drapeau pour indiquer si OPOriginal a été trouvé

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

                        // Si le nom du paramètre est égal au nom attendu, renvoyer l'Element ID
                        if (ParameterName == Indice_3D)
                        {
                            indice3D = Parameter;
                            LogMessage($"Paramètre '{Indice_3D}' trouvé.", System.Drawing.Color.Green);
                        }
                        if (ParameterName == Commentaire)
                        {
                            commentaireOriginal = Parameter;
                            LogMessage($"Paramètre '{Commentaire}' trouvé.", System.Drawing.Color.Green);
                        }
                        if (ParameterName == Designation)
                        {
                            designationOriginal = Parameter;
                            LogMessage($"Paramètre '{Designation}' trouvé.", System.Drawing.Color.Green);
                        }
                        if (ParameterName == nomElec)
                        {
                            nomElecOriginal = Parameter;
                            LogMessage($"Paramètre '{nomElec}' trouvé.", System.Drawing.Color.Green);
                        }
                        if (ParameterName == Nomdocu)
                        {
                            nomDocuOriginal = Parameter;
                            LogMessage($"Paramètre '{Nomdocu}' trouvé.", System.Drawing.Color.Green);
                        }

                        if (prepaTrouvé && !opOriginalFound)
                        {
                            bool dérivé = TSHD.Tools.IsDerived(PrepaDocument);
                            try
                            {
                                if (dérivé)
                                {
                                    List<ElementId> parameters = TSH.Parameters.GetParameters(PrepaDocument);

                                    foreach (ElementId parameter in parameters)
                                    {
                                        string parameterTxt = TSH.Elements.GetFriendlyName(parameter);

                                        if (parameterTxt == OP)
                                        {
                                            OPOriginal = parameter;
                                            LogMessage($"Paramètre '{OP}' trouvé.", System.Drawing.Color.Green);
                                            opOriginalFound = true; // Mettre le drapeau à true
                                            dérivé = false;
                                            break; // Sortir de la boucle interne
                                        }
                                    }
                                }
                                else
                                {
                                    foreach (ElementId ParameterPrepaDocument in ParameterPubliéListPrepaDocument)
                                    {
                                        string ParameterPrepaDocumentName = TSH.Elements.GetFriendlyName(ParameterPrepaDocument);

                                        if (ParameterPrepaDocumentName == OP)
                                        {
                                            OPOriginal = ParameterPrepaDocument;
                                            LogMessage($"Paramètre '{OP}' trouvé.", System.Drawing.Color.Green);
                                            opOriginalFound = true; // Mettre le drapeau à true
                                            break; // Sortir de la boucle interne
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                LogMessage($"Erreur : lors de la récupération des paramètres OP : {ex.Message}", System.Drawing.Color.Red);
                                MessageBox.Show("Erreur : lors de la récupération des paramètres OP " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogMessage($"Erreur : lors de la récupération des paramètres maître : {ex.Message}", System.Drawing.Color.Red);
                    MessageBox.Show("Erreur : lors de la récupération des paramètres maître " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                LogMessage("Erreur : La liste des paramètres publiés est vide dans le document maître.", System.Drawing.Color.Red);
                MessageBox.Show("Erreur : La liste des paramètres publiés est vide dans le document maître");
                indice3D = new ElementId();
                commentaireOriginal = new ElementId();
                designationOriginal = new ElementId();
            }
        }

        //Configure les parametres dans les documents sans electrodes
        private void SetSmartTxtParameter(DocumentId currentDoc, SmartText[] SmartTxtTable, ElementId OPOriginal)
        {
            ElementId OPPublishingElement = ElementId.Empty;

            // Démarre la modification dans l'application TopSolid
            if (!TopSolidHost.Application.StartModification("My Action", false)) return;

            currentDoc = DocumentCourant();

            bool CommentaireUpdated = false;
            bool DesignationUpdated = false;
            bool Indice_3DUpdated = false;
            bool NomElecUpdated = false;
            bool OPUpdated = false;
            bool OPDefinitionSet = false; // Variable pour suivre si la méthode a déjà été exécutée
            bool NomDocuSet = false;

            try
            {
                TopSolidHost.Documents.EnsureIsDirty(ref currentDoc);

                // Recherche des paramètres publiés dans le document courant
                List<ElementId> ParameterPubliéList = TSH.Parameters.GetParameters(currentDoc);
                string DocumentExt = String.Empty;

                // Vérifie si des paramètres ont été trouvés
                if (ParameterPubliéList.Count > 0)
                {
                    // Parcours de tous les paramètres publiés
                    foreach (ElementId ParameterPublié in ParameterPubliéList)
                    {
                        string Parametertype = TSH.Elements.GetTypeFullName(ParameterPublié);

                        // Récupération du PdmObjectId associé
                        PdmObjectId pdmObjectId = TSH.Documents.GetPdmObject(currentDoc);

                        // Récupération de l'extension du document
                        TSH.Pdm.GetType(pdmObjectId, out DocumentExt);

                        // Vérification du type de paramètre
                        if (Parametertype == "TopSolid.Kernel.DB.Parameters.TextParameterEntity")
                        {
                            string ParameterPubliéName = TSH.Elements.GetFriendlyName(ParameterPublié);

                            // Mise à jour du paramètre "Commentaire"
                            if (ParameterPubliéName == Commentaire)
                            {
                                ElementId ParameterPubliéOp = TSH.Elements.GetParent(ParameterPublié);
                                try
                                {
                                    TSH.Parameters.SetSmartTextParameterCreation(ParameterPubliéOp, SmartTxtTable[1]);
                                    CommentaireUpdated = true; // Marquer la méthode comme exécutée
                                    LogMessage($"Paramètre '{Commentaire}' mis à jour.", System.Drawing.Color.Green);
                                }
                                catch (Exception ex)
                                {
                                    LogMessage($"Erreur : lors de la mise à jour du paramètre 'Commentaire' : {ex.Message}", System.Drawing.Color.Red);
                                    throw; // Relancer l'exception pour s'assurer que le bloc finally est exécuté
                                }
                            }

                            // Mise à jour du paramètre "Designation"
                            if (ParameterPubliéName == Designation)
                            {
                                ElementId ParameterPubliéOp = TSH.Elements.GetParent(ParameterPublié);
                                try
                                {
                                    TSH.Parameters.SetSmartTextParameterCreation(ParameterPubliéOp, SmartTxtTable[2]);
                                    DesignationUpdated = true; // Marquer la méthode comme exécutée
                                    LogMessage($"Paramètre '{Designation}' mis à jour.", System.Drawing.Color.Green);
                                }
                                catch (Exception ex)
                                {
                                    LogMessage($"Erreur : lors de la mise à jour du paramètre 'Designation' : {ex.Message}", System.Drawing.Color.Red);
                                    throw; // Relancer l'exception pour s'assurer que le bloc finally est exécuté
                                }
                            }

                            // Mise à jour du paramètre "Indice_3D"
                            if (ParameterPubliéName == Indice_3D)
                            {
                                ElementId ParameterPubliéOp = TSH.Elements.GetParent(ParameterPublié);
                                try
                                {
                                    TSH.Parameters.SetSmartTextParameterCreation(ParameterPubliéOp, SmartTxtTable[3]);
                                    Indice_3DUpdated = true; // Marquer la méthode comme exécutée
                                    LogMessage($"Paramètre '{Indice_3D}' mis à jour.", System.Drawing.Color.Green);
                                }
                                catch (Exception ex)
                                {
                                    LogMessage($"Erreur : lors de la mise à jour du paramètre 'Indice_3D' : {ex.Message}", System.Drawing.Color.Red);
                                    throw; // Relancer l'exception pour s'assurer que le bloc finally est exécuté
                                }
                            }

                            // Traitement spécifique pour l'extension .TopMillTurn
                            if (DocumentExt == ".TopMillTurn")
                            {
                                if (ParameterPubliéName == OP)
                                {
                                    ElementId ParameterPubliéOp = TSH.Elements.GetParent(ParameterPublié);
                                    try
                                    {
                                        TSH.Parameters.SetSmartTextParameterCreation(ParameterPubliéOp, SmartTxtTable[0]);
                                        OPDefinitionSet = true; // Marquer la méthode comme exécutée
                                        LogMessage($"Paramètre '{OP}' mis à jour.", System.Drawing.Color.Green);
                                        OPUpdated = true;
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show("Erreur : lors de l'édition du paramètre OP " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        throw; // Relancer l'exception pour s'assurer que le bloc finally est exécuté
                                    }
                                }
                                // Mise à jour du paramètre "Commentaire"
                                if (ParameterPubliéName == Nomdocu)
                                {
                                    ElementId ParameterPubliéOp = TSH.Elements.GetParent(ParameterPublié);
                                    try
                                    {
                                        TSH.Parameters.SetSmartTextParameterCreation(ParameterPubliéOp, SmartTxtTable[5]);
                                        NomDocuSet = true; // Marquer la méthode comme exécutée
                                        LogMessage($"Paramètre '{Nomdocu}' mis à jour.", System.Drawing.Color.Green);
                                    }
                                    catch (Exception ex)
                                    {
                                        LogMessage($"Erreur : lors de la mise à jour du paramètre 'Commentaire' : {ex.Message}", System.Drawing.Color.Red);
                                        throw; // Relancer l'exception pour s'assurer que le bloc finally est exécuté
                                    }
                                }
                            }
                            // Traitement spécifique pour l'extension .TopNewPrtSet
                            else if (DocumentExt == ".TopNewPrtSet")
                            {
                                bool dérivé = TSHD.Tools.IsDerived(currentDoc);
                                if (dérivé)
                                {
                                    ElementId opCourantPrepa = RecupOpCourantPrepa(currentDoc);

                                    DocumentId docBase = TSHD.Tools.GetBaseDocument(currentDoc);
                                    List<ElementId> parameters = TSH.Parameters.GetParameters(docBase);
                                    foreach (ElementId parameter in parameters)
                                    {
                                        string parameterTxt = TSH.Elements.GetFriendlyName(parameter);

                                        if (parameterTxt == OP)
                                        {
                                            string parameterValue = TSH.Parameters.GetTextValue(parameter);
                                            int parameterValueInt = int.Parse(parameterValue);
                                            int newParameterValueInt = parameterValueInt + 1;
                                            string newParameterValue = newParameterValueInt.ToString();
                                            SmartText parameterValueIntSmartTxt = new SmartText(newParameterValue);

                                            TSH.Parameters.SetSmartTextParameterCreation(opCourantPrepa, parameterValueIntSmartTxt);

                                            LogMessage($"Paramètre '{OP}' trouvé et mis à jour.", System.Drawing.Color.Green);
                                            dérivé = false;
                                            break; // Sortir de la boucle interne
                                        }
                                    }
                                }
                                // Mise à jour du paramètre "Nom elec"
                                if (ParameterPubliéName == nomElec)
                                {
                                    ElementId ParameterPubliéNomElec = TSH.Elements.GetParent(ParameterPublié);
                                    try
                                    {
                                        TSH.Parameters.SetSmartTextParameterCreation(ParameterPubliéNomElec, SmartTxtTable[4]);
                                        NomElecUpdated = true; // Marquer la méthode comme exécutée
                                        LogMessage($"Paramètre '{nomElec}' mis à jour.", System.Drawing.Color.Green);
                                    }
                                    catch (Exception ex)
                                    {
                                        LogMessage($"Erreur : lors de la mise à jour du paramètre 'Designation' : {ex.Message}", System.Drawing.Color.Red);
                                        throw; // Relancer l'exception pour s'assurer que le bloc finally est exécuté
                                    }
                                }
                                // Mise à jour du paramètre "Commentaire"
                                if (ParameterPubliéName == Nomdocu)
                                {
                                    ElementId ParameterPubliéOp = TSH.Elements.GetParent(ParameterPublié);
                                    try
                                    {
                                        TSH.Parameters.SetSmartTextParameterCreation(ParameterPubliéOp, SmartTxtTable[5]);
                                        NomDocuSet = true; // Marquer la méthode comme exécutée
                                        LogMessage($"Paramètre '{Nomdocu}' mis à jour.", System.Drawing.Color.Green);
                                    }
                                    catch (Exception ex)
                                    {
                                        LogMessage($"Erreur : lors de la mise à jour du paramètre 'Commentaire' : {ex.Message}", System.Drawing.Color.Red);
                                        throw; // Relancer l'exception pour s'assurer que le bloc finally est exécuté
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    LogMessage("Erreur : La liste des paramètres courants est vide", System.Drawing.Color.Red);
                }

                // Construction du message de confirmation
                string confirmationMessage = $"Paramètres mis à jour :\n" +
                                            $"{Commentaire} : {(CommentaireUpdated ? "Oui" : "Non")}\n" +
                                            $"{Designation} : {(DesignationUpdated ? "Oui" : "Non")}\n" +
                                            $"{Indice_3D} : {(Indice_3DUpdated ? "Oui" : "Non")}\n" +
                                            $"{nomElec} : {(NomElecUpdated ? "Oui" : "Non")}\n" +
                                            $"{OP} : {(OPUpdated ? "Oui" : "Non")}";

                // Affichage de l'état des paramètres après la mise à jour
                if (DocumentExt == ".TopMillTurn" || !CommentaireUpdated || !DesignationUpdated || !Indice_3DUpdated || !OPUpdated)
                {
                    MessageBox.Show(confirmationMessage + "\n\n Attention, certains paramètres sont peut-être manquants.\n Vérifiez les documents parents.", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    LogMessage(confirmationMessage, System.Drawing.Color.Black);
                    MessageBox.Show(confirmationMessage, "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Erreur : lors de la récupération des paramètres courants : {ex.Message}", System.Drawing.Color.Red);
            }
            finally
            {
                // Assurer que le document est marqué comme modifié et fin de la modification dans l'application
                TopSolidHost.Documents.EnsureIsDirty(ref currentDoc);
                TopSolidHost.Application.EndModification(true, true);
            }
        }

        //Fonction pour recuperer le numero d'OP dans un document de prepa
        private ElementId RecupOpCourantPrepa(DocumentId currentDoc)
        {
            // Vérification que le document actuel est valide
            if (currentDoc == null)
            {
                // Log en rouge si le document est nul
                LogMessage("Erreur : Le document est nul.", System.Drawing.Color.Red);
                throw new ArgumentNullException("Le document est nul.");
            }

            // Récupération de la liste des paramètres (opérations) du document
            List<ElementId> parameters = TSH.Operations.GetOperations(currentDoc);

            // Vérification si des paramètres ont été récupérés
            if (parameters == null || parameters.Count == 0)
            {
                // Log en rouge si aucun paramètre n'est trouvé
                LogMessage("Erreur : Aucun paramètre trouvé dans le document.", System.Drawing.Color.Red);
                throw new InvalidOperationException("Aucun paramètre trouvé dans le document.");
            }

            // Recherche du paramètre "OP"
            foreach (ElementId parameter in parameters)
            {
                try
                {
                    // Récupération du nom du paramètre
                    string parameterName = TSH.Elements.GetFriendlyName(parameter);

                    // Vérification si le nom du paramètre est "Paramètre texte (OP)"
                    if (parameterName == "Paramètre texte (OP)")
                    {
                        // Log en vert pour indiquer que le paramètre a été trouvé
                        LogMessage("Paramètre 'OP' trouvé : " + parameter.ToString(), System.Drawing.Color.Green);
                        return parameter;
                    }
                }
                catch (Exception ex)
                {
                    // Log en rouge en cas d'erreur lors de la récupération du nom du paramètre
                    LogMessage("Erreur lors de la récupération du nom du paramètre : " + ex.Message, System.Drawing.Color.Red);
                }
            }

            // Si le paramètre "OP" n'a pas été trouvé, lever une exception
            LogMessage("Erreur : Aucun paramètre 'OP' trouvé.", System.Drawing.Color.Red);
            throw new InvalidOperationException("Aucun paramètre 'OP' trouvé.");
        }

        //Fonction de log
        private void LogMessage(string message, System.Drawing.Color color)
        {
            // Positionne le curseur à la fin du texte actuel dans le RichTextBox
            richTextBox1.SelectionStart = richTextBox1.TextLength;
            richTextBox1.SelectionLength = 0;

            // Définit la couleur du texte à ajouter
            richTextBox1.SelectionColor = color;

            // Ajoute le message au RichTextBox avec la couleur spécifiée
            richTextBox1.AppendText(message + Environment.NewLine);

            // Réinitialise la couleur du texte à la couleur par défaut du RichTextBox
            richTextBox1.SelectionColor = richTextBox1.ForeColor;

            // Fait défiler le RichTextBox pour s'assurer que le curseur est visible à la fin
            richTextBox1.ScrollToCaret();
        }

        //Fonction pour creer un smart text
        private SmartText CreateSmartTxt(ElementId elementId)
        {
            // Vérification que l'élément n'est pas vide avant de créer l'objet SmartText
            if (elementId == ElementId.Empty)
            {
                LogMessage("Erreur : Impossible de créer un SmartText avec un élément vide.", System.Drawing.Color.Red);
                return null; // Retourne null si l'élément est vide
            }

            // Création de l'objet SmartText
            SmartText smartTxtId = new SmartText(elementId);

            // Log en vert pour indiquer la création réussie de SmartText
            LogMessage($"SmartText créé pour l'élément : {elementId}", System.Drawing.Color.Green);

            return smartTxtId;
        }

        //Fonction pour creer un smart integer
        private SmartInteger CreateSmartInt(ElementId elementId)
        {
            // Vérification que l'élément n'est pas vide avant de créer l'objet SmartInteger
            if (elementId == ElementId.Empty)
            {
                LogMessage("Erreur : Impossible de créer un SmartInteger avec un élément vide.", System.Drawing.Color.Red);
                return null; // Retourne null si l'élément est vide
            }

            // Création de l'objet SmartInteger
            SmartInteger smartIntId = new SmartInteger(elementId);

            // Log en vert pour indiquer la création réussie de SmartInteger
            LogMessage($"SmartInteger créé pour l'élément : {elementId}", System.Drawing.Color.Green);

            return smartIntId;
        }

        //Fonction pour creer un smart real
        private SmartReal CreateSmartReal(ElementId real)
        {
            // Vérification que l'élément n'est pas vide avant de créer l'objet SmartReal
            if (real == ElementId.Empty)
            {
                LogMessage("Erreur : Impossible de créer un SmartReal avec un élément vide.", System.Drawing.Color.Red);
                return null; // Retourne null si l'élément est vide
            }

            // Création de l'objet SmartReal
            SmartReal smartRealId = new SmartReal(real);

            // Log en vert pour indiquer la création réussie de SmartReal
            LogMessage($"SmartReal créé pour l'élément : {real}", System.Drawing.Color.Green);

            return smartRealId;
        }

        //Configuration des parametres de derivations
        private void DerivationConfig(DocumentId currentDoc)
        {
            // Vérification que le document actuel est valide
            if (currentDoc != null)
            {
                // Vérification si le document est dérivé
                bool isDerived = TSHD.Tools.IsDerived(currentDoc);
                if (isDerived)
                {
                    // Tentative de démarrer la modification
                    if (TopSolidHost.Application.StartModification("My Modification", false))
                    {
                        // Récupérer à nouveau le document courant
                        currentDoc = DocumentCourant();
                        try
                        {
                            // Application des héritages de dérivation avec les paramètres spécifiés
                            TSHD.Tools.SetDerivationInheritances(
                                currentDoc,                  // DocumentId
                                false,                        // Nom
                                false,                        // Description
                                false,                        // Code
                                false,                        // Numéro de pièce
                                false,                        // Numéro de pièce complémentaire
                                false,                        // Fabricant
                                false,                        // Numéro de pièce du fabricant
                                false,                        // Commentaire
                                new List<ElementId>(),        // Autres paramètres du système
                                false,                        // Paramètres non systèmes
                                true,                         // Points
                                true,                         // Axes
                                true,                         // Plans
                                true,                         // Cadres
                                true,                         // Esquisses
                                true,                         // Formes
                                true,                         // Publications
                                true,                         // Fonctions
                                true,                         // Symétries
                                true,                         // Non sectionnabilité
                                false,                        // Représentations
                                false,                        // Ensembles
                                true                          // Caméras
                            );

                            // Marquer le document comme modifié
                            TopSolidHost.Documents.EnsureIsDirty(ref currentDoc);

                            // Finaliser la modification
                            TopSolidHost.Application.EndModification(true, true);

                            // Log en vert pour indiquer le succès
                            LogMessage("Modification réussie : Héritages de dérivation appliqués.", System.Drawing.Color.Green);
                        }
                        catch (Exception ex)
                        {
                            // Log en rouge en cas d'erreur
                            LogMessage("Erreur lors de l'application des héritages de dérivation : " + ex.Message, System.Drawing.Color.Red);

                            // Marquer le document comme modifié malgré l'erreur
                            TopSolidHost.Documents.EnsureIsDirty(ref currentDoc);

                            // Affichage de l'erreur à l'utilisateur
                            MessageBox.Show("Erreur : " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            // Annuler la modification
                            TopSolidHost.Application.EndModification(false, false);
                        }
                    }
                    else
                    {
                        // Log en rouge si la modification ne peut pas commencer
                        LogMessage("Erreur : Impossible de démarrer la modification.", System.Drawing.Color.Red);
                    }
                }
                else
                {
                    // Log en noir si le document n'est pas dérivé
                    LogMessage("Le document n'est pas dérivé.", System.Drawing.Color.Black);
                }
            }
            else
            {
                // Log en rouge si le document est nul
                LogMessage("Erreur : Le document est nul.", System.Drawing.Color.Red);
            }
        }

        //Configure les parametres des docus electrodes
        private void ParamElecListe(DocumentId currentDoc)
        {
            // Définition des paramètres de texte
            const string Indice_3DOp = "Paramètre texte (Indice 3D)";
            const string CommentaireOp = "Paramètre texte (Nom_docu)";
            const string DesignationOp = "Paramètre texte (Designation)";
            const string Nom_elecTxt = "Paramètre texte (Nom elec)";
            const string TotalBrutTxt = "Paramètre entier (Total brut)";

            // Récupération de la liste des opérations
            List<ElementId> operations = TSH.Operations.GetOperations(currentDoc);
            if (operations == null)
            {
                // Log en rouge si les opérations ne peuvent pas être récupérées
                LogMessage("Erreur : Impossible de récupérer les opérations.", System.Drawing.Color.Red);
                MessageBox.Show("Erreur : Impossible de récupérer les opérations.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Initialisation de la liste des paramètres électrodes
            List<ElementId> paramElecList = new List<ElementId>();
            string nom = "";

            ElementId nomElecOriginal = new ElementId();
            bool isElectrode = Iselectrode(currentDoc);

            if (isElectrode)
            {
                nomElecOriginal = TSH.Parameters.GetNameParameter(currentDoc);
                LogMessage($"Paramètre Nom_elec trouvé.", System.Drawing.Color.Green); // Succès en vert
            }

            // Recherche des éléments nécessaires
            ElementId nomDocu = TSH.Elements.SearchByName(currentDoc, "$TopSolid.Cad.Electrode.DB.Electrodes.ShapeToErodeName");
            ElementId designationPiece = TSH.Elements.SearchByName(currentDoc, "$TopSolid.Cad.Electrode.DB.Electrodes.ShapeToErodeDescription");
            ElementId indice3DElec = TSH.Elements.SearchByName(currentDoc, "Indice Elec");
            ElementId TotalBrutId = SearchParamByName(currentDoc, "Total brut");

            // Fonction Total brut
            int TotalBrut = TotalBrutCalcul(currentDoc);

            // Initialisation des objets SmartText
            SmartText nomDocuSmart = new SmartText("");
            SmartText nomElecSmart = new SmartText("");
            SmartText designationPieceSmart = new SmartText("");
            SmartText indice3DElecSmart = new SmartText("");
            SmartInteger TotalBrutSmart = new SmartInteger(TotalBrut);

            // Si les éléments sont trouvés, on crée les objets SmartText
            if (nomDocu != ElementId.Empty)
            {
                nomDocuSmart = CreateSmartTxt(nomDocu);
            }
            if (nomElecOriginal != ElementId.Empty)
            {
                nomElecSmart = CreateSmartTxt(nomElecOriginal);
            }
            if (designationPiece != ElementId.Empty)
            {
                designationPieceSmart = CreateSmartTxt(designationPiece);
            }
            if (indice3DElec != ElementId.Empty)
            {
                indice3DElecSmart = CreateSmartTxt(indice3DElec);
            }

            // Tableau des SmartText
            SmartText[] SmartTxtTable = new SmartText[4];
            SmartTxtTable[0] = nomDocuSmart;
            SmartTxtTable[1] = nomElecSmart;
            SmartTxtTable[2] = designationPieceSmart;
            SmartTxtTable[3] = indice3DElecSmart;

            // Tentative de démarrer la modification
            if (!TopSolidHost.Application.StartModification("My Action", false))
            {
                // Log en rouge si la modification ne peut pas commencer
                LogMessage("Erreur : Impossible de démarrer la modification.", System.Drawing.Color.Red);
                MessageBox.Show("Erreur : Impossible de démarrer la modification.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            currentDoc = DocumentCourant();
            try
            {
                // Marquer le document comme modifié
                TopSolidHost.Documents.EnsureIsDirty(ref currentDoc);

                // Traitement des opérations
                foreach (ElementId operation in operations)
                {
                    try
                    {
                        // Récupération du nom de l'opération
                        nom = TSH.Elements.GetFriendlyName(operation);
                    }
                    catch (Exception ex)
                    {
                        // Log en rouge en cas d'erreur lors de la récupération du nom
                        LogMessage("Erreur : Impossible de récupérer le nom des paramètres électrode. " + ex.Message, System.Drawing.Color.Red);
                        MessageBox.Show("Erreur : Impossible de récupérer le nom des paramètres électrode. " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        continue;
                    }

                    // Vérification du nom pour appliquer les paramètres correspondants
                    if (nom == CommentaireOp)
                    {
                        TSH.Parameters.SetSmartTextParameterCreation(operation, SmartTxtTable[0]);
                        LogMessage($"Paramètre '{CommentaireOp}' appliqué avec succès.", System.Drawing.Color.Green); // Succès en vert
                    }
                    else if (nom == DesignationOp)
                    {
                        TSH.Parameters.SetSmartTextParameterCreation(operation, SmartTxtTable[2]);
                        LogMessage($"Paramètre '{DesignationOp}' appliqué avec succès.", System.Drawing.Color.Green); // Succès en vert
                    }
                    else if (nom == Indice_3DOp)
                    {
                        TSH.Parameters.SetSmartTextParameterCreation(operation, SmartTxtTable[3]);
                        LogMessage($"Paramètre '{Indice_3DOp}' appliqué avec succès.", System.Drawing.Color.Green); // Succès en vert
                    }
                    else if (nom == Nom_elecTxt)
                    {
                        TSH.Parameters.SetSmartTextParameterCreation(operation, SmartTxtTable[1]);
                        LogMessage($"Paramètre '{Nom_elecTxt}' appliqué avec succès.", System.Drawing.Color.Green); // Succès en vert
                    }
                    else if (nom == TotalBrutTxt)
                    {
                        TSH.Parameters.SetSmartIntegerParameterCreation(operation, TotalBrutSmart);
                        LogMessage($"Paramètre '{TotalBrutTxt}' appliqué avec succès.", System.Drawing.Color.Green); // Succès en vert
                    }
                }

                // Finalisation de la modification
                TopSolidHost.Application.EndModification(true, true);
                LogMessage("Modification terminée avec succès.", System.Drawing.Color.Green); // Succès en vert
            }
            catch (Exception ex)
            {
                // Log en rouge en cas d'erreur lors de la modification
                LogMessage("Erreur : Une erreur s'est produite lors de la modification. " + ex.Message, System.Drawing.Color.Red);
                MessageBox.Show("Erreur : Une erreur s'est produite lors de la modification. " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // Annuler la modification en cas d'erreur
                TopSolidHost.Application.EndModification(false, false);
            }
        }

        //Verifie si le gap est publié
        private bool GapPublishExist(List<ElementId> publishingList, string NomRecherche)
        {
            // Vérification des paramètres d'entrée pour éviter les valeurs nulles
            if (publishingList == null || NomRecherche == null)
            {
                LogMessage("Erreur : publishingList ou NomRecherche ne peut pas être null.", System.Drawing.Color.Red); // Erreur en rouge
                throw new ArgumentNullException("publishingList ou NomRecherche ne peut pas être null.");
            }

            // Si la liste n'est pas vide, on parcourt chaque élément
            if (publishingList.Count > 0)
            {
                foreach (var item in publishingList)
                {
                    try
                    {
                        // Récupération du nom de l'élément en cours
                        string name = TSH.Elements.GetName(item);

                        // Log du nom récupéré pour débogage
                        LogMessage($"Nom de l'élément : {name}", System.Drawing.Color.Blue);

                        // Comparaison du nom avec celui recherché
                        if (name == NomRecherche)
                        {
                            LogMessage($"L'élément '{NomRecherche}' a été trouvé.", System.Drawing.Color.Green); // Succès en vert
                            return true;
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log de l'exception en rouge
                        LogMessage($"Erreur lors de la récupération du nom de l'élément : {ex.Message}", System.Drawing.Color.Red);
                    }
                }
            }

            // Si aucun élément n'a été trouvé, on le log et retourne false
            LogMessage($"L'élément '{NomRecherche}' n'a pas été trouvé.", System.Drawing.Color.Red); // Message d'absence en rouge
            return false;
        }

        //Recupere les parametres de Gap et les publie
        private void GapPublish(DocumentId currentDoc)
        {
            // Définition des noms des paramètres
            const string GapEb = "Gap EB";
            const string GapDemiFini = "Gap Demi-Fini";
            const string GapFini = "Gap Fini";

            const string GapEbExisteTxt = "00-Gap Eb";
            const string GapDemiFiniExisteTxt = "01-Gap Demi fini";
            const string GapFiniExisteTxt = "02-Gap Fini";

            LogMessage("Début de la publication des gaps...", System.Drawing.Color.Black); // Information en noir

            // Recherche des paramètres
            ElementId GapEbId = SearchParamByName(currentDoc, GapEbExisteTxt);
            ElementId GapDemiFiniId = SearchParamByName(currentDoc, GapDemiFiniExisteTxt);
            ElementId GapFiniId = SearchParamByName(currentDoc, GapFiniExisteTxt);

            // Initialisation des SmartReal
            SmartReal gapEbPublie = new SmartReal(ElementId.Empty);
            SmartReal gapDemiFiniPublie = new SmartReal(ElementId.Empty);
            SmartReal gapFiniPublie = new SmartReal(ElementId.Empty);

            // Récupération de la liste des publications existantes
            List<ElementId> publishingList = TSH.Entities.GetPublishings(currentDoc);

            if (publishingList == null)
            {
                string errorMsg = "Erreur : Impossible de récupérer la liste des publications.";
                LogMessage(errorMsg, System.Drawing.Color.Red); // Erreur en rouge
                MessageBox.Show(errorMsg, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            LogMessage($"Nombre de publications existantes : {publishingList.Count}", System.Drawing.Color.Green); // Succès en vert

            // Vérification de l'existence des publications
            bool gapEbExist = GapPublishExist(publishingList, GapEb);
            bool gapDemiFiniExist = GapPublishExist(publishingList, GapDemiFini);
            bool gapFiniExist = GapPublishExist(publishingList, GapFini);

            // Début de la modification dans TopSolid
            if (!TopSolidHost.Application.StartModification("Publication des gaps", false))
            {
                string errorMsg = "Erreur : Impossible de démarrer la modification.";
                LogMessage(errorMsg, System.Drawing.Color.Red); // Erreur en rouge
                MessageBox.Show(errorMsg, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            currentDoc = DocumentCourant();

            try
            {
                // Vérification que le document est bien modifiable
                TopSolidHost.Documents.EnsureIsDirty(ref currentDoc);

                // Publication des gaps si elles n'existent pas déjà
                if (!gapEbExist && GapEbId != ElementId.Empty)
                {
                    gapEbPublie = CreateSmartReal(GapEbId);
                    ElementId realPublie = TSH.Parameters.PublishReal(currentDoc, GapEb, gapEbPublie);
                    TSH.Elements.SetName(realPublie, GapEb);
                    LogMessage("Publication de Gap EB réalisée avec succès.", System.Drawing.Color.Green); // Succès en vert
                }

                if (!gapDemiFiniExist && GapDemiFiniId != ElementId.Empty)
                {
                    gapDemiFiniPublie = CreateSmartReal(GapDemiFiniId);
                    ElementId realPublie = TSH.Parameters.PublishReal(currentDoc, GapDemiFini, gapDemiFiniPublie);
                    TSH.Elements.SetName(realPublie, GapDemiFini);
                    LogMessage("Publication de Gap Demi-Fini réalisée avec succès.", System.Drawing.Color.Green); // Succès en vert
                }

                if (!gapFiniExist && GapFiniId != ElementId.Empty)
                {
                    gapFiniPublie = CreateSmartReal(GapFiniId);
                    ElementId realPublie = TSH.Parameters.PublishReal(currentDoc, GapFini, gapFiniPublie);
                    TSH.Elements.SetName(realPublie, GapFini);
                    LogMessage("Publication de Gap Fini réalisée avec succès.", System.Drawing.Color.Green); // Succès en vert
                }

                // Validation de la modification
                TopSolidHost.Application.EndModification(true, true);
                LogMessage("Modification terminée avec succès.", System.Drawing.Color.Green); // Succès en vert
            }
            catch (Exception ex)
            {
                string errorMsg = $"Erreur : Une erreur s'est produite lors de la modification. {ex.Message}";
                LogMessage(errorMsg, System.Drawing.Color.Red); // Erreur en rouge
                MessageBox.Show(errorMsg, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);

                // Annulation de la modification en cas d'erreur
                TopSolidHost.Application.EndModification(false, false);
            }
        }

        //Cherche un parametre par son nom dans un document
        private ElementId SearchParamByName(DocumentId currentDoc, string nomParam)
        {
            // Vérification que le document courant est valide
            if (currentDoc == null)
            {
                LogMessage("Erreur : Document invalide (null).", System.Drawing.Color.Red); // Erreur en rouge
                return ElementId.Empty;
            }

            List<ElementId> listeParam = new List<ElementId>();

            try
            {
                // Récupération de la liste des paramètres du document
                listeParam = TSH.Parameters.GetParameters(currentDoc);

                // Vérification si des paramètres existent
                if (listeParam == null || listeParam.Count == 0)
                {
                    string warningMsg = "Aucun paramètre trouvé dans le document.";
                    LogMessage(warningMsg, System.Drawing.Color.Red); // Erreur en rouge
                    MessageBox.Show(warningMsg, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return ElementId.Empty;
                }

                // Log du nombre de paramètres récupérés
                LogMessage($"Nombre de paramètres récupérés : {listeParam.Count}", System.Drawing.Color.Green); // Succès en vert
            }
            catch (Exception ex)
            {
                string errorMsg = $"Erreur : Impossible de récupérer la liste des paramètres. {ex.Message}";
                LogMessage(errorMsg, System.Drawing.Color.Red); // Erreur en rouge
                MessageBox.Show(errorMsg, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return ElementId.Empty;
            }

            // Parcours de la liste des paramètres pour rechercher celui correspondant au nom
            foreach (ElementId param in listeParam)
            {
                string paramName = string.Empty;

                try
                {
                    // Récupération du nom du paramètre
                    paramName = TSH.Elements.GetName(param);
                }
                catch (Exception ex)
                {
                    string errorMsg = $"Erreur : Impossible de récupérer le nom du paramètre. {ex.Message}";
                    LogMessage(errorMsg, System.Drawing.Color.Red); // Erreur en rouge
                    MessageBox.Show(errorMsg, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    continue; // Passe au paramètre suivant en cas d'erreur
                }

                // Vérification si le paramètre correspond au nom recherché
                if (!string.IsNullOrEmpty(paramName) && paramName.StartsWith(nomParam, StringComparison.OrdinalIgnoreCase))
                {
                    LogMessage($"Paramètre trouvé : {paramName}", System.Drawing.Color.Green); // Succès en vert
                    return param; // Retourne le paramètre trouvé
                }
            }

            // Si aucun paramètre correspondant n'a été trouvé
            string infoMsg = $"Aucun paramètre correspondant à '{nomParam}' trouvé.";
            LogMessage(infoMsg, System.Drawing.Color.Red); // Erreur en rouge
            return ElementId.Empty;
        }

        //Verifie si le document est une electrode
        public bool Iselectrode(DocumentId currentDoc)
        {
            try
            {
                // Récupération du dossier des paramètres système du document
                ElementId systemParametersFolder = TSH.Parameters.GetSystemParametersFolder(currentDoc);
                List<ElementId> parameterListe = TSH.Elements.GetConstituents(systemParametersFolder);

                // Vérification si des paramètres existent
                if (parameterListe == null || parameterListe.Count == 0)
                {
                    LogMessage("Aucun paramètre système trouvé dans le document.", System.Drawing.Color.Red); // Erreur en rouge
                    return false;
                }

                // Parcours des paramètres pour vérifier si c'est une électrode
                foreach (ElementId param in parameterListe)
                {
                    string nom = TSH.Elements.GetName(param);

                    // Log du nom du paramètre pour débogage
                    LogMessage($"Nom du paramètre : {nom}", System.Drawing.Color.Blue);

                    // Vérification des noms de paramètres correspondant à une électrode
                    if (nom == "$TopSolid.Kernel.TX.Properties.ElectrodeTemplateAssociation" ||
                        nom == "$TopSolid.Cad.Electrode.DB.Electrodes.ElectrodeDocument")
                    {
                        LogMessage("Le document contient une électrode.", System.Drawing.Color.Green); // Succès en vert
                        return true;
                    }
                }

                // Si aucun paramètre ne correspond à une électrode
                LogMessage("Le document ne contient pas d'électrode.", System.Drawing.Color.Green); // Succès en vert
                return false;
            }
            catch (Exception ex)
            {
                // Gestion des erreurs
                string errorMsg = $"Erreur lors de la vérification de l'électrode : {ex.Message}";
                LogMessage(errorMsg, System.Drawing.Color.Red); // Erreur en rouge
                return false;
            }
        }

        //trouver operation de propriété electrode
        private ElementId OperationByType(List<ElementId> operations, string typeCible)
        {
            // Vérification si la liste des opérations est nulle
            if (operations == null)
            {
                string errorMsg = "Erreur : La liste des opérations est nulle.";
                LogMessage(errorMsg, System.Drawing.Color.Red); // Erreur en rouge
                MessageBox.Show(errorMsg, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return ElementId.Empty;
            }

            // Parcours de la liste des opérations pour trouver celle du type recherché
            foreach (ElementId operation in operations)
            {
                try
                {
                    // Récupération du type de l'opération
                    string type = TSH.Elements.GetTypeFullName(operation);

                    // Log du type de l'opération pour débogage
                    LogMessage($"Type de l'opération : {type}", System.Drawing.Color.Blue);

                    // Vérification si l'opération correspond au type recherché
                    if (type == typeCible)
                    {
                        LogMessage($"Opération trouvée de type '{typeCible}'", System.Drawing.Color.Green); // Succès en vert
                        return operation; // Retourne l'opération trouvée
                    }
                }
                catch (Exception ex)
                {
                    // Log et affichage d'une erreur si la récupération du type échoue
                    string errorMsg = $"Erreur : Impossible de récupérer le type de l'opération. {ex.Message}";
                    LogMessage(errorMsg, System.Drawing.Color.Red); // Erreur en rouge
                    MessageBox.Show(errorMsg, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    continue; // Passe à l'opération suivante en cas d'erreur
                }
            }

            // Si aucune opération du type recherché n'a été trouvée
            string warningMsg = $"Aucune opération du type '{typeCible}' trouvée.";
            LogMessage(warningMsg, System.Drawing.Color.Red); // Erreur en rouge
            return ElementId.Empty;
        }

        //trouver position de l'operation de propriété electrode
        private int IndexOperation(ElementId operationCible, DocumentId currentDoc)
        {
            // Liste pour stocker toutes les opérations du document
            List<ElementId> operations = new List<ElementId>();

            // Récupération du stage de modélisation
            ElementId modelingStage = TSH.Operations.GetModelingStage(currentDoc);

            // Vérification si le stage de modélisation est vide
            if (modelingStage == ElementId.Empty)
            {
                LogMessage("Le stage de modélisation est vide.", System.Drawing.Color.Red); // Erreur en rouge
                return -1; // Retourne une valeur d'erreur si pas de stage valide
            }

            try
            {
                // Récupération de toutes les opérations du document
                operations = TSH.Operations.GetOperations(currentDoc);

                // Vérification si la liste des opérations est vide ou nulle
                if (operations == null || operations.Count == 0)
                {
                    string errorMsg = "Erreur : Impossible de récupérer les opérations du document.";
                    LogMessage(errorMsg, System.Drawing.Color.Red); // Erreur en rouge
                    MessageBox.Show(errorMsg, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return -1; // Retourne -1 en cas d'erreur
                }

                // Log du nombre d'opérations récupérées
                LogMessage($"Nombre d'opérations récupérées : {operations.Count}", System.Drawing.Color.Green); // Succès en vert
            }
            catch (Exception ex)
            {
                // Gestion des erreurs lors de la récupération des opérations
                string errorMsg = $"Erreur lors de la récupération des opérations : {ex.Message}";
                LogMessage(errorMsg, System.Drawing.Color.Red); // Erreur en rouge
                MessageBox.Show(errorMsg, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }

            // Filtrage des opérations qui appartiennent au modeling stage
            List<ElementId> modelingStageOperations = operations.Where(op => TSH.Elements.GetOwner(op).Equals(modelingStage)).ToList();

            // Log du nombre d'opérations dans le modeling stage
            LogMessage($"Nombre d'opérations dans le modeling stage : {modelingStageOperations.Count}", System.Drawing.Color.Green); // Succès en vert

            // Recherche de l'index de l'opération cible
            for (int i = 0; i < modelingStageOperations.Count; i++)
            {
                if (modelingStageOperations[i] == operationCible)
                {
                    LogMessage($"Opération cible trouvée à l'index {i}.", System.Drawing.Color.Green); // Succès en vert
                    return i; // Retourne l'index de l'opération trouvée
                }
            }

            // Si l'opération cible n'a pas été trouvée
            string warningMsg = "Opération cible non trouvée dans le modeling stage.";
            LogMessage(warningMsg, System.Drawing.Color.Red); // Erreur en rouge
            MessageBox.Show(warningMsg, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);

            return -1; // Retourne -1 si l'opération cible n'a pas été trouvée
        }

        //Fonction qui recupere extention du document
        private string Extention(DocumentId currentDoc)
        {
            PdmObjectId pdmObjectId = new PdmObjectId();
            string DocumentExt = "";

            try
            {
                // Récupération du PdmObjectId associé au document courant
                pdmObjectId = TSH.Documents.GetPdmObject(currentDoc);
                LogMessage("PdmObjectId récupéré avec succès.", System.Drawing.Color.Green); // Succès en vert
            }
            catch (Exception ex)
            {
                // Log et affichage d'une erreur si la récupération du PdmObjectId échoue
                LogMessage($"Erreur lors de la récupération du document courant : {ex.Message}", System.Drawing.Color.Red); // Erreur en rouge
                MessageBox.Show("Erreur lors de la récupération du document courant : " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return ""; // Retourne une chaîne vide en cas d'erreur
            }

            try
            {
                // Récupération de l'extension du document
                TSH.Pdm.GetType(pdmObjectId, out DocumentExt);
                LogMessage($"Extension du document récupérée : {DocumentExt}", System.Drawing.Color.Green); // Succès en vert
            }
            catch (Exception ex)
            {
                // Log et affichage d'une erreur si la récupération de l'extension échoue
                LogMessage($"Erreur lors de la récupération de l'extension du document : {ex.Message}", System.Drawing.Color.Red); // Erreur en rouge
                MessageBox.Show("Erreur lors de la récupération de l'extension du document : " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return ""; // Retourne une chaîne vide en cas d'erreur
            }

            return DocumentExt;
        }

        //Fonction qui recupere l'etape prepa d'un document d'assemblage electrode
        private ElementId PrepaElectrodeStage(DocumentId currentDoc)
        {
            // Récupération de la liste des étapes du document courant
            List<ElementId> stages = TSH.Operations.GetStages(currentDoc);

            // Parcours de chaque étape pour vérifier si c'est celle de préparation d'électrode
            foreach (ElementId stage in stages)
            {
                // Récupération du nom de l'étape
                string nomStage = TSH.Elements.GetName(stage);

                // Log du nom de l'étape pour débogage
                LogMessage($"Nom de l'étape : {nomStage}", System.Drawing.Color.Blue);

                // Vérification si l'étape correspond à celle de préparation d'électrode
                if (nomStage == "$TopSolid.Cad.Electrode.DB.Documents.ElementName.PreparationStage")
                {
                    LogMessage("Étape de préparation d'électrode trouvée.", System.Drawing.Color.Green); // Succès en vert
                                                                                                         // Retourne l'élément de l'étape
                    return stage;
                }
            }

            // Log si aucune étape de préparation n'est trouvée
            LogMessage("Aucune étape de préparation d'électrode trouvée.", System.Drawing.Color.Red); // Erreur en rouge

            // Retourne une valeur vide si aucune étape de préparation n'est trouvée
            return ElementId.Empty;
        }

        //Recherche elements par sont friendlyname
        private ElementId SearchByFriendlyName(DocumentId currentDoc, string FriendlyName)
        {
            // Récupération de la liste des éléments du document courant
            List<ElementId> elementIds = TSH.Parameters.GetParameters(currentDoc);

            // Parcours de chaque élément pour comparer son nom amical
            foreach (ElementId elementId in elementIds)
            {
                // Récupération du type de l'élément
                string type = TSH.Elements.GetTypeFullName(elementId);

                if (type != null)
                {
                    // Vérification si l'élément n'est pas un paramètre de texte localisable
                    if (type != "TopSolid.Kernel.DB.Parameters.LocalizableTextParameterEntity")
                    {
                        // Récupération du nom amical de l'élément
                        string FriendlyNameResult = TSH.Elements.GetFriendlyName(elementId);

                        // Log du nom amical récupéré pour débogage
                        LogMessage($"Nom amical récupéré : {FriendlyNameResult}", System.Drawing.Color.Blue);

                        // Si le nom amical correspond à celui recherché, retourne l'élément
                        if (FriendlyNameResult == FriendlyName)
                        {
                            LogMessage($"Élément trouvé : {FriendlyName}", System.Drawing.Color.Green); // Succès en vert
                            return elementId;
                        }
                    }
                }
            }

            // Log si aucun élément avec le nom amical donné n'est trouvé
            LogMessage($"Aucun élément avec le nom amical '{FriendlyName}' trouvé.", System.Drawing.Color.Red); // Erreur en rouge

            // Retourne un élément vide si aucun élément n'est trouvé
            return ElementId.Empty;
        }

        //fonction pour afficher le nom du document courant
        private void DisplayDocumentName(string NomCurrentDoc)
        {
            try
            {
                // Vérification si le nom du document est vide ou invalide
                if (string.IsNullOrEmpty(NomCurrentDoc))
                {
                    // Log et affichage d'un message d'erreur si le nom est vide
                    LogMessage("Erreur : Le nom du document est vide ou invalide.", System.Drawing.Color.Red);
                    labelDocumentName.Text = "Nom de document non disponible";
                }
                else
                {
                    // Affichage du nom du document dans le label
                    labelDocumentName.Text = NomCurrentDoc;
                    LogMessage($"Nom du document affiché : {NomCurrentDoc}", System.Drawing.Color.Green); // Succès en vert
                }
            }
            catch (Exception ex)
            {
                // Log et affichage d'une erreur si une exception est levée
                LogMessage($"Erreur dans DisplayDocumentName : {ex.Message}", System.Drawing.Color.Red);
                labelDocumentName.Text = "Erreur lors de l'affichage";
            }
        }

        //fonction pour afficher le nom du document maitre
        private bool DisplayMasterDocumentName()
        {
            bool prepaTrouvé = false; // Initialisation de la variable pour indiquer si un document de prépa est trouvé

            try
            {
                // Récupération de l'ID et du nom du document courant
                DocumentId documentCourantId = DocumentCourant();
                string documentCourantName = NomDocumentCourant(documentCourantId);

                // Appel de la fonction pour récupérer le document de prépa et le document maître
                var (prepaDocument, docMaster) = RecupDocuMaster(documentCourantId);

                // Vérification du document maître
                if (docMaster != DocumentId.Empty)
                {
                    // Récupération et affichage du nom du document maître
                    string documentMasterName = TSH.Documents.GetName(docMaster);
                    labelDocumentMasterName.Text = documentMasterName;
                    LogMessage($"Document maître trouvé : {documentMasterName}", System.Drawing.Color.Green); // Succès en vert
                }
                else
                {
                    // Affichage d'un message si aucun document maître n'est trouvé
                    labelDocumentMasterName.Text = "Aucun document maître trouvé.";
                    LogMessage("Aucun document maître trouvé.", System.Drawing.Color.Red); // Erreur en rouge
                }

                // Vérification du document de prépa
                if (prepaDocument != DocumentId.Empty)
                {
                    // Récupération et log du nom du document de prépa
                    string docPrepaName = TSH.Documents.GetName(prepaDocument);
                    LogMessage($"Document de prépa trouvé : {docPrepaName}", System.Drawing.Color.Green); // Succès en vert
                    prepaTrouvé = true;
                }
                else
                {
                    // Log si aucun document de prépa n'est trouvé
                    LogMessage("Aucun document de prépa trouvé.", System.Drawing.Color.Red); // Erreur en rouge
                    prepaTrouvé = false;
                }
            }
            catch (Exception ex)
            {
                // Log et affichage d'une erreur si la récupération des noms des documents échoue
                LogMessage($"Erreur lors de l'affichage des noms des documents : {ex.Message}", System.Drawing.Color.Red); // Erreur en rouge
                MessageBox.Show("Erreur : lors de l'affichage des noms des documents " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return prepaTrouvé; // Retourne la valeur indiquant si un document de prépa a été trouvé
        }

        //Fonction pour redemarrer l'app
        private void RestartApplication()
        {
            // Récupération du document courant
            try
            {
                if (currentDoc.DocId != DocumentId.Empty)
                {
                    // Récupération et log du nom du document courant
                    string nomDocumentCourant = NomDocumentCourant(currentDoc.DocId);
                    LogMessage($"Document courant : {nomDocumentCourant}", System.Drawing.Color.Green); // Succès en vert
                }
                else
                {
                    // Log si aucun document courant n'est trouvé
                    LogMessage("Aucun document courant trouvé.", System.Drawing.Color.Red); // Erreur en rouge
                }
            }
            catch (Exception ex)
            {
                // Log et affichage d'une erreur si la récupération du document échoue
                LogMessage($"Une erreur s'est produite : {ex.Message}", System.Drawing.Color.Red); // Erreur en rouge
                MessageBox.Show($"Une erreur s'est produite : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // Affichage des informations sur les documents
            DisplayDocumentName(currentDoc.DocNomTxt); // Affiche le nom du document courant
            DisplayMasterDocumentName(); // Affiche le nom du document maître

            // Optionnel : redémarrer l'application (commenté dans votre code original)
            // string applicationPath = Application.ExecutablePath;
            // Process.Start(applicationPath);
            // Application.Exit();
        }

        //Fonction qui recupere l'elementId de l'electrode dans l'insertion
        private ElementId ElecsEnInsertionId(DocumentId currentDoc, HashSet<DocumentId> visitedDocs = null)
        {
            // Initialisation du HashSet pour suivre les documents visités
            if (visitedDocs == null)
                visitedDocs = new HashSet<DocumentId>();

            // Vérification si le document courant est vide ou déjà visité
            if (currentDoc == DocumentId.Empty || visitedDocs.Contains(currentDoc))
                return ElementId.Empty;

            // Marquer le document courant comme visité
            visitedDocs.Add(currentDoc);

            // Récupération de la liste des opérations dans le document courant
            List<ElementId> operationIds = TSH.Operations.GetOperations(currentDoc);

            // Parcours de chaque opération
            foreach (ElementId operationId in operationIds)
            {
                try
                {
                    // Vérification si l'opération est une inclusion
                    if (TSHD.Assemblies.IsInclusion(operationId))
                    {
                        // Récupération de l'ID du document enfant inclus
                        ElementId documentChildId = TSHD.Assemblies.GetInclusionChildOccurrence(operationId);
                        DocumentId documentId = documentChildId.DocumentId;

                        // Vérification si le document enfant est vide
                        if (documentId == DocumentId.Empty)
                        {
                            LogMessage($"L'inclusion {operationId} a un document vide. Ignoré.", System.Drawing.Color.Orange);
                            continue;
                        }

                        // Log de l'analyse de l'inclusion
                        LogMessage($"Analyse de l'inclusion : {documentId}", System.Drawing.Color.Blue);

                        // Vérification si le document enfant est une électrode
                        if (Iselectrode(documentId))
                        {
                            LogMessage($"Électrode trouvée : {documentId}", System.Drawing.Color.Green);
                            return documentChildId; // Retourne l'ID de l'électrode trouvée
                        }

                        // Exploration récursive sur l'inclusion
                        ElementId recursiveResult = ElecsEnInsertionId(documentId, visitedDocs);
                        if (recursiveResult != ElementId.Empty)
                            return recursiveResult;
                    }
                    else
                    {
                        // Log si aucune inclusion n'est trouvée pour l'opération
                        LogMessage($"Aucune inclusion trouvée pour l'opération {operationId}", System.Drawing.Color.Black);
                    }
                }
                catch (Exception ex)
                {
                    // Log et affichage d'une erreur si le traitement des inclusions échoue
                    LogMessage($"Erreur lors du traitement des inclusions : {ex.Message}", System.Drawing.Color.Red);
                }
            }

            // Retourne un ID vide si aucune électrode n'est trouvée
            return ElementId.Empty;
        }

        //Fonction Total brut
        private int TotalBrutCalcul(DocumentId currentDoc)
        {
            // Constantes pour les noms des paramètres à rechercher
            const string nbrElecEbTxt = "Nbr d'élec Eb";
            const string nbrElecDemiFiniTxt = "Nbr d'élec Demi-fini";
            const string nbrElecFiniTxt = "Nbr d'élec Fini";

            int nbr = 0; // Variable pour stocker le nombre temporaire
            int totalBrut = 0; // Variable pour accumuler le total brut

            // Liste des noms des paramètres à rechercher
            List<string> listNbrElecTxt = new List<string> { nbrElecEbTxt, nbrElecDemiFiniTxt, nbrElecFiniTxt };

            // Parcours de chaque nom de paramètre dans la liste
            foreach (string nbrElecTxt in listNbrElecTxt)
            {
                // Recherche du paramètre par son nom
                ElementId param = SearchParamByName(currentDoc, nbrElecTxt);

                if (!param.Equals(ElementId.Empty))
                {
                    try
                    {
                        // Récupération de la valeur entière du paramètre
                        nbr = TSH.Parameters.GetIntegerValue(param);
                        totalBrut += nbr; // Ajoute la valeur de nbr à totalBrut
                        LogMessage($"Paramètre '{nbrElecTxt}' trouvé et ajouté au total brut : {nbr}", System.Drawing.Color.Green);
                    }
                    catch (Exception ex)
                    {
                        // Log et affichage d'une erreur si la récupération de la valeur échoue
                        LogMessage($"Erreur lors de la récupération de la valeur du paramètre '{nbrElecTxt}' : {ex.Message}", System.Drawing.Color.Red);
                        MessageBox.Show("Erreur : " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    // Log d'information si le paramètre n'est pas trouvé
                    LogMessage($"Paramètre '{nbrElecTxt}' non trouvé.", System.Drawing.Color.Black);
                }
            }

            // Retourne le total brut calculé
            return totalBrut;
        }

        //Formulaire des options---------------------------------------------------------------------------------------------------
        private void optionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Créez une instance de
            FormulaireConfig form2 = new FormulaireConfig();
            // Affichez le formulaire
            form2.Show();
        }

        //Bouton refresh
        private void buttonRestart_Click(object sender, EventArgs e)
        {
            try
            {
                // Appel à la méthode de redémarrage de l'application
                RestartApplication();

                // Log du succès du redémarrage en vert
                LogMessage("Application redémarrée avec succès.", System.Drawing.Color.Green);
            }
            catch (Exception ex)
            {
                // Log de l'erreur en rouge en cas d'échec du redémarrage
                LogMessage($"Erreur lors du redémarrage de l'application : {ex.Message}", System.Drawing.Color.Red);

                // Affichage d'une boîte de dialogue avec le message d'erreur
                MessageBox.Show("Erreur lors du redémarrage de l'application: " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //Bouton quiter-----------------------------------------------------------------------------------------------------------
        private void quitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }

        //Bouton invok
        private void button3_Click(object sender, EventArgs e)
        {
            

        }

        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // Gestionnaire de clic pour le bouton Fix.-----------------------------------------------------------------------------------------------------------------------------------------
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                // Vérification si un document courant est sélectionné
                if (currentDoc.DocId != DocumentId.Empty)
                {
                    LogMessage($"Document courant : {currentDoc.DocNomTxt}", System.Drawing.Color.Black);
                }
                else
                {
                    LogMessage("Aucun document courant trouvé.", System.Drawing.Color.Red);
                    return;
                }
            }
            catch (Exception ex)
            {
                // Log et affichage d'une erreur si la récupération du document échoue
                LogMessage($"Erreur lors de la récupération du document courant : {ex.Message}", System.Drawing.Color.Red);
                return;
            }

            try
            {
                // Récupération de l'extension du document
                string ext = Extention(currentDoc.DocId);

                if (!string.IsNullOrEmpty(ext))
                {
                    if (ext == ".TopPrt") // Vérifie si c'est un document d'électrode
                    {
                        if (currentDoc.DocIsElectrode)
                        {
                            LogMessage("Détection d'un document électrode.", System.Drawing.Color.Black);

                            // Variables pour la recherche d'opérations spécifiques
                            string proprieteElecTxt = "TopSolid.Cad.Electrode.DB.PropertyData.PropertyDataOperation";
                            string dossierType = "TopSolid.Kernel.DB.Operations.FolderOperation";

                            ElementId operationProprieteElec = new ElementId();
                            ElementId operationDossierType = new ElementId();

                            try
                            {
                                LogMessage($"Nombre d'opérations trouvées : {currentDoc.DocOperations.Count}", System.Drawing.Color.Black);
                            }
                            catch (Exception ex)
                            {
                                // Log et affichage d'une erreur si la récupération des opérations échoue
                                LogMessage($"Erreur lors de la récupération des opérations : {ex.Message}", System.Drawing.Color.Red);
                                return;
                            }

                            if (currentDoc.DocOperations.Count > 0)
                            {
                                operationProprieteElec = OperationByType(currentDoc.DocOperations, proprieteElecTxt);
                                operationDossierType = OperationByType(currentDoc.DocOperations, dossierType);
                            }

                            // Vérifie si une opération dossier est trouvée
                            if (operationDossierType != ElementId.Empty)
                            {
                                string nomDossier = TSH.Elements.GetFriendlyName(operationDossierType);
                                LogMessage($"Nom du dossier trouvé : {nomDossier}", System.Drawing.Color.Black);
                            }

                            // Vérifie si l'opération électrode a été trouvée
                            if (operationProprieteElec == ElementId.Empty)
                            {
                                MessageBox.Show("Propriétés électrode manquante.\n\n" +
                                                "Avez-vous configuré les propriétés électrode dans l'ensemble électrode ?\n\n" +
                                                "Relancez l'application une fois les propriétés électrode configurées pour actualiser les paramètres.",
                                                "Erreur : Propriétés électrode manquante",
                                                MessageBoxButtons.OK,
                                                MessageBoxIcon.Warning);
                                LogMessage("Propriétés électrode manquante.", System.Drawing.Color.Red);
                            }
                            else
                            {
                                ElementId parentOperationProprieteElec = TSH.Elements.GetOwner(operationProprieteElec);
                                int indexoperationProprieteElec = IndexOperation(operationProprieteElec, currentDoc.DocId);
                                string message = TSH.Elements.GetFriendlyName(parentOperationProprieteElec);
                                LogMessage($"Nom du parent de l'opération : {message}", System.Drawing.Color.Black);

                                int positionCible = indexoperationProprieteElec + 1;

                                // Début de la modification
                                if (!TopSolidHost.Application.StartModification("Déplacement d'opération", false))
                                {
                                    LogMessage("Impossible de démarrer la modification.", System.Drawing.Color.Red);
                                    return;
                                }

                                currentDoc.DocId = DocumentCourant();
                                DocumentId currentDocId = currentDoc.DocId;

                                try
                                {
                                    TopSolidHost.Documents.EnsureIsDirty(ref currentDocId);
                                    TSH.Operations.MoveOperation(operationDossierType, parentOperationProprieteElec, positionCible);
                                    TSH.Documents.Update(currentDoc.DocId, true);
                                    TopSolidHost.Application.EndModification(true, true);
                                    LogMessage("Opération déplacée avec succès.", System.Drawing.Color.Green); // Succès en vert
                                }
                                catch (Exception ex)
                                {
                                    TopSolidHost.Application.EndModification(false, false);
                                    LogMessage($"Erreur lors du déplacement de l'opération : {ex.Message}", System.Drawing.Color.Red);
                                }
                            }
                            // Récupération des paramètres électrode
                            ParamElecListe(currentDoc.DocId);
                            GapPublish(currentDoc.DocId);
                        }
                    }
                    else // Autres types de documents
                    {
                        LogMessage("Document non électrode détecté.", System.Drawing.Color.Black);

                        // Récupération du document maître et du document de préparation
                        var (prepaDocument, docMaster) = RecupDocuMaster(currentDoc.DocId);

                        // Configuration des paramètres de dérivation
                        DerivationConfig(currentDoc.DocId);

                        // Gestion du document maître
                        if (docMaster != DocumentId.Empty)
                        {
                            string docMasterName = TSH.Documents.GetName(docMaster);
                            LogMessage($"Document maître trouvé : {docMasterName}", System.Drawing.Color.Black);
                        }
                        else
                        {
                            LogMessage("Aucun document maître trouvé.", System.Drawing.Color.Red);
                        }

                        // Gestion du document de préparation
                        if (prepaDocument != DocumentId.Empty)
                        {
                            string docPrepaName = TSH.Documents.GetName(prepaDocument);
                            LogMessage($"Document de prépa trouvé : {docPrepaName}", System.Drawing.Color.Black);
                        }
                        else
                        {
                            LogMessage("Aucun document de prépa trouvé.", System.Drawing.Color.Red);
                        }

                        // Récupération des paramètres maîtres
                        ElementId indice3D, commentaireOriginal, designationOriginal, OPOriginal, nomElecOriginal, nomDocuOriginal;
                        ParametreMaster(in docMaster, in prepaDocument, out indice3D, out commentaireOriginal, out designationOriginal, out OPOriginal, out nomElecOriginal, out nomDocuOriginal);

                        // Création des SmartText
                        SmartText[] SmartTxtTable = new SmartText[6]
                        {
                    CreateSmartTxt(OPOriginal),         // Index 0
                    CreateSmartTxt(commentaireOriginal), // Index 1
                    CreateSmartTxt(designationOriginal), // Index 2
                    CreateSmartTxt(indice3D),            // Index 3
                    CreateSmartTxt(nomElecOriginal),    // Index 4
                    CreateSmartTxt(nomDocuOriginal)      // Index 5
                        };

                        // Appliquer les paramètres SmartText au document courant
                        SetSmartTxtParameter(currentDoc.DocId, SmartTxtTable, OPOriginal);
                        LogMessage("Paramètres SmartText appliqués avec succès.", System.Drawing.Color.Green); // Succès en vert
                    }
                }
            }
            catch (Exception ex)
            {
                // Log et affichage d'une erreur inattendue
                LogMessage($"Erreur inattendue : {ex.Message}", System.Drawing.Color.Red);
            }
        }

        //Bouton build ---------------------------------------------------------------------------------------------------------
        private void button2_Click(object sender, EventArgs e)
        {
            // Vérification de la validité du document courant
            if (currentDoc.DocId == null || currentDoc.DocId == DocumentId.Empty)
            {
                // Log et affichage d'une erreur si le document est invalide ou vide
                LogMessage("Erreur : Le document courant est invalide ou vide.", System.Drawing.Color.Red);
                MessageBox.Show("Erreur : Le document courant est invalide ou vide.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Récupération de l'extension du document
            string ext = Extention(currentDoc.DocId);

            // Configuration des paramètres de dérivation si le document est dérivé
            DerivationConfig(currentDoc.DocId);

            // Tentative de démarrage de la modification du document
            if (!TopSolidHost.Application.StartModification("My Action", false))
            {
                // Log et affichage d'une erreur si la modification ne peut pas démarrer
                LogMessage("Erreur : Impossible de démarrer la modification.", System.Drawing.Color.Red);
                MessageBox.Show("Erreur : Impossible de démarrer la modification.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Mise à jour de l'ID du document courant
            currentDoc.DocId = DocumentCourant();
            DocumentId currentDocID = currentDoc.DocId;

            try
            {
                // Assure que le document est marqué comme modifié
                TopSolidHost.Documents.EnsureIsDirty(ref currentDocID);

                // Initialisation des noms des paramètres à créer
                string nom_docuTxt = "Nom_docu";
                string CommentaireTxt = "Commentaire";
                string DesignationTxt = "Designation";
                string Indice_3DTxt = "Indice 3D";
                string OPIdTxt = "OP";
                string Nom_ElecTxt = "Nom elec";
                string TotalBrutTxt = "Total brut";
                string NomDocuTxt = "Nom_docu";

                // Initialisation des booléens pour suivre la création des paramètres
                bool Indice_3DCreated = false;
                bool DesignationCreated = false;
                bool CommentaireCreated = false;
                bool OPIdCreated = false;
                bool Nom_ElecCreated = false;
                bool TotalBrutCreated = false;
                bool NomDocuCreated = false;

                ElementId ModelingStage = new ElementId();

                // Traitement basé sur l'extension du document
                if (ext != null)
                {
                    ElementId prepaElectrodeStage = new ElementId();

                    if (ext == ".TopEld")
                    {
                        // Si le document est un fichier d'électrode, configure l'étape de préparation
                        prepaElectrodeStage = PrepaElectrodeStage(currentDoc.DocId);

                        try
                        {
                            // Définir l'étape de travail à l'étape de préparation de l'électrode
                            TSH.Operations.SetWorkingStage(prepaElectrodeStage);
                        }
                        catch (Exception ex)
                        {
                            // Log et affichage d'une erreur si l'activation de l'étape échoue
                            LogMessage($"Erreur : L'activation de l'étape préparation du document d'assemblage électrode a échoué : {ex.Message}", System.Drawing.Color.Red);
                            MessageBox.Show("Erreur : L'activation de l'étape préparation du document d'assemblage électrode a échoué" + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        // Pour d'autres types de documents, récupérer et définir l'étape de modélisation
                        try
                        {
                            ModelingStage = TSH.Operations.GetModelingStage(currentDoc.DocId);
                        }
                        catch (Exception ex)
                        {
                            // Log et affichage d'une erreur si la récupération de l'étape échoue
                            LogMessage($"Erreur : La récupération de l'étape modélisation a échoué : {ex.Message}", System.Drawing.Color.Red);
                            MessageBox.Show("Erreur : La récupération de l'étape modélisation a échoué" + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                        try
                        {
                            // Définir l'étape de travail à l'étape de modélisation
                            TSH.Operations.SetWorkingStage(ModelingStage);
                        }
                        catch (Exception ex)
                        {
                            // Log et affichage d'une erreur si l'activation de l'étape échoue
                            LogMessage($"Erreur : L'activation de l'étape modélisation a échoué : {ex.Message}", System.Drawing.Color.Red);
                            MessageBox.Show("Erreur : L'activation de l'étape modélisation a échoué" + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }

                // Recherche des éléments existants dans le document
                ElementId CommentaireExiste = SearchByFriendlyName(currentDoc.DocId, CommentaireTxt);
                ElementId Indice_3DExiste = TSH.Elements.SearchByName(currentDoc.DocId, Indice_3DTxt);
                ElementId Nom_docuExiste = SearchByFriendlyName(currentDoc.DocId, nom_docuTxt);
                ElementId DesignationExiste = TSH.Elements.SearchByName(currentDoc.DocId, DesignationTxt);
                ElementId Nom_ElecExiste = TSH.Elements.SearchByName(currentDoc.DocId, Nom_ElecTxt);
                ElementId TotalBrutExiste = TSH.Elements.SearchByName(currentDoc.DocId, TotalBrutTxt);

                // Vérification si le document est une électrode
                bool iselectrode = Iselectrode(docMaster);

                // Création des paramètres si nécessaire
                if (CommentaireExiste == ElementId.Empty)
                {
                    if (!currentDoc.DocIsElectrode && !iselectrode)
                    {
                        // Création du paramètre "Commentaire"
                        ElementId CommentaireParam = TSH.Parameters.CreateSmartTextParameter(currentDoc.DocId, new SmartText(""));
                        TSH.Elements.SetName(CommentaireParam, CommentaireTxt);
                        CommentaireCreated = true;
                        LogMessage($"Paramètre '{CommentaireTxt}' créé.", System.Drawing.Color.Green);
                    }
                    else
                    {
                        LogMessage($"Paramètre '{CommentaireTxt}' existe déjà.", System.Drawing.Color.Black);
                    }
                }
                else
                {
                    LogMessage($"Paramètre '{CommentaireTxt}' existe déjà.", System.Drawing.Color.Black);
                }

                if (Nom_docuExiste == ElementId.Empty)
                {
                    if (currentDoc.DocIsElectrode || iselectrode)
                    {
                        // Création du paramètre "Nom_docu"
                        ElementId Nom_docuParam = TSH.Parameters.CreateSmartTextParameter(currentDoc.DocId, new SmartText(""));
                        TSH.Elements.SetName(Nom_docuParam, nom_docuTxt);
                        NomDocuCreated = true;
                        LogMessage($"Paramètre '{nom_docuTxt}' créé.", System.Drawing.Color.Green);
                    }
                    else
                    {
                        LogMessage($"Paramètre '{nom_docuTxt}' existe déjà.", System.Drawing.Color.Black);
                    }
                }
                else
                {
                    LogMessage($"Paramètre '{nom_docuTxt}' existe déjà.", System.Drawing.Color.Black);
                }

                if (Indice_3DExiste == ElementId.Empty)
                {
                    // Création du paramètre "Indice 3D"
                    ElementId Indice_3DParam = TSH.Parameters.CreateSmartTextParameter(currentDoc.DocId, new SmartText(""));
                    TSH.Elements.SetName(Indice_3DParam, Indice_3DTxt);
                    Indice_3DCreated = true;
                    LogMessage($"Paramètre '{Indice_3DTxt}' créé.", System.Drawing.Color.Green);
                }
                else
                {
                    LogMessage($"Paramètre '{Indice_3DTxt}' existe déjà.", System.Drawing.Color.Black);
                }

                if (DesignationExiste == ElementId.Empty)
                {
                    // Création du paramètre "Designation"
                    ElementId DesignationParam = TSH.Parameters.CreateSmartTextParameter(currentDoc.DocId, new SmartText(""));
                    TSH.Elements.SetName(DesignationParam, DesignationTxt);
                    DesignationCreated = true;
                    LogMessage($"Paramètre '{DesignationTxt}' créé.", System.Drawing.Color.Green);
                }
                else
                {
                    LogMessage($"Paramètre '{DesignationTxt}' existe déjà.", System.Drawing.Color.Black);
                }

                if (Nom_ElecExiste == ElementId.Empty)
                {
                    if (currentDoc.DocIsElectrode || iselectrode)
                    {
                        // Création du paramètre "Nom_elec"
                        ElementId Nom_ElecParam = TSH.Parameters.CreateSmartTextParameter(currentDoc.DocId, new SmartText(""));
                        TSH.Elements.SetName(Nom_ElecParam, Nom_ElecTxt);
                        Nom_ElecCreated = true;
                        LogMessage($"Paramètre '{Nom_ElecTxt}' créé.", System.Drawing.Color.Green);
                    }
                    else
                    {
                        LogMessage($"Paramètre '{Nom_ElecTxt}' existe déjà.", System.Drawing.Color.Black);
                    }
                }
                else
                {
                    LogMessage($"Paramètre '{Nom_ElecTxt}' existe déjà.", System.Drawing.Color.Black);
                }

                if (TotalBrutExiste == ElementId.Empty)
                {
                    if (currentDoc.DocIsElectrode || iselectrode)
                    {
                        // Création du paramètre "Total brut"
                        ElementId TotalBrutParam = TSH.Parameters.CreateSmartIntegerParameter(currentDoc.DocId, new SmartInteger(0));
                        TSH.Elements.SetName(TotalBrutParam, TotalBrutTxt);
                        TotalBrutCreated = true;
                        LogMessage($"Paramètre '{TotalBrutTxt}' créé.", System.Drawing.Color.Green);
                    }
                    else
                    {
                        LogMessage($"Paramètre '{TotalBrutTxt}' existe déjà.", System.Drawing.Color.Black);
                    }
                }
                else
                {
                    LogMessage($"Paramètre '{TotalBrutTxt}' existe déjà.", System.Drawing.Color.Black);
                }

                // Obtention de la liste des publications pour vérifier l'existence de "OP"
                List<ElementId> PublishingListe = TSH.Entities.GetPublishings(currentDoc.DocId);
                if (PublishingListe != null)
                {
                    bool OPIdExisteOui = false;
                    foreach (ElementId PublishingListeEntities in PublishingListe)
                    {
                        string OPIdExiste = TSH.Elements.GetName(PublishingListeEntities);
                        if (OPIdExiste == OPIdTxt)
                        {
                            OPIdExisteOui = true;
                        }
                    }

                    string DocumentExt = "";
                    try
                    {
                        DocumentExt = Extention(currentDoc.DocId);
                    }
                    catch (Exception ex)
                    {
                        LogMessage($"Erreur : {ex.Message}", System.Drawing.Color.Red);
                        MessageBox.Show("Erreur : " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    if (!OPIdExisteOui)
                    {
                        if (DocumentExt != null)
                        {
                            if (DocumentExt == ".TopNewPrtSet")
                            {
                                bool derivé = TSHD.Tools.IsDerived(currentDoc.DocId);
                                if (derivé)
                                {
                                    // Création du paramètre "OP" pour un document dérivé
                                    ElementId OPParam = TSH.Parameters.CreateSmartTextParameter(currentDoc.DocId, new SmartText(""));
                                    TSH.Elements.SetName(OPParam, OPIdTxt);
                                    OPIdCreated = true;
                                    LogMessage($"Paramètre '{OPIdTxt}' créé.", System.Drawing.Color.Green);
                                }
                                else
                                {
                                    // Publication du paramètre "OP"
                                    ElementId OPParam = TSH.Parameters.PublishText(currentDoc.DocId, OPIdTxt, new SmartText("1"));
                                    TSH.Elements.SetName(OPParam, OPIdTxt);
                                    OPIdCreated = true;
                                    LogMessage($"Paramètre '{OPIdTxt}' créé.", System.Drawing.Color.Green);
                                }
                            }
                            if (DocumentExt == ".TopMillTurn")
                            {
                                ElementId OPExiste = TSH.Elements.SearchByName(currentDoc.DocId, OPIdTxt);
                                if (OPExiste == ElementId.Empty)
                                {
                                    // Création du paramètre "OP" pour un document de type ".TopMillTurn"
                                    ElementId OPParam = TSH.Parameters.CreateSmartTextParameter(currentDoc.DocId, new SmartText(""));
                                    TSH.Elements.SetName(OPParam, OPIdTxt);
                                    OPIdCreated = true;
                                    LogMessage($"Paramètre '{OPIdTxt}' créé.", System.Drawing.Color.Green);
                                }
                                else
                                {
                                    LogMessage($"Paramètre '{OPIdTxt}' existe déjà.", System.Drawing.Color.Black);
                                }
                            }
                        }
                        else
                        {
                            LogMessage($"Paramètre '{OPIdTxt}' existe déjà.", System.Drawing.Color.Black);
                        }
                    }
                    else
                    {
                        LogMessage($"Paramètre '{OPIdTxt}' existe déjà.", System.Drawing.Color.Black);
                    }

                    if (DocumentExt == ".TopNewPrtSet")
                    {
                        Nom_ElecExiste = TSH.Elements.SearchByName(currentDoc.DocId, Nom_ElecTxt);

                        if (currentDoc.DocIsElectrode || iselectrode)
                        {
                            if (Nom_ElecExiste == ElementId.Empty)
                            {
                                // Création du paramètre "Nom_elec" pour un document de type ".TopNewPrtSet"
                                ElementId Nom_ElecParam = TSH.Parameters.CreateSmartTextParameter(currentDoc.DocId, new SmartText(""));
                                TSH.Elements.SetName(Nom_ElecParam, Nom_ElecTxt);
                                Nom_ElecCreated = true;
                                LogMessage($"Paramètre '{Nom_ElecTxt}' créé.", System.Drawing.Color.Green);
                            }
                            else
                            {
                                LogMessage($"Paramètre '{Nom_ElecTxt}' existe déjà.", System.Drawing.Color.Black);
                            }

                            if (Nom_docuExiste == ElementId.Empty)
                            {
                                // Création du paramètre "Nom_docu" pour un document de type ".TopNewPrtSet"
                                ElementId Nom_DocuParam = TSH.Parameters.CreateSmartTextParameter(currentDoc.DocId, new SmartText(""));
                                TSH.Elements.SetName(Nom_DocuParam, NomDocuTxt);
                                NomDocuCreated = true;
                                LogMessage($"Paramètre '{NomDocuTxt}' créé.", System.Drawing.Color.Green);
                            }
                            else
                            {
                                LogMessage($"Paramètre '{NomDocuTxt}' existe déjà.", System.Drawing.Color.Black);
                            }
                        }
                    }
                    else
                    {
                        LogMessage($"Paramètre '{Nom_ElecTxt}' non trouvé, la liste des paramètres est vide", System.Drawing.Color.Black);
                    }
                }

                // Construction du message de confirmation
                string confirmationMessage = "Opérations de création de paramètres :\n";

                // Ajout des résultats de création pour chaque paramètre
                if (Indice_3DCreated)
                {
                    confirmationMessage += $"{Indice_3DTxt} : Créé avec succès.\n";
                    LogMessage($"Paramètre '{Indice_3DTxt}' créé.", System.Drawing.Color.Green);
                }
                else
                {
                    confirmationMessage += $"{Indice_3DTxt} : Non créé, existe déjà.\n";
                    LogMessage($"Paramètre '{Indice_3DTxt}' existe déjà.", System.Drawing.Color.Black);
                }

                if (DesignationCreated)
                {
                    confirmationMessage += $"{DesignationTxt} : Créé avec succès.\n";
                    LogMessage($"Paramètre '{DesignationTxt}' créé.", System.Drawing.Color.Green);
                }
                else
                {
                    confirmationMessage += $"{DesignationTxt} : Non créé, existe déjà.\n";
                    LogMessage($"Paramètre '{DesignationTxt}' existe déjà.", System.Drawing.Color.Black);
                }

                if (CommentaireCreated)
                {
                    confirmationMessage += $"{CommentaireTxt} : Créé avec succès.\n";
                    LogMessage($"Paramètre '{CommentaireTxt}' créé.", System.Drawing.Color.Green);
                }
                else
                {
                    confirmationMessage += $"{CommentaireTxt} : Non créé, existe déjà.\n";
                    LogMessage($"Paramètre '{CommentaireTxt}' existe déjà.", System.Drawing.Color.Black);
                }

                if (OPIdCreated)
                {
                    confirmationMessage += $"{OPIdTxt} : Créé avec succès.\n";
                    LogMessage($"Paramètre '{OPIdTxt}' créé.", System.Drawing.Color.Green);
                }
                else
                {
                    confirmationMessage += $"{OPIdTxt} : Non créé, existe déjà.\n";
                    LogMessage($"Paramètre '{OPIdTxt}' existe déjà.", System.Drawing.Color.Black);
                }

                if (Nom_ElecCreated)
                {
                    confirmationMessage += $"{Nom_ElecTxt} : Créé avec succès.\n";
                    LogMessage($"Paramètre '{Nom_ElecTxt}' créé.", System.Drawing.Color.Green);
                }
                else
                {
                    confirmationMessage += $"{Nom_ElecTxt} : Non créé, existe déjà.\n";
                    LogMessage($"Paramètre '{Nom_ElecTxt}' existe déjà.", System.Drawing.Color.Black);
                }

                if (TotalBrutCreated)
                {
                    confirmationMessage += $"{TotalBrutTxt} : Créé avec succès.\n";
                    LogMessage($"Paramètre '{TotalBrutTxt}' créé.", System.Drawing.Color.Green);
                }
                else
                {
                    confirmationMessage += $"{TotalBrutTxt} : Non créé, existe déjà.\n";
                    LogMessage($"Paramètre '{TotalBrutTxt}' existe déjà.", System.Drawing.Color.Black);
                }

                if (NomDocuCreated)
                {
                    confirmationMessage += $"{NomDocuTxt} : Créé avec succès.\n";
                    LogMessage($"Paramètre '{NomDocuTxt}' créé.", System.Drawing.Color.Green);
                }
                else
                {
                    confirmationMessage += $"{NomDocuTxt} : Non créé, existe déjà.\n";
                    LogMessage($"Paramètre '{NomDocuTxt}' existe déjà.", System.Drawing.Color.Black);
                }

                // Affichage du message de confirmation
                if (Indice_3DCreated || DesignationCreated || CommentaireCreated || OPIdCreated || Nom_ElecCreated || TotalBrutCreated || NomDocuCreated)
                {
                    LogMessage(confirmationMessage, System.Drawing.Color.Blue);
                    MessageBox.Show(confirmationMessage, "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    LogMessage("Aucun paramètre n'a été créé, tous existent déjà.", System.Drawing.Color.Black);
                    MessageBox.Show("Aucun paramètre n'a été créé, tous existent déjà.", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                // Log et affichage d'une erreur en cas d'exception
                LogMessage($"Erreur : {ex.Message}", System.Drawing.Color.Red);
                MessageBox.Show("Erreur : " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TopSolidHost.Application.EndModification(false, false);
            }
            finally
            {
                // Terminer la modification du document
                TopSolidHost.Application.EndModification(true, true);
            }
        }

    }

    internal class Elementid
    {
    }
}
