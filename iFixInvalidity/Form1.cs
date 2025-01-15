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
            DisplayMasterDocumentName();
        }

        bool prepaTrouvé = false;

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
                        LogMessage("Connexion réussie à TopSolid." , System.Drawing.Color.Green);
                        //MessageBox.Show("Connexion réussie à TopSolid.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                LogMessage($"Erreur lors de la connexion à TopSolid : {ex.Message}" , System.Drawing.Color.Red);
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
                        LogMessage("Connexion réussie à TopSolid module design." , System.Drawing.Color.Green);
                        //MessageBox.Show("Connexion réussie à TopSolid module design.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        LogMessage("Connexion échouée à TopSolid module design." , System.Drawing.Color.Red);
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
                DocumentId documentId = TopSolidHost.Documents.EditedDocument;
                LogMessage("Document courant récupéré avec succès.", System.Drawing.Color.Green);
                return documentId;
            }
            catch (Exception ex)
            {
                LogMessage($"Erreur lors de la récupération du document courant : {ex.Message}", System.Drawing.Color.Red);
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
                LogMessage($"Nom du document courant récupéré avec succès : {nomDocument}", System.Drawing.Color.Green);
                return nomDocument;
            }
            catch (Exception ex)
            {
                LogMessage($"Erreur lors de la récupération du nom du document : {ex.Message}", System.Drawing.Color.Red);
                MessageBox.Show($"Erreur lors de la récupération du nom du document : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return "Nom inconnu";
            }
        }





        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //Fonction recurcive pour remonter au document piece original-------------------------------------------------------------------------------------------------------------------
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        private (DocumentId PrepaDocument, DocumentId docMaster) RecupDocuMaster(DocumentId documentCourantId, bool firstPrepaDocumentFound = false)
        {
            DocumentId prepaDocument = DocumentId.Empty;
            DocumentId docMaster = DocumentId.Empty;
            bool documentDerivé = TSHD.Tools.IsDerived(documentCourantId);

            if (!documentDerivé)
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
                //MessageBox.Show("Le document est un document dérivé.");

                DocumentId docBaseDerivation = TSHD.Tools.GetBaseDocument(documentCourantId);
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
        private List<ElementId> OperationsList(DocumentId documentCourantId)
        {
            if (documentCourantId != null && documentCourantId != DocumentId.Empty)
            {
                try
                {
                    List<ElementId> operationsList = TSH.Operations.GetOperations(documentCourantId);
                    LogMessage("Liste des opérations récupérée avec succès.", System.Drawing.Color.Green);
                    return operationsList;
                }
                catch (Exception ex)
                {
                    LogMessage($"Erreur : lors de la récupération de la liste des opérations : {ex.Message}", System.Drawing.Color.Red);
                    MessageBox.Show("Erreur : lors de la récupération de la liste des opérations " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                LogMessage("Document courant non valide.", System.Drawing.Color.Red);
            }
            return new List<ElementId>();
        }

        // Déclaration des noms des paramètres texte à rechercher dans les publications
        string Commentaire = "Commentaire"; // Paramètre commentaire
        string Designation = "Designation";
        string Indice_3D = "Indice 3D";
        string OP = "OP";

        //Recuperation des parametre dans le document maitre
        private void ParametreMaster(in DocumentId docMaster, in DocumentId PrepaDocument, out ElementId indice3D, out ElementId commentaireOriginal, out ElementId designationOriginal, out ElementId OPOriginal)
        {
            if (PrepaDocument != DocumentId.Empty)
            {

                string nomDocPrepa = TSH.Documents.GetName(PrepaDocument);
                MessageBox.Show(nomDocPrepa);
            }

            // Initialisation des paramètres out avec des valeurs par défaut
            indice3D = new ElementId();
            commentaireOriginal = new ElementId();
            designationOriginal = new ElementId();
            OPOriginal = new ElementId();

            List<ElementId> ParameterPubliéList=new List<ElementId>();
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
                        //MessageBox.Show(ParameterName);

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
                                            OPOriginal = ParameterPrepaDocument; // Correction ici
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

        private void SetSmartTxtParameter(DocumentId documentCourantId, SmartText[] SmartTxtTable, ElementId OPOriginal)
        {
            ElementId OPPublishingElement = ElementId.Empty;

            if (!TopSolidHost.Application.StartModification("My Action", false)) return;

            bool CommentaireUpdated = false;
            bool DesignationUpdated = false;
            bool Indice_3DUpdated = false;
            bool OPUpdated = false;
            bool OPDefinitionSet = false; // Variable pour suivre si la méthode a déjà été exécutée

            try
            {
                TopSolidHost.Documents.EnsureIsDirty(ref documentCourantId);

                // Recherche des paramètres publiés dans le document courant
                List<ElementId> ParameterPubliéList = TSH.Parameters.GetParameters(documentCourantId);

                string DocumentExt = String.Empty;
                // Si la liste des paramètres publiés n'est pas vide
                if (ParameterPubliéList.Count > 0)
                {
                    foreach (ElementId ParameterPublié in ParameterPubliéList)
                    {
                        string Parametertype = TSH.Elements.GetTypeFullName(ParameterPublié);
                        // MessageBox.Show(Parametertype);

                        // Récupération du PdmObjectId associé
                        PdmObjectId pdmObjectId = TSH.Documents.GetPdmObject(documentCourantId);

                        // Récupération de l'extension du document
                        TSH.Pdm.GetType(pdmObjectId, out DocumentExt);

                        if (Parametertype == "TopSolid.Kernel.DB.Parameters.TextParameterEntity")
                        {
                            string ParameterPubliéName = TSH.Elements.GetFriendlyName(ParameterPublié);

                            if (ParameterPubliéName == Commentaire)
                            {
                                ElementId ParameterPubliéOp = TSH.Elements.GetParent(ParameterPublié);
                                try
                                {
                                    TSH.Parameters.SetSmartTextParameterCreation(ParameterPubliéOp, SmartTxtTable[1]);
                                    CommentaireUpdated = true;// Marquer la méthode comme exécutée
                                    LogMessage($"Paramètre '{Commentaire}' mis à jour.", System.Drawing.Color.Green);
                                }
                                catch (Exception ex)
                                {
                                    LogMessage($"Erreur : lors de la mise à jour du paramètre 'Commentaire' : {ex.Message}", System.Drawing.Color.Red);
                                    throw; // Relancer l'exception pour s'assurer que le bloc finally est exécuté
                                }
                            }
                            if (ParameterPubliéName == Designation)
                            {
                                ElementId ParameterPubliéOp = TSH.Elements.GetParent(ParameterPublié);
                                try
                                {
                                    TSH.Parameters.SetSmartTextParameterCreation(ParameterPubliéOp, SmartTxtTable[2]);
                                    DesignationUpdated = true;// Marquer la méthode comme exécutée
                                    LogMessage($"Paramètre '{Designation}' mis à jour.", System.Drawing.Color.Green);
                                }
                                catch (Exception ex)
                                {
                                    LogMessage($"Erreur : lors de la mise à jour du paramètre 'Designation' : {ex.Message}", System.Drawing.Color.Red);
                                    throw; // Relancer l'exception pour s'assurer que le bloc finally est exécuté
                                }
                            }
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
                            }
                            else if (DocumentExt == ".TopNewPrtSet")
                            {
                                bool dérivé = TSHD.Tools.IsDerived(documentCourantId);
                                if (dérivé)
                                {
                                    ElementId opCourantPrepa = RecupOpCourantPrepa(documentCourantId);

                                    DocumentId docBase = TSHD.Tools.GetBaseDocument(documentCourantId);
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

                                            LogMessage($"Paramètre '{OP}' trouvé.", System.Drawing.Color.Green);
                                            //opOriginalFound = true; // Mettre le drapeau à true
                                            dérivé = false;
                                            break; // Sortir de la boucle interne
                                        }
                     
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
                                              $"{OP} : {(OPUpdated ? "Oui" : "Non")}";

                if (DocumentExt == ".TopMillTurn" || CommentaireUpdated == false || DesignationUpdated == false || Indice_3DUpdated == false || OPUpdated == false)
                {
                    MessageBox.Show(confirmationMessage + "\n\n Attention, certains paramètres sont peut-être manquants.\n Verifiez les documents parents.", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                TopSolidHost.Application.EndModification(true, true);
            }
        }

        private ElementId RecupOpCourantPrepa(DocumentId documentCourantId)
        {
            List<ElementId> parameters = TSH.Operations.GetOperations(documentCourantId);
            foreach (ElementId parameter in parameters)
            {
                string parameterName = TSH.Elements.GetFriendlyName(parameter);
                if (parameterName == "Paramètre texte (OP)")
                {
                    return parameter;
                }
            }
            throw new InvalidOperationException("Aucun paramètre 'OP' trouvé.");

        }

        private void LogMessage(string message, System.Drawing.Color color)
        {
            // Ajouter le message à la fin du texte actuel
            richTextBox1.SelectionStart = richTextBox1.TextLength;
            richTextBox1.SelectionLength = 0;

            // Définir la couleur du texte
            richTextBox1.SelectionColor = color;

            // Ajouter le message
            richTextBox1.AppendText(message + Environment.NewLine);

            // Réinitialiser la couleur du texte à la couleur par défaut
            richTextBox1.SelectionColor = richTextBox1.ForeColor;

            // Faire défiler le RichTextBox pour que le curseur soit visible
            richTextBox1.ScrollToCaret();
        }

        private SmartText CreateSmartTxt(ElementId smartTxt)
        {
            SmartText SmartTxtId = new SmartText(smartTxt);
            LogMessage($"SmartText créé pour l'élément : {smartTxt}", System.Drawing.Color.Green);
            return SmartTxtId;
        }

        private void DerivationConfig(DocumentId documentCourantId)
        {
            if (documentCourantId != null) 
            {
                bool isDerived = TSHD.Tools.IsDerived(documentCourantId);
                if (isDerived)
                {
                    if (TopSolidHost.Application.StartModification("My Modification", false))
                    {
                        try
                        {
                            TSHD.Tools.SetDerivationInheritances
                            (
                                documentCourantId,//DocumentId inDocumentId
                                false,//bool inName, 
                                false,//bool inDescription,
                                false,//bool inCode,
                                false,//bool inPartNumber,
                                false,//bool inComplementaryPartNumber,
                                false,//bool inManufacturer,
                                false,//bool inManufacturerPartNumber,
                                false,//bool inComment,
                                new List<ElementId>(),//List < ElementId > inOtherSystemParameters,
                                false,//bool inNonSystemParameters,  
                                true,//bool inPoints,    
                                true,//bool inAxes,   
                                true,//bool inPlanes,  
                                true,//bool inFrames,   
                                true,// bool inSketches, 
                                true,//bool inShapes,
                                true,//bool inPublishings,
                                true,// bool inFunctions,  
                                true,//bool inSymmetries,   
                                true,//bool inUnsectionabilities,   
                                false,//bool inRepresentations,   
                                false,//bool inSets,
                                true//bool inCameras
                            );

                            TopSolidHost.Application.EndModification(true, true);
                        }
                        catch(Exception ex)
                        {
                            MessageBox.Show("Erreur : " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            TopSolidHost.Application.EndModification(false, false);
                        }
                    }

                }
            }

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
                    LogMessage($"Document courant : {nomDocumentCourant}", System.Drawing.Color.Green);
                    //MessageBox.Show($"Document courant : {nomDocumentCourant}", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    LogMessage("Aucun document courant trouvé.", System.Drawing.Color.Red);
                    //MessageBox.Show("Aucun document courant trouvé.", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Une erreur s'est produite : {ex.Message}", System.Drawing.Color.Red);
                MessageBox.Show($"Une erreur s'est produite : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // Fonction recup docu master
            var (prepaDocument, docMaster) = RecupDocuMaster(documentCourantId);

            //Configuration parametre de derivation si c'est un document derivé
            DerivationConfig(documentCourantId);

            if (docMaster != DocumentId.Empty)
            {
                string docMasterName = TSH.Documents.GetName(docMaster);
                LogMessage($"Document maître trouvé : {docMasterName}", System.Drawing.Color.Green);
                //MessageBox.Show("Document pièce trouvé : " + docMasterName);
            }
            else
            {
                LogMessage("Aucun document maître trouvé.", System.Drawing.Color.Red);
                //MessageBox.Show("Aucun document maître trouvé.", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            if (prepaDocument != DocumentId.Empty)
            {
                string docPrepaName = TSH.Documents.GetName(prepaDocument);
                LogMessage($"Document de prépa trouvé : {docPrepaName}", System.Drawing.Color.Green);
                //MessageBox.Show("Document de prépa trouvé : " + docPrepaName);
            }
            else
            {
                LogMessage("Aucun document de prépa trouvé.", System.Drawing.Color.Red);
                MessageBox.Show("Aucun document de prépa trouvé.", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }


            // Récupération des paramètres de base originaux depuis l'opération d'inclusion
            // Déclaration des variables pour recevoir les valeurs de retour
            ElementId indice3D;
            ElementId commentaireOriginal;
            ElementId designationOriginal;
            ElementId OPOriginal;

            // Récupération des paramètres maître
            ParametreMaster(in docMaster, in prepaDocument, out indice3D, out commentaireOriginal, out designationOriginal , out OPOriginal);

            SmartText SmartTxtCommentaireId = CreateSmartTxt(commentaireOriginal);
            SmartText SmartTxtDesignationId = CreateSmartTxt(designationOriginal);
            SmartText SmartTxtIndice_3DId = CreateSmartTxt(indice3D);
            SmartText SmartTxtIndice_OPId = CreateSmartTxt(OPOriginal);

            // Déclarer et initialiser un tableau de SmartText
            SmartText[] SmartTxtTable = new SmartText[4];

            

            SmartTxtTable[1] = SmartTxtCommentaireId; // Index 1
            SmartTxtTable[2] = SmartTxtDesignationId; // Index 2
            SmartTxtTable[3] = SmartTxtIndice_3DId; // Index 3
            SmartTxtTable[0] = SmartTxtIndice_OPId; // Index 0

            SetSmartTxtParameter(documentCourantId, SmartTxtTable, OPOriginal);

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

        //Fonction qui recupere extention du document
        private string Extention (DocumentId documentCourantId)
        {
            // Récupération du PdmObjectId associé
            PdmObjectId pdmObjectId = TSH.Documents.GetPdmObject(documentCourantId);

            // Récupération de l'extension du document
            TSH.Pdm.GetType(pdmObjectId, out string DocumentExt);
            return DocumentExt;
        }

        //Bouton build ---------------------------------------------------------------------------------------------------------
        private void button2_Click(object sender, EventArgs e)
        {
            DocumentId documentCourantId = DocumentCourant();



            // Vérification de documentCourantId
            if (documentCourantId == null || documentCourantId == DocumentId.Empty)
            {
                LogMessage("Erreur : Le document courant est invalide ou vide.", System.Drawing.Color.Red);
                MessageBox.Show("Erreur : Le document courant est invalide ou vide.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //Configuration parametre de derivation si c'est un document derivé
            DerivationConfig(documentCourantId);

            if (!TopSolidHost.Application.StartModification("My Action", false))
            {
                LogMessage("Erreur : Impossible de démarrer la modification.", System.Drawing.Color.Red);
                MessageBox.Show("Erreur : Impossible de démarrer la modification.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            

            try
            {
                TopSolidHost.Documents.EnsureIsDirty(ref documentCourantId);

                string CommentaireTxt = "Commentaire";
                string DesignationTxt = "Designation";
                string Indice_3DTxt = "Indice 3D";
                string OPIdTxt = "OP";

                bool Indice_3DCreated = false;
                bool DesignationCreated = false;
                bool CommentaireCreated = false;
                bool OPIdCreated = false;
                ElementId ModelingStage = new ElementId();

                //Recup etape modelisation
                try
                {
                 ModelingStage = TSH.Operations.GetModelingStage(documentCourantId);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erreur : La recuperation de l'étape modélisation a échoué" + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                //passe à l'etape modelisation
                try
                {
                 TSH.Operations.SetWorkingStage(ModelingStage);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erreur : L'activation de l'étape modélisation a échoué" + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                //cherche un element appelé Indice 3D
                ElementId Indice_3DExiste = TSH.Elements.SearchByName(documentCourantId, Indice_3DTxt);
                if (Indice_3DExiste == ElementId.Empty)
                {
                    //Creation parametre smart texte puis renommange en Indice 3D
                    ElementId Indice_3DParam = TSH.Parameters.CreateSmartTextParameter(documentCourantId, new SmartText(""));
                    TSH.Elements.SetName(Indice_3DParam, Indice_3DTxt);
                    Indice_3DCreated = true;
                    LogMessage($"Paramètre '{Indice_3DTxt}' créé.", System.Drawing.Color.Green);
                }
                else
                {
                    LogMessage($"Paramètre '{Indice_3DTxt}' existe déjà.", System.Drawing.Color.Black);
                }
                //cherche un element appelé Designation
                ElementId DesignationExiste = TSH.Elements.SearchByName(documentCourantId, DesignationTxt);
                if (DesignationExiste == ElementId.Empty)
                {
                    //Creation parametre smart texte puis renommage en Designation
                    ElementId DesignationParam = TSH.Parameters.CreateSmartTextParameter(documentCourantId, new SmartText(""));
                    TSH.Elements.SetName(DesignationParam, DesignationTxt);
                    DesignationCreated = true; //bool le Designation a bien été créé
                    LogMessage($"Paramètre '{DesignationTxt}' créé.", System.Drawing.Color.Black);
                }
                else
                {
                    LogMessage($"Paramètre '{DesignationTxt}' existe déjà.", System.Drawing.Color.Black);
                }
                //cherche un element appelé Commentaire
                ElementId CommentaireExiste = TSH.Elements.SearchByName(documentCourantId, CommentaireTxt);
                if (CommentaireExiste == ElementId.Empty)
                {
                    //Creation parametre smart texte puis renommage en Commentaire
                    ElementId CommentaireParam = TSH.Parameters.CreateSmartTextParameter(documentCourantId, new SmartText(""));
                    TSH.Elements.SetName(CommentaireParam, CommentaireTxt);
                    CommentaireCreated = true; //bool le commentaire a bien été créé
                    LogMessage($"Paramètre '{CommentaireTxt}' créé.", System.Drawing.Color.Black);
                }
                else
                {
                    LogMessage($"Paramètre '{CommentaireTxt}' existe déjà.", System.Drawing.Color.Black);
                }
                //Obtient la liste des publications pour chercher OP 
                List<ElementId> PublishingListe = TSH.Entities.GetPublishings(documentCourantId);
                if (PublishingListe != null)
                {
                    bool OPIdExisteOui = false; //Initialisation bool existe
                    foreach (ElementId PublishingListeEntities in PublishingListe)
                    {
                        string OPIdExiste = TSH.Elements.GetName(PublishingListeEntities);//obtient le nom des publications
                        //MessageBox.Show(OPIdExiste);

                        if (OPIdExiste == OPIdTxt)
                        {
                            OPIdExisteOui = true; //OP existe
                        }

                    }

                    if (OPIdExisteOui != true)
                    {
                        //Recupere extention du document
                        string DocumentExt = Extention(documentCourantId);

                        //

                        if (DocumentExt != null)
                        {
                            if (DocumentExt == ".TopNewPrtSet")
                            {
                                bool derivé = TSHD.Tools.IsDerived(documentCourantId);
                                if (derivé)
                                {
                                    //Creation parametre smart texte puis renommage en OP
                                    ElementId OPParam = TSH.Parameters.CreateSmartTextParameter(documentCourantId, new SmartText(""));
                                    TSH.Elements.SetName(OPParam, OPIdTxt);
                                    OPIdCreated = true; //bool le commentaire a bien été créé
                                    LogMessage($"Paramètre '{OPIdTxt}' créé.", System.Drawing.Color.Black);
                                    return;
                                }
                                else 
                                { 
                                    ElementId OPParam = TSH.Parameters.PublishText(documentCourantId, OPIdTxt, new SmartText("1"));
                                    TSH.Elements.SetName(OPParam, OPIdTxt);
                                    OPIdCreated = true;
                                    LogMessage($"Paramètre '{OPIdTxt}' créé.", System.Drawing.Color.Black);
                                    return;
                                }
                            }
                            if (DocumentExt == ".TopMillTurn")
                            {
                                ElementId OPExiste = TSH.Elements.SearchByName(documentCourantId, OPIdTxt);
                                if (OPExiste == ElementId.Empty)
                                {
                                    //Creation parametre smart texte puis renommage en OP
                                    ElementId OPParam = TSH.Parameters.CreateSmartTextParameter(documentCourantId, new SmartText(""));
                                    TSH.Elements.SetName(OPParam, OPIdTxt);
                                    OPIdCreated = true; //bool le commentaire a bien été créé
                                    LogMessage($"Paramètre '{OPIdTxt}' créé.", System.Drawing.Color.Black);
                                }
                                else
                                {
                                    LogMessage($"Paramètre '{CommentaireTxt}' existe déjà.", System.Drawing.Color.Black);
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
                        LogMessage($"Paramètre '{OPIdTxt}' non trouvé, la liste des parametres est vide", System.Drawing.Color.Black);
                    }

                    // Construction du message de confirmation
                    string confirmationMessage = $"Paramètres créés :\n" +
                                                  $"{Indice_3DTxt} : {(Indice_3DCreated ? "Oui" : "Non!")}\n" +
                                                  $"{DesignationTxt} : {(DesignationCreated ? "Oui" : "Non!")}\n" +
                                                  $"{CommentaireTxt} : {(CommentaireCreated ? "Oui" : "Non!")}\n" +
                                                  $"{OPIdTxt} : {(OPIdCreated ? "Oui" : "Non!")}";

                    if (Indice_3DCreated == false || DesignationCreated == false || CommentaireCreated == false || OPIdCreated == false)
                    {
                        MessageBox.Show(confirmationMessage + "\n\n certains paramètres existent deja", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        LogMessage(confirmationMessage, System.Drawing.Color.Black);
                        //MessageBox.Show(confirmationMessage, "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                }
            }
            catch (Exception ex)
            {
                LogMessage($"Erreur : {ex.Message}", System.Drawing.Color.Red);
                MessageBox.Show("Erreur : " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TopSolidHost.Application.EndModification(false, false);
            }
            finally
            {
                TopSolidHost.Application.EndModification(true, true);
            }
        }

        //fonction pour afficher le nom du document courant
        private DocumentId DisplayDocumentName()
        {
            DocumentId documentCourantId = DocumentCourant();
           string documentCourantName = NomDocumentCourant(documentCourantId);
          
            labelDocumentName.Text = documentCourantName;
            return documentCourantId;
        }

        //fonction pour afficher le nom du document maitre
        private bool DisplayMasterDocumentName()
        {
            //bool prepaTrouvé = false; // Déclaration de la variable en dehors des blocs if et else

            try
            {
                DocumentId documentCourantId = DocumentCourant();
                string documentCourantName = NomDocumentCourant(documentCourantId);

                // Appel de la fonction et déstructuration du tuple
                var (prepaDocument, docMaster) = RecupDocuMaster(documentCourantId);

                // Vérification du document maître
                if (docMaster != DocumentId.Empty)
                {
                    string documentMasterName = TSH.Documents.GetName(docMaster);
                    labelDocumentMasterName.Text = documentMasterName;
                    LogMessage($"Document maître trouvé : {documentMasterName}", System.Drawing.Color.Green);
                }
                else
                {
                    labelDocumentMasterName.Text = "Aucun document maître trouvé.";
                    LogMessage("Aucun document maître trouvé.", System.Drawing.Color.Red);
                    //MessageBox.Show("Aucun document maître trouvé.", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                // Vérification du document de prépa
                if (prepaDocument != DocumentId.Empty)
                {
                    string docPrepaName = TSH.Documents.GetName(prepaDocument);
                    LogMessage($"Document de prépa trouvé : {docPrepaName}", System.Drawing.Color.Green);
                    prepaTrouvé = true;
                }
                else
                {
                    LogMessage("Aucun document de prépa trouvé.", System.Drawing.Color.Red);
                    //MessageBox.Show("Aucun document de prépa trouvé.", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    prepaTrouvé = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur : lors de l'affichage des noms des documents " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return prepaTrouvé; // Retour de la valeur de prepaTrouvé
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

        //Bouton invok
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                //// Lancer la commande
                //// Créer une instance de UserQuestion
                //var userQuestion = new UserQuestion
                //{
                //    AcceptsAuto = true // Activer la réponse automatique
                //};
                UserQuestion userQuestion = new UserQuestion();
                bool okAuto = true;
                okAuto = userQuestion.AcceptsAuto;
                TSH.Application.InvokeCommand("TopSolid.Cad.Design.UI.Publishings.PublishingsInheritingCommand");

                //// Réinitialiser AcceptsAuto après l'exécution si nécessaire
                //userQuestion.AcceptsAuto = false;
            }
            catch (Exception ex)
            {
                // Gérer les exceptions
                MessageBox.Show("Erreur lors de l'exécution de la commande : " + ex.Message);
            }
        }
    }

    internal class Elementid
    {
    }
}
