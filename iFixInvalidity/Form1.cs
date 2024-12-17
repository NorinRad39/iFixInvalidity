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
                        MessageBox.Show("Connexion réussie à TopSolid.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Connexion échouée à TopSolid.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("TopSolid est déjà connecté.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (InvalidOperationException ex)
            {
                // Gérer une exception spécifique si nécessaire
                MessageBox.Show($"Problème opérationnel : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                // Gérer d'autres exceptions
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
                        MessageBox.Show("Connexion réussie à TopSolid module design.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Connexion échouée à TopSolid module design.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("TopSolid module design est déjà connecté.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (InvalidOperationException ex)
            {
                // Gérer une exception spécifique si nécessaire
                MessageBox.Show($"Problème opérationnel : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                // Gérer d'autres exceptions
                MessageBox.Show($"Erreur lors de la connexion à TopSolid module design : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    //Liste les operations du document
                    List<ElementId> operationsList = OperationsList(documentCourantId);
                    if (operationsList.Count != 0)
                    {
                        //Pour chaque elements de la liste
                        foreach (ElementId operation in operationsList)
                        {
                            //Si c'est une inclusion
                            if (TSHD.Assemblies.IsInclusion(operation))
                            {
                                //Recuperation de l'element de l'inclusion
                                ElementId insertedElementId = TSHD.Assemblies.GetInclusionChildOccurrence(operation);
                                //Recuperation de l'instance du document associé
                                DocumentId instanceDocument = TSHD.Assemblies.GetOccurrenceDefinition(insertedElementId);

                                // Récupération du PdmObjectId associé
                                PdmObjectId pdmObjectId = TSH.Documents.GetPdmObject(instanceDocument);

                                // Récupération de l'extension du document
                                TSH.Pdm.GetType(pdmObjectId, out String DocumentExt);
                                //Si c'est un fichier piece, renvoyer le document ID
                                if (DocumentExt == ".TopPrt")
                                {
                                    //DocumentId documentDansPrepa = RecupDocuMaster(instanceDocument);
                                    if (instanceDocument != DocumentId.Empty)
                                    {
                                        return instanceDocument;
                                    }
                                }
                                //Si c'est un document de prepa, relancer la fonction
                                if (DocumentExt == ".TopNewPrtSet")
                                {
                                    DocumentId result = RecupDocuMaster(instanceDocument);
                                    if (result != DocumentId.Empty)
                                    {
                                        return result;
                                    }
                                }
                                
                            }
                        }
                    }
                    return DocumentId.Empty;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erreur : " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    // Retourner DocumentId.Empty en cas d'exception
                    return DocumentId.Empty;
                }
            }
            // Retourner DocumentId.Empty si documentCourantId est null ou empty
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
                    return operationsList;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erreur : lors de la récupération de la liste des opérations " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            return new List<ElementId>();
        }

        //Déclaration des noms des parametres texte a rechercher dans les publications
        string Commentaire = "Commentaire"; //Parametre commentaire 
        string Designation = "Designation";
        string Indice_3D = "Indice 3D";

        private void ParametreMaster(in DocumentId docMaster, out ElementId indice3D, out ElementId commentaireOriginal, out ElementId designationOriginal)
        {
            // Initialisation des paramètres out avec des valeurs par défaut
            indice3D = new ElementId();
            commentaireOriginal = new ElementId();
            designationOriginal = new ElementId();

            //Recherche des parametre publié dans le document maitre
            List<ElementId>ParameterPubliéList = TSH.Entities.GetPublishings(docMaster);

            //Si la liste de parametre publié n'est pas vide
            if (ParameterPubliéList.Count > 0)
            {
                try
                {
                    //Pour chaque parametre publié
                    foreach (ElementId Parameter in ParameterPubliéList)
                    {
                        //Récuperation du nom de chaque parametre publié pour comparaison
                        string ParameterName = TSH.Elements.GetFriendlyName(Parameter);
                        MessageBox.Show(ParameterName);

                        //Si le nom du parametre est egale au nom attendu renvoyé l'element ID
                        if (ParameterName == Indice_3D)
                        {
                            indice3D = Parameter;
                        }
                        if (ParameterName == Commentaire)
                        {
                            commentaireOriginal = Parameter;
                        }
                        if (ParameterName == Designation)
                        {
                            designationOriginal = Parameter;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erreur : lors de la recuperation des parametre maitre" + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else 
            {
                MessageBox.Show("Erreur : La liste des parametre publié est vide dans le document maitre");
                indice3D = new ElementId();
                commentaireOriginal = new ElementId();
                designationOriginal = new ElementId();
            }
        }

        //Recuperation des parametres a modifier.
        private void SetSmartTxtParameter(DocumentId documentCourantId, SmartText[] SmartTxtTable)
        {
            if (!TopSolidHost.Application.StartModification("My Action", false)) return;

            try
            {
                TopSolidHost.Documents.EnsureIsDirty(ref documentCourantId);

                //Recherche des parametre publié dans le document courant
                List<ElementId> ParameterPubliéList = TSH.Parameters.GetParameters(documentCourantId);

                //Si la liste des parametres publié n'est pas vide
                if (ParameterPubliéList.Count > 0)
                {
                    foreach(ElementId ParameterPublié in ParameterPubliéList)
                    {
                        string Parametertype = TSH.Elements.GetTypeFullName(ParameterPublié);
                        //MessageBox.Show(Parametertype);

                        if (Parametertype == "TopSolid.Kernel.DB.Parameters.TextParameterEntity")
                        {
                            string ParameterPubliéName = TSH.Elements.GetFriendlyName(ParameterPublié);

                            if (ParameterPubliéName == Commentaire)
                            {
                                ElementId ParameterPubliéOp = TSH.Elements.GetParent(ParameterPublié);
                                TSH.Parameters.SetSmartTextParameterCreation(ParameterPubliéOp, SmartTxtTable[1]);
                            }
                            if (ParameterPubliéName == Designation)
                            {
                                ElementId ParameterPubliéOp = TSH.Elements.GetParent(ParameterPublié);
                                TSH.Parameters.SetSmartTextParameterCreation(ParameterPubliéOp, SmartTxtTable[2]);
                            }
                            if (ParameterPubliéName == Indice_3D)
                            {
                                ElementId ParameterPubliéOp = TSH.Elements.GetParent(ParameterPublié);
                                TSH.Parameters.SetSmartTextParameterCreation(ParameterPubliéOp, SmartTxtTable[3]);
                            }

                        }
                    }
                }
                else
                {

                MessageBox.Show("Erreur : La liste des parametres courant est vide");
                ElementId indice3DCourant = new ElementId();
                ElementId commentaireCourant = new ElementId();
                ElementId designationCourant = new ElementId();
                }

                TopSolidHost.Application.EndModification(true, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur : lors de la recuperation des parametres courant" + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TopSolidHost.Application.EndModification(false, false);
            }

        }

        private SmartText CreateSmartTxt(ElementId smartTxt)
            {
            SmartText SmartTxtId = new SmartText(smartTxt);
            return SmartTxtId;
            }

        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // Gestionnaire de clic pour le bouton.-----------------------------------------------------------------------------------------------------------------------------------------
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        private void button1_Click_1(object sender, EventArgs e)
        {
            //Déclaration variable docu courant
            DocumentId documentCourantId = new DocumentId();

            //Recupération docu courant
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

            //Fonction recup docu master
            DocumentId docMaster = RecupDocuMaster(documentCourantId);

            String docMasterName = TSH.Documents.GetName(docMaster);
            MessageBox.Show("Document piece trouvé : " + docMasterName);

            //Recuperation des parametres de base originaux depuis l'operation d'inclusion
            //Déclaration des variables pour recevoir les valeurs de retour
            ElementId indice3D;
            ElementId nomOriginal;
            ElementId commentaireOriginal;
            ElementId designationOriginal;



            //Recuperation des parametres master
            ParametreMaster(in docMaster, out indice3D, out commentaireOriginal, out designationOriginal);

            SmartText SmartTxtCommentaireId = CreateSmartTxt(commentaireOriginal);
            SmartText SmartTxtDesignationId = CreateSmartTxt(designationOriginal);
            SmartText SmartTxtIndice_3DId = CreateSmartTxt(indice3D);

            // Déclarer et initialiser un tableau de SmartText
            SmartText[] SmartTxtTable = new SmartText[4];

            //SmartTxtTable[0] = SmartTxtEmpty;

            SmartTxtTable[1] = SmartTxtCommentaireId; //Index 1

            SmartTxtTable[2] = SmartTxtDesignationId;//Index 2

            SmartTxtTable[3] = SmartTxtIndice_3DId; //Index 3

            SetSmartTxtParameter(documentCourantId, SmartTxtTable);






            TSH.Disconnect();
            Environment.Exit(0);


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
            DocumentId documentCourantId = new DocumentId();

            if (!TopSolidHost.Application.StartModification("My Action", false)) return;
            try
            {
                documentCourantId = DocumentCourant();
                if (documentCourantId != DocumentId.Empty)
                {
                    string CommentaireTxt = "Commentaire";
                    string DesignationTxt = "Designation";
                    string Indice_3DTxt = "Indice 3D";


                    ElementId Indice_3DExiste = TSH.Elements.SearchByName(documentCourantId, Indice_3DTxt);
                    if (Indice_3DExiste == ElementId.Empty)
                    {
                        ElementId Indice_3DParam = TSH.Parameters.CreateTextParameter(documentCourantId, "");
                        TSH.Elements.SetName(Indice_3DParam, Indice_3DTxt);
                    }

                    ElementId DesignationExiste = TSH.Elements.SearchByName(documentCourantId, DesignationTxt);
                    if (DesignationExiste == ElementId.Empty)
                    {
                        ElementId DesignationParam = TSH.Parameters.CreateTextParameter(documentCourantId, "");
                        TSH.Elements.SetName(DesignationParam, DesignationTxt);
                    }

                    ElementId CommentaireExiste = TSH.Elements.SearchByName(documentCourantId, CommentaireTxt);
                    if (CommentaireExiste == ElementId.Empty)
                    {
                        ElementId CommentaireParam = TSH.Parameters.CreateTextParameter(documentCourantId, "");
                        TSH.Elements.SetName(CommentaireParam, CommentaireTxt);

                    }
                }
                else
                {
                    MessageBox.Show("Pas de document courant");

                }
            }
            catch
            {
                // End modification (failure).
                TopSolidHost.Application.EndModification(false, false);
                TSH.Disconnect();
                return;
            }
            finally
            {
                // End modification (success).
                TopSolidHost.Application.EndModification(true, true);
                TSH.Disconnect();
            }
        }

    }

    internal class Elementid
    {
    }
}
