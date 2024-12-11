﻿using System;
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
using System.Security;

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




        //récupere element id du dossier publication
        private ElementId PublishFolder(DocumentId documentCourantId)
        {

            if (documentCourantId != null)
            {

                try
                {
                    ElementId PublishFolderId = TSH.Elements.SearchByName(documentCourantId, "$TopSolid.Kernel.DB.Documents.ElementName.Publishings");
                    return PublishFolderId;

                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Erreur lors de la récupération dossiers publication : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return ElementId.Empty;

                }

            }
            else
            {
                MessageBox.Show("Le dossier Publications n'existe pas ou est vide");
                return new ElementId();
            }

        }


        //recupere le contenu du dossier publication
        private ElementId PublicationsId (ElementId PublishFolderId, String NomParametreTxt)
        {

            if (PublishFolderId != null)
            {
                try
                {
                    List<ElementId> ElementsPubliedIds = TSH.Elements.GetConstituents(PublishFolderId);

                    if (ElementsPubliedIds != null)
                    {


                            foreach (ElementId ElementsPubliedId in ElementsPubliedIds)
                            {
                                string NameElementsPubliedId = TSH.Elements.GetName(ElementsPubliedId);

                                if (NameElementsPubliedId == NomParametreTxt)
                                {
                                    
                                   return ElementsPubliedId;
                                }
                               
                            }

                                return ElementId.Empty;
                    }
                    else 
                    {
                        return ElementId.Empty;
                        
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Erreur lors de la récupération du contenu du dossier publication : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return ElementId.Empty;
                }
            }
            else
            {
                MessageBox.Show("Le dossier Publications n'existe pas ou est vide");
                return ElementId.Empty;
            }

        }


        //creer une liste du contenu du dossier publication
        private List<ElementId> PublishedElements(ElementId PublishFolderId)
        {
            if (PublishFolderId != null)
            {
                try
                {
                    List<ElementId> ElementsPubliedIds = TSH.Elements.GetConstituents(PublishFolderId);

                    return ElementsPubliedIds;

                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Erreur lors de la récupération du contenu du dossier publication : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return new List<ElementId>();
                }
            }
            else 
            {
                MessageBox.Show("Le dossier Publications n'existe pas ou est vide");
                return new List<ElementId>();          
            }
  
        }


        //récupere le nom des parametres du dossier publication
        private List<string> NamesPublishedElements(List<ElementId> ElementsPubliedIds)
        {

            List<string> namesPublishedElementsTxts = new List<string>();

            try
            {
                foreach (ElementId ElementsPubliedId in ElementsPubliedIds)
                {

                    string namesPublishedElementsTxt = TSH.Elements.GetName(ElementsPubliedId);
                    namesPublishedElementsTxts.Add(namesPublishedElementsTxt);
                }
            }
            catch (Exception ex) 
            {
                MessageBox.Show($"Erreur lors de la récupération des noms des éléments publiés : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error); 
            }

            return namesPublishedElementsTxts;


        }

        //récupere dossier parametre
        private ElementId ParametresFolder(DocumentId documentCourantId)
        {

            if (documentCourantId != null)
            {

                try
                {
                    ElementId ParametresFolderId = TSH.Elements.SearchByName(documentCourantId, "$TopSolid.Kernel.DB.Documents.ElementName.SystemParameters");
                    return ParametresFolderId;

                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Erreur lors de la récupération dossiers parametre : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return ElementId.Empty;

                }

            }
            else
            {
                MessageBox.Show("Le dossier Parametre n'existe pas ou est vide");
                return new ElementId();
            }

        }


        //récupere une liste du contenu du dossier parametre
        private List<ElementId> ParametresElements(ElementId ParametreFolderId)
        {
            if (ParametreFolderId != null)
            {
                try
                {
                    List<ElementId> ParametresdIds = TSH.Elements.GetConstituents(ParametreFolderId);

                    return ParametresdIds;

                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Erreur lors de la récupération du contenu du dossier parametre : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return new List<ElementId>();
                }
            }
            else
            {
                MessageBox.Show("Le dossier parametre n'existe pas ou est vide");
                return new List<ElementId>();
            }

        }

        
        
        
       
       //verifie si les parametre du dossier parametre sont invalide
        
        private ElementId ParametreIsInValide (List<ElementId> ParametresdIds)
        {
                bool IsInvalide = false;
            if (ParametresdIds != null)
            {
                foreach (ElementId ParametresdId in ParametresdIds)
                {
                    
                    IsInvalide = TSH.Elements.IsInvalid(ParametresdId);

                    if (IsInvalide)
                    {
                        return ParametresdId;
                    }
                }
                    return ElementId.Empty;

            }
            return ElementId.Empty;
        }

        
        //verifie si les operations de l'arbre des operation sont valide
        public ElementId ParamInvalide (DocumentId documentCourantId)
        {

            List<ElementId> OperationsId = TSH.Operations.GetOperations(documentCourantId);



            foreach (ElementId OperationId in OperationsId)
            {
                string namesOperationsTxt = TSH.Elements.GetName(OperationId);
                
                MessageBox.Show(namesOperationsTxt);

                bool IsInvalide = TSH.Elements.IsInvalid(OperationId);
                if (IsInvalide)
                {

                    return OperationId;
                }
               
            }
            return ElementId.Empty;

        }
   
        //recupere le nom des enfant d'une operation et l'affiche dans une txtbox


    private void NomElementOperation(DocumentId documentCourantId)
        {
            try
            {

                List<ElementId> ElementsId = TSH.Operations.GetOperations(documentCourantId);

                if(ElementsId != null)
                {


                    foreach(ElementId ElementId in ElementsId)
                    {
                        List<ElementId> ElementIdNames = TSH.Operations.GetChildren(ElementId);

                        foreach (ElementId ElementIdName in ElementIdNames)
                        {
                            string ElementIdNamTxt = TSH.Elements.GetName(ElementIdName);

                            MessageBox.Show(ElementIdNamTxt);



                        }


                    }
            
                }
                else
                {
                    MessageBox.Show("la liste est vide");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de la récupération nom de l'entité de l'operation : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);

                
            }

        }


        //reconecte le parametre suivant l'indexe
        private void Fixit(ElementId OperationId, List<SmartText> ParametresSmartTxts, int IndexSmartTxt)
        {
           
            
            if (OperationId != null)
            {
                
                
                    if (!TopSolidHost.Application.StartModification("My Action", false)) return;
                    try
                    {
                        SmartText SmartTxt = new SmartText("");
                        // Assurez-vous que la liste n'est pas vide avant d'accéder au premier élément
                        if (ParametresSmartTxts.Count > IndexSmartTxt)
                        {
                            SmartTxt = ParametresSmartTxts[IndexSmartTxt];
                            TSH.Parameters.SetSmartTextParameterCreation(OperationId, SmartTxt);

                            
                        }
                        else
                        {
                            MessageBox.Show("La liste est vide", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }


                       // TSH.Parameters.SetSmartTextParameterCreation(OperationId, SmartTxt);

                        TopSolidHost.Application.EndModification(true, true);
                    }
                    catch
                    {
                        // End modification (failure).
                        TopSolidHost.Application.EndModification(false, false);
                    }


                
               
            }

         }

        //private ElementId ParametreId(DocumentId ) {

        //Récuperation des parametres de derivation---------------------------------------------------------------------------------------------------------------------------------

        private void ObtenirDerivation(DocumentId documentCourantId) {

            bool derivé =  TopSolidDesignHost.Tools.IsDerived(documentCourantId);



            if (derivé) { 

            TopSolidDesignHost.Tools.GetDerivationInheritances(documentCourantId,
                                            out bool outName,
                                            out bool outDescription,
                                            out bool outCode,
                                            out bool outPartNumber,
                                            out bool outComplementaryPartNumber,
                                            out bool outManufacturer,
                                            out bool outManufacturerPartNumber,
                                            out bool outComment,
                                            out List<ElementId> outOtherSystemParameters,
                                            out bool outNonSystemParameters,
                                            out bool outPoints,
                                            out bool outAxes,
                                            out bool outPlanes,
                                            out bool outFrames,
                                            out bool outSketches,
                                            out bool outShapes,
                                            out bool outPublishings,
                                            out bool outFunctions,
                                            out bool outSymmetries,
                                            out bool outUnsectionabilities,
                                            out bool outRepresentations,
                                            out bool outSets,
                                            out bool outCameras
                                        );
                string message = "Variables et leurs valeurs :\n" + $"outName: {outName}\n" + $"outDescription: {outDescription}\n" + $"outCode: {outCode}\n" + $"outPartNumber: {outPartNumber}\n" + $"outComplementaryPartNumber: {outComplementaryPartNumber}\n" + $"outManufacturer: {outManufacturer}\n" + $"outManufacturerPartNumber: {outManufacturerPartNumber}\n" + $"outComment: {outComment}\n" + $"outOtherSystemParameters: {string.Join(", ", outOtherSystemParameters ?? new List<ElementId>())}\n" + $"outNonSystemParameters: {outNonSystemParameters}\n" + $"outPoints: {outPoints}\n" + $"outAxes: {outAxes}\n" + $"outPlanes: {outPlanes}\n" + $"outFrames: {outFrames}\n" + $"outSketches: {outSketches}\n" + $"outShapes: {outShapes}\n" + $"outPublishings: {outPublishings}\n" + $"outFunctions: {outFunctions}\n" + $"outSymmetries: {outSymmetries}\n" + $"outUnsectionabilities: {outUnsectionabilities}\n" + $"outRepresentations: {outRepresentations}\n" + $"outSets: {outSets}\n" + $"outCameras: {outCameras}";
                MessageBox.Show(message, "Valeurs des Variables", MessageBoxButtons.OK, MessageBoxIcon.Information);


            }
            else
            {
                MessageBox.Show("le document n'est pas adapté a la derivation");
            } 
        }


        private void GetCommentaire(DocumentId documentCourantId)
        {

            ElementId commentId = TSH.Parameters.GetCommentParameter(documentCourantId);


            string commentTxt = TSH.Parameters.GetTextValue(commentId);


            MessageBox.Show(commentTxt);

        }




        private string TxtParametrePublié (SmartText SmartTxtParametrePublié)
        {
            string TxtParametrePubliéString = SmartTxtParametrePublié.ToString();
            return TxtParametrePubliéString;
        }

    
    // Gestionnaire de clic pour le bouton.-----------------------------------------------------------------------------------------------------------------------------------------
    //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


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

            //GetCommentaire(documentCourantId);
            //ObtenirDerivation(documentCourantId);




            ElementId PublishFolderId = PublishFolder(documentCourantId);

            List<ElementId> ElementsPubliedIds = PublishedElements(PublishFolderId);

            List<string> namesPublishedElementsTxts = NamesPublishedElements(ElementsPubliedIds);

            // Convertir la liste en une chaîne avec chaque élément sur une nouvelle ligne
            string message = string.Join("\n", namesPublishedElementsTxts); // Afficher la chaîne dans un MessageBox
            MessageBox.Show(message, "Noms des éléments publiés", MessageBoxButtons.OK, MessageBoxIcon.Information);

            string Commentaire = "Commentaire";
            string Designation = "Designation";
            string Indice_3D = "Indice 3D";
            string Nom_docu = "Nom_docu";

            ElementId CommentaireId = PublicationsId(PublishFolderId, Commentaire);
            ElementId DesignationId = PublicationsId(PublishFolderId, Designation);
            ElementId Indice_3DId = PublicationsId(PublishFolderId, Indice_3D);
            ElementId Nom_docuId = PublicationsId(PublishFolderId, Nom_docu);

            SmartText SmartTxtCommentaireId = new SmartText(CommentaireId);
            SmartText SmartTxtDesignationId = new SmartText(DesignationId);
            SmartText SmartTxtIndice_3DId = new SmartText(Indice_3DId);
            SmartText SmartTxtNom_docuId = new SmartText(Nom_docuId);

            List<SmartText> SmartTxtList = new List<SmartText>();
            SmartTxtList.Add(SmartTxtCommentaireId); //Index 0
            int number0 = 0;
            SmartTxtList.Add(SmartTxtDesignationId);//Index 1
            int number1 = 1;
            SmartTxtList.Add(SmartTxtIndice_3DId); //Index 2
            int number2 = 2;
            SmartTxtList.Add(SmartTxtNom_docuId); //Index 3
            int number3 = 3;

            ElementId ParametresFolderId =  ParametresFolder(documentCourantId);

            List<ElementId> ParametresdIds =  ParametresElements(ParametresFolderId);

            
            ElementId OperationId = ParamInvalide(documentCourantId);

            NomElementOperation(documentCourantId);


            Fixit(OperationId, SmartTxtList, number0);
            OperationId = ParamInvalide(documentCourantId);
            NomElementOperation(documentCourantId);
            Fixit(OperationId, SmartTxtList, number1);
            OperationId = ParamInvalide(documentCourantId);
            // Fixit(OperationId, SmartTxtList, number2);
            Fixit(OperationId, SmartTxtList, number3);




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
    }

}