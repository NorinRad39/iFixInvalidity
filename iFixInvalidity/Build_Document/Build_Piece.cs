using OutilsTs;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Security;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using TopSolid.Cad.Design.Automating;
using TopSolid.Cad.Drafting.Automating;
using TopSolid.Cad.Electrode.Automating;
using TopSolid.Kernel.Automating;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;
using TSEH = TopSolid.Cad.Electrode.Automating.TopSolidElectrodeHost;
using TSH = TopSolid.Kernel.Automating.TopSolidHost;
using TSHD = TopSolid.Cad.Design.Automating.TopSolidDesignHost;


namespace iFixInvalidity.Build_Document
{
    internal static class Build_Piece
    {
        private static readonly List<string> parametresTxt = new List<string> { "Indice 3D", "Matiere plan", "Nombre de pieces", "Traitement" };

        public static void BuildPiece(DocumentId currentDoc)
        {
            // Étape 1 : Valider que l'identifiant du document fourni n'est pas vide
            if (currentDoc.IsEmpty)
                throw new InvalidOperationException("Aucun document ouvert");

            // Étape 2 : Démarrer une transaction de modification TopSolid
            if (!TSH.Application.StartModification("Creer des parametres texte de base et les publier", false)) return;

            try
            {
                // Étape 3 : Marquer le document comme modifié (dirty) pour forcer la sauvegarde
                TSH.Documents.EnsureIsDirty(ref currentDoc);

                                // Étape 4 : Créer/publier les paramètres texte standards du document.
                foreach (var parametreTxt in parametresTxt)
                {
                    // Étape 4.1 : Vérifier si un élément homonyme existe déjà (évite les doublons).
                    ElementId existingId = TSH.Elements.SearchByName(currentDoc, parametreTxt);
                    if (existingId != ElementId.Empty)
                        continue;

                    // Étape 4.2 : Créer un SmartText vide (valeur initiale du paramètre).
                    SmartText smartText = new SmartText(string.Empty);

                    // Étape 4.3 : Créer le paramètre texte dans le document.
                    ElementId smartTextId = TSH.Parameters.CreateSmartTextParameter(currentDoc, smartText);
                    if (smartTextId == ElementId.Empty)
                        continue;

                    // Étape 4.4 : Renseigner le nom et la description du paramètre.
                    TSH.Elements.SetName(smartTextId, parametreTxt);
                    TSH.Elements.SetDescription(smartTextId, parametreTxt);

                    // Étape 4.5 : Publier le texte et nommer l'élément publié si création réussie.
                    ElementId publishTextId = TSH.Parameters.PublishText(currentDoc, parametreTxt, smartText);
                    if (publishTextId != ElementId.Empty)
                    {
                        TSH.Elements.SetName(publishTextId, parametreTxt);
                    }
                }

                var currentDocNameId = TSH.Parameters.GetNameParameter(currentDoc);
               // var currentDocNameValue = TSH.Parameters.GetTextValue(currentDocNameId);
                SmartText currentDocNameSmartText = new SmartText(currentDocNameId);
                ElementId publishNameId = TSH.Parameters.PublishText(currentDoc, "Nom_docu", currentDocNameSmartText);
                TSH.Elements.SetName(publishNameId, "Nom_docu");

                var currentDocDescriptionId = TSH.Parameters.GetDescriptionParameter(currentDoc);
                //var currentDocDescriptionValue = TSH.Parameters.GetTextValue(currentDocDescriptionId);
                SmartText currentDocDescriptionSmartText = new SmartText(currentDocDescriptionId);
                ElementId publishDescriptionId = TSH.Parameters.PublishText(currentDoc, "Designation", currentDocDescriptionSmartText);
                TSH.Elements.SetName(publishDescriptionId, "Designation");

                var currentDocCommentId = TSH.Parameters.GetCommentParameter(currentDoc);
                //var currentDocCommentValue = TSH.Parameters.GetTextValue(currentDocCommentId);
                SmartText currentDocCommentSmartText = new SmartText(currentDocCommentId);
                ElementId publishCommentId = TSH.Parameters.PublishText(currentDoc, "Commentaire", currentDocCommentSmartText);
                TSH.Elements.SetName(publishCommentId, "Commentaire");

                // Étape 7 : Valider la transaction TopSolid et appliquer les changements.
                TSH.Application.EndModification(true, true);
            }
            catch (Exception)
            {
                // Étape 8 : Annuler la transaction en cas d'erreur pour éviter toute corruption.
                TSH.Application.EndModification(false, false);
            }
        }
    }
}
