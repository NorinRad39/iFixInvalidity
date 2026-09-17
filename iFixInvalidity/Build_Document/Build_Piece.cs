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

        public static void BuildPiece(Document currentDoc)
        {
            // Étape 1 : Valider que l'identifiant du document fourni n'est pas vide
            if (currentDoc.IsEmpty)
                throw new InvalidOperationException("Aucun document ouvert");

            // Étape 2 : Démarrer une transaction de modification TopSolid
            if (!TSH.Application.StartModification("Creer des parametres texte de base et les publier", false)) return;

            try
            {
                DocumentId currentDocId = currentDoc.DocId;
                // Étape 3 : Marquer le document comme modifié (dirty) pour forcer la sauvegarde
                TSH.Documents.EnsureIsDirty(ref currentDocId);

                                // Étape 4 : Créer/publier les paramètres texte standards du document.
                foreach (var parametreTxt in parametresTxt)
                {
                    // Étape 4.1 : Vérifier si un élément homonyme existe déjà (évite les doublons).
                    ElementId existingId = TSH.Elements.SearchByName(currentDoc.DocId, parametreTxt);
                    if (existingId != ElementId.Empty)
                        continue;

                    // Étape 4.2 : Créer un SmartText vide (valeur initiale du paramètre).
                    SmartText smartText = new SmartText(string.Empty);

                    // Étape 4.3 : Créer le paramètre texte dans le document.
                    ElementId smartTextId = TSH.Parameters.CreateSmartTextParameter(currentDoc.DocId, smartText);
                    if (smartTextId == ElementId.Empty)
                        continue;

                    // Étape 4.4 : Renseigner le nom et la description du paramètre.
                    TSH.Elements.SetName(smartTextId, parametreTxt);
                    TSH.Elements.SetDescription(smartTextId, parametreTxt);

                    // Étape 4.5 : Publier le texte et nommer l'élément publié si création réussie.
                    ElementId publishTextId = TSH.Parameters.PublishText(currentDoc.DocId, parametreTxt, smartText);
                    if (publishTextId != ElementId.Empty)
                    {
                        TSH.Elements.SetName(publishTextId, parametreTxt);
                    }
                }

                // Étape 5 : Publier les métadonnées principales du document (nom, description, commentaire)
                /// <summary>
                /// Extrait et publie les métadonnées principales du document.
                /// </summary>
                /// <remarks>
                /// Traite trois métadonnées du document courant :
                /// <list type="bullet">
                ///   <item><description><b>Nom</b> : Extrait via <see cref="TopSolid.Kernel.Automating.Parameters.GetNameParameter"/>, publié sous la clé "Nom_docu"</description></item>
                ///   <item><description><b>Description</b> : Extrait via <see cref="TopSolid.Kernel.Automating.Parameters.GetDescriptionParameter"/>, publiée sous "Designation"</description></item>
                ///   <item><description><b>Commentaire</b> : Extrait via <see cref="TopSolid.Kernel.Automating.Parameters.GetCommentParameter"/>, publié sous "Commentaire"</description></item>
                /// </list>
                /// Pour chacune : conversion en <see cref="SmartText"/>, publication et nommage de l'élément résultant.
                /// </remarks>
                
                // Étape 5.1 : Récupérer l'identifiant du paramètre de nom du document
                var currentDocNameId = TSH.Parameters.GetNameParameter(currentDoc.DocId);
                
                // Étape 5.2 : Créer un objet SmartText encapsulant la valeur du nom
                SmartText currentDocNameSmartText = new SmartText(currentDocNameId);
                
                // Étape 5.3 : Publier le texte du nom avec la clé "Nom_docu" et nommer l'élément
                ElementId publishNameId = TSH.Parameters.PublishText(currentDoc.DocId, "Nom_docu", currentDocNameSmartText);
                TSH.Elements.SetName(publishNameId, "Nom_docu");

                // Étape 5.4 : Récupérer l'identifiant du paramètre de description du document
                var currentDocDescriptionId = TSH.Parameters.GetDescriptionParameter(currentDoc.DocId);
                
                // Étape 5.5 : Créer un objet SmartText encapsulant la valeur de la description
                SmartText currentDocDescriptionSmartText = new SmartText(currentDocDescriptionId);
                
                // Étape 5.6 : Publier le texte de la description avec la clé "Designation" et nommer l'élément
                ElementId publishDescriptionId = TSH.Parameters.PublishText(currentDoc.DocId, "Designation", currentDocDescriptionSmartText);
                TSH.Elements.SetName(publishDescriptionId, "Designation");

                // Étape 5.7 : Récupérer l'identifiant du paramètre de commentaire du document
                var currentDocCommentId = TSH.Parameters.GetCommentParameter(currentDoc.DocId);
                
                // Étape 5.8 : Créer un objet SmartText encapsulant la valeur du commentaire
                SmartText currentDocCommentSmartText = new SmartText(currentDocCommentId);
                
                // Étape 5.9 : Publier le texte du commentaire avec la clé "Commentaire" et nommer l'élément
                ElementId publishCommentId = TSH.Parameters.PublishText(currentDoc.DocId, "Commentaire", currentDocCommentSmartText);
                TSH.Elements.SetName(publishCommentId, "Commentaire");

                // Étape 6 : Valider la transaction TopSolid et appliquer les changements.
                TSH.Application.EndModification(true, true);
            }
            catch (Exception)
            {
                // Étape 7 : Annuler la transaction en cas d'erreur pour éviter toute corruption.
                TSH.Application.EndModification(false, false);
            }
        }
    }
}
