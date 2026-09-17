using OutilsTs;
using System;
using System.Collections.Generic;
using System.Linq;
using TopSolid.Cad.Design.Automating;
using TopSolid.Kernel.Automating;
using TSH = TopSolid.Kernel.Automating.TopSolidHost;


namespace iFixInvalidity.Build_Document
{
    /// <summary>
    /// Classe statique responsable de la construction et de la gestion des documents
    /// dans le contexte PDM TopSolid.
    /// </summary>
    internal static class MainDocBuild
    {
        /// <summary>
        /// Point d'entrée : s'assure que la bibliothèque PDM nommée <c>docuType</c>
        /// et le document d'énumération <c>enumDocuType</c> existent.
        /// Si nécessaire, crée la bibliothèque ou le document puis renomme l'objet PDM.
        /// </summary>
        /// <remarks>
        /// <para>
        /// Le traitement se déroule en deux phases :
        /// </para>
        /// <list type="number">
        ///   <item>
        ///     <term>Bibliothèque <c>docuType</c></term>
        ///     <description>
        ///       Appelle <see cref="BibliothequeEnum.CheckLib"/> pour vérifier ou créer la bibliothèque,
        ///       puis effectue une mise au coffre (<c>CheckIn</c>) et valide son cycle de vie
        ///       en la passant à l'état <see cref="PdmLifeCycleMainState.Validated"/>.
        ///     </description>
        ///   </item>
        ///   <item>
        ///     <term>Document d'énumération <c>enumDocuType</c></term>
        ///     <description>
        ///       Appelle <see cref="BibliothequeEnum.CheckEnum(PdmObjectId)"/> pour vérifier ou créer
        ///       le document d'énumération dans la bibliothèque, renomme l'objet PDM en
        ///       <c>enumDocuType</c>, appelle <see cref="BibliothequeEnum.EditEnumDocument"/> pour
        ///       y insérer les différents types de documents, puis effectue une mise au coffre
        ///       et valide son cycle de vie.
        ///     </description>
        ///   </item>
        /// </list>
        /// <para>
        /// Cette méthode appelle <see cref="BibliothequeEnum.CheckLib"/> et
        /// <see cref="BibliothequeEnum.CheckEnum(PdmObjectId)"/>.
        /// </para>
        /// </remarks>
        public static void GestionBibliothequeEnum()
        {
            #region gestion bibliothèque docuType
            //Gestion de creation de la librairie docuType avec mise au coffre et validation du cycle de vie
            var libDocuType = BibliothequeEnum.CheckLib();
            //Mise au coffre
            if (!PDM.IsCheckedIn(libDocuType))
            {
                TSH.Pdm.CheckIn(libDocuType, false);
            }

            //Validation cycle de vie
            if (!PDM.IsValidated(libDocuType))
            {
                TSH.Pdm.SetLifeCycleMainState(libDocuType, PdmLifeCycleMainState.Validated);
            }
            #endregion

            #region gestion document d'enumeration des type de document enumDocuType
            // Vérifie l'existence du document d'énumération associé à la bibliothèque `docuType`.
            // Si nécessaire, le document est créé afin de centraliser les valeurs métier
            // utilisées pour typer les documents du projet.
            var enumDoc = BibliothequeEnum.CheckEnum(libDocuType);

            // Lit le nom courant de l'objet PDM afin de déterminer si une normalisation
            // du document est encore nécessaire.
            var currentPdmName = TSH.Pdm.GetName(enumDoc.DocPdmObject);

            // Lit les valeurs courantes de l'énumération dans le document existant.
            BibliothequeEnum.GetEnumValuesFromDocument(enumDoc.DocId, out _, out List<string> currentEnumTextValues);

            // Récupère la liste attendue depuis la source métier unique.
            List<string> expectedEnumTextValues = BibliothequeEnum.GetEnumTextValues();

            bool mustRenameEnumDocument = !string.Equals(currentPdmName, "enumDocuType", StringComparison.Ordinal);
            bool mustUpdateEnumContent =
                currentEnumTextValues == null ||
                !currentEnumTextValues.SequenceEqual(expectedEnumTextValues);

            // Si le nom du document n'est pas conforme, il est normalisé.
            if (mustRenameEnumDocument)
            {
                TSH.Pdm.SetName(enumDoc.DocPdmObject, "enumDocuType");
            }

            // Si le nom a changé ou si le contenu diffère de `DocuTypeStr`,
            // le document est régénéré puis sécurisé.
            if (mustRenameEnumDocument || mustUpdateEnumContent)
            {
                BibliothequeEnum.EditEnumDocument(enumDoc.DocId);

                if (!PDM.IsCheckedIn(enumDoc.DocPdmObject))
                {
                    TSH.Pdm.CheckIn(enumDoc.DocPdmObject, false);
                }

                if (!PDM.IsValidated(enumDoc.DocPdmObject))
                {
                    TSH.Pdm.SetLifeCycleMainState(enumDoc.DocPdmObject, PdmLifeCycleMainState.Validated);
                }
            }

            // Référence la bibliothèque `docuType` dans le projet courant pour la rendre accessible
            // dans le contexte PDM de l'application.
            PDM.RefLibrary(PDM.GetCurrentProjectPdmObject(), libDocuType);
            #endregion

            #region gestion document de propriété utilisateur userPropertyDocuType
            // Récupère tous les éléments portés par le document d'énumération.
            // L'index `1` correspond ici à la définition exploitable de l'énumération
            // utilisée ensuite pour le paramétrage métier.
            List<ElementId> enumElements = TSH.Elements.GetElements(enumDoc.DocId);

            // Conserve le type réel du document d'énumération.
            // Cette information peut servir à vérifier que le document manipulé
            // correspond bien à une définition d'énumération attendue.
            Guid enumDocGuid = TSH.Documents.GetTypeGuid(enumDoc.DocId);

            // Référence l'élément de définition de l'énumération dans le document.
            ElementId enumDefinitionId = enumElements[1];

            // Cible le document actuellement édité dans lequel le paramètre sera créé.
            Document currentDoc = new Document();

            //DocBuild.CreateEnumParam(enumDocId.DocId);
            // Ouvre une fenêtre permettant à l'utilisateur de choisir une valeur dans l'énumération
            // `enumDocuType` et récupère un tuple :
            //  - `intEnumValue`  : identifiant entier PDM de la valeur sélectionnée (-1 si l'utilisateur annule)
            //  - `textEnumValue` : libellé textuel associé (chaîne vide si l'utilisateur annule)
            // Remarque : la méthode `BibliothequeEnum.AfficherEnumValues()` retourne `(-1, string.Empty)`
            // si la boîte de dialogue est annulée — le code appelant doit vérifier `intValue` pour
            // détecter une annulation avant d'utiliser les valeurs.
            var (intEnumValue, textEnumValue) = BibliothequeEnum.AfficherEnumValues();

            if (intEnumValue < 0)
                return;

            DocBuild.CreateEnumParam(currentDoc.DocId, enumDoc.DocId, intEnumValue, textEnumValue);

            switch (textEnumValue)
            {
                case BibliothequeEnum.DocuTypes.Piece:
                  break;

                case BibliothequeEnum.DocuTypes.LiasseDePlans:
                    // Build_Assemblage.BuildAssemblage(currentDoc.DocId);
                    break;

                case BibliothequeEnum.DocuTypes.Electrode:
                case BibliothequeEnum.DocuTypes.ElectrodeParallelisee:
                case BibliothequeEnum.DocuTypes.BrutElectrode:
                    // Build_Electrode.BuildElectrode(currentDoc.DocId);
                    break;

                default:
                    break;
            }


            #endregion


        }
    }
}
