using OutilsTs;
using System;
using System.Collections.Generic;
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
            if(!PDM.IsValidated(libDocuType))
            {
                TSH.Pdm.SetLifeCycleMainState(libDocuType, PdmLifeCycleMainState.Validated);
            }
            #endregion

            #region gestion document d'enumeration des type de document enumDocuType
            //Gestion de creation du document enumDocuType avec mise au coffre et validation du cycle de vie
            var enumDoc = BibliothequeEnum.CheckEnum(libDocuType);
            //On renomme le document créé en enumDocuType
            TSH.Pdm.SetName(enumDoc.DocPdmObject, "enumDocuType");
            //insértion des différents types de documents
            BibliothequeEnum.EditEnumDocument(enumDoc.DocId);
            //Mise au coffre
            if(!PDM.IsCheckedIn(enumDoc.DocPdmObject))
            {
                TSH.Pdm.CheckIn(enumDoc.DocPdmObject, false);
            }
            //Validation cycle de vie
            if(!PDM.IsValidated(enumDoc.DocPdmObject))
            {
                TSH.Pdm.SetLifeCycleMainState(enumDoc.DocPdmObject, PdmLifeCycleMainState.Validated);
            }

            PDM.RefLibrary(PDM.GetCurrentProjectPdmObject(), libDocuType);

           

            // Récupérer le premier élément à l'intérieur du fichier d'énumération (la définition de l'enum)
            List<ElementId> enumElements = TSH.Elements.GetElements(enumDoc.DocId);
            Guid enumDocGuid = TSH.Documents.GetTypeGuid(enumDoc.DocId);
            ElementId enumDefinitionId = enumElements[1];
            DocumentId currentDoc = TSH.Documents.EditedDocument;

            if (TopSolidHost.Application.StartModification("Création Paramètre via Enum Directe", false))
            {
                try
                {
                    // 2. Création d'un paramètre entier standard
                    ElementId paramId = TSH.Parameters.CreateIntegerParameter(currentDoc, 0 );

                    TSH.Parameters.SetEnumerationValue(paramId, 1);

                    TopSolidHost.Application.EndModification(true, true);
                }
                catch (Exception ex)
                {
                    TopSolidHost.Application.EndModification(false, false);
                }
            }


            //DocBuild.CreateEnumParam(enumDocId.DocId);

            var (intValue, textValue) = BibliothequeEnum.AfficherEnumValues();

            #endregion





        }
    }
}
