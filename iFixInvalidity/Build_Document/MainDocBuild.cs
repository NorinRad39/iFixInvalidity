using OutilsTs;
using System;
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
            TopSolidHost.Pdm.CheckIn(libDocuType, false);
            //Validation cycle de vie
            TopSolidHost.Pdm.SetLifeCycleMainState(libDocuType, PdmLifeCycleMainState.Validated);
            #endregion

            #region gestion document enumDocuType
            //Gestion de creation du document enumDocuType avec mise au coffre et validation du cycle de vie
            var enumDocId = BibliothequeEnum.CheckEnum(libDocuType);
            //On renomme le document créé en enumDocuType
            TSH.Pdm.SetName(enumDocId.DocPdmObject, "enumDocuType");
            //insértion des différents types de documents
            BibliothequeEnum.EditEnumDocument(enumDocId.DocId);
            //Mise au coffre
            TopSolidHost.Pdm.CheckIn(enumDocId.DocPdmObject, false);
            //Validation cycle de vie
            TopSolidHost.Pdm.SetLifeCycleMainState(enumDocId.DocPdmObject, PdmLifeCycleMainState.Validated);
            #endregion
        }
    }
}
