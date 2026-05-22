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
    /// <summary>
    /// Fournit des utilitaires pour vérifier et créer la bibliothèque PDM
    /// et le document d'énumération utilisé pour la génération de documents.
    /// </summary>
    internal class BibliothequeEnum
    {
        /// <summary>
        /// Point d'entrée : s'assure que la bibliothèque PDM nommée <c>docuType</c>
        /// et le document d'énumération <c>enumDocuType</c> existent.
        /// Si nécessaire, crée la bibliothèque ou le document puis renomme l'objet PDM.
        /// </summary>
        /// <remarks>
        /// Cette méthode appelle <see cref="CheckLib"/> et <see cref="CheckEnum(ProjetPDM)"/>.
        /// Elle dépend de l'API PDM exposée via <c>TSH.Pdm</c> (TopSolid).
        /// </remarks>
        public void GestionBibliothequeEnum()
        {
            var libDocuType = CheckLib();
            var enumDocId = CheckEnum(libDocuType);
            TSH.Pdm.SetName(enumDocId, "enumDocuType");
        }

        /// <summary>
        /// Vérifie l'existence d'une bibliothèque PDM nommée <c>docuType</c>.
        /// Si la bibliothèque est absente, tente de la créer.
        /// </summary>
        /// <returns>
        /// Le <c>ProjetPDM</c> correspondant à la bibliothèque <c>docuType</c> si trouvé,
        /// sinon la valeur retournée par l'appel à <c>ProjetPDM.TryGetLibrary</c>.
        /// </returns>
        /// <remarks>
        /// Attention : si <c>ProjetPDM.TryGetLibrary</c> renvoie <c>false</c>, la variable
        /// locale <c>libDocuType</c> peut rester <c>null</c> même après l'appel à
        /// <c>TSH.Pdm.CreateProject</c>. Le code appelant doit prendre en compte ce cas
        /// ou la méthode devrait récupérer explicitement le projet créé.
        /// </remarks>
        private ProjetPDM CheckLib()
        {
            if (!ProjetPDM.TryGetLibrary("docuType", out var libDocuType, true))
            {
                // Tentative de création si la bibliothèque n'existe pas
                TSH.Pdm.CreateProject("docuType", true);

                // Réessayer de récupérer le projet créé afin de retourner
                // un Objet ProjetPDM non-null si la création a réussi.
                ProjetPDM.TryGetLibrary("docuType", out libDocuType, true);
            }

            return libDocuType;
        }

        /// <summary>
        /// Recherche dans la bibliothèque fournie un document nommé <c>enumDocuType</c>.
        /// Si aucun document correspondant n'est trouvé, crée un nouveau document de type <c>TopEnu</c>.
        /// </summary>
        /// <param name="libDocuType">La bibliothèque PDM à inspecter. Peut être <c>null</c>.</param>
        /// <returns>
        /// L'identifiant PDM (<c>PdmObjectId</c>) du document existant ou nouvellement créé.
        /// </returns>
        /// <remarks>
        /// Si <paramref name="libDocuType"/> est <c>null</c>, l'appel à
        /// <c>TSH.Pdm.CreateDocument(libDocuType.PdmObjectId, ...)</c> provoquera une exception.
        /// Il est recommandé de valider que la bibliothèque est non nulle avant l'appel,
        /// ou de modifier <see cref="CheckLib"/> pour retourner systématiquement le projet créé.
        /// </remarks>
        private PdmObjectId CheckEnum(ProjetPDM libDocuType)
        {
            // Certains API TopSolid peuvent renvoyer un objet ProjetPDM non-null mais
            // contenant un PdmObjectId vide. On vérifie donc à la fois la nullité
            // et l'état Empty de l'identifiant PDM.
            if (libDocuType == null || libDocuType.PdmObjectId.IsEmpty)
            {
                throw new InvalidOperationException("La bibliothèque PDM 'docuType' est introuvable ou invalide : impossible de créer ou retrouver 'enumDocuType'.");
            }

            foreach (var doc in libDocuType.Documents)
            {
                string docName = TSH.Pdm.GetName(doc);
                if (docName == "enumDocuType")
                {
                    return doc;
                }
            }

            // Aucun document trouvé : création d'un nouveau document d'énumération
            var enumDocId = TSH.Pdm.CreateDocument(libDocuType.PdmObjectId, ".TopEnu", false);
            return enumDocId;
        }       
    } 
}