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
    /// Fournit des utilitaires statiques pour vérifier et créer la bibliothèque PDM
    /// et le document d'énumération utilisé pour la génération de documents.
    /// </summary>
    internal static class BibliothequeEnum
    {
        /// <summary>
        /// Vérifie l'existence d'une bibliothèque PDM nommée <c>docuType</c> parmi toutes les
        /// bibliothèques disponibles dans le PDM.
        /// <list type="bullet">
        ///   <item>
        ///     <description>
        ///       Si la bibliothèque <c>docuType</c> est trouvée et active (hors corbeille),
        ///       son identifiant est immédiatement retourné.
        ///     </description>
        ///   </item>
        ///   <item>
        ///     <description>
        ///       Si la bibliothèque <c>docuType</c> est trouvée mais se trouve dans la corbeille,
        ///       un message d'information est affiché et la recherche se poursuit.
        ///     </description>
        ///   </item>
        ///   <item>
        ///     <description>
        ///       Si aucune bibliothèque <c>docuType</c> active n'est trouvée (liste vide ou absente),
        ///       une nouvelle bibliothèque est créée via <c>TSH.Pdm.CreateProject</c>.
        ///     </description>
        ///   </item>
        /// </list>
        /// </summary>
        /// <returns>
        /// Le <see cref="PdmObjectId"/> de la bibliothèque PDM <c>docuType</c> existante ou
        /// nouvellement créée. Retourne <see cref="PdmObjectId.Empty"/> si la bibliothèque
        /// <c>docuType</c> n'existe que dans la corbeille et qu'aucune création n'a été possible.
        /// </returns>
        /// <remarks>
        /// Des boîtes de dialogue (<see cref="System.Windows.Forms.MessageBox"/>) sont affichées
        /// pour informer l'utilisateur du résultat de chaque scénario (existence, présence en
        /// corbeille ou création réussie).
        /// </remarks>
        internal static PdmObjectId CheckLib()
        {
            // Étape 1 : Récupérer toutes les bibliothèques disponibles dans le PDM
            var alLib = PDM.GetLibraries();

            // Indicateur pour savoir si une bibliothèque 'docuType' active a été trouvée
            bool docuTypeExiste = false;

            // Étape 2 : Vérifier si la liste des bibliothèques est non nulle et non vide
            if (alLib != null && alLib.Count > 0)
            {
                // Étape 3 : Parcourir chaque bibliothèque pour chercher 'docuType'
                foreach (var lib in alLib)
                {
                    // Étape 4 : Comparer le nom de la bibliothèque avec 'docuType'
                    if (TSH.Pdm.GetName(lib) == "docuType")
                    {
                        // Étape 5 : Vérifier que la bibliothèque n'est pas dans la corbeille
                        if (!PDM.IsInRecycleBin(lib))
                        {
                            // Bibliothèque active trouvée : on la retourne immédiatement
                            //MessageBox.Show("La bibliothèque PDM 'docuType' existe déjà.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            docuTypeExiste = true;
                            return lib;
                        }
                        // Sinon : la bibliothèque est en corbeille, on continue la recherche
                        //else
                        //{
                        //    MessageBox.Show("Une bibliothèque PDM 'docuType' existe dans la corbeille.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        //}
                    }
                }

                // Étape 6 : Aucune bibliothèque 'docuType' active trouvée — on la crée
                if (!docuTypeExiste)
                {
                    // Création de la bibliothèque 'docuType' dans le PDM
                    var libDocuType = TSH.Pdm.CreateProject("docuType", true);
                    MessageBox.Show("La bibliothèque PDM 'docuType' a été créée avec succès.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return libDocuType;
                }
            }
            else
            {
                // Étape 7 : La liste est vide ou nulle — aucune bibliothèque dans le PDM, on crée 'docuType'
                var libDocuType = TSH.Pdm.CreateProject("docuType", true);
                MessageBox.Show("La bibliothèque PDM 'docuType' a été créée avec succès.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return libDocuType;
            }

            // Étape 8 : Cas de repli — la bibliothèque n'existe qu'en corbeille et n'a pas pu être créée
            return PdmObjectId.Empty;
        }

        /// <summary>
        /// Vérifie l'existence d'un document d'énumération nommé <c>enumDocuType</c> dans la bibliothèque PDM spécifiée.
        /// Si aucun document correspondant n'est trouvé, un nouveau document de type <c>.TopEnu</c> est créé.
        /// </summary>
        /// <param name="libDocuType">
        /// Le <see cref="PdmObjectId"/> de la bibliothèque PDM à inspecter.
        /// Ne doit pas être vide (<see cref="PdmObjectId.IsEmpty"/> doit être <c>false</c>).
        /// </param>
        /// <returns>
        /// Le <see cref="PdmObjectId"/> du document <c>enumDocuType</c> existant s'il est trouvé,
        /// ou celui du document nouvellement créé (<c>.TopEnu</c>) dans le cas contraire.
        /// </returns>
        /// <exception cref="InvalidOperationException">
        /// Levée si <paramref name="libDocuType"/> est vide (<see cref="PdmObjectId.IsEmpty"/> est <c>true</c>).
        /// </exception>
        internal static Document CheckEnum(PdmObjectId libDocuType)
        {
            Document enumDocument;

            // Étape 1 : Valider que l'identifiant de bibliothèque fourni n'est pas vide
            if (libDocuType.IsEmpty)
                throw new InvalidOperationException("La bibliothèque PDM 'docuType' doit être fournie pour vérifier ou créer le document d'énumération.");

            // Étape 2 : Parcourir tous les documents de la bibliothèque pour chercher 'enumDocuType'
            foreach (var doc in PDM.GetAllProjectDocuments(libDocuType))
            {
                // Étape 3 : Comparer le nom du document avec 'enumDocuType'
                if (TSH.Pdm.GetName(doc) == "enumDocuType")
                {
                    // Document trouvé : on le retourne encapsulé dans un objet Document
                    return enumDocument = new Document(TSH.Documents.GetDocument(doc));
                }
            }

            // Étape 4 : Aucun document 'enumDocuType' trouvé — on le crée avec l'extension .TopEnu
            var enumDocumentPdmId = TSH.Pdm.CreateDocument(libDocuType, ".TopEnu", false);

            // Étape 5 : Retourner le document nouvellement créé encapsulé dans un objet Document
            return enumDocument = new Document(TSH.Documents.GetDocument(enumDocumentPdmId));
        }

        /// <summary>
        /// Édite le document d'énumération spécifié en y injectant des valeurs et libellés prédéfinis
        /// via l'API TopSolid.
        /// <para>
        /// La méthode ouvre une transaction de modification TopSolid, marque le document comme modifié,
        /// puis appelle <c>TSH.Parameters.SetUserEnumerationValues</c> pour définir les couples
        /// (valeur entière / texte) de l'énumération utilisateur.
        /// La transaction est validée en cas de succès, ou annulée en cas d'exception afin de ne pas
        /// corrompre le document.
        /// </para>
        /// </summary>
        /// <param name="enumDocument">
        /// Le <see cref="DocumentId"/> du document d'énumération TopSolid (<c>.TopEnu</c>) à éditer.
        /// Ne doit pas être vide (<see cref="DocumentId.IsEmpty"/> doit être <c>false</c>).
        /// </param>
        /// <exception cref="InvalidOperationException">
        /// Levée si <paramref name="enumDocument"/> est vide (<see cref="DocumentId.IsEmpty"/> est <c>true</c>).
        /// </exception>
        /// <remarks>
        /// <list type="bullet">
        ///   <item>
        ///     <description>
        ///       Si <c>TSH.Application.StartModification</c> retourne <c>false</c>, la méthode retourne
        ///       immédiatement sans effectuer aucune modification.
        ///     </description>
        ///   </item>
        ///   <item>
        ///     <description>
        ///       En cas d'exception, <c>TopSolidHost.Application.EndModification(false, false)</c> est
        ///       appelé pour annuler la transaction et préserver l'intégrité du document.
        ///     </description>
        ///   </item>
        /// </list>
        /// </remarks>
        internal static void EditEnumDocument(DocumentId enumDocument)
        {
            // Étape 1 : Valider que l'identifiant du document fourni n'est pas vide
            if (enumDocument.IsEmpty)
                throw new InvalidOperationException("Le document d'énumération 'enumDocuType' doit être fourni pour être édité.");

            // Étape 2 : Démarrer une transaction de modification TopSolid
            // Si la transaction ne peut pas démarrer, on abandonne immédiatement
            if (!TSH.Application.StartModification("Ajout valeurs Enum", false)) return;

            try
            {
                // Étape 3 : Marquer le document comme modifié (dirty) pour forcer la sauvegarde
                TSH.Documents.EnsureIsDirty(ref enumDocument);

                // Étape 4 : Préparer la liste des valeurs entières de l'énumération
                List<int> intValues = "1,2,3".Split(',').Select(int.Parse).ToList();

                // Étape 5 : Préparer la liste des libellés textuels correspondants
                List<string> text = "TypeA,TypeB,TypeC".Split(',').ToList();

                // Étape 6 : Injecter les couples (valeur / libellé) dans le document d'énumération
                TSH.Parameters.SetUserEnumerationValues(enumDocument, intValues, text);

                // Étape 7 : Valider et clôturer la transaction avec succès
                TopSolidHost.Application.EndModification(true, true);
            }
            catch (Exception ex)
            {
                // Étape 8 (erreur) : Annuler la transaction pour ne pas corrompre le document
                TopSolidHost.Application.EndModification(false, false);
            }

        }
    }
}