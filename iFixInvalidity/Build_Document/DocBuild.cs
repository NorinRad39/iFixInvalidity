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
    internal class DocBuild
    {
                /// <summary>
        /// Crée (ou remplace) le paramètre d'énumération utilisateur <c>DocuType</c> dans le document courant.
        /// </summary>
        /// <param name="currentDoc">Identifiant du document TopSolid actuellement édité.</param>
        /// <param name="enumDocId">Identifiant du document portant la définition de l'énumération.</param>
        /// <param name="intEnumValue">Valeur entière de l'élément d'énumération à appliquer.</param>
        /// <param name="textEnumValue">Texte descriptif associé à la valeur d'énumération.</param>
        /// <returns>
        /// L'identifiant du paramètre créé si l'opération aboutit ; <c>default(ElementId)</c> si la modification
        /// ne démarre pas ou si l'utilisateur refuse le remplacement d'un paramètre existant.
        /// </returns>
        /// <exception cref="InvalidOperationException">
        /// Levée lorsque <paramref name="currentDoc"/> est vide.
        /// </exception>
        /// <remarks>
        /// La méthode encapsule les opérations dans une transaction TopSolid :
        /// <list type="number">
        /// <item><description>Démarre la modification.</description></item>
        /// <item><description>Recherche un paramètre nommé <c>DocuType</c> déjà présent.</description></item>
        /// <item><description>Demande confirmation à l'utilisateur avant remplacement.</description></item>
        /// <item><description>Marque le document comme modifié puis crée et renseigne le paramètre.</description></item>
        /// <item><description>Valide la transaction en cas de succès, sinon annule et relaie l'exception.</description></item>
        /// </list>
        /// </remarks>
        public static ElementId CreateEnumParam(DocumentId currentDoc, DocumentId enumDocId, int intEnumValue, string textEnumValue)
        {
            if (currentDoc.IsEmpty)
                throw new InvalidOperationException("Aucun document en cours d'édition.");

            if (!TSH.Application.StartModification("Ajout parametre enum dans document courant", false))
                return default(ElementId);

            try
            {
                var docuTypeExiste = TSH.Elements.SearchByName(currentDoc, "DocuType");

                if (!docuTypeExiste.IsEmpty)
                {
                    DialogResult reponse = MessageBox.Show(
                        "Le paramètre DocuType existe dans le document courant. Voulez-vous le remplacer ?",
                        "Confirmation",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (reponse == DialogResult.No)
                    {
                        TSH.Application.EndModification(false, false);
                        return default(ElementId);
                    }
                    else if (reponse == DialogResult.Yes)
                    {
                        TSH.Elements.Delete(docuTypeExiste);
                    }
                }

                TSH.Documents.EnsureIsDirty(ref currentDoc);

                var userEnumParameter = TSH.Parameters.CreateUserEnumParameter(currentDoc, enumDocId);

                TSH.Parameters.SetUserEnumerationValue(userEnumParameter, intEnumValue);
                TSH.Elements.SetName(userEnumParameter, "DocuType");
                TSH.Elements.SetDescription(userEnumParameter, textEnumValue);

                TSH.Application.EndModification(true, true);
                return userEnumParameter;
            }
            catch (Exception)
            {
                TSH.Application.EndModification(false, false);
                throw;
            }
        }



            
    }   
}
