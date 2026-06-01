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

namespace iFixInvalidity
{
    internal class DocBuild
    {
        public static ElementId CreateEnumParam(DocumentId enumDocId)
        {
            DocumentId docId = TSH.Documents.EditedDocument;

            // Étape 1 : Valider que l'identifiant du document fourni n'est pas vide
            if (docId.IsEmpty)
                throw new InvalidOperationException("Le document actuel est vide.");

            // Étape 2 : Démarrer une transaction de modification TopSolid
            if (!TSH.Application.StartModification("Ajout parametre enum dans document courant", false)) 
                return default(ElementId);

            try
            {
                // Étape 3 : Marquer le document comme modifié (dirty) pour forcer la sauvegarde
                TSH.Documents.EnsureIsDirty(ref docId);
                var elementId = TSH.Parameters.CreateUserPropertyParameter(docId, enumDocId);
                TSH.Parameters.SetIntegerValue(elementId, 0);

                // Étape 7 : Valider et clôturer la transaction avec succès
                TSH.Application.EndModification(true, true);
                return elementId;
            }
            catch (Exception ex)
            {
                // Étape 8 (erreur) : Annuler la transaction pour ne pas corrompre le document
                TSH.Application.EndModification(false, false);
                throw;
            }
        }
    }
}
