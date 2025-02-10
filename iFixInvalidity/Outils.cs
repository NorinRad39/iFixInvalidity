
using System;
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
using TSHD = TopSolid.Cad.Design.Automating.TopSolidDesignHost;
using System.IO;
using System.Security;
using System.Security.Cryptography;
using System.Diagnostics;
using static System.Net.Mime.MediaTypeNames;
using TopSolid.Cad.Electrode.Automating;
using TSEH = TopSolid.Cad.Electrode.Automating.TopSolidElectrodeHost;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;
using System.Net.Sockets;



namespace iFixInvalidity
{
    internal class Document
    {
        private Form1 formInstance;

        // Constructeur qui accepte une instance de Form1
        public Document(Form1 form)
        {
            formInstance = form;
        }

        // Identifiant du document
        private DocumentId docId;

        // Nom du document sans extension
        private string docNomTxt;

        // Extension du document (ex: .TopPrt, .TopAsm)
        private string docExtention;

        // Identifiant PDM du document
        private PdmObjectId docPdmObject;

        // Liste des paramètres du document
        private List<ElementId> docParameters = new List<ElementId>();

        // Liste des opérations associées au document
        private List<ElementId> docOperations = new List<ElementId>();

        // Identifiants pour les champs système "Commentaire" et "Description"
        private ElementId docCommentaireSysId;
        private ElementId docDescriptionSysId;

        // Indique si le document est dérivé
        private bool docDerived;

        // Indique si le document est une électrode
        private bool docIsElectrode;

        /// <summary>
        /// Identifiant unique du document.
        /// Lorsque défini, met automatiquement à jour les autres propriétés associées.
        /// </summary>
        public DocumentId DocId
        {
            get => docId; 
            set
            {
                docId = value; // Affectation de l'identifiant du document

                try
                {
                    // Récupération du nom du document
                    docNomTxt = TSH.Documents.GetName(docId) ?? "Nom inconnu";
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erreur lors de la récupération du nom du document : {ex.Message}");
                    docNomTxt = "Erreur";
                }

                try
                {
                    // Récupération de l'objet PDM du document
                    docPdmObject = TSH.Documents.GetPdmObject(docId);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erreur lors de la récupération de l'objet PDM : {ex.Message}");
                    docPdmObject = PdmObjectId.Empty; // Valeur par défaut en cas d'erreur
                }

                try
                {
                    // Récupération des paramètres du document
                    docParameters = TSH.Parameters.GetParameters(docId) ?? new List<ElementId>();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erreur lors de la récupération des paramètres : {ex.Message}");
                    docParameters = new List<ElementId>(); // Liste vide en cas d'erreur
                }

                try
                {
                    // Récupération des opérations associées au document
                    docOperations = TSH.Operations.GetOperations(docId) ?? new List<ElementId>();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erreur lors de la récupération des opérations : {ex.Message}");
                    docOperations = new List<ElementId>(); // Liste vide en cas d'erreur
                }

                try
                {
                    // Récupération de l'ID du commentaire système
                    docCommentaireSysId = TSH.Parameters.GetCommentParameter(docId);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erreur lors de la récupération du commentaire système : {ex.Message}");
                    docCommentaireSysId = ElementId.Empty; // Valeur par défaut en cas d'erreur
                }

                try
                {
                    // Récupération de l'ID de la description système
                    docDescriptionSysId = TSH.Parameters.GetDescriptionParameter(docId);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erreur lors de la récupération de la description système : {ex.Message}");
                    docDescriptionSysId = ElementId.Empty; // Valeur par défaut en cas d'erreur
                }

                try
                {
                    // Vérification si le document est dérivé
                    docDerived = TSHD.Tools.IsDerived(docId);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erreur lors de la vérification du caractère dérivé du document : {ex.Message}");
                    docDerived = false; // Valeur par défaut en cas d'erreur
                }

                try
                {
                    
                    // Vérification si le document est une électrode
                    docIsElectrode = formInstance.Iselectrode(docId);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erreur lors de la vérification si le document est une électrode : {ex.Message}");
                    docIsElectrode = false; // Valeur par défaut en cas d'erreur
                }
                
            }
        }

        /// <summary>
        /// Nom du document sans extension (lecture seule).
        /// </summary>
        public string DocNomTxt { get => docNomTxt; private set => docNomTxt = value; }

        /// <summary>
        /// Extension du document (ex: .TopPrt, .TopAsm) (lecture seule).
        /// </summary>
        public string DocExtention { get => docExtention; private set => docExtention = value; }

        /// <summary>
        /// Objet PDM associé au document. Met à jour l'extension du document lors de l'affectation.
        /// </summary>
        public PdmObjectId DocPdmObject
        {
            get => docPdmObject;
            set
            {
                docPdmObject = value;
                TSH.Pdm.GetType(docPdmObject, out docExtention);
            }
        }

        /// <summary>
        /// Liste des paramètres du document (lecture seule).
        /// </summary>
        public List<ElementId> DocParameters { get => docParameters; private set => docParameters = value; }

        /// <summary>
        /// Liste des opérations associées au document (lecture seule).
        /// </summary>
        public List<ElementId> DocOperations { get => docOperations; private set => docOperations = value; }

        /// <summary>
        /// Identifiant du paramètre système pour les commentaires (lecture seule).
        /// </summary>
        public ElementId DocCommentaireSysId { get => docCommentaireSysId; private set => docCommentaireSysId = value; }

        /// <summary>
        /// Identifiant du paramètre système pour la description (lecture seule).
        /// </summary>
        public ElementId DocDescriptionSysId { get => docDescriptionSysId; private set => docDescriptionSysId = value; }

        /// <summary>
        /// Indique si le document est dérivé (lecture seule).
        /// </summary>
        public bool DocDerived { get => docDerived; private set => docDerived = value; }

        /// <summary>
        /// Indique si le document est une électrode (lecture seule).
        /// </summary>
        public bool DocIsElectrode { get => docIsElectrode; private set => docIsElectrode = value; }

        public Document()
        {
            // Constructeur par défaut
        } 
    }
}
