
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


namespace iFixInvalidity
{
    internal class Document
    {
        private DocumentId docId;
        private string docNomTxt;
        private string docExtention;
        private PdmObjectId docPdmObject;
        private List<ElementId> docParameters;
        private List<ElementId> docOperations;
        private ElementId docCommentaireSysId;
        private ElementId docDescriptionSysId;
        private bool docDerived;
        private bool docIsElectrode;


        public DocumentId DocId
        {
            get { return docId; }
            set
            {
                docId = value;
                docNomTxt = TSH.Documents.GetName(docId); // Mise à jour automatique de Nom en utilisant la méthode externe
                docPdmObject = TSH.Documents.GetPdmObject(docId); // Mise à jour automatique de Nom en utilisant la méthode externe
                docParameters = TSH.Parameters.GetParameters(docId);
                docOperations = TSH.Operations.GetOperations(docId);
                docCommentaireSysId = TSH.Parameters.GetCommentParameter(docId);
                docDescriptionSysId = TSH.Parameters.GetDescriptionParameter(docId);
                docDerived = TSHD.Tools.IsDerived(docId);
                docIsElectrode = TSEH.Electrodes.IsElectrodes(docId);
            }
        }
    
        public string DocNomTxt
        {
            get { return docNomTxt; }
            private set { docNomTxt = value; }
        }
        public string DocExtention
        {
            get { return docExtention; }
            private set { docExtention = value; }
        }

        public PdmObjectId DocPdmObject
        {
            get { return docPdmObject; }
            set 
            { 
                docPdmObject = value;
                TSH.Pdm.GetType(docPdmObject,out docExtention);
            }
        }

        public List<ElementId> DocParameters
        {
            get { return docParameters; }
            private set { docParameters = value; }
        }

        public List<ElementId> DocOperations
        {
            get { return docOperations; }
            private set { docOperations = value; }
        }

        public ElementId DocCommentaireSysId
        {
            get { return docCommentaireSysId; }
            private set { docCommentaireSysId = value; }
        }

        public ElementId DocDescriptionSysId
        {
            get { return docDescriptionSysId; }
            private set { docDescriptionSysId = value; }
        }
        public bool DocDerived
        {
            get { return docDerived; }
            private set { docDerived = value; }
        }
        public bool DocIsElectrode
        {
            get { return docIsElectrode; }
            private set { docIsElectrode = value; }
        }

        public Document()
        {
            // Constructeur par défaut
        }
    }
}
