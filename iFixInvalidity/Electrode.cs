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
using OutilsTs;

namespace iFixInvalidity
{
    public class Electrode
    {
        // Méthode publique à appeler
        public void TraiterElectrodes()
        {
            Document document = new Document();

            document.DocId = TSH.Documents.EditedDocument;

            PdmObjectId Projet = TSH.Pdm.GetProject(document.DocPdmObject);
           
            //List<PdmObjectId> pdmObjectIds = TSH.Pdm.GetProjectDocuments(Projet);







        }
    }
}
