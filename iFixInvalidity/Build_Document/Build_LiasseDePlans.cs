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
    internal class Build_LiasseDePlans
    {
        
        public static void BuildLiasseDePlans(Document currentDoc)
        {
            
            var ProjetId = TSH.Pdm.GetProject(currentDoc.DocPdmObject);
           // var mainDocPdmId = TSH.Pdm.GetProjectMainDocument(ProjetId);
            var mainDocumentId = TSH.Documents.GetDocument(ProjetId);
            var namePdmProject = TSH.Parameters.GetNameParameter(mainDocumentId);

            TopSolidHost.Application.StartModification("Création paramètre relais nom projet", false);
            try
            {
                var currentDocId = currentDoc.DocId;
                TopSolidHost.Documents.EnsureIsDirty(ref currentDocId);

                ElementId relayParamId = TopSolidHost.Parameters.CreateTextRelayedParameter(
                    currentDoc.DocId,
                    namePdmProject,
                    ParameterRelayType.Project
                );

                TopSolidHost.Application.EndModification(true, true);
            }
            catch
            {
                TopSolidHost.Application.EndModification(false, false);
            }
        }
    }
}
