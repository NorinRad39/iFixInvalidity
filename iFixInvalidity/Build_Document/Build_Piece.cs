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
    internal static class Build_Piece
    {
        private static readonly List<string> parametresTxt = new List<string> { "Indice 3D", "Matiere plan", "Nombre de pieces", "Traitement" };

        public static void BuildPiece(DocumentId currentDoc)
        {
            // Code to build the piece using the parameters in parametresTxt
            foreach (var parametreTxt in parametresTxt)
            {

                ElementId parametreTxtId = TSH.Parameters.CreateTextParameter(currentDoc, parametreTxt);
                SmartText SmartText = new SmartText("");
                TSH.Parameters.CreateSmartTextParameter(currentDoc, SmartText);
                TSH.Parameters.PublishText(currentDoc, parametreTxt, SmartText);
            }
        }
    }
}
