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
using System.IO;
using System.Net;
using System.Net.Cache;
using System.Reflection;
using System.Xml.Serialization;
using AutoUpdaterDotNET;

namespace iFixInvalidity
{
    internal static class Program
    {
        /// <summary>
        /// Point d'entrée principal de l'application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Controle de version avant toute fenetre : une version en retard ne doit meme pas
            // s'ouvrir sur un document.
            if (!PeutDemarrer()) return;

            Application.Run(new Form1());
        }

        /// <summary>Adresse du descripteur de mise a jour, sur le partage reseau.</summary>
        /// <remarks>
        /// Chemin UNC et non une lettre de lecteur, qui designerait pourtant le meme dossier : un
        /// mappage est propre a la session, et sur un poste ou il manque les mises a jour
        /// cesseraient sans que personne ne s'en apercoive.
        ///
        /// Cette adresse ne doit plus bouger. Chaque poste la porte en dur dans l'executable qu'il
        /// a installe : la deplacer coupe des mises a jour tous ceux deja en service, et sans le
        /// moindre signe, puisqu'un partage injoignable laisse l'application demarrer.
        /// </remarks>
        private const string DescripteurMiseAJour =
            @"\\jbtec-be\meca$\topsolid\iFixInvalidity\update.xml";

        /// <summary>
        /// Indique si l'application est autorisee a demarrer, apres controle de sa version.
        /// </summary>
        /// <remarks>
        /// La mise a jour peut etre refusee, mais l'application ne demarre pas tant qu'elle n'est
        /// pas faite : tous les postes doivent tourner sur la meme version, car la correction des
        /// invalidites ecrit dans le PDM et deux versions differentes n'y appliquent pas forcement
        /// les memes regles.
        ///
        /// Le controle est fait ici, avant toute fenetre, et non par <c>AutoUpdater.Start</c> :
        /// celui-ci mene le deroule de bout en bout et ne dit pas si l'utilisateur a refuse. Ses
        /// boites de dialogue s'affichent tres bien sans boucle de messages, elles pompent la leur.
        /// </remarks>
        private static bool PeutDemarrer()
        {
            UpdateInfoEventArgs descripteur;

            try
            {
                descripteur = LireDescripteurMiseAJour();
            }
            catch (Exception ex)
            {
                // Partage injoignable, poste hors reseau : on laisse l'atelier travailler plutot
                // que de l'immobiliser sur un incident de reseau. Ne pas savoir n'est pas la meme
                // chose que savoir qu'une version manque.
                Console.WriteLine($"[Mise a jour] Verification impossible : {ex.Message}");
                return true;
            }

            Version installee = Assembly.GetExecutingAssembly().GetName().Version;
            if (descripteur == null || string.IsNullOrWhiteSpace(descripteur.CurrentVersion)) return true;
            if (new Version(descripteur.CurrentVersion) <= installee) return true;

            DialogResult reponse = MessageBox.Show(
                $"La version {descripteur.CurrentVersion} est disponible ; ce poste utilise la {installee}.\n\n"
                + "iFixInvalidity ne peut pas s'ouvrir tant que la mise a jour n'est pas faite.",
                "Mise a jour requise",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Information);

            if (reponse != DialogResult.OK)
            {
                Console.WriteLine("[Mise a jour] Refusee : l'application ne demarre pas.");
                return false;
            }

            // L'installation se fait par utilisateur, dans %LOCALAPPDATA% : aucune elevation n'est
            // necessaire. Sans ce reglage, AutoUpdater lance le programme d'installation avec le
            // verbe « runas » et declenche une invite UAC pour rien.
            AutoUpdater.RunUpdateAsAdmin = false;

            // Rend la main une fois le programme d'installation lance : celui-ci remplace les
            // fichiers puis relance l'application, il ne faut donc pas la laisser demarrer ici.
            if (AutoUpdater.DownloadUpdate(descripteur)) return false;

            MessageBox.Show(
                "Le telechargement de la mise a jour n'a pas abouti.\n\n"
                + "Verifiez l'acces au reseau, puis relancez iFixInvalidity.",
                "Mise a jour",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);

            return false;
        }

        /// <summary>
        /// Lit le descripteur publie sur le partage.
        /// </summary>
        /// <remarks>
        /// Meme format et meme mecanisme que ceux d'AutoUpdater — WebClient sait lire une URI
        /// file://, donc un chemin UNC — mais lu ici pour garder la decision de demarrer.
        /// </remarks>
        private static UpdateInfoEventArgs LireDescripteurMiseAJour()
        {
            using (WebClient client = new WebClient())
            {
                // Sans cela, un update.xml fraichement publie peut rester masque par le cache le
                // temps que les postes le voient.
                client.CachePolicy = new RequestCachePolicy(RequestCacheLevel.NoCacheNoStore);

                string xml = client.DownloadString(new Uri(DescripteurMiseAJour));

                XmlSerializer serialiseur = new XmlSerializer(typeof(UpdateInfoEventArgs));
                using (StringReader lecteur = new StringReader(xml))
                {
                    return (UpdateInfoEventArgs)serialiseur.Deserialize(lecteur);
                }
            }
        }
    }
}
