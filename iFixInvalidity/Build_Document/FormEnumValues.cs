using OutilsTs;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using TopSolid.Kernel.Automating;
using TSH = TopSolid.Kernel.Automating.TopSolidHost;

namespace iFixInvalidity.Build_Document
{
    /// <summary>
    /// Formulaire permettant à l'utilisateur de sélectionner une valeur textuelle
    /// depuis l'énumération <c>enumDocuType</c> du PDM TopSolid et de récupérer
    /// l'identifiant entier PDM correspondant.
    /// </summary>
    internal class FormEnumValues : Form
    {
        private Label _label;
        private ComboBox _comboBox;
        private Button _btnOK;
        private Button _btnAnnuler;

        private List<int> _intValues;
        private List<string> _textValues;

        /// <summary>Valeur entière (id PDM) de la sélection. Vaut -1 si annulé.</summary>
        internal int SelectedIntValue { get; private set; } = -1;

        /// <summary>Libellé textuel de la sélection. Vide si annulé.</summary>
        internal string SelectedTextValue { get; private set; } = string.Empty;

        internal FormEnumValues()
        {
            InitializeComponents();
            ChargerValeurs();
        }

        private void InitializeComponents()
        {
            Text = "Sélectionner un type de document";
            Width = 420;
            Height = 180;
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            _label = new Label
            {
                Text = "Type de document :",
                Left = 12,
                Top = 16,
                Width = 160,
                Font = new Font("Segoe UI", 10f)
            };

            _comboBox = new ComboBox
            {
                Left = 12,
                Top = 40,
                Width = 374,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 10f)
            };

            _btnOK = new Button
            {
                Text = "OK",
                Left = 230,
                Top = 90,
                Width = 75,
                Height = 30,
                DialogResult = DialogResult.OK
            };
            _btnOK.Click += BtnOK_Click;

            _btnAnnuler = new Button
            {
                Text = "Annuler",
                Left = 312,
                Top = 90,
                Width = 75,
                Height = 30,
                DialogResult = DialogResult.Cancel
            };

            AcceptButton = _btnOK;
            CancelButton = _btnAnnuler;

            Controls.AddRange(new Control[] { _label, _comboBox, _btnOK, _btnAnnuler });
        }

        private void ChargerValeurs()
        {
            try
            {
                DocumentId enumDocId = TrouverEnumDocuType();
                if (enumDocId.IsEmpty)
                {
                    MessageBox.Show("Le document 'enumDocuType' est introuvable dans le PDM.",
                        "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                BibliothequeEnum.GetEnumValuesFromDocument(enumDocId, out _intValues, out _textValues);

                if (_textValues != null)
                    foreach (var t in _textValues)
                        _comboBox.Items.Add(t);

                if (_comboBox.Items.Count > 0)
                    _comboBox.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors du chargement des valeurs : {ex.Message}",
                    "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Recherche le <see cref="DocumentId"/> du fichier <c>enumDocuType</c>
        /// dans la bibliothèque <c>docuType</c> du PDM.
        /// </summary>
        private DocumentId TrouverEnumDocuType()
        {
            PdmObjectId libDocuType = BibliothequeEnum.CheckLib();
            if (libDocuType.IsEmpty) return DocumentId.Empty;

            foreach (var doc in PDM.GetAllProjectDocuments(libDocuType))
            {
                if (TSH.Pdm.GetName(doc) == "enumDocuType")
                    return TSH.Documents.GetDocument(doc);
            }
            return DocumentId.Empty;
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            int idx = _comboBox.SelectedIndex;
            if (idx < 0 || _intValues == null || _textValues == null)
            {
                MessageBox.Show("Veuillez sélectionner une valeur.", "Sélection requise",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            SelectedIntValue = _intValues[idx];
            SelectedTextValue = _textValues[idx];
        }
    }
}