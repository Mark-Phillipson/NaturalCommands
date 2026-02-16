using System;
using System.Drawing;
using System.Windows.Forms;
using NaturalCommands.Models;

namespace NaturalCommands
{
    /// <summary>
    /// Modal dialog to create a QuickClick region (name, click-type, monitor-scope).
    /// </summary>
    public class CreateQuickClickForm : Form
    {
        private readonly TextBox _txtName;
        private readonly ComboBox _cmbClickType;
        private readonly CheckBox _chkApplyToMonitor;
        private readonly Button _btnOk;
        private readonly Button _btnCancel;

        public string RegionName => _txtName.Text.Trim();
        public QuickClickClickType ClickType => Enum.TryParse<QuickClickClickType>(_cmbClickType.SelectedItem?.ToString() ?? "Left", true, out var ct) ? ct : QuickClickClickType.Left;
        public bool ApplyToMonitorOnly => _chkApplyToMonitor.Checked;

        public CreateQuickClickForm(string defaultName = "New Region", QuickClickClickType defaultClick = QuickClickClickType.Left, bool defaultApplyToMonitorOnly = true)
        {
            Text = "Create Quick Click";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(380, 150);

            var lblName = new Label { Text = "Name:", AutoSize = true, Location = new Point(12, 16) };
            _txtName = new TextBox { Location = new Point(80, 12), Width = 280, Text = defaultName };

            var lblType = new Label { Text = "Click type:", AutoSize = true, Location = new Point(12, 52) };
            _cmbClickType = new ComboBox { Location = new Point(80, 48), Width = 160, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbClickType.Items.AddRange(Enum.GetNames(typeof(QuickClickClickType)));
            _cmbClickType.SelectedItem = defaultClick.ToString();

            _chkApplyToMonitor = new CheckBox { Text = "Apply to this monitor only", Location = new Point(80, 82), AutoSize = true, Checked = defaultApplyToMonitorOnly };

            _btnOk = new Button { Text = "OK", DialogResult = DialogResult.OK, Location = new Point(200, 110), Size = new Size(75, 26) };
            _btnCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Location = new Point(285, 110), Size = new Size(75, 26) };

            Controls.Add(lblName);
            Controls.Add(_txtName);
            Controls.Add(lblType);
            Controls.Add(_cmbClickType);
            Controls.Add(_chkApplyToMonitor);
            Controls.Add(_btnOk);
            Controls.Add(_btnCancel);

            AcceptButton = _btnOk;
            CancelButton = _btnCancel;

            _txtName.SelectAll();
        }

        /// <summary>
        /// Show helper — centers on provided point when given (screen coords).
        /// </summary>
        public static bool ShowDialog(out string name, out QuickClickClickType clickType, out bool applyToMonitorOnly, System.Drawing.Point? centerOn = null, string defaultName = "New Region", QuickClickClickType defaultClick = QuickClickClickType.Left, bool defaultApplyToMonitorOnly = true)
        {
            using var dlg = new CreateQuickClickForm(defaultName, defaultClick, defaultApplyToMonitorOnly);
            if (centerOn.HasValue)
            {
                dlg.StartPosition = FormStartPosition.Manual;
                var sz = dlg.Size;
                dlg.Location = new Point(Math.Max(0, centerOn.Value.X - sz.Width / 2), Math.Max(0, centerOn.Value.Y - sz.Height / 2));
            }

            var dr = dlg.ShowDialog();
            name = dlg.RegionName;
            clickType = dlg.ClickType;
            applyToMonitorOnly = dlg.ApplyToMonitorOnly;
            return dr == DialogResult.OK;
        }
    }
}
