using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using OutlookAI.Services;

namespace OutlookAI
{
    public sealed class SettingsForm : Form
    {
        private readonly LiteLlmCredentialService _credentials;

        private TextBox _txtPassword;
        private Label _lblError;
        private Panel _panelSettings;
        private TextBox _txtApiKey;
        private Label _lblCredentialStatus;
        private TextBox _txtNewPassword;
        private ComboBox _cmbReasoningEffort;
        private CheckedListBox _clbWriteTools;
        private Label _lblSaved;
        private bool _authenticated;

        public SettingsForm()
            : this(Globals.ThisAddIn != null ? Globals.ThisAddIn.CredentialService : null)
        {
        }

        public SettingsForm(LiteLlmCredentialService credentials)
        {
            _credentials = credentials;

            Text = "OutlookAI Settings";
            Size = new Size(460, 560);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            BackColor = Color.FromArgb(250, 249, 248);

            BuildLoginUi();
            BuildSettingsPanel();

            if (_credentials != null)
            {
                _credentials.StatusChanged += OnCredentialStatusChanged;
                FormClosed += (s, e) => _credentials.StatusChanged -= OnCredentialStatusChanged;
            }
        }

        private void BuildLoginUi()
        {
            var lblPassword = new Label
            {
                Text = "Admin Password:",
                Location = new Point(20, 20),
                AutoSize = true
            };

            _txtPassword = new TextBox
            {
                Location = new Point(20, 45),
                Width = 380,
                PasswordChar = '*'
            };

            var btnLogin = new Button
            {
                Text = "Login",
                Location = new Point(320, 75),
                Width = 80
            };
            btnLogin.Click += BtnLogin_Click;

            _lblError = new Label
            {
                Location = new Point(20, 80),
                AutoSize = true,
                ForeColor = Color.DarkRed,
                Visible = false
            };

            Controls.AddRange(new Control[] { lblPassword, _txtPassword, btnLogin, _lblError });
        }

        private void BuildSettingsPanel()
        {
            _panelSettings = new Panel
            {
                Location = new Point(0, 110),
                Size = new Size(460, 420),
                Visible = true,
                AutoScroll = true
            };

            BuildConnectorGroup();
            BuildAdminGroup();
            BuildAiBehaviorGroup();
            Controls.Add(_panelSettings);
            UpdateCredentialUi(GetCurrentStatus());
        }

        private void BuildConnectorGroup()
        {
            var grp = new GroupBox
            {
                Text = "LiteLLM Connector",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Location = new Point(20, 10),
                Size = new Size(400, 170)
            };

            var lblEndpoint = new Label
            {
                Text = "Endpoint: " + Config.NormalizeBaseUrl(Config.LiteLlmBaseUrl),
                Location = new Point(15, 25),
                Size = new Size(370, 20),
                Font = new Font("Segoe UI", 8.5F, FontStyle.Regular)
            };

            var lblModel = new Label
            {
                Text = "Model: " + Config.Model,
                Location = new Point(15, 47),
                Size = new Size(370, 20),
                Font = new Font("Segoe UI", 8.5F, FontStyle.Regular)
            };

            var lblVoiceModel = new Label
            {
                Text = "Voice model: " + (string.IsNullOrWhiteSpace(Config.VoiceModel) ? "(disabled)" : Config.VoiceModel),
                Location = new Point(15, 69),
                Size = new Size(370, 20),
                Font = new Font("Segoe UI", 8.5F, FontStyle.Regular)
            };

            var lblApiKey = new Label
            {
                Text = "API key:",
                Location = new Point(15, 96),
                AutoSize = true,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular)
            };

            _txtApiKey = new TextBox
            {
                Location = new Point(75, 93),
                Width = 210,
                PasswordChar = '*',
                Text = Config.LiteLlmApiKey ?? ""
            };

            var btnSave = new Button
            {
                Text = "Save",
                Location = new Point(295, 91),
                Width = 80
            };
            btnSave.Click += BtnSaveApiKey_Click;

            var btnClear = new Button
            {
                Text = "Clear Key",
                Location = new Point(295, 123),
                Width = 80
            };
            btnClear.Click += BtnClearApiKey_Click;

            _lblCredentialStatus = new Label
            {
                Location = new Point(15, 128),
                Size = new Size(270, 30),
                Font = new Font("Segoe UI", 8.5F, FontStyle.Regular)
            };

            grp.Controls.AddRange(new Control[]
            {
                lblEndpoint, lblModel, lblVoiceModel,
                lblApiKey, _txtApiKey, btnSave, btnClear, _lblCredentialStatus
            });
            _panelSettings.Controls.Add(grp);
        }

        private void BuildAdminGroup()
        {
            var lblNewPassword = new Label
            {
                Text = "New Admin Password (leave blank to keep):",
                Location = new Point(20, 195),
                AutoSize = true
            };

            _txtNewPassword = new TextBox
            {
                Location = new Point(20, 215),
                Width = 240,
                PasswordChar = '*'
            };

            var btnSavePassword = new Button
            {
                Text = "Save Password",
                Location = new Point(280, 213),
                Width = 120
            };
            btnSavePassword.Click += BtnSavePassword_Click;

            _panelSettings.Controls.AddRange(new Control[]
            {
                lblNewPassword, _txtNewPassword, btnSavePassword
            });
        }

        private void BuildAiBehaviorGroup()
        {
            var grpAi = new GroupBox
            {
                Text = "AI Behavior",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Location = new Point(20, 255),
                Size = new Size(400, 155)
            };

            var lblReasoning = new Label
            {
                Text = "Reasoning effort:",
                Location = new Point(15, 28),
                AutoSize = true,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular)
            };

            _cmbReasoningEffort = new ComboBox
            {
                Location = new Point(145, 25),
                Width = 230,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular)
            };
            _cmbReasoningEffort.Items.AddRange(Config.AvailableReasoningEfforts);
            var effortIdx = Array.IndexOf(Config.AvailableReasoningEfforts, Config.ReasoningEffort);
            _cmbReasoningEffort.SelectedIndex = effortIdx >= 0 ? effortIdx : 0;

            var lblWriteTools = new Label
            {
                Text = "Allowed write tools:",
                Location = new Point(15, 60),
                AutoSize = true,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular)
            };

            _clbWriteTools = new CheckedListBox
            {
                Location = new Point(145, 58),
                Size = new Size(230, 60),
                CheckOnClick = true,
                Font = new Font("Segoe UI", 8.5F, FontStyle.Regular),
                IntegralHeight = false
            };

            foreach (var tool in Config.AllWriteTools)
            {
                var index = _clbWriteTools.Items.Add(tool);
                _clbWriteTools.SetItemChecked(index, Config.EnabledWriteTools.Contains(tool));
            }

            var btnSaveAi = new Button
            {
                Text = "Save AI Settings",
                Location = new Point(255, 123),
                Width = 120
            };
            btnSaveAi.Click += BtnSaveAiSettings_Click;

            _lblSaved = new Label
            {
                Location = new Point(15, 128),
                AutoSize = true,
                ForeColor = Color.DarkGreen,
                Font = new Font("Segoe UI", 8F, FontStyle.Italic),
                Visible = false,
                Text = "Saved."
            };

            grpAi.Controls.AddRange(new Control[]
            {
                lblReasoning, _cmbReasoningEffort,
                lblWriteTools, _clbWriteTools,
                btnSaveAi, _lblSaved
            });
            _panelSettings.Controls.Add(grpAi);
        }

        private void BtnLogin_Click(object sender, EventArgs e)
        {
            if (_txtPassword.Text == Config.AdminPassword)
            {
                _authenticated = true;
                _lblError.Visible = false;
                _txtPassword.Enabled = false;
                UpdateCredentialUi(GetCurrentStatus());
            }
            else
            {
                _lblError.Text = "Invalid password";
                _lblError.Visible = true;
            }
        }

        private void BtnSaveApiKey_Click(object sender, EventArgs e)
        {
            if (_credentials == null)
            {
                return;
            }
            _credentials.SaveApiKey(_txtApiKey.Text);
            UpdateCredentialUi(_credentials.GetStatus());
        }

        private async void BtnClearApiKey_Click(object sender, EventArgs e)
        {
            if (_credentials == null)
            {
                return;
            }
            _txtApiKey.Text = "";
            await _credentials.ClearApiKeyAsync().ConfigureAwait(true);
            UpdateCredentialUi(_credentials.GetStatus());
        }

        private void BtnSavePassword_Click(object sender, EventArgs e)
        {
            if (!_authenticated || string.IsNullOrWhiteSpace(_txtNewPassword.Text))
            {
                return;
            }
            Config.AdminPassword = _txtNewPassword.Text;
            Config.SaveConfig();
            _txtNewPassword.Text = "";
            MessageBox.Show(
                this,
                "Admin password updated.",
                "OutlookAI Settings",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void BtnSaveAiSettings_Click(object sender, EventArgs e)
        {
            if (!_authenticated)
            {
                return;
            }

            var pickedEffort = _cmbReasoningEffort.SelectedItem as string;
            if (!string.IsNullOrEmpty(pickedEffort))
            {
                Config.ReasoningEffort = pickedEffort;
            }

            var newSet = new HashSet<string>(StringComparer.Ordinal);
            for (int i = 0; i < _clbWriteTools.Items.Count; i++)
            {
                if (_clbWriteTools.GetItemChecked(i))
                {
                    newSet.Add(_clbWriteTools.Items[i].ToString());
                }
            }
            Config.EnabledWriteTools = newSet;
            Config.WriteToolsEnabled = newSet.Count > 0;
            Config.SaveConfig();

            _lblSaved.Visible = true;
            var t = new Timer { Interval = 2500 };
            t.Tick += (s, e2) => { _lblSaved.Visible = false; t.Stop(); t.Dispose(); };
            t.Start();
        }

        private CredentialStatus GetCurrentStatus()
        {
            return _credentials != null
                ? _credentials.GetStatus()
                : CredentialStatus.Error("Credential service unavailable");
        }

        private void OnCredentialStatusChanged(object sender, CredentialStatus status)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => UpdateCredentialUi(status)));
                return;
            }
            UpdateCredentialUi(status);
        }

        private void UpdateCredentialUi(CredentialStatus status)
        {
            switch (status.State)
            {
                case CredentialState.Configured:
                    _lblCredentialStatus.ForeColor = Color.DarkGreen;
                    _lblCredentialStatus.Text = "API key configured.";
                    break;
                case CredentialState.Error:
                    _lblCredentialStatus.ForeColor = Color.DarkRed;
                    _lblCredentialStatus.Text = "Credential error: " + status.Message;
                    break;
                default:
                    _lblCredentialStatus.ForeColor = Color.DarkSlateGray;
                    _lblCredentialStatus.Text = status.Message;
                    break;
            }
        }
    }
}
