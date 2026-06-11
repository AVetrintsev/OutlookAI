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

        private TextBox _txtApiKey;
        private Label _lblCredentialStatus;
        private ComboBox _cmbReasoningEffort;
        private CheckedListBox _clbWriteTools;
        private Label _lblSaved;
        private CheckBox _chkLlmDebugLogEnabled;
        private TextBox _txtLlmDebugLogPath;
        private Label _lblDebugLogSaved;

        public SettingsForm()
            : this(Globals.ThisAddIn != null ? Globals.ThisAddIn.CredentialService : null)
        {
        }

        public SettingsForm(LiteLlmCredentialService credentials)
        {
            _credentials = credentials;

            Text = "Настройки OutlookAI";
            Size = new Size(460, 545);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            BackColor = Color.FromArgb(250, 249, 248);

            BuildSettingsPanel();

            if (_credentials != null)
            {
                _credentials.StatusChanged += OnCredentialStatusChanged;
                FormClosed += (s, e) => _credentials.StatusChanged -= OnCredentialStatusChanged;
            }
        }

        private void BuildSettingsPanel()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true
            };

            BuildConnectorGroup(panel);
            BuildAiBehaviorGroup(panel);
            BuildDebugLoggingGroup(panel);
            Controls.Add(panel);
            UpdateCredentialUi(GetCurrentStatus());
        }

        private void BuildConnectorGroup(Control parent)
        {
            var grp = new GroupBox
            {
                Text = "Коннектор LiteLLM",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Location = new Point(20, 15),
                Size = new Size(400, 170)
            };

            var lblEndpoint = new Label
            {
                Text = "Адрес: " + Config.NormalizeBaseUrl(Config.LiteLlmBaseUrl),
                Location = new Point(15, 25),
                Size = new Size(370, 20),
                Font = new Font("Segoe UI", 8.5F, FontStyle.Regular)
            };

            var lblModel = new Label
            {
                Text = "Модель: " + Config.Model,
                Location = new Point(15, 47),
                Size = new Size(370, 20),
                Font = new Font("Segoe UI", 8.5F, FontStyle.Regular)
            };

            var lblVoiceModel = new Label
            {
                Text = "Модель транскрибации: "
                    + (string.IsNullOrWhiteSpace(Config.VoiceModel) ? "(отключена)" : Config.VoiceModel),
                Location = new Point(15, 69),
                Size = new Size(370, 20),
                Font = new Font("Segoe UI", 8.5F, FontStyle.Regular)
            };

            var lblApiKey = new Label
            {
                Text = "Ключ API:",
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
                Text = "Сохранить",
                Location = new Point(295, 91),
                Width = 80
            };
            btnSave.Click += BtnSaveApiKey_Click;

            var btnClear = new Button
            {
                Text = "Удалить ключ",
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
            parent.Controls.Add(grp);
        }

        private void BuildAiBehaviorGroup(Control parent)
        {
            var grpAi = new GroupBox
            {
                Text = "Поведение AI",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Location = new Point(20, 200),
                Size = new Size(400, 155)
            };

            var lblReasoning = new Label
            {
                Text = "Уровень рассуждений:",
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
                Text = "Разрешённые действия:",
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
                Text = "Сохранить",
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
                Text = "Сохранено."
            };

            grpAi.Controls.AddRange(new Control[]
            {
                lblReasoning, _cmbReasoningEffort,
                lblWriteTools, _clbWriteTools,
                btnSaveAi, _lblSaved
            });
            parent.Controls.Add(grpAi);
        }

        private void BuildDebugLoggingGroup(Control parent)
        {
            var grp = new GroupBox
            {
                Text = "Отладочный лог LLM",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Location = new Point(20, 370),
                Size = new Size(400, 105)
            };

            _chkLlmDebugLogEnabled = new CheckBox
            {
                Text = "Логировать запросы и ответы LiteLLM",
                Location = new Point(15, 24),
                AutoSize = true,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Checked = Config.LlmDebugLogEnabled
            };

            _txtLlmDebugLogPath = new TextBox
            {
                Location = new Point(15, 52),
                Width = 280,
                Text = Config.LlmDebugLogPath ?? ""
            };

            var btnBrowse = new Button
            {
                Text = "...",
                Location = new Point(300, 50),
                Width = 32
            };
            btnBrowse.Click += BtnBrowseDebugLog_Click;

            var btnSave = new Button
            {
                Text = "Сохранить",
                Location = new Point(335, 50),
                Width = 55
            };
            btnSave.Click += BtnSaveDebugLog_Click;

            _lblDebugLogSaved = new Label
            {
                Location = new Point(15, 78),
                AutoSize = true,
                ForeColor = Color.DarkGreen,
                Font = new Font("Segoe UI", 8F, FontStyle.Italic),
                Visible = false,
                Text = "Сохранено."
            };

            grp.Controls.AddRange(new Control[]
            {
                _chkLlmDebugLogEnabled,
                _txtLlmDebugLogPath,
                btnBrowse,
                btnSave,
                _lblDebugLogSaved
            });
            parent.Controls.Add(grp);
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

        private void BtnSaveAiSettings_Click(object sender, EventArgs e)
        {
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

        private void BtnBrowseDebugLog_Click(object sender, EventArgs e)
        {
            using (var dlg = new SaveFileDialog())
            {
                dlg.Title = "Файл отладочного лога LLM";
                dlg.Filter = "Log files (*.log)|*.log|Text files (*.txt)|*.txt|All files (*.*)|*.*";
                dlg.FileName = string.IsNullOrWhiteSpace(_txtLlmDebugLogPath.Text)
                    ? "outlookai-llm-debug.log"
                    : System.IO.Path.GetFileName(_txtLlmDebugLogPath.Text);
                if (!string.IsNullOrWhiteSpace(_txtLlmDebugLogPath.Text))
                {
                    try
                    {
                        dlg.InitialDirectory = System.IO.Path.GetDirectoryName(_txtLlmDebugLogPath.Text);
                    }
                    catch { }
                }

                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    _txtLlmDebugLogPath.Text = dlg.FileName;
                    _chkLlmDebugLogEnabled.Checked = true;
                }
            }
        }

        private void BtnSaveDebugLog_Click(object sender, EventArgs e)
        {
            Config.LlmDebugLogEnabled = _chkLlmDebugLogEnabled.Checked;
            Config.LlmDebugLogPath = (_txtLlmDebugLogPath.Text ?? "").Trim();
            Config.SaveConfig();

            _lblDebugLogSaved.Visible = true;
            var t = new Timer { Interval = 2500 };
            t.Tick += (s, e2) => { _lblDebugLogSaved.Visible = false; t.Stop(); t.Dispose(); };
            t.Start();
        }

        private CredentialStatus GetCurrentStatus()
        {
            return _credentials != null
                ? _credentials.GetStatus()
                : CredentialStatus.Error("Сервис учётных данных недоступен");
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
                    _lblCredentialStatus.Text = "Ключ API настроен.";
                    break;
                case CredentialState.Error:
                    _lblCredentialStatus.ForeColor = Color.DarkRed;
                    _lblCredentialStatus.Text = "Ошибка учётных данных: " + status.Message;
                    break;
                default:
                    _lblCredentialStatus.ForeColor = Color.DarkSlateGray;
                    _lblCredentialStatus.Text = status.Message;
                    break;
            }
        }
    }
}
