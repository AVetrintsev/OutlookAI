using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using OutlookAI.Services;
using OutlookAI.Services.CustomActions;

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
        private TabControl _tabSettings;
        private Panel _panelActions;
        private ListBox _lstCustomActions;
        private TextBox _txtActionTitle;
        private TextBox _txtActionDescription;
        private TextBox _txtActionPrompt;
        private ComboBox _cmbActionSource;
        private ComboBox _cmbActionFilter;
        private ComboBox _cmbActionPeriod;
        private NumericUpDown _numActionMaxItems;
        private CheckBox _chkActionFullBodies;
        private CheckBox _chkActionAttachments;
        private ComboBox _cmbActionOutput;
        private CheckBox _chkActionAllowTools;
        private Label _lblActionSaved;
        private readonly CustomActionStore _customActionStore = new CustomActionStore();
        private bool _authenticated;

        public SettingsForm()
            : this(Globals.ThisAddIn != null ? Globals.ThisAddIn.CredentialService : null)
        {
        }

        public SettingsForm(LiteLlmCredentialService credentials)
        {
            _credentials = credentials;

            Text = "Настройки OutlookAI";
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
                Text = "Пароль администратора:",
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
                Text = "Войти",
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
            _tabSettings = new TabControl
            {
                Location = new Point(0, 110),
                Size = new Size(460, 420),
                Visible = true
            };

            var tabMain = new TabPage("Основные");
            var tabActions = new TabPage("Действия");

            _panelSettings = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true
            };
            _panelActions = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true
            };

            tabMain.Controls.Add(_panelSettings);
            tabActions.Controls.Add(_panelActions);
            _tabSettings.TabPages.Add(tabMain);
            _tabSettings.TabPages.Add(tabActions);

            BuildConnectorGroup();
            BuildAdminGroup();
            BuildAiBehaviorGroup();
            BuildActionsTab();
            Controls.Add(_tabSettings);
            UpdateCredentialUi(GetCurrentStatus());
        }

        private void BuildConnectorGroup()
        {
            var grp = new GroupBox
            {
                Text = "Коннектор LiteLLM",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Location = new Point(20, 10),
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
                Text = "Модель транскрибации: " + (string.IsNullOrWhiteSpace(Config.VoiceModel) ? "(отключена)" : Config.VoiceModel),
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
            _panelSettings.Controls.Add(grp);
        }

        private void BuildAdminGroup()
        {
            var lblNewPassword = new Label
            {
                Text = "Новый пароль администратора (оставьте пустым, чтобы не менять):",
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
                Text = "Сохранить пароль",
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
                Text = "Поведение AI",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Location = new Point(20, 255),
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
                Text = "Сохранить настройки AI",
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
            _panelSettings.Controls.Add(grpAi);
        }

        private void BuildActionsTab()
        {
            _lstCustomActions = new ListBox
            {
                Location = new Point(15, 15),
                Size = new Size(400, 70),
                DisplayMember = "Title"
            };
            _lstCustomActions.SelectedIndexChanged += (s, e) => LoadSelectedActionIntoEditor();

            var lblTitle = new Label { Text = "Название:", Location = new Point(15, 95), AutoSize = true };
            _txtActionTitle = new TextBox { Location = new Point(120, 92), Width = 295 };

            var lblDescription = new Label { Text = "Описание:", Location = new Point(15, 123), AutoSize = true };
            _txtActionDescription = new TextBox { Location = new Point(120, 120), Width = 295 };

            var lblPrompt = new Label { Text = "Промпт:", Location = new Point(15, 151), AutoSize = true };
            _txtActionPrompt = new TextBox
            {
                Location = new Point(120, 148),
                Size = new Size(295, 70),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical
            };

            var lblSource = new Label { Text = "Источник:", Location = new Point(15, 228), AutoSize = true };
            _cmbActionSource = new ComboBox
            {
                Location = new Point(120, 225),
                Width = 130,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _cmbActionSource.Items.AddRange(new object[]
            {
                "current_open_message",
                "current_selection",
                "related_thread",
                "current_folder",
                "all_folders"
            });

            var lblFilter = new Label { Text = "Фильтр:", Location = new Point(260, 228), AutoSize = true };
            _cmbActionFilter = new ComboBox
            {
                Location = new Point(315, 225),
                Width = 100,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _cmbActionFilter.Items.AddRange(new object[] { "all", "unread", "read" });

            var lblPeriod = new Label { Text = "Период:", Location = new Point(15, 258), AutoSize = true };
            _cmbActionPeriod = new ComboBox
            {
                Location = new Point(120, 255),
                Width = 130,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _cmbActionPeriod.Items.AddRange(new object[]
            {
                "last_hour",
                "today",
                "yesterday",
                "since_last_run",
                "manual"
            });

            var lblMax = new Label { Text = "Max items:", Location = new Point(260, 258), AutoSize = true };
            _numActionMaxItems = new NumericUpDown
            {
                Location = new Point(335, 255),
                Width = 80,
                Minimum = 1,
                Maximum = 100,
                Value = 20
            };

            _chkActionFullBodies = new CheckBox
            {
                Text = "Полные тела писем",
                Location = new Point(120, 285),
                AutoSize = true
            };
            _chkActionAttachments = new CheckBox
            {
                Text = "Вложения/метаданные",
                Location = new Point(260, 285),
                AutoSize = true
            };

            var lblOutput = new Label { Text = "Результат:", Location = new Point(15, 315), AutoSize = true };
            _cmbActionOutput = new ComboBox
            {
                Location = new Point(120, 312),
                Width = 130,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _cmbActionOutput.Items.AddRange(new object[] { "chat", "create_draft", "export_pdf", "export_excel" });

            _chkActionAllowTools = new CheckBox
            {
                Text = "Разрешить tools",
                Location = new Point(260, 314),
                AutoSize = true
            };

            var btnNew = new Button { Text = "Новое", Location = new Point(120, 345), Width = 70 };
            btnNew.Click += (s, e) => ClearActionEditor();

            var btnSave = new Button { Text = "Сохранить", Location = new Point(200, 345), Width = 90 };
            btnSave.Click += BtnSaveCustomAction_Click;

            _lblActionSaved = new Label
            {
                Text = "Сохранено.",
                Location = new Point(300, 350),
                AutoSize = true,
                ForeColor = Color.DarkGreen,
                Visible = false
            };

            _panelActions.Controls.AddRange(new Control[]
            {
                _lstCustomActions,
                lblTitle, _txtActionTitle,
                lblDescription, _txtActionDescription,
                lblPrompt, _txtActionPrompt,
                lblSource, _cmbActionSource,
                lblFilter, _cmbActionFilter,
                lblPeriod, _cmbActionPeriod,
                lblMax, _numActionMaxItems,
                _chkActionFullBodies, _chkActionAttachments,
                lblOutput, _cmbActionOutput,
                _chkActionAllowTools,
                btnNew, btnSave, _lblActionSaved
            });

            ReloadCustomActions();
            ClearActionEditor();
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
                _lblError.Text = "Неверный пароль";
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
                "Пароль администратора обновлён.",
                "Настройки OutlookAI",
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

        private void ReloadCustomActions()
        {
            if (_lstCustomActions == null)
            {
                return;
            }

            _lstCustomActions.Items.Clear();
            foreach (var action in _customActionStore.Load())
            {
                _lstCustomActions.Items.Add(action);
            }
        }

        private void LoadSelectedActionIntoEditor()
        {
            var action = _lstCustomActions.SelectedItem as CustomActionDefinition;
            if (action == null)
            {
                return;
            }

            var context = action.Context ?? new CustomActionContext();
            _txtActionTitle.Text = action.Title ?? "";
            _txtActionDescription.Text = action.Description ?? "";
            _txtActionPrompt.Text = action.Prompt ?? "";
            SelectComboValue(_cmbActionSource, context.Source ?? "current_selection");
            SelectComboValue(_cmbActionFilter, context.ReadFilter ?? "all");
            SelectComboValue(_cmbActionPeriod, context.TimeRange ?? "today");
            var maxItems = (decimal)(context.MaxItems <= 0 ? 20 : context.MaxItems);
            _numActionMaxItems.Value = Math.Max(
                _numActionMaxItems.Minimum,
                Math.Min(_numActionMaxItems.Maximum, maxItems));
            _chkActionFullBodies.Checked = context.IncludeFullBodies;
            _chkActionAttachments.Checked = context.IncludeAttachments;
            SelectComboValue(_cmbActionOutput, action.Output ?? "chat");
            _chkActionAllowTools.Checked = action.AllowTools;
        }

        private void ClearActionEditor()
        {
            if (_txtActionTitle == null)
            {
                return;
            }

            _lstCustomActions.ClearSelected();
            _txtActionTitle.Text = "";
            _txtActionDescription.Text = "";
            _txtActionPrompt.Text = "";
            SelectComboValue(_cmbActionSource, "current_selection");
            SelectComboValue(_cmbActionFilter, "all");
            SelectComboValue(_cmbActionPeriod, "today");
            _numActionMaxItems.Value = 20;
            _chkActionFullBodies.Checked = true;
            _chkActionAttachments.Checked = false;
            SelectComboValue(_cmbActionOutput, "chat");
            _chkActionAllowTools.Checked = false;
        }

        private void BtnSaveCustomAction_Click(object sender, EventArgs e)
        {
            var title = (_txtActionTitle.Text ?? "").Trim();
            var prompt = (_txtActionPrompt.Text ?? "").Trim();
            if (string.IsNullOrWhiteSpace(title) || string.IsNullOrWhiteSpace(prompt))
            {
                MessageBox.Show(this, "Заполните название и промпт действия.", "OutlookAI",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var selected = _lstCustomActions.SelectedItem as CustomActionDefinition;
            var id = selected != null
                ? selected.Id
                : MakeActionId(title);

            var action = new CustomActionDefinition
            {
                Id = id,
                Title = title,
                Description = (_txtActionDescription.Text ?? "").Trim(),
                Prompt = prompt,
                Context = new CustomActionContext
                {
                    Source = ComboValue(_cmbActionSource, "current_selection"),
                    MessageScope = ComboValue(_cmbActionSource, "current_selection") == "related_thread" ? "thread" : "selected",
                    FolderScope = ComboValue(_cmbActionSource, "current_selection") == "all_folders" ? "all_folders" : "current_folder",
                    ReadFilter = ComboValue(_cmbActionFilter, "all"),
                    TimeRange = ComboValue(_cmbActionPeriod, "today"),
                    IncludeFullBodies = _chkActionFullBodies.Checked,
                    IncludeAttachments = _chkActionAttachments.Checked,
                    MaxItems = (int)_numActionMaxItems.Value
                },
                Output = ComboValue(_cmbActionOutput, "chat"),
                AllowTools = _chkActionAllowTools.Checked
            };

            _customActionStore.Upsert(action);
            ReloadCustomActions();
            _lblActionSaved.Visible = true;
            var t = new Timer { Interval = 2500 };
            t.Tick += (s, e2) => { _lblActionSaved.Visible = false; t.Stop(); t.Dispose(); };
            t.Start();
        }

        private static void SelectComboValue(ComboBox combo, string value)
        {
            if (combo == null)
            {
                return;
            }
            var idx = combo.Items.IndexOf(value);
            combo.SelectedIndex = idx >= 0 ? idx : 0;
        }

        private static string ComboValue(ComboBox combo, string fallback)
        {
            return combo != null && combo.SelectedItem != null
                ? combo.SelectedItem.ToString()
                : fallback;
        }

        private static string MakeActionId(string title)
        {
            var chars = (title ?? "").Trim().ToLowerInvariant()
                .Select(ch => char.IsLetterOrDigit(ch) ? ch : '_')
                .ToArray();
            var id = new string(chars).Trim('_');
            while (id.Contains("__"))
            {
                id = id.Replace("__", "_");
            }
            return string.IsNullOrWhiteSpace(id)
                ? "custom_action_" + DateTime.UtcNow.ToString("yyyyMMddHHmmss")
                : id;
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
