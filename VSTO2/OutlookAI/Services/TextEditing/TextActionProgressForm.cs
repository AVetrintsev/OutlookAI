using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace OutlookAI.Services.TextEditing
{
    /// <summary>
    /// Small modeless status window used while a selected-text action runs.
    /// The controller owns all threading decisions; this form is only touched
    /// from the Outlook UI thread.
    /// </summary>
    public sealed class TextActionProgressForm : Form
    {
        private readonly Label _titleLabel;
        private readonly Label _statusLabel;
        private readonly ProgressBar _progressBar;
        private readonly Button _cancelButton;
        private readonly Button _undoButton;
        private readonly Button _closeButton;
        private readonly TextBox _loadedSkills;
        private bool _running;

        public TextActionProgressForm()
        {
            Text = "OutlookAI";
            AutoScaleDimensions = new SizeF(96F, 96F);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(440, 174);
            MinimumSize = new Size(400, 210);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowIcon = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);

            _titleLabel = new Label
            {
                AutoEllipsis = true,
                Dock = DockStyle.Top,
                Font = new Font(Font, FontStyle.Bold),
                Height = 24,
                TextAlign = ContentAlignment.MiddleLeft
            };
            _statusLabel = new Label
            {
                AutoEllipsis = true,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.TopLeft,
                Padding = new Padding(0, 6, 0, 0)
            };
            _progressBar = new ProgressBar
            {
                Dock = DockStyle.Bottom,
                Height = 8,
                MarqueeAnimationSpeed = 24,
                Style = ProgressBarStyle.Marquee
            };
            _loadedSkills = new TextBox
            {
                Name = "loadedSkills",
                AccessibleName = "Загруженные скиллы",
                Dock = DockStyle.Fill,
                Height = 66,
                ReadOnly = true,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                TabStop = false,
                Visible = false
            };

            _cancelButton = new Button
            {
                AutoSize = true,
                Text = "Отмена",
                UseVisualStyleBackColor = true
            };
            _undoButton = new Button
            {
                AutoSize = true,
                Text = "Отменить изменение",
                UseVisualStyleBackColor = true,
                Visible = false
            };
            _closeButton = new Button
            {
                AutoSize = true,
                Text = "Закрыть",
                UseVisualStyleBackColor = true,
                Visible = false
            };

            _cancelButton.Click += (sender, args) => RaiseCancelRequested();
            _undoButton.Click += (sender, args) =>
            {
                var handler = UndoRequested;
                if (handler != null) handler(this, EventArgs.Empty);
            };
            _closeButton.Click += (sender, args) => Close();

            var content = new TableLayoutPanel
            {
                ColumnCount = 1,
                Dock = DockStyle.Fill,
                Padding = new Padding(18, 14, 18, 12),
                RowCount = 4
            };
            content.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            content.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            content.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            content.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var statusPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 4, 0, 12)
            };
            statusPanel.Controls.Add(_statusLabel);
            statusPanel.Controls.Add(_progressBar);

            var buttons = new FlowLayoutPanel
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                Margin = new Padding(0),
                WrapContents = false
            };
            buttons.Controls.Add(_closeButton);
            buttons.Controls.Add(_undoButton);
            buttons.Controls.Add(_cancelButton);

            content.Controls.Add(_titleLabel, 0, 0);
            content.Controls.Add(statusPanel, 0, 1);
            content.Controls.Add(_loadedSkills, 0, 2);
            content.Controls.Add(buttons, 0, 3);
            Controls.Add(content);

            CancelButton = _cancelButton;
            ShowRunning(null);
        }

        public event EventHandler CancelRequested;

        public event EventHandler UndoRequested;

        public bool IsRunning => _running;

        public void ShowLoadedSkills(IEnumerable<string> names)
        {
            var values = (names ?? new string[0]).Where(name => !string.IsNullOrWhiteSpace(name))
                .Select(name => name.Replace('\r', ' ').Replace('\n', ' ').Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            var show = values.Length > 0;
            var wasShown = _loadedSkills.Text.Length > 0;
            _loadedSkills.Text = show ? "Загружены скиллы: " + string.Join(", ", values) : "";
            _loadedSkills.Visible = show;
            if (show != wasShown) ClientSize = new Size(ClientSize.Width, ClientSize.Height + (show ? 72 : -72));
            // Do not call Activate/Focus: the Word selection and native Ctrl+Z keep their focus.
        }

        public void ShowRunning(string actionTitle)
        {
            _running = true;
            _titleLabel.Text = string.IsNullOrWhiteSpace(actionTitle)
                ? "AI-ассистент"
                : actionTitle;
            _statusLabel.Text = "Обрабатываю выделенный текст…";
            _progressBar.Style = ProgressBarStyle.Marquee;
            _progressBar.Visible = true;
            _cancelButton.Enabled = true;
            _cancelButton.Visible = true;
            _undoButton.Enabled = false;
            _undoButton.Visible = false;
            _closeButton.Visible = false;
            CancelButton = _cancelButton;
        }

        public void ShowCompleted()
        {
            _running = false;
            _statusLabel.Text = "Текст заменён, выделение сохранено. Ctrl+Z в письме отменяет всю правку вместе с форматированием.";
            _progressBar.Visible = false;
            _cancelButton.Visible = false;
            _undoButton.Enabled = true;
            _undoButton.Visible = true;
            _closeButton.Visible = true;
            CancelButton = _closeButton;
        }

        public void ShowFinalizing()
        {
            _running = false;
            _statusLabel.Text = "Текст заменён. Сохраняю запись в истории…";
            _progressBar.Style = ProgressBarStyle.Marquee;
            _progressBar.Visible = true;
            _cancelButton.Visible = false;
            _undoButton.Visible = false;
            _closeButton.Visible = false;
            CancelButton = null;
        }

        public void ShowNoChange()
        {
            _running = false;
            _statusLabel.Text = "Текст уже соответствует выбранному действию и не был изменён.";
            _progressBar.Visible = false;
            _cancelButton.Visible = false;
            _undoButton.Visible = false;
            _closeButton.Visible = true;
            CancelButton = _closeButton;
            Activate();
        }

        public void ShowUndoInProgress()
        {
            _running = true;
            _statusLabel.Text = "Отменяю изменение…";
            _progressBar.Style = ProgressBarStyle.Marquee;
            _progressBar.Visible = true;
            _cancelButton.Visible = false;
            _undoButton.Enabled = false;
            _closeButton.Visible = false;
            CancelButton = null;
        }

        public void ShowUndoCompleted(bool historyUpdated)
        {
            _running = false;
            _statusLabel.Text = historyUpdated
                ? "Изменение отменено и отмечено в истории."
                : "Изменение отменено. Не удалось обновить запись в истории.";
            _progressBar.Visible = false;
            _cancelButton.Visible = false;
            _undoButton.Visible = false;
            _closeButton.Visible = true;
            CancelButton = _closeButton;
            Activate();
        }

        public void ShowUndoUnavailable(string message)
        {
            _running = false;
            _statusLabel.Text = string.IsNullOrWhiteSpace(message)
                ? "Не удалось отменить изменение."
                : message;
            _progressBar.Visible = false;
            _cancelButton.Visible = false;
            _undoButton.Enabled = false;
            _undoButton.Visible = false;
            _closeButton.Visible = true;
            CancelButton = _closeButton;
            Activate();
        }

        public void ShowCancelled()
        {
            _running = false;
            _statusLabel.Text = "Обработка отменена. Исходный текст не изменён.";
            _progressBar.Visible = false;
            _cancelButton.Visible = false;
            _undoButton.Visible = false;
            _closeButton.Visible = true;
            CancelButton = _closeButton;
        }

        public void ShowFailure(string message)
        {
            _running = false;
            _statusLabel.Text = string.IsNullOrWhiteSpace(message)
                ? "Не удалось обработать выделенный текст."
                : message;
            _progressBar.Visible = false;
            _cancelButton.Visible = false;
            _undoButton.Visible = false;
            _closeButton.Visible = true;
            CancelButton = _closeButton;
            Activate();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_running) RaiseCancelRequested();
            base.OnFormClosing(e);
        }

        private void RaiseCancelRequested()
        {
            if (!_running) return;
            _cancelButton.Enabled = false;
            _statusLabel.Text = "Отменяю обработку…";
            var handler = CancelRequested;
            if (handler != null) handler(this, EventArgs.Empty);
        }
    }
}
