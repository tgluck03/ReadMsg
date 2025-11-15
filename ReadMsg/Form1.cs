using MsgReader.Outlook;
using MimeKit; // For .eml files
using System.Text;
using System.Drawing.Drawing2D;
using Image = System.Drawing.Image;
using Rectangle = System.Drawing.Rectangle;
using iText.Html2pdf;
using System.Diagnostics;
using System.Net;
using System.IO;
using System.Windows.Forms; // Notwendig für RichTextBox
using System.Linq; // Für LINQ-Erweiterungsmethoden wie .Any() und .Select()

namespace ReadMsg
{
    public partial class Form1 : Form
    {
        string filePath;

        public Form1()
        {
            // Stellt sicher, dass CodePagesEncodingProvider registriert ist, um ältere Encodings zu unterstützen
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            InitializeComponent();
            panel1.DragEnter += new DragEventHandler(panel1_DragEnter);
            panel1.DragDrop += new DragEventHandler(panel1_DragDrop);
            panel1.Paint += new PaintEventHandler(panel1_Paint);
            LoadImageToPictureBox();
            label1.TextAlign = ContentAlignment.MiddleCenter;
        }

        private void LoadImageToPictureBox()
        {
            string projectDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string imagePath = Path.Combine(projectDirectory, "images", "upload.png");

            if (File.Exists(imagePath))
            {
                pictureBox1.Image = Image.FromFile(imagePath);
            }
            else
            {
                MessageBox.Show("Bilddatei nicht gefunden: " + imagePath, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {
            int borderRadius = 20;
            int borderWidth = 2;
            Color borderColor = Color.White;

            Rectangle panelRect = panel1.ClientRectangle;
            panelRect.Inflate(-borderWidth, -borderWidth);

            GraphicsPath path = new GraphicsPath();
            path.AddArc(panelRect.X, panelRect.Y, borderRadius, borderRadius, 180, 90);
            path.AddArc(panelRect.Right - borderRadius, panelRect.Y, borderRadius, borderRadius, 270, 90);
            path.AddArc(panelRect.Right - borderRadius, panelRect.Bottom - borderRadius, borderRadius, borderRadius, 0, 90);
            path.AddArc(panelRect.X, panelRect.Bottom - borderRadius, borderRadius, borderRadius, 90, 90);
            path.CloseFigure();

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            using (Brush brush = new SolidBrush(panel1.BackColor))
            {
                e.Graphics.FillPath(brush, path);
            }

            using (Pen pen = new Pen(borderColor, borderWidth))
            {
                e.Graphics.DrawPath(pen, path);
            }
        }

        private void panel1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                string extension = Path.GetExtension(files[0]).ToLower();

                if (extension == ".msg" || extension == ".eml")
                {
                    e.Effect = DragDropEffects.Copy;
                }
                else
                {
                    e.Effect = DragDropEffects.None;
                }
            }
        }

        private void panel1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            filePath = files[0];
            button1.Text = Path.GetFileName(filePath);

            string extension = Path.GetExtension(filePath).ToLower();
            if (extension == ".msg")
            {
                createMSG();
            }
            else if (extension == ".eml")
            {
                createEML();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Email files (*.msg, *.eml)|*.msg;*.eml";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = openFileDialog.FileName;
                    button1.Text = Path.GetFileName(filePath);

                    string extension = Path.GetExtension(filePath).ToLower();
                    if (extension == ".msg")
                    {
                        Task.Run(() => createMSG());
                    }
                    else if (extension == ".eml")
                    {
                        Task.Run(() => createEML());
                    }
                }
            }
        }

        private void SetLabel(string text, Color color)
        {
            if (label1.InvokeRequired)
            {
                label1.Invoke((MethodInvoker)delegate
                {
                    label1.Text = text;
                    label1.ForeColor = color;
                });
            }
            else
            {
                label1.Text = text;
                label1.ForeColor = color;
            }
        }

        private void createMSG()
        {
            SetLabel("in Bearbeitung...", Color.Yellow);

            try
            {
                using (var msg = new Storage.Message(filePath))
                {
                    string recipients = string.Join(", ", msg.Recipients.Select(r => string.IsNullOrWhiteSpace(r.DisplayName) ? r.Email : $"{r.Email} ({r.DisplayName})"));
                    ProcessEmailContent(msg.Subject, msg.Sender.DisplayName, msg.Sender.Email, msg.SentOn, recipients, msg.BodyHtml, msg.BodyText, msg.BodyRtf, msg.Attachments);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Verarbeiten der MSG-Datei: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                SetLabel("Fehler!", Color.Red);
            }
        }

        private void createEML()
        {
            SetLabel("in Bearbeitung...", Color.Yellow);

            try
            {
                var message = MimeMessage.Load(filePath);
                string subject = message.Subject;
                string senderName = message.From.Mailboxes.FirstOrDefault()?.Name;
                string senderEmail = message.From.Mailboxes.FirstOrDefault()?.Address;
                DateTimeOffset? sentOn = message.Date;
                var recipients = message.To.Mailboxes.Select(m => $"{m.Name} ({m.Address})").ToList();
                string recipientsStr = string.Join(", ", recipients.Select(r => r.ToString()));
                string bodyHtml = message.HtmlBody;
                string bodyText = message.TextBody;

                // Filter attachments to include only MimePart objects
                var attachments = message.Attachments
                    .Where(a => a is MimePart)
                    .Select(a => new AttachmentWrapper((MimePart)a))
                    .ToList();

                ProcessEmailContent(subject, senderName, senderEmail, sentOn, recipientsStr, bodyHtml, bodyText, null, attachments);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Verarbeiten der EML-Datei: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                SetLabel("Fehler!", Color.Red);
            }
        }

        private void ProcessEmailContent(string subject, string senderName, string senderEmail, DateTimeOffset? sentOn, string recipients, string bodyHtml, string bodyText, string bodyRtf, IEnumerable<object> attachments)
        {
            string type = senderEmail.Contains("bc-wb.de") ? "A" : "E";
            string sentOnStr = sentOn.HasValue ? sentOn.Value.DateTime.ToString("dddd, d. MMMM yyyy HH:mm") : "Datum unbekannt";
            string cleanSubject = CleanFileName(subject);
            string datum = sentOn.HasValue ? sentOn.Value.DateTime.ToString("yyMMdd") : "Datum unbekannt";
            string defaultFileName = $"{datum}_Email{type}_{cleanSubject}".Replace(" ", "_");

            string folderPath = null;
            string userFileName = null;
            string currentFinalFolderPath = null; // Deklaration hier
            string currentOutputPdfPath = null;   // Deklaration hier
            bool needsPathReevaluation = true;

            while (needsPathReevaluation)
            {
                needsPathReevaluation = false;

                // 1. Basispfad für den Speicherort abrufen oder neu abrufen
                if (string.IsNullOrEmpty(folderPath)) // Prüfen, ob folderPath noch nicht gesetzt ist oder zurückgesetzt wurde
                {
                    folderPath = GetFolderPath();
                    if (string.IsNullOrEmpty(folderPath)) { SetLabel("Abgebrochen!", Color.Red); return; }
                }

                // 2. Benutzerdefinierten E-Mail-Ordner-/PDF-Namen abrufen oder neu abrufen
                // Wenn userFileName null ist (erste Iteration oder zurückgesetzt), wird PromptForFileName aufgerufen.
                // Ansonsten wird der bestehende userFileName verwendet.
                userFileName = PromptForFileName(userFileName ?? defaultFileName, folderPath, attachments.Any());
                if (string.IsNullOrEmpty(userFileName)) { SetLabel("Abgebrochen!", Color.Red); return; }

                // 3. Aktuellen finalen Ordnerpfad und PDF-Ausgabepfad konstruieren
                if (attachments.Any())
                {
                    currentFinalFolderPath = Path.Combine(folderPath, userFileName);
                    currentOutputPdfPath = Path.Combine(currentFinalFolderPath, $"{userFileName}.pdf");
                }
                else
                {
                    currentFinalFolderPath = folderPath;
                    currentOutputPdfPath = Path.Combine(folderPath, $"{userFileName}.pdf");
                }

                // 4. Pfadlänge der Haupt-PDF-Datei validieren
                if (currentOutputPdfPath.Length > 250)
                {
                    var retryResult = MessageBox.Show(
                        "Der gewählte Pfad für die E-Mail-PDF ist zu lang (>250 Zeichen).\nMöchten Sie einen anderen Ordner oder einen kürzeren Dateinamen wählen?",
                        "Pfad oder Dateiname zu lang",
                        MessageBoxButtons.RetryCancel,
                        MessageBoxIcon.Warning);

                    if (retryResult == DialogResult.Retry)
                    {
                        using (SaveFileDialog retryDialog = new SaveFileDialog())
                        {
                            retryDialog.FileName = Path.GetFileName(currentOutputPdfPath);
                            retryDialog.Filter = "PDF-Datei|*.pdf";

                            string initialDirectoryForDialog;
                            try
                            {
                                initialDirectoryForDialog = Path.GetDirectoryName(currentOutputPdfPath);
                                if (!Directory.Exists(initialDirectoryForDialog))
                                {
                                    initialDirectoryForDialog = folderPath;
                                }
                            }
                            catch { initialDirectoryForDialog = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); }
                            retryDialog.InitialDirectory = initialDirectoryForDialog;

                            int currentAllowedChars = CalculateAllowedChars(initialDirectoryForDialog, attachments.Any());

                            if (currentAllowedChars == 0)
                            {
                                retryDialog.Title = "Der Startordner ist bereits zu lang. Bitte wählen Sie einen DEUTLICH kürzeren Speicherort.";
                            }
                            else
                            {
                                retryDialog.Title = $"Wählen Sie einen kürzeren Speicherort oder Dateinamen (max. {currentAllowedChars} Zeichen für den Dateinamen)";
                            }

                            if (retryDialog.ShowDialog() == DialogResult.OK)
                            {
                                folderPath = Path.GetDirectoryName(retryDialog.FileName) ?? folderPath; // Update folderPath
                                userFileName = Path.GetFileNameWithoutExtension(retryDialog.FileName); // Update userFileName
                                needsPathReevaluation = true; // Hauptschleife neu starten, um alles neu zu validieren
                                continue;
                            }
                            else
                            {
                                SetLabel("Abgebrochen!", Color.Red); return;
                            }
                        }
                    }
                    else
                    {
                        SetLabel("Abgebrochen!", Color.Red); return;
                    }
                }

                // 5. Anhangspfade validieren (nur wenn Anhänge vorhanden sind)
                if (attachments.Any())
                {
                    string attachmentsBaseFolder = Path.Combine(currentFinalFolderPath, "Anhänge");
                    List<string> longAttachmentFileNames = new List<string>();

                    foreach (var attachment in attachments)
                    {
                        string attachmentFileName = "";
                        if (attachment is Storage.Attachment msgAtt)
                        {
                            attachmentFileName = msgAtt.FileName;
                        }
                        else if (attachment is AttachmentWrapper emlAtt)
                        {
                            attachmentFileName = emlAtt.FileName;
                        }
                        // Wenn attachmentFileName immer noch leer ist (unerwarteter Typ), überspringen.
                        if (string.IsNullOrEmpty(attachmentFileName)) continue;

                        string fullAttachmentPath = Path.Combine(attachmentsBaseFolder, attachmentFileName);

                        if (fullAttachmentPath.Length > 250)
                        {
                            longAttachmentFileNames.Add(attachmentFileName);
                        }
                    }

                    if (longAttachmentFileNames.Any())
                    {
                        string message = $"Die Pfade für folgende Anhänge sind zu lang (>250 Zeichen):\n\n" +
                                         $"{string.Join("\n", longAttachmentFileNames.Select(n => $" - {n}"))}\n\n" +
                                         "Möchten Sie trotzdem fortfahren?\n\n" +
                                         "(Wenn Sie fortfahren, könnten Anhänge nicht korrekt angezeigt werden.)";

                        var attachmentPathResult = MessageBox.Show(message,
                                                                   "Anhangspfad zu lang",
                                                                   MessageBoxButtons.YesNoCancel,
                                                                   MessageBoxIcon.Warning);

                        if (attachmentPathResult == DialogResult.No)
                        {
                            folderPath = null; // Basispfad-Auswahl erzwingen
                            userFileName = null; // E-Mail-Namen-Auswahl erzwingen
                            needsPathReevaluation = true; // Hauptschleife neu starten
                            continue;
                        }
                        else if (attachmentPathResult == DialogResult.Cancel)
                        {
                            SetLabel("Abgebrochen!", Color.Red);
                            return;
                        }
                    }
                }
            } // Ende der while (needsPathReevaluation)-Schleife

            // --- HIER beginnt die tatsächliche Dateisystem-Operation ---
            List<string> attachmentNames = new List<string>();

            // 1. Hauptordner für die E-Mail erstellen (nur wenn Anhänge vorhanden)
            if (attachments.Any())
            {
                if (!Directory.Exists(currentFinalFolderPath))
                {
                    Directory.CreateDirectory(currentFinalFolderPath);
                }
            }

            // 2. Anhänge speichern (erstellt auch den "Anhänge"-Unterordner, falls nötig)
            if (attachments.Any())
            {
                string attachmentsFolder = Path.Combine(currentFinalFolderPath, "Anhänge");
                Directory.CreateDirectory(attachmentsFolder);

                foreach (var attachment in attachments)
                {
                    try
                    {
                        string attachmentFileName = "";
                        if (attachment is Storage.Attachment msgAtt)
                        {
                            attachmentFileName = msgAtt.FileName;
                        }
                        else if (attachment is AttachmentWrapper emlAtt)
                        {
                            attachmentFileName = emlAtt.FileName;
                        }
                        if (string.IsNullOrEmpty(attachmentFileName)) continue;

                        string attachmentFilePath = Path.Combine(attachmentsFolder, attachmentFileName);

                        //// Erneute Pfadlängenprüfung vor dem Speichern, falls der Benutzer "Trotzdem fortfahren" gewählt hat
                        //if (attachmentFilePath.Length > 250)
                        //{
                        //    MessageBox.Show($"Der Pfad für den Anhang '{attachmentFileName}' ist immer noch zu lang ({attachmentFilePath.Length} Zeichen > 250). Das Speichern könnte fehlschlagen.",
                        //                    "Warnung: Anhangspfad zu lang", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        //}

                        if (attachment is Storage.Attachment msgAttToSave)
                        {
                            File.WriteAllBytes(attachmentFilePath, msgAttToSave.Data);
                            attachmentNames.Add(msgAttToSave.FileName);
                        }
                        else if (attachment is AttachmentWrapper emlAttToSave)
                        {
                            using (var stream = File.Create(attachmentFilePath))
                            {
                                emlAttToSave.MimePart.Content.DecodeTo(stream);
                            }
                            attachmentNames.Add(emlAttToSave.FileName);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Fehler beim Speichern des Anhangs '{attachment.GetType().GetProperty("FileName")?.GetValue(attachment) ?? "Unbekannt"}': {ex.Message}", "Fehler beim Anhang", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }

            string attaches = string.Join(", ", attachmentNames);
            // 3. PDF erstellen
            try
            {
                AddBodyToPdf(bodyHtml, bodyText, bodyRtf, senderName, sentOnStr, recipients, subject, senderEmail, currentOutputPdfPath, attaches);
                SetLabel("Abgeschlossen!", Color.LimeGreen);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Erstellen der PDF-Datei: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                SetLabel("Fehler!", Color.Red);
            }
        }

        private string ConvertRtfToText(string rtf)
        {
            using (var rtb = new RichTextBox())
            {
                try
                {
                    rtb.Rtf = rtf;
                    return rtb.Text;
                }
                catch (ArgumentException)
                {
                    return "RTF-Inhalt konnte nicht gelesen werden.";
                }
            }
        }

        private string GetFolderPath()
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Wählen Sie den Speicherort für die Dateien aus";
                folderDialog.ShowNewFolderButton = true;

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    return folderDialog.SelectedPath;
                }
            }
            return null;
        }

        // Hilfsmethode zur Berechnung der erlaubten Zeichen für den E-Mail-Ordner-/PDF-Namen
        private int CalculateAllowedChars(string basePath, bool hasAttachments)
        {
            if (string.IsNullOrEmpty(basePath)) return 0;

            if (hasAttachments)
            {
                // Voller Pfad wird sein: basePath + "\" + userFileName (Ordner) + "\" + userFileName (PDF) + ".pdf"
                // Benötigt: basePath.Length + 1 + userFileName.Length + 1 + userFileName.Length + 4 <= 250
                // Vereinfacht: 2 * userFileName.Length <= 250 - basePath.Length - 6
                return Math.Max(0, (250 - basePath.Length - 6) / 2);
            }
            else
            {
                // Voller Pfad wird sein: basePath + "\" + userFileName (PDF) + ".pdf"
                // Benötigt: basePath.Length + 1 + userFileName.Length + 4 <= 250
                // Vereinfacht: userFileName.Length <= 250 - basePath.Length - 5
                return Math.Max(0, 250 - basePath.Length - 5);
            }
        }

        private string PromptForFileName(string defaultFileName, string folderPath, bool hasAttachments)
        {
            int allowedChars = CalculateAllowedChars(folderPath, hasAttachments);

            string dialogTitle;
            if (allowedChars == 0)
            {
                dialogTitle = "Der gewählte Ordnerpfad ist bereits zu lang. Bitte wählen Sie einen DEUTLICH kürzeren Speicherort.";
            }
            else
            {
                dialogTitle = $"Geben Sie einen Namen für den Ordner bzw. die PDF-Datei ein (max. {allowedChars} Zeichen für den Dateinamen)";
            }

            bool defaultFileNameTooLong = defaultFileName.Length > allowedChars;
            bool defaultFullPathTooLong = false;

            string tempFullPath;
            if (hasAttachments)
            {
                tempFullPath = Path.Combine(folderPath, defaultFileName, $"{defaultFileName}.pdf");
            }
            else
            {
                tempFullPath = Path.Combine(folderPath, $"{defaultFileName}.pdf");
            }
            defaultFullPathTooLong = tempFullPath.Length > 250;

            // Der Dialog soll nur erscheinen, wenn der Standardname zu lang ist oder wenn allowedChars 0 ist
            // In diesem Fall, wenn allowedChars 0 ist, wird der Titel die entsprechende Meldung anzeigen.
            if (defaultFileNameTooLong || defaultFullPathTooLong || allowedChars == 0)
            {
                using (SaveFileDialog saveDialog = new SaveFileDialog())
                {
                    saveDialog.FileName = defaultFileName;
                    saveDialog.Filter = "PDF-Datei|*.pdf";
                    saveDialog.Title = dialogTitle;

                    try
                    {
                        saveDialog.InitialDirectory = Directory.Exists(folderPath) ? folderPath : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    }
                    catch
                    {
                        saveDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    }

                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        return Path.GetFileNameWithoutExtension(saveDialog.FileName);
                    }
                }
                return null; // Benutzer hat abgebrochen
            }

            return defaultFileName; // Standard-Dateiname ist in Ordnung
        }

        private void AddBodyToPdf(string bodyHtml, string bodyText, string bodyRtf, string n, string o, string r, string sub, string m, string output, string a)
        {
            string from = n + " (" + m + ")";
            string sentDate = o;
            string to = r;
            string subject = sub;

            string bodyContent = !string.IsNullOrEmpty(bodyHtml) ? bodyHtml : bodyText;

            if (string.IsNullOrEmpty(bodyContent) && !string.IsNullOrEmpty(bodyRtf))
            {
                bodyContent = ConvertRtfToText(bodyRtf);
            }

            if (string.IsNullOrEmpty(bodyContent))
            {
                bodyContent = "Kein Inhalt verfügbar.";
            }
            else if (!bodyContent.TrimStart().StartsWith("<") && !bodyContent.TrimStart().StartsWith("<!DOCTYPE", StringComparison.OrdinalIgnoreCase))
            {
                bodyContent = "<pre>" + WebUtility.HtmlEncode(bodyContent) + "</pre>";
            }

            string newContent = $@"
    <html>
        <head>
            <meta charset=""UTF-8"">
            <style>
                body {{ font-family: sans-serif; font-size: 10pt; }}
                table {{ border-collapse: collapse; width: 100%; margin-bottom: 15px; }}
                td {{ border: 1px solid #ddd; padding: 8px; vertical-align: top; }}
                th {{ border: 1px solid #ddd; padding: 8px; text-align: left; background-color: #f2f2f2; }}
                hr {{ border: 0; height: 1px; background-color: #ccc; margin: 20px 0; }}
                pre {{ white-space: pre-wrap; word-wrap: break-word; font-family: monospace; }}
            </style>
        </head>
        <body>
            <table cellpadding=""5"" cellspacing=""0"">
                <tr>
                    <td style=""width: 100px;""><strong>Von:</strong></td>
                    <td>{WebUtility.HtmlEncode(from)}</td>
                </tr>
                <tr>
                    <td><strong>Gesendet:</strong></td>
                    <td>{WebUtility.HtmlEncode(sentDate)}</td>
                </tr>
                <tr>
                    <td><strong>An:</strong></td>
                    <td>{WebUtility.HtmlEncode(to)}</td>
                </tr>
                <tr>
                    <td><strong>Betreff:</strong></td>
                    <td>{WebUtility.HtmlEncode(subject)}</td>
                </tr>
                <tr>
                    <td><strong>Anlagen:</strong></td>
                    <td>{WebUtility.HtmlEncode(a)}</td>
                </tr>
            </table>

            <hr>
            {bodyContent}
        </body>
    </html>";

            using (FileStream pdfStream = new FileStream(output, FileMode.Create, FileAccess.Write))
            {
                HtmlConverter.ConvertToPdf(newContent, pdfStream);
            }
        }

        private string CleanFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return "Unbenannt";

            foreach (char c in Path.GetInvalidFileNameChars())
            {
                fileName = fileName.Replace(c, '_');
            }
            fileName = fileName.Replace(':', '_').Replace('/', '_').Replace('\\', '_').Replace('|', '_').Replace('*', '_').Replace('?', '_').Replace('"', '_').Replace('<', '_').Replace('>', '_');
            fileName = fileName.Trim('_', ' ');
            if (string.IsNullOrEmpty(fileName)) return "Unbenannt";
            return fileName;
        }

        private void label2_Click(object sender, EventArgs e)
        {
            // Leerer Event-Handler, kann entfernt werden, wenn nicht verwendet
        }

        private void label3_Click(object sender, EventArgs e)
        {
            // Leerer Event-Handler, kann entfernt werden, wenn nicht verwendet
        }
    }

    public class AttachmentWrapper
    {
        public MimePart MimePart { get; }
        public string FileName => MimePart.FileName;

        public AttachmentWrapper(MimePart mimePart)
        {
            MimePart = mimePart ?? throw new ArgumentNullException(nameof(mimePart));
        }
    }
}