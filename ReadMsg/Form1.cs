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
                // Task.Run wird hier nicht benötigt, da Drag & Drop bereits im UI-Thread stattfindet
                // und createMSG SetLabel aufruft, das InvokeRequired prüft.
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
            string recipientsStr = recipients;
            string cleanSubject = CleanFileName(subject);
            string datum = sentOn.HasValue ? sentOn.Value.DateTime.ToString("yyMMdd") : "Datum unbekannt";
            string defaultFileName = $"{datum}_Email{type}_{cleanSubject}".Replace(" ", "_");

            string folderPath = GetFolderPath();
            if (string.IsNullOrEmpty(folderPath))
            {
                SetLabel("Abgebrochen!", Color.Red);
                return;
            }

            // Prompt for filename, potentially shortening it if default is too long
            string userFileName = PromptForFileName(defaultFileName, folderPath, attachments.Any());
            if (string.IsNullOrEmpty(userFileName))
            {
                SetLabel("Abgebrochen!", Color.Red);
                return;
            }

            // --- Vorläufige Pfade initialisieren (noch NICHT erstellen) ---
            // Diese Variablen werden in der Schleife angepasst
            string currentFinalFolderPath; // Der Ordner für die E-Mail (falls Anhänge)
            string currentOutputPdfPath;   // Der vollständige Pfad zur PDF-Datei

            if (attachments.Any())
            {
                currentFinalFolderPath = Path.Combine(folderPath, userFileName);
                currentOutputPdfPath = Path.Combine(currentFinalFolderPath, $"{userFileName}.pdf");
            }
            else
            {
                // Wenn keine Anhänge, ist der "finale Ordner" einfach der gewählte Speicherort
                currentFinalFolderPath = folderPath; // Dies ist der Basisordner, der von GetFolderPath gewählt wurde
                currentOutputPdfPath = Path.Combine(folderPath, $"{userFileName}.pdf");
            }

            // --- Pfadlängenprüfung und Benutzerinteraktion ---
            while (currentOutputPdfPath.Length > 250)
            {
                var retryResult = MessageBox.Show(
                    "Der gewählte Pfad ist weiterhin zu lang (>250 Zeichen).\nMöchten Sie einen anderen Ordner oder einen kürzeren Dateinamen wählen?",
                    "Pfad oder Dateiname zu lang",
                    MessageBoxButtons.RetryCancel,
                    MessageBoxIcon.Warning);

                if (retryResult == DialogResult.Retry)
                {
                    using (SaveFileDialog retryDialog = new SaveFileDialog())
                    {
                        // Der FileName im Dialog sollte immer den Namen der PDF enthalten, auch wenn der Pfad geändert wird
                        retryDialog.FileName = Path.GetFileName(currentOutputPdfPath);
                        retryDialog.Filter = "PDF-Datei|*.pdf";

                        // Berechne die erlaubten Zeichen für den Dateinamen basierend auf dem aktuellen Verzeichnis und Anhängen
                        string baseDir = Path.GetDirectoryName(currentOutputPdfPath) ?? folderPath;
                        int allowedChars;
                        if (attachments.Any())
                        {
                            // Ungleichung: baseDir.Length + 1 (für \) + name + 1 (für \) + name + 4 (.pdf) <= 250
                            // 2 * name <= 250 - baseDir.Length - 6
                            allowedChars = Math.Max(0, (250 - baseDir.Length - 6) / 2);
                        }
                        else
                        {
                            // Ungleichung: baseDir.Length + 1 (für \) + name + 4 (.pdf) <= 250
                            // name <= 250 - baseDir.Length - 5
                            allowedChars = Math.Max(0, 250 - baseDir.Length - 5);
                        }
                        retryDialog.Title = $"Wählen Sie einen kürzeren Speicherort oder Dateinamen (max. {allowedChars} Zeichen für den Dateinamen)";

                        try
                        {
                            retryDialog.InitialDirectory = Directory.Exists(baseDir) ? baseDir : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                        }
                        catch
                        {
                            retryDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                        }

                        if (retryDialog.ShowDialog() == DialogResult.OK)
                        {
                            // NEUE LOGIK HIER:
                            // Extrahiere den vom Benutzer gewählten BASIS-Ordner und den DATEINAMEN aus der Eingabe
                            string chosenBaseDirectory = Path.GetDirectoryName(retryDialog.FileName) ?? folderPath; // Der Ordner, in dem der E-Mail-Containerordner liegen soll
                            string chosenEmailName = Path.GetFileNameWithoutExtension(retryDialog.FileName); // Der Name für den E-Mail-Containerordner und die PDF

                            if (attachments.Any())
                            {
                                // Wenn Anhänge vorhanden sind, wird der E-Mail-Containerordner erstellt
                                currentFinalFolderPath = Path.Combine(chosenBaseDirectory, chosenEmailName);
                                // Die PDF wird INNERHALB dieses Containerordners gespeichert
                                currentOutputPdfPath = Path.Combine(currentFinalFolderPath, $"{chosenEmailName}.pdf");
                            }
                            else
                            {
                                // Wenn keine Anhänge, wird die PDF direkt im gewählten Basisordner gespeichert
                                currentFinalFolderPath = chosenBaseDirectory; // Kein separater Containerordner
                                currentOutputPdfPath = Path.Combine(chosenBaseDirectory, $"{chosenEmailName}.pdf");
                            }
                            continue; // Schleife erneut durchlaufen, um die Länge zu prüfen
                        }
                        else
                        {
                            // Benutzer hat die erneute Auswahl abgebrochen -> Vorgang abbrechen
                            SetLabel("Abgebrochen!", Color.Red);
                            return;
                        }
                    }
                }
                else
                {
                    // Benutzer hat die erste Warnung abgebrochen -> Vorgang abbrechen
                    SetLabel("Abgebrochen!", Color.Red);
                    return;
                }
            }

            // --- HIER beginnt die tatsächliche Dateisystem-Operation ---
            // Wenn wir hier ankommen, ist der Pfad gültig und der Benutzer hat nicht abgebrochen.

            List<string> attachmentNames = new List<string>();

            // 1. Hauptordner für die E-Mail erstellen (nur wenn Anhänge vorhanden)
            // Dieser Schritt wird nur ausgeführt, wenn es Anhänge gibt und der Pfad validiert wurde.
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
                Directory.CreateDirectory(attachmentsFolder); // Stellt sicher, dass sowohl currentFinalFolderPath als auch attachmentsFolder existieren

                foreach (var attachment in attachments)
                {
                    try
                    {
                        if (attachment is Storage.Attachment msgAttachment)
                        {
                            string attachmentFilePath = Path.Combine(attachmentsFolder, msgAttachment.FileName);
                            File.WriteAllBytes(attachmentFilePath, msgAttachment.Data);
                            attachmentNames.Add(msgAttachment.FileName);
                        }
                        else if (attachment is AttachmentWrapper emlAttachment)
                        {
                            string attachmentFilePath = Path.Combine(attachmentsFolder, emlAttachment.FileName);
                            using (var stream = File.Create(attachmentFilePath))
                            {
                                emlAttachment.MimePart.Content.DecodeTo(stream);
                            }
                            attachmentNames.Add(emlAttachment.FileName);
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
                AddBodyToPdf(bodyHtml, bodyText, bodyRtf, senderName, sentOnStr, recipientsStr, subject, senderEmail, currentOutputPdfPath, attaches);
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
            // Verwendet System.Windows.Forms.RichTextBox für die RTF-Konvertierung
            // Stelle sicher, dass dein Projekt eine Referenz auf System.Windows.Forms hat.
            using (var rtb = new RichTextBox())
            {
                try
                {
                    rtb.Rtf = rtf;
                    return rtb.Text;
                }
                catch (ArgumentException)
                {
                    // Behandle Fälle, in denen der RTF-String ungültig ist
                    return "RTF-Inhalt konnte nicht gelesen werden.";
                }
            }
        }

        private string GetFolderPath()
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Wählen Sie den Speicherort für die Dateien aus";
                folderDialog.ShowNewFolderButton = true; // Ermöglicht das Erstellen neuer Ordner direkt im Dialog

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    return folderDialog.SelectedPath;
                }
            }
            return null;
        }

        private string PromptForFileName(string defaultFileName, string folderPath, bool hasAttachments)
        {
            // Berechne die maximal erlaubten Zeichen für den Dateinamen-Teil
            int allowedChars;
            if (hasAttachments)
            {
                // Pfadstruktur: folderPath + "\" + userFileName (Ordner) + "\" + userFileName (PDF) + ".pdf"
                // Länge: folderPath.Length + 1 + userFileName.Length + 1 + userFileName.Length + 4
                // 2 * userFileName.Length <= 250 - folderPath.Length - 6
                allowedChars = Math.Max(0, (250 - folderPath.Length - 6) / 2);
            }
            else
            {
                // Pfadstruktur: folderPath + "\" + userFileName (PDF) + ".pdf"
                // Länge: folderPath.Length + 1 + userFileName.Length + 4
                // userFileName.Length <= 250 - folderPath.Length - 5
                allowedChars = Math.Max(0, 250 - folderPath.Length - 5);
            }

            // Prüfe, ob der Standard-Dateiname bereits zu lang für den erlaubten Bereich ist ODER
            // ob der vollständige Standardpfad (mit Standard-Dateinamen) 250 Zeichen überschreiten würde.
            bool defaultPathTooLong = false;
            if (hasAttachments)
            {
                defaultPathTooLong = (folderPath.Length + 1 + defaultFileName.Length + 1 + defaultFileName.Length + 4) > 250;
            }
            else
            {
                defaultPathTooLong = (folderPath.Length + 1 + defaultFileName.Length + 4) > 250;
            }


            if (defaultFileName.Length > allowedChars || defaultPathTooLong)
            {
                using (SaveFileDialog saveDialog = new SaveFileDialog())
                {
                    saveDialog.FileName = defaultFileName;
                    saveDialog.Filter = "PDF-Datei|*.pdf"; // Filter ist für die PDF, aber der Name ist für den Ordner/die Datei

                    saveDialog.Title = $"Geben Sie einen Namen für den Ordner bzw. die PDF-Datei ein (max. {allowedChars} Zeichen für den Dateinamen)";

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
            // Prüfe, ob es bereits HTML ist. Wenn nicht, HTML-kodieren und in <pre> einpacken.
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
            if (string.IsNullOrEmpty(fileName)) return "Unbenannt"; // Behandle null oder leere Betreffzeilen

            // Ungültige Zeichen durch Unterstrich ersetzen
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                fileName = fileName.Replace(c, '_');
            }
            // Ersetze auch Zeichen, die oft in Pfaden/Dateinamen problematisch sind, wie ':' '/' '\'
            fileName = fileName.Replace(':', '_').Replace('/', '_').Replace('\\', '_').Replace('|', '_').Replace('*', '_').Replace('?', '_').Replace('"', '_').Replace('<', '_').Replace('>', '_');
            // Führende/nachfolgende Leerzeichen oder Unterstriche entfernen
            fileName = fileName.Trim('_', ' ');
            // Sicherstellen, dass es nach der Bereinigung nicht leer ist
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