using MsgReader.Outlook;
using MimeKit; // For .eml files
using System.Text;
using System.Drawing.Drawing2D;
using Image = System.Drawing.Image;
using Rectangle = System.Drawing.Rectangle;
using iText.Html2pdf;
using System.Diagnostics;
using System.Net;
using RtfPipe;
using System.IO;

namespace ReadMsg
{
    public partial class Form1 : Form
    {
        string filePath;

        public Form1()
        {
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
                MessageBox.Show("Bilddatei nicht gefunden: " + imagePath);
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

            using (var msg = new Storage.Message(filePath))
            {
                string recipients = string.Join(", ", msg.Recipients.Select(r => string.IsNullOrWhiteSpace(r.DisplayName) ? r.Email : $"{r.Email} ({r.DisplayName})"));
                ProcessEmailContent(msg.Subject, msg.Sender.DisplayName, msg.Sender.Email, msg.SentOn, recipients, msg.BodyHtml, msg.BodyText, msg.BodyRtf, msg.Attachments);
            }

            SetLabel("Abgeschlossen!", Color.LimeGreen);
        }

        private void createEML()
        {
            SetLabel("in Bearbeitung...", Color.Yellow);

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

            SetLabel("Abgeschlossen!", Color.LimeGreen);
        }

        private void ProcessEmailContent(string subject, string senderName, string senderEmail, DateTimeOffset? sentOn, string recipients, string bodyHtml, string bodyText, string bodyRtf, IEnumerable<object> attachments)
        {
            string sentOnStr = sentOn.HasValue ? sentOn.Value.DateTime.ToString("dddd, d. MMMM yyyy HH:mm") : "Datum unbekannt";
            string recipientsStr = recipients;
            string cleanSubject = CleanFileName(subject);
            string datum = sentOn.HasValue ? sentOn.Value.DateTime.ToString("yyMMdd") : "Datum unbekannt";
            string defaultFileName = $"{datum}_Email_{cleanSubject}".Replace(" ", "_");

            string folderPath = GetFolderPath();
            if (string.IsNullOrEmpty(folderPath))
            {
                SetLabel("Abgebrochen!", Color.Red);
                return;
            }

            string userFileName = PromptForFileName(defaultFileName);
            if (string.IsNullOrEmpty(userFileName))
            {
                SetLabel("Abgebrochen!", Color.Red);
                return;
            }

            string finalFolderPath = Path.Combine(folderPath, userFileName);
            string pdfFileName = $"{userFileName}.pdf";
            string outputPdfPath;

            if (attachments.Any())
            {
                if (!Directory.Exists(finalFolderPath))
                {
                    Directory.CreateDirectory(finalFolderPath);
                }
                outputPdfPath = Path.Combine(finalFolderPath, pdfFileName);
            }
            else
            {
                outputPdfPath = Path.Combine(folderPath, pdfFileName);
            }

            List<string> attachmentNames = new List<string>();

            if (attachments.Any())
            {
                string attachmentsFolder = Path.Combine(finalFolderPath, "Anhänge");
                Directory.CreateDirectory(attachmentsFolder);

                foreach (var attachment in attachments)
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
            }

            string attaches = string.Join(", ", attachmentNames);
            AddBodyToPdf(bodyHtml, bodyText, bodyRtf, senderName, sentOnStr, recipientsStr, subject, senderEmail, outputPdfPath, attaches);
        }

        private string ConvertRtfToText(string rtf)
        {
            using (var rtb = new RichTextBox())
            {
                rtb.Rtf = rtf;
                return rtb.Text;
            }
        }

        private string GetFolderPath()
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Wählen Sie den Speicherort für die Dateien aus";

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    return folderDialog.SelectedPath;
                }
            }
            return null;
        }

        private string PromptForFileName(string defaultFileName)
        {
            if (defaultFileName.Length > 50)
            {
                using (SaveFileDialog saveDialog = new SaveFileDialog())
                {
                    saveDialog.FileName = defaultFileName;
                    saveDialog.Filter = "PDF-Datei|*.pdf";
                    saveDialog.Title = "Geben Sie einen Namen für den Ordner bzw. die PDF-Datei ein";

                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        return Path.GetFileNameWithoutExtension(saveDialog.FileName);
                    }
                }
                return null;
            }

            return defaultFileName;
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
            else if (!bodyContent.TrimStart().StartsWith("<"))
            {
                bodyContent = "<pre>" + WebUtility.HtmlEncode(bodyContent) + "</pre>";
            }

            string newContent = $@"
    <html>
        <body>
            <table cellpadding=""5"" cellspacing=""0"" style=""border-collapse: collapse; width: 100%;"">
                <tr>
                    <td><strong>Von:</strong></td>
                    <td>{from}</td>
                </tr>
                <tr>
                    <td><strong>Gesendet:</strong></td>
                    <td>{sentDate}</td>
                </tr>
                <tr>
                    <td><strong>An:</strong></td>
                    <td>{to}</td>
                </tr>
                <tr>
                    <td><strong>Betreff:</strong></td>
                    <td>{subject}</td>
                </tr>
                <tr>
                    <td><strong>Anlagen:</strong></td>
                    <td>{a}</td>
                </tr>
            </table>

            <hr>
            <p>{bodyContent}</p>
        </body>
    </html>";

            using (FileStream pdfStream = new FileStream(output, FileMode.Create, FileAccess.Write))
            {
                HtmlConverter.ConvertToPdf(newContent, pdfStream);
            }
        }

        private string CleanFileName(string fileName)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                fileName = fileName.Replace(c, '_');
            }
            return fileName;
        }

        private void label2_Click(object sender, EventArgs e)
        {

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