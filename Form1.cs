using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Printing;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using Newtonsoft.Json;
using System.IO;
using iText.Layout.Properties;

namespace Matiz2024
{
    public partial class Form1 : Form
    {
        List<string> labelsOnPage = new List<string>(new string[24]);

        private const int LabelsPerRow = 3;
        private const int LabelsPerColumn = 8;
        private PrintDocument printDocument = new PrintDocument();
        private PrintPreviewDialog printPreviewDialog = new PrintPreviewDialog();

    public Form1()
        {
            InitializeComponent();

            // Initialize PrintDocument and attach the PrintPage event
            printDocument.PrintPage += PrintPage;

            // Initialize PrintPreviewDialog
            printPreviewDialog.Document = printDocument;

            // Attach events
            saveSablonButton.Click += SaveSablon;
            loadSablonButton.Click += LoadSablon;
            addToA4Button.Click += AddLabelToA4;
            printButton.Click += PrintButton_Click;

            string tooltiptext = "Sačuvaj";
            string tooltiptext1 = "Obriši";

            // Attach events
            toolTip1.SetToolTip(saveUvoznikButton, tooltiptext);
            toolTip1.SetToolTip(saveUverenjeButton, tooltiptext);
            toolTip1.SetToolTip(saveSrpsButton, tooltiptext);
            toolTip1.SetToolTip(saveProizvodjacButton, tooltiptext);
            toolTip1.SetToolTip(savePostavaButton, tooltiptext);
            toolTip1.SetToolTip(savePorekloButton, tooltiptext);
            toolTip1.SetToolTip(saveOdrzavanjeButton, tooltiptext);
            toolTip1.SetToolTip(saveNamenaButton, tooltiptext);
            toolTip1.SetToolTip(saveLiceButton, tooltiptext);
            toolTip1.SetToolTip(saveIzradaButton, tooltiptext);
            toolTip1.SetToolTip(saveDjonButton, tooltiptext);
            toolTip1.SetToolTip(saveArtikalButton, tooltiptext);
            toolTip1.SetToolTip(saveNazivButton, tooltiptext);

            toolTip2.SetToolTip(removeUvoznikButton, tooltiptext1);
            toolTip2.SetToolTip(removeUverenjeButton, tooltiptext1);
            toolTip2.SetToolTip(removeSrpsButton, tooltiptext1);
            toolTip2.SetToolTip(removeProizvodjacButton, tooltiptext1);
            toolTip2.SetToolTip(removePostavaButton, tooltiptext1);
            toolTip2.SetToolTip(removePorekloButton, tooltiptext1);
            toolTip2.SetToolTip(removeOdrzavanjeButton, tooltiptext1);
            toolTip2.SetToolTip(removeNamenaButton, tooltiptext1);
            toolTip2.SetToolTip(removeLiceButton, tooltiptext1);
            toolTip2.SetToolTip(removeIzradaButton, tooltiptext1);
            toolTip2.SetToolTip(removeArtikalButton, tooltiptext1);
            toolTip2.SetToolTip(removeNazivButton, tooltiptext1);
            toolTip2.SetToolTip(removeDjonButton, tooltiptext1);

            comboBoxPoreklo.SelectedIndexChanged += (s, ev) => UpdateLabelCloseup();
            comboBoxUvoznik.SelectedIndexChanged += (s, ev) => UpdateLabelCloseup();
            comboBoxProizvodjac.SelectedIndexChanged += (s, ev) => UpdateLabelCloseup();
            comboBoxUverenje.SelectedIndexChanged += (s, ev) => UpdateLabelCloseup();
            comboBoxNaziv.SelectedIndexChanged += (s, ev) => UpdateLabelCloseup();
            comboBoxArtikal.SelectedIndexChanged += (s, ev) => UpdateLabelCloseup();
            comboBoxLice.SelectedIndexChanged += (s, ev) => UpdateLabelCloseup();
            comboBoxPostava.SelectedIndexChanged += (s, ev) => UpdateLabelCloseup();
            comboBoxDjon.SelectedIndexChanged += (s, ev) => UpdateLabelCloseup();
            comboBoxSrps.SelectedIndexChanged += (s, ev) => UpdateLabelCloseup();
            comboBoxIzrada.SelectedIndexChanged += (s, ev) => UpdateLabelCloseup();
            comboBoxNamena.SelectedIndexChanged += (s, ev) => UpdateLabelCloseup();
            comboBoxOdrzavanje.SelectedIndexChanged += (s, ev) => UpdateLabelCloseup();

            // Attach events

            saveUvoznikButton.Click += saveUvoznikButton_Click;
            saveUverenjeButton.Click += saveUverenjeButton_Click;
            saveSrpsButton.Click += saveSrpsButton_Click;
            saveProizvodjacButton.Click += saveProizvodjacButton_Click;
            savePostavaButton.Click += savePostavaButton_Click;
            savePorekloButton.Click += savePorekloButton_Click;
            saveOdrzavanjeButton.Click += saveOdrzavanjeButton_Click;
            saveNamenaButton.Click += saveNamenaButton_Click;
            saveLiceButton.Click += saveLiceButton_Click;
            saveIzradaButton.Click += saveIzradaButton_Click;
            saveDjonButton.Click += saveDjonButton_Click;
            saveArtikalButton.Click += saveArtikalButton_Click;
            saveNazivButton.Click += saveNazivButton_Click;

            removeUvoznikButton.Click += removeUvoznikButton_Click;
            removeUverenjeButton.Click += removeUverenjeButton_Click;
            removeSrpsButton.Click += removeSrpsButton_Click;
            removeProizvodjacButton.Click += removeProizvodjacButton_Click;
            removePostavaButton.Click += removePostavaButton_Click;
            removePorekloButton.Click += removePorekloButton_Click;
            removeOdrzavanjeButton.Click += removeOdrzavanjeButton_Click;
            removeNamenaButton.Click += removeNamenaButton_Click;
            removeLiceButton.Click += removeLiceButton_Click;
            removeIzradaButton.Click += removeIzradaButton_Click;
            removeDjonButton.Click += removeDjonButton_Click;
            removeArtikalButton.Click += removeArtikalButton_Click;
            removeNazivButton.Click += removeNazivButton_Click;

            deleteLabelButton.Click += DeleteLabelButton_Click;


            // Load saved values
            LoadComboBoxValue(comboBoxUvoznik, "uvoznik.json");
            LoadComboBoxValue(comboBoxUverenje, "uverenje.json");
            LoadComboBoxValue(comboBoxSrps, "srps.json");
            LoadComboBoxValue(comboBoxProizvodjac, "proizvodjac.json");
            LoadComboBoxValue(comboBoxPostava, "postava.json");
            LoadComboBoxValue(comboBoxPoreklo, "poreklo.json");
            LoadComboBoxValue(comboBoxOdrzavanje, "odrzavanje.json");
            LoadComboBoxValue(comboBoxNamena, "namena.json");
            LoadComboBoxValue(comboBoxLice, "lice.json");
            LoadComboBoxValue(comboBoxIzrada, "izrada.json");
            LoadComboBoxValue(comboBoxDjon, "djon.json");
            LoadComboBoxValue(comboBoxArtikal, "artikal.json");
            LoadComboBoxValue(comboBoxNaziv, "naziv.json");

            // Attach Paint event handler
            previewPanel.Paint += previewPanel_Paint;

            labelCloseupPanel.Paint += labelCloseupPanel_Paint;
        }


        private void SaveSablon(object sender, EventArgs e)
        {
            var sablon = new
            {
                nazivSablona = txtNaziv.Text,
                uvoznik = comboBoxUvoznik.SelectedItem?.ToString() ?? "",
                uverenje = comboBoxUverenje.SelectedItem?.ToString() ?? "",
                srps = comboBoxSrps.SelectedItem?.ToString() ?? "",
                proizvodjac = comboBoxProizvodjac.SelectedItem?.ToString() ?? "",
                postava = comboBoxPostava.SelectedItem?.ToString() ?? "",
                poreklo = comboBoxPoreklo.SelectedItem?.ToString() ?? "",
                odrzavanje = comboBoxOdrzavanje.SelectedItem?.ToString() ?? "",
                namena = comboBoxNamena.SelectedItem?.ToString() ?? "",
                lice = comboBoxLice.SelectedItem?.ToString() ?? "",
                izrada = comboBoxIzrada.SelectedItem?.ToString() ?? "",
                djon = comboBoxDjon.SelectedItem?.ToString() ?? "",
                artikal = comboBoxArtikal.SelectedItem?.ToString() ?? "",
                naziv = comboBoxNaziv.SelectedItem?.ToString() ?? ""
            };

            try
            {
                // Ensure the directory exists
                string directoryPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sabloni");
                if (!System.IO.Directory.Exists(directoryPath))
                {
                    System.IO.Directory.CreateDirectory(directoryPath);
                }

                // Sanitize file name
                string sanitizedFileName = System.IO.Path.GetInvalidFileNameChars().Aggregate(sablon.nazivSablona, (current, c) => current.Replace(c.ToString(), ""));
                string fileName = System.IO.Path.Combine(directoryPath, $"{sanitizedFileName}.json");

                // Serialize the object to JSON
                string json = JsonConvert.SerializeObject(sablon, Formatting.Indented);

                // Save the JSON to a file
                System.IO.File.WriteAllText(fileName, json);

                MessageBox.Show("Sablon sacuvan uspesno!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Greska prilikom cuvanja sablona: {ex.Message}");
            }
        }


        private void LoadSablon(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "JSON Files (*.json)|*.json";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Read the JSON file
                string json = File.ReadAllText(openFileDialog.FileName);

                // Deserialize the JSON to an object
                dynamic sablon = JsonConvert.DeserializeObject<dynamic>(json);

                // Check if sablon is not null and contains all necessary fields
                if (sablon != null)
                {
                    // Populate the ComboBoxes and TextBox with the values
                    txtNaziv.Text = (string)sablon.nazivSablona;
                    SetComboBoxValue(comboBoxUvoznik, (string)sablon.uvoznik);
                    SetComboBoxValue(comboBoxUverenje, (string)sablon.uverenje);
                    SetComboBoxValue(comboBoxSrps, (string)sablon.srps);
                    SetComboBoxValue(comboBoxProizvodjac, (string)sablon.proizvodjac);
                    SetComboBoxValue(comboBoxPostava, (string)sablon.postava);
                    SetComboBoxValue(comboBoxPoreklo, (string)sablon.poreklo);
                    SetComboBoxValue(comboBoxOdrzavanje, (string)sablon.odrzavanje);
                    SetComboBoxValue(comboBoxNamena, (string)sablon.namena);
                    SetComboBoxValue(comboBoxLice, (string)sablon.lice);
                    SetComboBoxValue(comboBoxIzrada, (string)sablon.izrada);
                    SetComboBoxValue(comboBoxDjon, (string)sablon.djon);
                    SetComboBoxValue(comboBoxArtikal, (string)sablon.artikal);
                    SetComboBoxValue(comboBoxNaziv, (string)sablon.naziv);

                    UpdateLabelCloseup();

                    MessageBox.Show("Sablon uvezen uspesno!");
                }
                else
                {
                    MessageBox.Show("Greska prilikom ucitavanja sablona! Fajl je corrupted ili je neodgovarajuceg formata!");
                }
            }
        }

        private void SetComboBoxValue(ComboBox comboBox, string value)
        {
            if (comboBox != null && !string.IsNullOrEmpty(value))
            {
                // Add value to the ComboBox items if it's not already present
                if (!comboBox.Items.Contains(value))
                {
                    comboBox.Items.Add(value);
                }
                comboBox.SelectedItem = value;
            }
        }
        private void AddLabelToA4(object sender, EventArgs e)
        {
            int startPosition = (int)numericUpDownPosition.Value - 1;
            int amount = (int)numericUpDownLabels.Value;

            // Validate the startPosition and amount
            if (startPosition < 0 || startPosition >= 24)
            {
                MessageBox.Show("Pozicija mora biti između 1 i 24.");
                return;
            }

            if (amount < 1)
            {
                MessageBox.Show("Broj deklaracija mora biti veći od 0.");
                return;
            }

            // Construct the label text
            string labelText = $"ZEMLJA POREKLA: {comboBoxPoreklo.Text}\n" +
                               $"UVOZNIK: {comboBoxUvoznik.Text}\n" +
                               $"PROIZVOĐAČ: {comboBoxProizvodjac.Text}\n" +
                               $"UVERENJE BR: {comboBoxUverenje.Text}\n" +
                               $"NAZIV ROBE: {comboBoxNaziv.Text}\n" +
                               $"ARTIKAL: {comboBoxArtikal.Text}\n" +
                               $"SIROVINSKI SASTAV: LICE-{comboBoxLice.Text}, POSTAVA-{comboBoxPostava.Text}\n" +
                               $"{new string(' ', 40)}ĐON-{comboBoxDjon.Text}\n" +
                               $"SRPS: {comboBoxSrps.Text}\n" +
                               $"NAČIN IZRADE: {comboBoxIzrada.Text}\n" +
                               $"NAMENA: {comboBoxNamena.Text}\n" +
                               $"ODRŽAVANJE: {comboBoxOdrzavanje.Text}\n";

            // Update labels in the specified range
            for (int i = startPosition; i < startPosition + amount; i++)
            {
                if (i < 24) // Ensure the index is within the valid range
                {
                    labelsOnPage[i] = labelText; // Update label at the specified position
                }
            }

            // Update the ListBox with the new label information
            UpdateLabelsListBox();

            // Update the preview with the new labels
            UpdateA4Preview();
        }
        // Method to extract the value of ARTIKAL from labelText
        // Method to extract the value of ARTIKAL from labelText
        private string ExtractArtikalValue(string labelText)
        {
            if (string.IsNullOrEmpty(labelText))
            {
                return string.Empty;
            }

            var lines = labelText.Split('\n');

            foreach (var line in lines)
            {
                if (line.StartsWith("ARTIKAL:"))
                {
                    return line.Substring("ARTIKAL:".Length).Trim();
                }
            }

            return string.Empty;
        }

        // Method to update the ListBox with grouped ARTIKAL values
        private void UpdateLabelsListBox()
        {
            if (labelsOnPage == null)
            {
                return; // Or handle the error as needed
            }

            var artikalCounts = new Dictionary<string, int>();

            // Count occurrences of each ARTIKAL value
            foreach (var labelText in labelsOnPage)
            {
                string artikalValue = ExtractArtikalValue(labelText);

                if (!string.IsNullOrEmpty(artikalValue))
                {
                    if (artikalCounts.ContainsKey(artikalValue))
                    {
                        artikalCounts[artikalValue]++;
                    }
                    else
                    {
                        artikalCounts[artikalValue] = 1;
                    }
                }
            }

            // Update ListBox with grouped ARTIKAL values
            labelsListBox.Items.Clear();
            foreach (var kvp in artikalCounts)
            {
                string displayText = $"{kvp.Key} - {kvp.Value} komada";
                labelsListBox.Items.Add(displayText);
            }
        }

        private void DeleteLabelButton_Click(object sender, EventArgs e)
        {
            if (labelsListBox.SelectedItem == null)
            {
                MessageBox.Show("Izaberite deklaraciju za brisanje!");
                return;
            }

            string selectedEntry = labelsListBox.SelectedItem.ToString();

            // Extract the ARTIKAL value and the number of labels to delete
            string artikalValue = selectedEntry.Split('-')[0].Trim();
            int numberOfLabelsToDelete = 0;

            var artikalCounts = new Dictionary<string, int>();

            foreach (var labelText in labelsOnPage)
            {
                string artikal = ExtractArtikalValue(labelText);
                if (!string.IsNullOrEmpty(artikal))
                {
                    if (artikalCounts.ContainsKey(artikal))
                    {
                        artikalCounts[artikal]++;
                    }
                    else
                    {
                        artikalCounts[artikal] = 1;
                    }
                }
            }

            if (artikalCounts.ContainsKey(artikalValue))
            {
                numberOfLabelsToDelete = artikalCounts[artikalValue];
            }

            // Remove the label(s) from labelsOnPage
            for (int i = 0; i < labelsOnPage.Count; i++)
            {
                if (numberOfLabelsToDelete <= 0) break;

                string artikal = ExtractArtikalValue(labelsOnPage[i]);

                if (artikal == artikalValue)
                {
                    labelsOnPage[i] = null; // Or assign an empty string
                    numberOfLabelsToDelete--;
                }
            }

            // Remove the entry from the ListBox
            labelsListBox.Items.Remove(labelsListBox.SelectedItem);

            // Rerender the A4 sheet
            UpdateA4Preview();
        }

        private void UpdateA4Preview()
        {
            previewPanel.Invalidate(); // Forces the panel to repaint
        }

        private void previewPanel_Paint(object sender, PaintEventArgs e)
        {
            if (e == null || e.Graphics == null) return;

            Graphics g = e.Graphics;
            g.Clear(Color.White);

            float scaleFactor = 0.75f; // Scaling factor for the preview

            float mmToPixel = 96 / 25.4f * scaleFactor;
            float labelWidth = 70 * mmToPixel;
            float labelHeight = 37 * mmToPixel;

            int labelsPerRow = LabelsPerRow;
            int labelsPerColumn = LabelsPerColumn;

            Font headerFont = new Font("Arial", 16 * scaleFactor, FontStyle.Bold);
            Font footerFont = new Font("Arial", 10 * scaleFactor, FontStyle.Bold);
            Font textFont = new Font("Arial", 12 * scaleFactor, FontStyle.Regular);

            float textMargin = 5 * mmToPixel;

            for (int row = 0; row < labelsPerColumn; row++)
            {
                for (int col = 0; col < labelsPerRow; col++)
                {
                    int index = row * labelsPerRow + col;
                    if (index < labelsOnPage?.Count)
                    {
                        float x = col * labelWidth;
                        float y = row * labelHeight;

                        g.DrawRectangle(Pens.Black, x, y, labelWidth, labelHeight);

                        string labelText = labelsOnPage[index] ?? string.Empty;
                        if (!string.IsNullOrEmpty(labelText))
                        {
                            string header = "D E K L A R A C I J A";
                            float headerWidth = labelWidth - 2 * textMargin;
                            float headerX = x + textMargin;
                            float headerY = y + textMargin;
                            DrawText(g, header, headerX, headerY, headerWidth, 14 * scaleFactor, headerFont, centerText: true);

                            float headerHeight = g.MeasureString(header, headerFont).Height;
                            float spaceBetweenHeaderAndText = 3 * mmToPixel;
                            float remainingTextY = y + headerHeight + spaceBetweenHeaderAndText;
                            float remainingTextHeight = labelHeight - headerHeight - spaceBetweenHeaderAndText - textMargin - 5 * scaleFactor;

                            remainingTextHeight = Math.Max(remainingTextHeight, 0);

                            float remainingTextWidth = labelWidth - 2 * textMargin;
                            DrawText(g, labelText, x + textMargin, remainingTextY, remainingTextWidth, remainingTextHeight, textFont);

                            string footer = "KVALITET KONTROLISAO JUGOINSPEKT BEOGRAD";
                            float footerWidth = labelWidth - 2 * textMargin;
                            float footerX = x + textMargin;
                            float footerY = y + labelHeight - textMargin - 5 * scaleFactor;
                            DrawText(g, footer, footerX, footerY, footerWidth, 5 * scaleFactor, footerFont, centerText: true);
                        }
                    }
                }
            }
        }


        private void DrawText(Graphics g, string text, float x, float y, float width, float height, Font font, bool centerText = false)
        {
            SizeF textSize = g.MeasureString(text, font);

            // Adjust font size if needed
            while (textSize.Width > width || textSize.Height > height)
            {
                font = new Font(font.FontFamily, font.Size - 0.5f, font.Style);
                textSize = g.MeasureString(text, font);
                if (font.Size <= 0) break;
            }

            if (centerText)
            {
                x += (width - textSize.Width) / 2;
                y += (height - textSize.Height) / 2;
            }

            string[] lines = SplitTextIntoLines(text, font, width);
            float lineHeight = g.MeasureString("A", font).Height;

            float textY = y; // Start from the top of the sticker
            foreach (var line in lines)
            {
                g.DrawString(line, font, Brushes.Black, new PointF(x, textY)); // Draw text
                textY += lineHeight;
                if (textY + lineHeight > y + height)
                    break;
            }
        }

        private string[] SplitTextIntoLines(string text, Font font, float maxWidth)
        {
            var lines = new List<string>();
            var words = text.Split(' ');
            var currentLine = new StringBuilder();

            foreach (var word in words)
            {
                var testLine = currentLine.Length > 0 ? currentLine + " " + word : word;
                if (MeasureStringWidth(testLine, font) > maxWidth)
                {
                    lines.Add(currentLine.ToString());
                    currentLine.Clear();
                    currentLine.Append(word);
                }
                else
                {
                    currentLine.Append(currentLine.Length > 0 ? " " + word : word);
                }
            }

            // Add the last line
            if (currentLine.Length > 0)
            {
                lines.Add(currentLine.ToString());
            }

            return lines.ToArray();
        }

        private float MeasureStringWidth(string text, Font font)
        {
            using (var g = Graphics.FromImage(new Bitmap(1, 1)))
            {
                return g.MeasureString(text, font).Width;
            }
        }

        private void PrintPage(object sender, PrintPageEventArgs e)
        {
            const int dpi = 300;
            float mmToPixel = dpi / 25.4f;
            float labelWidth = 70 * mmToPixel;
            float labelHeight = 37 * mmToPixel;

            int labelsPerRow = LabelsPerRow;
            int labelsPerColumn = LabelsPerColumn;
            int labelsPerPage = labelsPerRow * labelsPerColumn;

            int currentPage = e.PageSettings.PrinterSettings.ToPage; // 0-based page index
            int startLabelIndex = currentPage * labelsPerPage;
            int endLabelIndex = Math.Min(startLabelIndex + labelsPerPage, labelsOnPage.Count);

            if (startLabelIndex >= labelsOnPage.Count)
            {
                e.HasMorePages = false;
                return;
            }

            float scaleX = (float)e.PageBounds.Width / (labelsPerRow * labelWidth);
            float scaleY = (float)e.PageBounds.Height / (labelsPerColumn * labelHeight);
            float scale = Math.Min(scaleX, scaleY);

            float xOffset = (e.PageBounds.Width - labelsPerRow * labelWidth * scale) / 2;
            float yOffset = (e.PageBounds.Height - labelsPerColumn * labelHeight * scale) / 2;

            Font headerFont = new Font("Arial", 18 * scale, FontStyle.Bold);
            Font footerFont = new Font("Arial", 10 * scale, FontStyle.Bold);
            Font textFont = new Font("Arial", 12 * scale, FontStyle.Regular);

            for (int row = 0; row < labelsPerColumn; row++)
            {
                for (int col = 0; col < labelsPerRow; col++)
                {
                    int index = startLabelIndex + row * labelsPerRow + col;
                    if (index < endLabelIndex && index < labelsOnPage.Count)
                    {
                        float x = xOffset + col * labelWidth * scale;
                        float y = yOffset + row * labelHeight * scale;

                        e.Graphics.DrawRectangle(Pens.Black, x, y, labelWidth * scale, labelHeight * scale);

                        string labelText = labelsOnPage[index] ?? string.Empty;
                        if (!string.IsNullOrEmpty(labelText))
                        {
                            float textMargin = 5 * mmToPixel * scale;
                            float headerHeight = e.Graphics.MeasureString("D E K L A R A C I J A", headerFont).Height;
                            float footerHeight = e.Graphics.MeasureString("KVALITET KONTROLISAO JUGOINSPEKT BEOGRAD", footerFont).Height;

                            // Center header and footer
                            float headerWidth = e.Graphics.MeasureString("D E K L A R A C I J A", headerFont).Width;
                            float footerWidth = e.Graphics.MeasureString("KVALITET KONTROLISAO JUGOINSPEKT BEOGRAD", footerFont).Width;

                            e.Graphics.DrawString("D E K L A R A C I J A", headerFont, Brushes.Black, x + (labelWidth * scale - headerWidth) / 2, y + textMargin);
                            e.Graphics.DrawString(labelText, textFont, Brushes.Black, x + textMargin, y + textMargin + headerHeight);
                            e.Graphics.DrawString("KVALITET KONTROLISAO JUGOINSPEKT BEOGRAD", footerFont, Brushes.Black, x + (labelWidth * scale - footerWidth) / 2, y + labelHeight * scale - textMargin - footerHeight);
                        }
                    }
                }
            }

            e.HasMorePages = endLabelIndex < labelsOnPage.Count;
        }

        private void LoadComboBoxValue(ComboBox comboBox, string fileName)
        {
            if (File.Exists(fileName))
            {
                try
                {
                    var json = File.ReadAllText(fileName);
                    var values = JsonConvert.DeserializeObject<List<string>>(json);

                    if (values != null)
                    {
                        comboBox.Items.Clear();
                        foreach (var value in values)
                        {
                            comboBox.Items.Add(value);
                        }

                        // Ensure no item is selected
                        comboBox.SelectedIndex = -1;
                    }
                }
                catch (JsonSerializationException)
                {
                    MessageBox.Show($"Greska prilikom ucitavanja podataka iz {fileName}. Fajl je mozda corrupted.");
                }
            }
        }

        private void SaveComboBoxValue(ComboBox comboBox, string fileName)
        {
            List<string> values = new List<string>();

            // Load existing values if the file exists
            if (File.Exists(fileName))
            {
                try
                {
                    var existingJson = File.ReadAllText(fileName);
                    values = JsonConvert.DeserializeObject<List<string>>(existingJson) ?? new List<string>();
                }
                catch (JsonSerializationException)
                {
                    // Handle the case where the file has corrupted or unexpected JSON data
                    MessageBox.Show($"Greska prilikom ucitavanja podataka iz {fileName}. Fajl je mozda corrupted.");
                }
            }

            // Add the new value if it's not empty and not already in the list
            var value = comboBox.Text;
            if (!string.IsNullOrWhiteSpace(value) && !values.Contains(value))
            {
                values.Add(value);
                var json = JsonConvert.SerializeObject(values, Formatting.Indented);
                File.WriteAllText(fileName, json);
            }

            // Refresh the ComboBox to display updated values
            LoadComboBoxValue(comboBox, fileName);
        }

        private void SaveComboBoxValues(ComboBox comboBox, string fileName)
        {
            // Save the current items in the ComboBox to the JSON file
            var items = comboBox.Items.Cast<string>().ToList();
            File.WriteAllText(fileName, JsonConvert.SerializeObject(items));
        }

        private void RemoveComboBoxValue(ComboBox comboBox, string fileName)
        {
            if (comboBox.SelectedItem != null)
            {
                // Remove the selected item from the ComboBox
                string selectedItem = comboBox.SelectedItem.ToString();
                comboBox.Items.Remove(selectedItem);

                // Save updated values back to JSON
                SaveComboBoxValues(comboBox, fileName);
            }
        }

        private void UpdateToolTipSave(ToolTip toolTip, Button button, string text, Color color)
        {
            toolTip.SetToolTip(button, text);
            Timer timer = new Timer { Interval = 2000 };
            timer.Tick += (s, e) =>
            {
                toolTip.SetToolTip(button, "Sačuvaj");
                timer.Stop();
            };
            timer.Start();
        }

        private void UpdateToolTipRemove(ToolTip toolTip, Button button, string text, Color color)
        {
            toolTip.SetToolTip(button, text);
            Timer timer = new Timer { Interval = 2000 };
            timer.Tick += (s, e) =>
            {
                toolTip.SetToolTip(button, "Obriši");
                timer.Stop();
            };
            timer.Start();
        }

        private void saveUvoznikButton_Click(object sender, EventArgs e)
        {
            SaveComboBoxValue(comboBoxUvoznik, "uvoznik.json");
            UpdateToolTipSave(toolTip1, saveUvoznikButton, "Sačuvano", Color.Green);
        }

        private void saveUverenjeButton_Click(object sender, EventArgs e)
        {
            SaveComboBoxValue(comboBoxUverenje, "uverenje.json");
            UpdateToolTipSave(toolTip1, saveUverenjeButton, "Sačuvano", Color.Green);
        }

        private void saveSrpsButton_Click(object sender, EventArgs e)
        {
            SaveComboBoxValue(comboBoxSrps, "srps.json");
            UpdateToolTipSave(toolTip1, saveSrpsButton, "Sačuvano", Color.Green);
        }

        private void saveProizvodjacButton_Click(object sender, EventArgs e)
        {
            SaveComboBoxValue(comboBoxProizvodjac, "proizvodjac.json");
            UpdateToolTipSave(toolTip1, saveProizvodjacButton, "Sačuvano", Color.Green);
        }

        private void savePostavaButton_Click(object sender, EventArgs e)
        {
            SaveComboBoxValue(comboBoxPostava, "postava.json");
            UpdateToolTipSave(toolTip1, savePostavaButton, "Sačuvano", Color.Green);
        }

        private void savePorekloButton_Click(object sender, EventArgs e)
        {
            SaveComboBoxValue(comboBoxPoreklo, "poreklo.json");
            UpdateToolTipSave(toolTip1, savePorekloButton, "Sačuvano", Color.Green);
        }

        private void saveOdrzavanjeButton_Click(object sender, EventArgs e)
        {
            SaveComboBoxValue(comboBoxOdrzavanje, "odrzavanje.json");
            UpdateToolTipSave(toolTip1, saveOdrzavanjeButton, "Sačuvano", Color.Green);
        }

        private void saveNamenaButton_Click(object sender, EventArgs e)
        {
            SaveComboBoxValue(comboBoxNamena, "namena.json");
            UpdateToolTipSave(toolTip1, saveNamenaButton, "Sačuvano", Color.Green);
        }

        private void saveLiceButton_Click(object sender, EventArgs e)
        {
            SaveComboBoxValue(comboBoxLice, "lice.json");
            UpdateToolTipSave(toolTip1, saveLiceButton, "Sačuvano", Color.Green);
        }

        private void saveIzradaButton_Click(object sender, EventArgs e)
        {
            SaveComboBoxValue(comboBoxIzrada, "izrada.json");
            UpdateToolTipSave(toolTip1, saveIzradaButton, "Sačuvano", Color.Green);
        }

        private void saveDjonButton_Click(object sender, EventArgs e)
        {
            SaveComboBoxValue(comboBoxDjon, "djon.json");
            UpdateToolTipSave(toolTip1, saveDjonButton, "Sačuvano", Color.Green);
        }

        private void saveArtikalButton_Click(object sender, EventArgs e)
        {
            SaveComboBoxValue(comboBoxArtikal, "artikal.json");
            UpdateToolTipSave(toolTip1, saveArtikalButton, "Sačuvano", Color.Green);
        }

        private void saveNazivButton_Click(object sender, EventArgs e)
        {
            SaveComboBoxValue(comboBoxNaziv, "naziv.json");
            UpdateToolTipSave(toolTip1, saveNazivButton, "Sačuvano", Color.Green);
        }

        private void removeUvoznikButton_Click(object sender, EventArgs e)
        {
            RemoveComboBoxValue(comboBoxUvoznik, "uvoznik.json");
            UpdateToolTipRemove(toolTip2, removeUvoznikButton, "Obrisano", Color.Red);
        }

        private void removeUverenjeButton_Click(object sender, EventArgs e)
        {
            RemoveComboBoxValue(comboBoxUverenje, "uverenje.json");
            UpdateToolTipRemove(toolTip2, removeUverenjeButton, "Obrisano", Color.Red);
        }

        private void removeSrpsButton_Click(object sender, EventArgs e)
        {
            RemoveComboBoxValue(comboBoxSrps, "srps.json");
            UpdateToolTipRemove(toolTip2, removeSrpsButton, "Obrisano", Color.Red);
        }

        private void removeProizvodjacButton_Click(object sender, EventArgs e)
        {
            RemoveComboBoxValue(comboBoxProizvodjac, "proizvodjac.json");
            UpdateToolTipRemove(toolTip2, removeProizvodjacButton, "Obrisano", Color.Red);
        }

        private void removePostavaButton_Click(object sender, EventArgs e)
        {
            RemoveComboBoxValue(comboBoxPostava, "postava.json");
            UpdateToolTipRemove(toolTip2, removePostavaButton, "Obrisano", Color.Red);
        }

        private void removePorekloButton_Click(object sender, EventArgs e)
        {
            RemoveComboBoxValue(comboBoxPoreklo, "poreklo.json");
            UpdateToolTipRemove(toolTip2, removePorekloButton, "Obrisano", Color.Red);
        }

        private void removeOdrzavanjeButton_Click(object sender, EventArgs e)
        {
            RemoveComboBoxValue(comboBoxOdrzavanje, "odrzavanje.json");
            UpdateToolTipRemove(toolTip2, removeOdrzavanjeButton, "Obrisano", Color.Red);
        }

        private void removeNamenaButton_Click(object sender, EventArgs e)
        {
            RemoveComboBoxValue(comboBoxNamena, "namena.json");
            UpdateToolTipRemove(toolTip2, removeNamenaButton, "Obrisano", Color.Red);
        }

        private void removeLiceButton_Click(object sender, EventArgs e)
        {
            RemoveComboBoxValue(comboBoxLice, "lice.json");
            UpdateToolTipRemove(toolTip2, removeLiceButton, "Obrisano", Color.Red);
        }

        private void removeIzradaButton_Click(object sender, EventArgs e)
        {
            RemoveComboBoxValue(comboBoxIzrada, "izrada.json");
            UpdateToolTipRemove(toolTip2, removeIzradaButton, "Obrisano", Color.Red);
        }

        private void removeDjonButton_Click(object sender, EventArgs e)
        {
            RemoveComboBoxValue(comboBoxDjon, "djon.json");
            UpdateToolTipRemove(toolTip2, removeDjonButton, "Obrisano", Color.Red);
        }

        private void removeArtikalButton_Click(object sender, EventArgs e)
        {
            RemoveComboBoxValue(comboBoxArtikal, "artikal.json");
            UpdateToolTipRemove(toolTip2, removeArtikalButton, "Obrisano", Color.Red);
        }

        private void removeNazivButton_Click(object sender, EventArgs e)
        {
            RemoveComboBoxValue(comboBoxNaziv, "naziv.json");
            UpdateToolTipRemove(toolTip2, removeNazivButton, "Obrisano", Color.Red);
        }
        private void PrintButton_Click(object sender, EventArgs e)
        {
            printPreviewDialog.ShowDialog();
        }

        private void UpdateLabelCloseup()
        {
            if (labelCloseupPanel == null) return;

            // Format the label text from combobox values
            string labelText = $"ZEMLJA POREKLA: {comboBoxPoreklo.Text}\n" +
                               $"UVOZNIK: {comboBoxUvoznik.Text}\n" +
                               $"PROIZVOĐAČ: {comboBoxProizvodjac.Text}\n" +
                               $"UVERENJE BR: {comboBoxUverenje.Text}\n" +
                               $"NAZIV ROBE: {comboBoxNaziv.Text}\n" +
                               $"ARTIKAL: {comboBoxArtikal.Text}\n" +
                               $"SIROVINSKI SASTAV: LICE-{comboBoxLice.Text}, POSTAVA-{comboBoxPostava.Text}\n" +
                               $"{new string(' ', 40)}ĐON-{comboBoxDjon.Text}\n" +
                               $"SRPS: {comboBoxSrps.Text}\n" +
                               $"NAČIN IZRADE: {comboBoxIzrada.Text}\n" +
                               $"NAMENA: {comboBoxNamena.Text}\n" +
                               $"ODRŽAVANJE: {comboBoxOdrzavanje.Text}\n";

            // Set the formatted text as the Tag of the panel
            labelCloseupPanel.Tag = labelText;

            // Invalidate the panel to trigger a repaint
            labelCloseupPanel.Invalidate();
        }

        private void labelCloseupPanel_Paint(object sender, PaintEventArgs e)
        {
            if (e == null || e.Graphics == null) return;

            Graphics g = e.Graphics;
            g.Clear(Color.White);

            // Get the labelText from the panel Tag
            string labelText = labelCloseupPanel.Tag as string;
            if (string.IsNullOrEmpty(labelText)) return;

            float scaleFactor = 2.0f; // Scaling factor for the closeup label

            // Convert millimeters to pixels using the scaling factor
            float mmToPixel = 96 / 25.4f * scaleFactor;
            float labelWidth = 70 * mmToPixel;
            float labelHeight = 37 * mmToPixel;

            // Font sizes with scaling factor
            Font headerFont = new Font("Arial", 16 * scaleFactor, FontStyle.Bold);
            Font footerFont = new Font("Arial", 10 * scaleFactor, FontStyle.Bold);
            Font textFont = new Font("Arial", 12 * scaleFactor, FontStyle.Regular);

            // Margins and spacing with scaling factor
            float textMargin = 5 * mmToPixel;
            float spaceBetweenHeaderAndText = 0;

            // Draw header
            string header = "D E K L A R A C I J A";
            float headerWidth = labelWidth - 2 * textMargin;
            float headerX = textMargin;
            float headerY = textMargin;
            DrawTextCl(g, header, headerX, headerY, headerWidth, 14 * scaleFactor, headerFont, centerText: true);

            // Measure header height
            float headerHeight = g.MeasureString(header, headerFont).Height - 15;

            // Draw label text
            float remainingTextY = headerY + headerHeight + spaceBetweenHeaderAndText;
            float remainingTextHeight = labelHeight - headerHeight - spaceBetweenHeaderAndText - textMargin - 10 * mmToPixel; // Adjust as needed

            remainingTextHeight = Math.Max(remainingTextHeight, 0);
            float remainingTextWidth = labelWidth - 2 * textMargin;
            DrawTextCl(g, labelText, textMargin, remainingTextY, remainingTextWidth, remainingTextHeight, textFont);

            // Draw footer
            string footer = "KVALITET KONTROLISAO JUGOINSPEKT BEOGRAD";
            float footerWidth = labelWidth - 2 * textMargin;
            float footerX = textMargin;
            float footerY = labelHeight - textMargin - g.MeasureString(footer, footerFont).Height; // Adjusted to fit in the bottom

            DrawTextCl(g, footer, footerX, footerY, footerWidth, 5 * scaleFactor, footerFont, centerText: true);
        }

        private void DrawTextCl(Graphics g, string text, float x, float y, float width, float height, Font font, bool centerText = false)
        {
            // Define a minimum font size to avoid extremely small fonts
            const float minFontSize = 7.0f;

            SizeF textSize = g.MeasureString(text, font);

            // Adjust font size if needed
            while (textSize.Width > width || textSize.Height > height)
            {
                font = new Font(font.FontFamily, font.Size - 0.5f, font.Style);
                textSize = g.MeasureString(text, font);

                // Exit loop if font size is below minimum threshold
                if (font.Size <= minFontSize) break;
            }

            if (centerText)
            {
                x += (width - textSize.Width) / 2;
                y += (height - textSize.Height) / 2;
            }

            string[] lines = SplitTextIntoLines(text, font, width);
            float lineHeight = g.MeasureString("A", font).Height;

            float textY = y; // Start from the top of the sticker
            foreach (var line in lines)
            {
                g.DrawString(line, font, Brushes.Black, new PointF(x, textY)); // Draw text
                textY += lineHeight;
                if (textY + lineHeight > y + height)
                    break;
            }
        }


    }
}
