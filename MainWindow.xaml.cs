using System;
using System.Windows;
using Microsoft.Win32;
using System.Windows.Media.Imaging;
using Tesseract;
using System.Drawing;
using System.Drawing.Imaging;
using ImageFormat = System.Drawing.Imaging.ImageFormat;
using System.Windows.Input;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
using System.Windows.Controls;

namespace DataExtractor
{
    public partial class MainWindow : Window
    {
        private string imagePath;

        private List<string> pastedImagePath = new List<string>();
        private const int MaxImages = 5;

        private string lastSavedCsvPath;


        private readonly Dictionary<string, string> corrections = new Dictionary<string, string>
        {
            /*{ "Tooi", "Tool" }, { "Too0", "Tool" }, { "To00", "Tool" }, { "Tooit", "Tool1" }, { "Tooi1", "Tool" },{"Tool0S","Tool05"},
            {"Tool", "Tool" }, {"SCREWO1","SCREW01"}, {"Tool03","Tool03"},  {"82B01","B2B01"},
            { "82801", "B2B01" }, { "82802", "B2B02" }, { "82803", "B2B03" }, { "82804", "B2B04" }, { "82B02", "B2B02" },
            { "82805", "B2B05" }, { "82806", "B2B06" }, { "82807", "B2B07" }, { "82808", "B2B08" },
            { "82809", "B2B09" }, { "82810", "B2B10" }, { "82811", "B2B11" },{"82B02_SS_1", "B2B02_SS_1" },{"82B01_H_DIFF","B2B01_H_DIFF"},            
            { "55", "SS" }, { "HN", "HN" }, { "H_DIFF", "H_DIFF" },{"5CREW","SCREW"}, {"SCREWO3","SCREW03"}, {"5CREWO1", "SCREW01"},
            {"COAXO01", "COAX01"}, {"5CREWO3","5CREW03"}, {"5CREWO01", "SCREW01"}, {"5CREWO02", "SCREW02"},
            { "Too102", "Tool02" }, { "Too103", "Tool03" }, { "Too104", "Tool04" },
            { "Too105", "Tool05" }, { "Too106", "Tool06" }, { "Too107", "Tool07" },
            { "Too108", "Tool08" }, { "8Z801", "B2B01" }, { "88ZO2", "B2B02" },
            { "SS_Z", "SS_2" }, { "SS_J", "SS_3" }, { "S5", "SS" },*/
            
            //tool variants
            { "Tooi", "Tool" }, { "Too0", "Tool" }, { "To00", "Tool" }, { "Tooit", "Tool1" }, { "Tooi1", "Tool" },
            { "Tool0S", "Tool05" }, { "Tool", "Tool" }, { "Tooll", "Tool" }, { "Toool", "Tool" },
            { "To0l", "Tool" }, { "T00l", "Tool" }, { "T0ol", "Tool" }, { "ToolO1", "Tool01" }, { "Tool01S", "Tool01" },

            // Specific tool numbers
            { "Too102", "Tool02" }, { "Too103", "Tool03" }, { "Too104", "Tool04" },
            { "Too105", "Tool05" }, { "Too106", "Tool06" }, { "Too107", "Tool07" },
            { "Too108", "Tool08" }, { "Tool03", "Tool03" },

            // B2B variants
            { "82801", "B2B01" }, { "82802", "B2B02" }, { "82803", "B2B03" }, { "82804", "B2B04" },
            { "82805", "B2B05" }, { "82806", "B2B06" }, { "82807", "B2B07" }, { "82808", "B2B08" },
            { "82809", "B2B09" }, { "82810", "B2B10" }, { "82811", "B2B11" },
            { "82B01", "B2B01" }, { "82B02", "B2B02" }, { "82B03", "B2B03" }, { "82B04", "B2B04" }, { "82B05", "B2B05" },
            { "8Z801", "B2B01" }, { "8Z802", "B2B02" }, { "88ZO2", "B2B02" }, { "88Z03", "B2B03" },

            // B2B test variants
            { "82B02_SS_1", "B2B02_SS_1" }, { "82B01_H_DIFF", "B2B01_H_DIFF" },

            // Screw variants
            { "SCREWO1", "SCREW01" }, { "SCREWO3", "SCREW03" }, { "SCREWOO1", "SCREW01" },
            { "SCREW1", "SCREW01" }, { "SCREW02A", "SCREW02" },
            { "5CREW", "SCREW" }, { "5CREWO1", "SCREW01" }, { "5CREWO3", "SCREW03" },
            { "5CREWO01", "SCREW01" }, { "5CREWO02", "SCREW02" },

            // Coax variants
            { "COAXO01", "COAX01" },

            // SS variants
            { "SS_Z", "SS_2" }, { "SS_J", "SS_3" }, { "S5", "SS" }, { "55", "SS" },
            { "SS1", "SS_1" }, { "SS2", "SS_2" }, { "SS3", "SS_3" }, { "SS4", "SS_4" },

            // HN and H_DIFF variants
            { "HN", "HN" }, { "H_DIFF", "H_DIFF" }, { "HDIFF", "H_DIFF" }, { "H-DIFF", "H_DIFF" },
            { "HN1", "HN_1" }, { "HN2", "HN_2" }


        };

        private readonly HashSet<string> validTests = new HashSet<string>
        {
            "B2B01_H_DIFF", "B2B01_HN_1", "B2B01_HN_2", "B2B01_SS_1", "B2B01_SS_2", "B2B01_SS_3", "B2B01_SS_4",
            "B2B02_H_DIFF", "B2B02_HN_1", "B2B02_HN_2", "B2B02_SS_1", "B2B02_SS_2", "B2B02_SS_3", "B2B02_SS_4",
            "B2B03_H_DIFF", "B2B03_HN_1", "B2B03_HN_2", "B2B03_SS_1", "B2B03_SS_2", "B2B03_SS_3", "B2B03_SS_4",
            "B2B04_H_DIFF", "B2B04_HN_1", "B2B04_HN_2", "B2B04_SS_1", "B2B04_SS_2", "B2B04_SS_3", "B2B04_SS_4",
            "B2B05_H_DIFF", "B2B05_HN_1", "B2B05_HN_2", "B2B05_SS_1", "B2B05_SS_2", "B2B05_SS_3", "B2B05_SS_4",
            "B2B06_H_DIFF", "B2B06_HN_1", "B2B06_HN_2", "B2B06_SS_1", "B2B06_SS_2", "B2B06_SS_3", "B2B06_SS_4",
            "B2B07_H_DIFF", "B2B07_HN_1", "B2B07_HN_2", "B2B07_SS_1", "B2B07_SS_2", "B2B07_SS_3", "B2B07_SS_4",
            "B2B08_H_DIFF", "B2B08_HN_1", "B2B08_HN_2", "B2B08_SS_1", "B2B08_SS_2", "B2B08_SS_3", "B2B08_SS_4",
            "B2B09_H_DIFF", "B2B09_HN_1", "B2B09_HN_2", "B2B09_SS_1", "B2B09_SS_2", "B2B09_SS_3", "B2B09_SS_4",
            "B2B10_H_DIFF", "B2B10_HN_1", "B2B10_HN_2", "B2B10_SS_1", "B2B10_SS_2", "B2B10_SS_3", "B2B10_SS_4",
            "B2B11_H_DIFF", "B2B11_HN_1", "B2B11_HN_2", "B2B11_SS_1", "B2B11_SS_2", "B2B11_SS_3", "B2B11_SS_4",
            "COAX01_SS", "COAX02_SS", "SCREW01_AH_N", "SCREW02_AH_N", "SCREW03_AH_N", "SCREW04_AH_N", "SCREW05_AH_N"
        };


        public MainWindow()
        {
            InitializeComponent();
        }


        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.Focus(); // Ensure the window can receive key events
        }


        private void Window_KeyDown2(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.V && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {

                if (Clipboard.ContainsImage())
                {
                    if (pastedImagePath.Count >= MaxImages)
                    {
                        MessageBox.Show("You can only paste up to 5 images.");
                        return;
                    }

                    BitmapSource bitmapSource = Clipboard.GetImage();
                    UploadedImage.Source = bitmapSource;

                    string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"pastedImage_{pastedImagePath.Count + 1}.png");
                    using (var fileStream = new FileStream(tempPath, FileMode.Create))
                    {
                        BitmapEncoder encoder = new PngBitmapEncoder();
                        encoder.Frames.Add(BitmapFrame.Create(bitmapSource));
                        encoder.Save(fileStream);
                    }

                    pastedImagePath.Add(tempPath);
                }
            }
        }


        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.V && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                if (Clipboard.ContainsImage())
                {
                    if (pastedImagePath.Count >= MaxImages)
                    {
                        MessageBox.Show("You can only paste up to 5 images.");
                        return;
                    }

                    BitmapSource bitmapSource = Clipboard.GetImage();

                    string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"pastedImage_{pastedImagePath.Count + 1}.png");
                    using (var fileStream = new FileStream(tempPath, FileMode.Create))
                    {
                        BitmapEncoder encoder = new PngBitmapEncoder();
                        encoder.Frames.Add(BitmapFrame.Create(bitmapSource));
                        encoder.Save(fileStream);
                    }

                    pastedImagePath.Add(tempPath);

                    // Automatically assign to next available slot
                    switch (pastedImagePath.Count)
                    {
                        case 1:
                            ImageSlot1.Source = bitmapSource;
                            break;
                        case 2:
                            ImageSlot2.Source = bitmapSource;
                            break;
                        case 3:
                            ImageSlot3.Source = bitmapSource;
                            break;
                        case 4:
                            ImageSlot4.Source = bitmapSource;
                            break;
                        case 5:
                            ImageSlot5.Source = bitmapSource;
                            break;
                    }
                }
                else
                {
                    MessageBox.Show("No image found in clipboard.");
                }
            }
        }

        private void UploadImage_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();//abrir caixa de dialogo
            dlg.Filter = "Image Files (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp";//define os tipos de imagem, podia ser só bmp que daria certo, mas coloquei mais

            if (dlg.ShowDialog() == true)//caso a caixa abra
            {
                
                imagePath = dlg.FileName;//caminho do arquivo
                UploadedImage.Source = new BitmapImage(new Uri(imagePath));

                /*
                string originalPath = dlg.FileName;
                imagePath = InvertImageColors(originalPath); // Save and use inverted image
                UploadedImage.Source = new BitmapImage(new Uri(imagePath));*/

                ExtractedTextBox.Clear();//limpa a caixa de texto
            }
        }

        public void extractText()
        {
            try
            {

                string projectRoot = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName;
                string tessDataPath = Path.Combine(projectRoot, "packages", "Tesseract.5.2.0", "tessdata");

                if (!Directory.Exists(tessDataPath))
                {
                    throw new DirectoryNotFoundException($"Tessdata folder not found at: {tessDataPath}");
                }

                string csvColumns = "LOOKUP2, LOOKUP1, HEADER IDX, PROGRAM NO, TOOL NO, TOOL NAME, REF_DES, LOW LIMIT, HIGH LIMIT, RESULT, TOOL STATUS, TEST, DEVICE NUMBER";

                string toolStatus = "Pass";
                ExtractedTextBox.Text = csvColumns + Environment.NewLine;

                
                var validImages = pastedImagePath
                            .Where(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path))
                            .ToList();

                if (validImages.Count == 0)
                {
                    MessageBox.Show("No valid images to extract. Please paste at least one image.", "Extraction Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    
                    return;
                }


                if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))

                //foreach (var imagePath in pastedImagePath)
                for (int i = 0; i < pastedImagePath.Count; i++)
                {
                    var imagePath = pastedImagePath[i];

                    string selectedProg = "UNKNOWN";
                    string selectedTest = "UNKNOWN";
                    string selectedDevice = "UNKNOWN";

                    switch (i)
                    {
                        case 0: 
                            { 
                                selectedTest = (testList1.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "UNKNOWN";
                                selectedProg = (progList1.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "UNKNOWN";
                                selectedDevice = (deviceList1.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "UNKNOWN";
                                break;
                            }

                        case 1:
                            {
                                selectedTest = (testList2.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "UNKNOWN";
                                selectedProg = (progList2.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "UNKNOWN";
                                selectedDevice = (deviceList2.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "UNKNOWN";
                                break;
                            }
                        case 2:
                            {
                                selectedTest = (testList3.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "UNKNOWN";
                                selectedProg = (progList3.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "UNKNOWN";
                                selectedDevice = (deviceList3.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "UNKNOWN";
                                break;
                            }

                        case 3:
                            {
                                selectedTest = (testList4.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "UNKNOWN";
                                selectedProg = (progList4.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "UNKNOWN";
                                selectedDevice = (deviceList4.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "UNKNOWN";
                                break;
                            }
                        case 4:
                            {
                                selectedTest = (testList5.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "UNKNOWN";
                                selectedProg = (progList5.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "UNKNOWN";
                                selectedDevice = (deviceList5.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "UNKNOWN";
                                break;
                            }

                        }

                    if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))
                    continue;

                    using (var engine = new TesseractEngine(tessDataPath, "eng", EngineMode.Default))
                    using (var img = Pix.LoadFromFile(imagePath))
                    using (var page = engine.Process(img))
                    {
                        var iterator = page.GetIterator();
                        iterator.Begin();

                        string allWords = "";
                        do
                        {
                            if (iterator.IsAtBeginningOf(PageIteratorLevel.Word))
                            {
                                string word = iterator.GetText(PageIteratorLevel.Word);
                                if (!string.IsNullOrWhiteSpace(word))
                                {
                                    allWords += word + Environment.NewLine;
                                }
                            }
                        } while (iterator.Next(PageIteratorLevel.Word));

                        string processed = ProcessExtractedText(allWords,  selectedProg, selectedTest, toolStatus, selectedDevice);
                        ExtractedTextBox.Text += processed + Environment.NewLine + Environment.NewLine;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error during OCR: " + ex.Message);
            }
        }


        private string ProcessExtractedText(string rawText, string selectedProg,string selectedTest, string toolStatus, string selectedDevice)
        {
            var lines = rawText.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            var processedResults = new List<string>();

            

            for (int i = 0; i < lines.Length - 1; i++)
            {
                string originalLine = lines[i].Trim();

                // Check if the line contains "Tool" or any known variation
                bool isToolLine = corrections.Keys.Any(key => originalLine.Contains(key)) || originalLine.Contains("Tool");

                if (isToolLine)
                {
                    // Apply corrections to the line
                    string correctedLine = originalLine;
                    foreach (var kvp in corrections)
                    {
                        correctedLine = correctedLine.Replace(kvp.Key, kvp.Value);
                    }

                    // Extract tool name
                    Match toolMatch = Regex.Match(correctedLine, @"Tool\d{2}|Too\d{2}|To0\d{2}");
                    string tool = toolMatch.Success ? toolMatch.Value : "[UNKNOWN TOOL]";

                    // Extract test name
                    
                    string toolName = correctedLine;

                    if (toolMatch.Success)
                    {
                        int index = correctedLine.IndexOf(toolMatch.Value);
                        toolName = correctedLine.Substring(index + toolMatch.Value.Length).Trim(':', ' ', '(', ')');

                        // Apply corrections to test name
                        foreach (var kvp in corrections)
                        {
                            toolName = toolName.Replace(kvp.Key, kvp.Value);
                        }

                        
                    }
                    // Try to extract value from next few lines
                    string value = "[NO VALUE FOUND]";
                    int linesUsed = 1;
                    for (int j = 1; j <= 3 && (i + j) < lines.Length; j++)
                    {
                        string candidate = string.Join("", lines.Skip(i + 1).Take(j).Select(l => l.Trim()));
                        foreach (var kvp in corrections)
                        {
                            candidate = candidate.Replace(kvp.Key, kvp.Value);
                        }

                        Match match = Regex.Match(candidate, @"-?\d+\.\d+");

                        if (match.Success)
                        {
                            value = match.Value;
                            linesUsed = j;
                            break;
                        }
                    }


                    int testNum = int.Parse(Regex.Match(selectedTest, @"\d+").Value);
                    int deviceNum = int.Parse(selectedDevice);
                    int headerIdx = testNum + 10 * (deviceNum - 1);


                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    ///
                    toolName = Regex.Replace(toolName, @"\(.*", "").Trim();// esse \( significa literalmente '(', o '.*' pega qualquer valor. então estou procurando
                    // por qualquer coisa depois de toolname que comece com (, que é o que mais da problema na hora de passar
                    //tool = Regex.Replace(tool, @"Tool", "").Trim();
                    tool = Regex.Match(tool, @"\d+").Value;

                    int ll = -20;
                    int ul = 20;

                    string lookUp2 = $"{headerIdx}_{toolName}_{toolStatus}";
                    string lookUp1 = $"{headerIdx}{toolName}";

                    // string csvColumns = "TOOL NO, TOOL NAME, RESULT, TEST, PROGRAM NO., TOOL STATUS, LOW LIMIT, HIGH LIMIT";
                    //processedResults.Add($"{tool}, {toolName}, {value}, {selectedTest}, {selectedProg}, {toolStatus}, {ll}, {ul}");
                    processedResults.Add($"{lookUp2},{lookUp1},{headerIdx}, {selectedProg},{tool}, {toolName}, {toolName}, {ll}, {ul},{value}, {toolStatus},{selectedTest}, {selectedDevice}");




                    i += linesUsed; // Skip value lines
                }
            }
            return string.Join(Environment.NewLine, processedResults);
        }


        private void PasteImageToSlot(int index, System.Windows.Controls.Image imageControl)
        {
            if (index < pastedImagePath.Count)
            {
                imageControl.Source = new BitmapImage(new Uri(pastedImagePath[index]));
            }
            else
            {
                MessageBox.Show($"No image available for slot {index + 1}.");
            }
        }
        private void ExtractText_Click(object sender, RoutedEventArgs e)
        {
            extractText();
        }

        private void PasteToSlot1_Click(object sender, RoutedEventArgs e)
        {
            if (Clipboard.ContainsImage())
            {
                BitmapSource bitmapSource = Clipboard.GetImage();
                ImageSlot1.Source = bitmapSource;

                string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "pastedImage_1.png");
                using (var fileStream = new FileStream(tempPath, FileMode.Create))
                {
                    BitmapEncoder encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(bitmapSource));
                    encoder.Save(fileStream);
                }

                if (pastedImagePath.Count >= 1)
                {
                    pastedImagePath[0] = tempPath;
                }
                else
                {
                    pastedImagePath.Add(tempPath);
                }
            }
            else
            {
                MessageBox.Show("No image found in clipboard.");
            }
            //foPasteImageToSlot(0, ImageSlot1);
        }

        private void PasteToSlot2_Click(object sender, RoutedEventArgs e)
        {
            if (Clipboard.ContainsImage())
            {
                BitmapSource bitmapSource = Clipboard.GetImage();
                ImageSlot2.Source = bitmapSource;

                string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "pastedImage_2.png");
                using (var fileStream = new FileStream(tempPath, FileMode.Create))
                {
                    BitmapEncoder encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(bitmapSource));
                    encoder.Save(fileStream);
                }

                if (pastedImagePath.Count >= 2)
                {
                    pastedImagePath[1] = tempPath;
                }
                else
                {
                    pastedImagePath.Add(tempPath);
                }
            }
            else
            {
                MessageBox.Show("No image found in clipboard.");
            }
        }

        private void PasteToSlot3_Click(object sender, RoutedEventArgs e)
        {
            if (Clipboard.ContainsImage())
            {
                BitmapSource bitmapSource = Clipboard.GetImage();
                ImageSlot3.Source = bitmapSource;

                string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "pastedImage_3.png");
                using (var fileStream = new FileStream(tempPath, FileMode.Create))
                {
                    BitmapEncoder encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(bitmapSource));
                    encoder.Save(fileStream);
                }

                if (pastedImagePath.Count >= 3)
                {
                    pastedImagePath[2] = tempPath;
                }
                else
                {
                    pastedImagePath.Add(tempPath);
                }
            }
            else
            {
                MessageBox.Show("No image found in clipboard.");
            }
        }

        private void PasteToSlot4_Click(object sender, RoutedEventArgs e)
        {

            if (Clipboard.ContainsImage())
            {
                BitmapSource bitmapSource = Clipboard.GetImage();
                ImageSlot4.Source = bitmapSource;

                string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "pastedImage_4.png");
                using (var fileStream = new FileStream(tempPath, FileMode.Create))
                {
                    BitmapEncoder encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(bitmapSource));
                    encoder.Save(fileStream);
                }

                if (pastedImagePath.Count >= 4)
                {
                    pastedImagePath[3] = tempPath;
                }
                else
                {
                    pastedImagePath.Add(tempPath);
                }
            }
            else
            {
                MessageBox.Show("No image found in clipboard.");
            }
        }

        private void PasteToSlot5_Click(object sender, RoutedEventArgs e)
        {

            if (Clipboard.ContainsImage())
            {
                BitmapSource bitmapSource = Clipboard.GetImage();
                ImageSlot5.Source = bitmapSource;

                string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "pastedImage_5.png");
                using (var fileStream = new FileStream(tempPath, FileMode.Create))
                {
                    BitmapEncoder encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(bitmapSource));
                    encoder.Save(fileStream);
                }

                if (pastedImagePath.Count >= 5)
                {
                    pastedImagePath[4] = tempPath;
                }
                else
                {
                    pastedImagePath.Add(tempPath);
                }
            }
            else
            {
                MessageBox.Show("No image found in clipboard.");
            }
        }



        public static string ReplaceInsensitive(string input, string oldValue, string newValue)
        {
            return System.Text.RegularExpressions.Regex.Replace(input, System.Text.RegularExpressions.Regex.Escape(oldValue), newValue, RegexOptions.IgnoreCase);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

        }

        private void ComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {

        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void ClearSlot1_Click(object sender, RoutedEventArgs e)
        {
            ImageSlot1.Source = null;

            if (pastedImagePath.Count >= 1)
            {
                pastedImagePath[0] = null;
            }
        }

        private void ClearSlot2_Click(object sender, RoutedEventArgs e)
        {
            ImageSlot2.Source = null;

            if (pastedImagePath.Count >= 2)
            {
                pastedImagePath[1] = null;
            }
        }

        private void ClearSlot3_Click(object sender, RoutedEventArgs e)
        {
            ImageSlot3.Source = null;

            if (pastedImagePath.Count >= 3)
            {
                pastedImagePath[2] = null;
            }
        }

        private void ClearSlot4_Click(object sender, RoutedEventArgs e)
        {
            ImageSlot4.Source = null;

            if (pastedImagePath.Count >= 4)
            {
                pastedImagePath[3] = null;
            }
        }

        private void ClearSlot5_Click(object sender, RoutedEventArgs e)
        {
            ImageSlot5.Source = null;

            if (pastedImagePath.Count >= 5)
            {
                pastedImagePath[4] = null;
            }
        }

        public void csvExport_method(string filePath)
        {
           
            try
            {
                var lines = ExtractedTextBox.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                bool appendToLast = lastCsv.IsChecked == true && !string.IsNullOrEmpty(lastSavedCsvPath) && File.Exists(lastSavedCsvPath);

                using (var writer = new StreamWriter(appendToLast ? lastSavedCsvPath : filePath, appendToLast))
                {
                    if (!appendToLast)
                    {
                        // Write header only if creating a new file
                        string header = lines[0];
                        writer.WriteLine(header);
                    }

                    foreach (var line in lines.Skip(1))
                    {
                        writer.WriteLine(line);
                    }
                }

                if (!appendToLast)
                {
                    lastSavedCsvPath = filePath; // Save the path for future appends
                }

                if (!lastCsv.IsChecked == true)
                {
                    MessageBox.Show("CSV file exported successfully!");
                }

                lastSavedCsvName.Content = System.IO.Path.GetFileName(appendToLast ? lastSavedCsvPath : filePath);
            }

            catch (Exception ex)
                {
                    MessageBox.Show("Error exporting CSV: " + ex.Message);
                }
        }

        

        private void csvExport_click(object sender, RoutedEventArgs e)
        {
            if(lastCsv.IsChecked == true && !string.IsNullOrEmpty(lastSavedCsvPath) && File.Exists(lastSavedCsvPath))
            {
                csvExport_method(lastSavedCsvPath);

                MessageBox.Show($"File saved as '{System.IO.Path.GetFileName(lastSavedCsvPath)}'","CSV Updated", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                var dialog = new Microsoft.Win32.SaveFileDialog
                {
                    Title = "Save CSV file",
                    Filter = "CSV files (*.csv)|*.csv",
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    FileName = "Exported_csv.csv"
                };

                if (dialog.ShowDialog() == true)
                {
                    csvExport_method(dialog.FileName);
                }
            }
        }

        private void addCsv_click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Select CSV file to append",
                Filter = "CSV files (*.csv)|*.csv",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };

            if (dialog.ShowDialog() == true)
            {
                var lines = ExtractedTextBox.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

                try
                {
                    using (var writer = new StreamWriter(dialog.FileName, true))
                    {
                        foreach (var line in lines.Skip(1)) // Skip header
                        {
                            //writer.WriteLine(line + "," + selectedProg);
                            writer.WriteLine(line);
                        }
                    }

                    lastSavedCsvPath = dialog.FileName;
                    lastSavedCsvName.Content = System.IO.Path.GetFileName(dialog.FileName); // Update the TextBox

                    MessageBox.Show($"Data appended to '{System.IO.Path.GetFileName(dialog.FileName)}'", "CSV Updated", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error appending to CSV: " + ex.Message);
                }
            }

        }

        private void processTxt_Click(object sender, RoutedEventArgs e)
        {
            var lines = ExtractedTextBox.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);//separar em linhas o conteudo dentro do textbox
            var cleanedLines = new List<string>();//armazena uma lista contendo as linhsa pos processamento

            foreach ( var line in lines)
            {
                var parts = line.Split(',');//separa baseado nas virgulas

                if (parts.Length >=2)//caso tenha mais de duas colunas
                {
                    parts[1]= Regex.Replace(parts[1], @"\(.*", "").Trim();//remove caso tenha parenteses

                    parts[1] = Regex.Replace(parts[1], @"\b\w+\b", match =>
                    {
                        var word = match.Value;
                        return corrections.ContainsKey(word) ? corrections[word] : word;
                    });

                }
                cleanedLines.Add(string.Join(",", parts));
            }

            ExtractedTextBox.Text = string.Join(Environment.NewLine, cleanedLines);
        }
    }
}
