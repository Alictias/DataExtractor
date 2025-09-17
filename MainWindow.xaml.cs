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

namespace DataExtractor
{
    public partial class MainWindow : Window
    {
        private string imagePath;

        private readonly Dictionary<string, string> corrections = new Dictionary<string, string>
        {
            { "Tooi", "Tool" }, { "Too0", "Tool" }, { "To00", "Tool" }, { "Tooit", "Tool1" }, { "Tooi1", "Tool" },{"Tool0S","Tool05"},
            {"Tool", "Tool" }, {"SCREWO1","SCREW01"}, {"Tool03","Tool03"},
            { "82801", "B2B01" }, { "82802", "B2B02" }, { "82803", "B2B03" }, { "82804", "B2B04" },
            { "82805", "B2B05" }, { "82806", "B2B06" }, { "82807", "B2B07" }, { "82808", "B2B08" },
            { "82809", "B2B09" }, { "82810", "B2B10" }, { "82811", "B2B11" },{"82B02_SS_1", "B2B02_SS_1" },{"82B01_H_DIFF","B2B01_H_DIFF"},            
            { "55", "SS" }, { "HN", "HN" }, { "H_DIFF", "H_DIFF" }, { "Height", "HEIGHT" }, { "Calcul.", "Calculations"},
            { "step", "STEP" }, { "acstep", "STEP" }, { "Calcul..", "Calculations"}, { "Calcul", "Calculations"}, {"Aver...", "Average"}, {"aver...", "Average"}, 
            {"Aver..", "Average"}, {"aver..", "Average"},{ "Calcul...", "Calculations"},
            {"mmy", ""},{"mmp", ""},{"mmg", ""}, {"5CREW","SCREW"}, {"SCREWO3","SCREW03"}, {"5CREWO1", "SCREW01"},
            {"COAXO01", "COAX01"}, {"5CREWO3","5CREW03"}, {"5CREWO01", "SCREW01"}, {"5CREWO02", "SCREW02"},
            { "Too102", "Tool02" }, { "Too103", "Tool03" }, { "Too104", "Tool04" },
            { "Too105", "Tool05" }, { "Too106", "Tool06" }, { "Too107", "Tool07" },
            { "Too108", "Tool08" }, { "8Z801", "B2B01" }, { "88ZO2", "B2B02" },
            { "SS_Z", "SS_2" }, { "SS_J", "SS_3" }, { "S5", "SS" }


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


        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.V && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                if (Clipboard.ContainsImage())
                {
                    BitmapSource bitmapSource = Clipboard.GetImage();
                    UploadedImage.Source = bitmapSource;

                    //save to temp file for ocr
                    string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "pastedImage.png");
                    using (var fileStream = new FileStream(tempPath, FileMode.Create))
                    {

                        BitmapEncoder encoder = new PngBitmapEncoder();
                        encoder.Frames.Add(BitmapFrame.Create(bitmapSource));
                        encoder.Save(fileStream);

                    }

                    imagePath = tempPath;
                    ExtractedTextBox.Clear();

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


        private void ExtractText_Click(object sender, RoutedEventArgs e)//botão extract text
        {
            try
            {
                string tessDataPath = @"C:\Users\jagaminh\Desktop\DataExtractor\packages\Tesseract.5.2.0\tessdata";//caminho onde o tessdata se encontra, ferramenta responsavel pela leitura de imagens
                //deve ser alterado depois para outros computadores
                using (var engine = new TesseractEngine(tessDataPath, "eng", EngineMode.Default))  // inicializa o tesseract


                using (var img = Pix.LoadFromFile(imagePath)) //carrega a imagem selecionada num formato que o tesseract consegue ler
                {
                    using (var page = engine.Process(img)) //processa a imagem
                    {
                        var iterator = page.GetIterator(); //serve para ler o resultado 
                        iterator.Begin(); //inicializa a leitura

                        string allWords = "";//define uma variavel para guardar o resultado da leitura
                        do
                        {
                            if (iterator.IsAtBeginningOf(PageIteratorLevel.Word))//checa se o iterator está no inicio de uma palavra
                            {
                                string word = iterator.GetText(PageIteratorLevel.Word);//captura a palavra
                                if (!string.IsNullOrWhiteSpace(word))//se for diferente de null ou espaço vazio
                                {
                                    allWords += word + Environment.NewLine; //adiciona na variavel
                                }
                            }
                        } while (iterator.Next(PageIteratorLevel.Word));


                        //ExtractedTextBox.Text = allWords;//atualiza a caixa de texto 

                        string csvColumns = "Tool, Test, Type, Reading";
                        string processed = ProcessExtractedText(allWords);
                        MessageBox.Show(allWords);
                        ExtractedTextBox.Text = csvColumns;
                        ExtractedTextBox.Text += Environment.NewLine + processed;
                        
                        //MessageBox.Show($"Processed output:\n{processed}");


                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error during OCR: " + ex.Message);
            }
        }


        private string ProcessExtractedText(string rawText)
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
                    
                    string testName = correctedLine;

                    if (toolMatch.Success)
                    {
                        int index = correctedLine.IndexOf(toolMatch.Value);
                        testName = correctedLine.Substring(index + toolMatch.Value.Length).Trim(':', ' ', '(', ')');

                        // Apply corrections to test name
                        foreach (var kvp in corrections)
                        {
                            testName = testName.Replace(kvp.Key, kvp.Value);
                        }

                        if (!testName.EndsWith(")"))
                        {
                            testName += ")";
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

                    //processedResults.Add($"{tool}, {testName}, {value}");

//###########################################################################################
                    string testType = "[UNKNOWN TYPE]";
                    Match typeMatch = Regex.Match(testName, @"\((.*?)\)");
                    if (typeMatch.Success)
                    {
                        testType = typeMatch.Groups[1].Value;

                        // Normalize test type if it starts with "Calcu"
                        if (Regex.IsMatch(testType, @"^Calcu", RegexOptions.IgnoreCase))
                        {
                            testType = "Calculations";
                        }
                        else if (Regex.IsMatch(testType, @"^Step", RegexOptions.IgnoreCase))
                        {
                            testType = "Step";
                        }
                        else if (Regex.IsMatch(testType, @"^Height", RegexOptions.IgnoreCase))
                        {
                            testType = "Height";
                        }
                        else if (Regex.IsMatch(testType, @"^Average", RegexOptions.IgnoreCase))
                        {
                            testType = "Average";
                        }

                        // Remove the type from the test name
                        testName = Regex.Replace(testName, @"\s*\(.*?\)", "");
                    }

                    processedResults.Add($"{tool}, {testName}, {testType}, {value}");





                    i += linesUsed; // Skip value lines
                }
            }

            return string.Join(Environment.NewLine, processedResults);
        }


        private string ProcessExtractedText2(string rawText)
        {
            var lines = rawText.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            var processedResults = new List<string>();

            for (int i = 0; i < lines.Length - 1; i++)
            {
                string originalLine = lines[i].Trim();
                string nextLine = lines[i + 1].Trim();

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
                    //Match toolMatch = Regex.Match(correctedLine, @"Tool\d{2}"); /primeira versão
                    Match toolMatch = Regex.Match(correctedLine, @"Tool\d{2}|Too\d{2}|To0\d{2}");
                    
                    string tool = toolMatch.Success ? toolMatch.Value : "[UNKNOWN TOOL]";

                    // Try to extract test name (just take the part after the tool name)
                    string testName = correctedLine;
                    string testType="";

                    if (toolMatch.Success)
                    {
                        int index = correctedLine.IndexOf(toolMatch.Value);
                        testName = correctedLine.Substring(index + toolMatch.Value.Length).Trim(':', ' ', '(', ')');

                        // Ensure it ends with a closing parenthesis
                        if (!testName.EndsWith(")"))
                        {
                            testName += ")";
                        }

                        if (testName.EndsWith("ations"))
                        {
                            testName = "Calculations)";
                        }
                        if (Regex.IsMatch(testName, @"Calcu\w*", RegexOptions.IgnoreCase))
                        {
                            testName = "Calculations";
                        }

                    }

                    // Clean and extract value from next line
                    string cleanedValueLine = nextLine;

                    string valueCandidate = nextLine;
                    if (i + 2 < lines.Length)
                        valueCandidate += lines[i + 2].Trim();
                    if (i + 3 < lines.Length)
                        valueCandidate += lines[i + 3].Trim();

                    foreach (var kvp in corrections)
                    {
                        //cleanedValueLine = cleanedValueLine.Replace(kvp.Key, kvp.Value);
                        testName = testName.Replace(kvp.Key, kvp.Value);//segunda versão
                        valueCandidate = valueCandidate.Replace(kvp.Key, kvp.Value);
                    }

                    //Match valueMatch = Regex.Match(cleanedValueLine, @"-?\d+(\.\d+)?");
                    Match valueMatch = Regex.Match(valueCandidate, @"-?\d+(\.\d+)?");

                    string value = valueMatch.Success ? valueMatch.Value : "[NO VALUE FOUND]";

                    processedResults.Add($"{tool}, {testName}, {value}");
                    i++; // Skip value line
                }
            }

            return string.Join(Environment.NewLine, processedResults);
        }

        public static string ReplaceInsensitive(string input, string oldValue, string newValue)
        {
            return System.Text.RegularExpressions.Regex.Replace(input, System.Text.RegularExpressions.Regex.Escape(oldValue), newValue, RegexOptions.IgnoreCase);
        }


    }
}
