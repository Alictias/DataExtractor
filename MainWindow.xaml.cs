using System;
using System.Windows;
using Microsoft.Win32;
using System.Windows.Media.Imaging;
using Tesseract;

namespace DataExtractor
{
    public partial class MainWindow : Window
    {
        private string imagePath;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void UploadImage_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();//abrir caixa de dialogo
            dlg.Filter = "Image Files (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp";//define os tipos de imagem, podia ser só bmp que daria certo, mas coloquei mais

            if (dlg.ShowDialog() == true)//caso a caixa abra
            {
                imagePath = dlg.FileName;//caminho do arquivo
                UploadedImage.Source = new BitmapImage(new Uri(imagePath));//
                ExtractedTextBox.Clear();//limpa a caixa de texto
            }
        }

        private void ExtractText_Click(object sender, RoutedEventArgs e)//botão extract text
        {
            try
            {
                string tessDataPath = @"C:\Users\jagaminh\Desktop\DataExtractor\packages\Tesseract.5.2.0\tessdata";
                using (var engine = new TesseractEngine(tessDataPath, "eng", EngineMode.Default))


                using (var img = Pix.LoadFromFile(imagePath))
                {
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

                        ExtractedTextBox.Text = allWords;
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error during OCR: " + ex.Message);
            }
        }
    }
}
