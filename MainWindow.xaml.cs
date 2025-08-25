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
                UploadedImage.Source = new BitmapImage(new Uri(imagePath));
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

                        ExtractedTextBox.Text = allWords;//atualiza a caixa de texto 
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