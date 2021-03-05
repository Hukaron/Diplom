using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Collections.ObjectModel;
using System.Data;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.Serialization.Formatters;
using System.Runtime.Serialization;
using System.Xml.Serialization;
using System.Xml;
using System.IO;
using Microsoft.Win32;
using System.Collections;
using System.Runtime.Serialization.Formatters.Binary;
using Word = Microsoft.Office.Interop.Word;

namespace CTT_PROG
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    //Проверка работы гита и сохранения им данных. Надеюсь всё работает
    //Тест работы гита. Должна появиться эта строчка в репе!
    //Задачи продублирую здесь:
    //1. В ворде настроить правильное форматирование
    //2. Оформить раздел 3 так, как он должен быть оформлен полностью
    //3. Промасштабировать всю программу, или на каждый раздел записать свою функцию... не знаю насколько это верно... надо будет подумать
    [Serializable]
    public class Products : INotifyPropertyChanged
    {  

        public Products() { }
        private string name;
        private int amount;
        private double price;
        private double total;

        public string Name { get { return name; } set { name=value; OnPropertyChanged("Name"); } }
        public int Amount { get { return amount; } set { amount = value; total = amount * price; OnPropertyChanged("Amount"); OnPropertyChanged("Total"); } }
        public double Price { get { return price; } set { price = value; total = amount * price; OnPropertyChanged("Price"); OnPropertyChanged("Total"); } }
        public double Total { get { return total; } set { total = price * amount; OnPropertyChanged("Total"); } }
        
        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }

           
            public void Serialize(string FileName, ObservableCollection<Products> products)
            { 
                XmlSerializer xml = new XmlSerializer(typeof(ObservableCollection<Products>));
                using (FileStream fs = new FileStream(FileName, FileMode.Create))
                {

                    xml.Serialize(fs, products);

                }
            }

        public ObservableCollection<Products> Deserialize(string FileName)
        {
            ObservableCollection < Products > products= new ObservableCollection<Products>();
            FileName.Trim();
            XmlSerializer xml = new XmlSerializer(typeof(ObservableCollection<Products>));
            using (FileStream fs = new FileStream(FileName, FileMode.Open))
            {
                ObservableCollection<Products> collection = (ObservableCollection<Products>)xml.Deserialize(fs);
                foreach (Products product in collection)
                {
                    products.Add(product);

                }

            }
            return products;
        }


        [NonSerialized]
       private Word.Application wordapp;
        [NonSerialized]
       private Word.Document worddocument;
        [NonSerialized]
       private Word.Paragraphs wordparagraphs;
        [NonSerialized]
       private Word.Paragraph wordparagraph;
        //Подробнее изучить работу с Word. Сделать форматирование текста.
        public void SaveInWord(string fileName, ObservableCollection<Products> products, DataGrid table)
        {
            try
            {
                int i = 0;
                wordapp = new Word.Application();
                wordapp.Visible = false;
                Object template = Type.Missing;
                Object newTemplate = false;
                Object documentType = Word.WdNewDocumentType.wdNewBlankDocument;
                Object visible = true;
                Object oMissing = System.Reflection.Missing.Value;
                Object saveChanges = Word.WdSaveOptions.wdPromptToSaveChanges;
                Object originalFormat = Word.WdOriginalFormat.wdWordDocument;
                Object routeDocument = Type.Missing;
                //Создание документа
                worddocument= wordapp.Documents.Add(ref template, ref newTemplate, ref documentType, ref visible);
                worddocument.Content.ParagraphFormat.FirstLineIndent = worddocument.Content.Application.CentimetersToPoints((float)1);
                worddocument.Content.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                wordparagraphs = worddocument.Paragraphs;
                for ( i = 0; i < 8; i++) worddocument.Paragraphs.Add(ref oMissing);

                //Переходим к первому добавленному параграфу
                wordparagraph = worddocument.Paragraphs[2];
                Word.Range wordrange = wordparagraph.Range;
                //Добавляем таблицу в начало второго параграфа
                Object defaultTableBehavior = Word.WdDefaultTableBehavior.wdWord9TableBehavior;
                Object autoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitWindow;
                //Создаем таблицу
                Word.Table wordtable1 = worddocument.Tables.Add(wordrange, products.Count+1, table.Columns.Count, ref defaultTableBehavior, ref autoFitBehavior);
                //Заполнение таблицы
                Word.Range wordcellrange = wordtable1.Cell(1, 1).Range;
                wordcellrange.Text = "Название товара";
                wordcellrange = wordtable1.Cell(1, 2).Range;
                wordcellrange.Text = "Количество товара";
                wordcellrange = wordtable1.Cell(1, 3).Range;
                wordcellrange.Text = "Цена товара";
                wordcellrange = wordtable1.Cell(1, 4).Range;
                wordcellrange.Text = "Итоговая стоимость товара";
                i = 2;
                foreach (Products p in products)
                {
                        wordcellrange = wordtable1.Cell(i, 1).Range;
                        wordcellrange.Text = p.Name.ToString();
                        wordcellrange = wordtable1.Cell(i, 2).Range;
                        wordcellrange.Text = p.Amount.ToString()+" шт.";
                        wordcellrange = wordtable1.Cell(i, 3).Range; 
                        wordcellrange.Text = p.Price.ToString()+" бел. руб.";
                        wordcellrange = wordtable1.Cell(i, 4).Range;
                        wordcellrange.Text = p.Total.ToString()+" бел. руб";
                    i++;
                }

                worddocument.SaveAs(fileName);
                wordapp.Quit(ref saveChanges, ref originalFormat, ref routeDocument);
                MessageBox.Show("Сохранения завершено. Можете продолжить работу","Сохранение успешно!", MessageBoxButton.OK, MessageBoxImage.Information);
                wordapp = null;


            }
            catch (Exception ex)
            {
                string str= ex.Message.ToString();
                MessageBox.Show(str);
                Object saveChanges = Word.WdSaveOptions.wdPromptToSaveChanges;
                Object originalFormat = Word.WdOriginalFormat.wdWordDocument;
                Object routeDocument = Type.Missing;
                wordapp.Quit(ref saveChanges, ref originalFormat, ref routeDocument);
            }
        }




    }

    public partial class MainWindow : Window
    {
        ObservableCollection<Products> productList = new ObservableCollection<Products>();
        public Products p = new Products();

        public MainWindow()
        {
            Binding binding = new Binding();
            InitializeComponent();
            Table1.ItemsSource = productList;

        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.DefaultExt = ".xml";
            saveFileDialog.Filter = "XML-документ|*.xml";
            saveFileDialog.AddExtension=true;
            if (saveFileDialog.ShowDialog() == true)
            {

                p.Serialize(saveFileDialog.FileName, productList);
            } 
        }

        private void LoadButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter= "XML-документ|*.xml";
            
            if (openFileDialog.ShowDialog()==true)
            {
               productList=p.Deserialize(openFileDialog.FileName);
            }
            Table1.ItemsSource = productList;
        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.DefaultExt = ".docx";
            saveFileDialog.Filter = "Документ Word|*.docx";
            if (saveFileDialog.ShowDialog()==true)
            {
                p.SaveInWord(saveFileDialog.FileName, productList, Table1);
            }    
        }

    }
}
