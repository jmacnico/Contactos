using System;
using System.Collections.Generic;
using System.IO;
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
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfApplication1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            

        }

        private void btnCreateXLS_Click(object sender, RoutedEventArgs e)
        {

            Excel.Workbook workbook = null;
            try
            {
                List<Contactos> lstcontato = new List<Contactos>();

                Contactos contato = null;
                using (StreamReader sr = new StreamReader(@"C:\Users\jmalmeida\Desktop\Contactos_002.vcf"))
                {
                    string line;
                    // Read and display lines from the file until the end of  
                    // the file is reached. 
                    while ((line = sr.ReadLine()) != null)
                    {
                        if (line.Contains("BEGIN:VCARD"))
                            contato = new Contactos();
                        else if (line.Contains("FN:"))
                            contato.Nome = line.Split(':')[1];
                        else if (line.Contains("PREF:") || line.Contains("CELL:") || line.Contains("TEL:"))
                        {
                            string[] Numeros = line.Split(';');
                            foreach (string item in Numeros)
                            {
                                if (item.Contains(":"))
                                    contato.Numero.Add(item.Split(':')[1]);
                            }

                        }
                        else if (line.Contains("END:VCARD"))
                        {
                            if (contato.Numero.Count > 0)
                            {
                                bool exist = false;
                                foreach (string item in contato.Numero)
                                {
                                    Contactos contExist = (from t in lstcontato
                                                           where t.Numero.FirstOrDefault(y => y.Equals(item)) != null
                                                           select t).FirstOrDefault();
                                    if (contExist != null)
                                    {
                                        if (string.IsNullOrEmpty(contExist.Nome) && !string.IsNullOrEmpty(contato.Nome))
                                        {
                                            contExist.Nome = contato.Nome;
                                        }
                                        foreach (string numero in contato.Numero)
                                        {
                                            if (contato.Numero.FirstOrDefault(x => x == numero) == null)
                                                contExist.Numero.Add(numero);
                                        }
                                        exist = true;
                                        break;
                                    }
                                }
                                if (!exist)
                                    lstcontato.Add(contato);

                            }
                        }
                    }
                }
                Excel.Application xlApp = new Excel.Application();
                workbook = xlApp.Workbooks.Add();
                Excel._Worksheet workSheet = workbook.Worksheets[1];
                int index = 1;
                Excel.Range ThisRange = workSheet.get_Range("B1:Z10000", System.Type.Missing);
                ThisRange.NumberFormat = "@";


                foreach (Contactos item in lstcontato)
                {
                    int x = 2;
                    workSheet.Cells[index, 1] = item.Nome;
                    foreach (var item1 in item.Numero)
                    {
                        var cell = workSheet.Cells[index, x] = item1;
                        x++;
                    }
                    index++;
                }
                if (File.Exists(@"c:\your-file-name.xls"))
                    File.Delete(@"c:\your-file-name.xls");
                workbook.SaveAs(@"c:\your-file-name.xls");
                workbook.Close();
                lstteste.ItemsSource = lstcontato;
                MessageBox.Show("Ficheiro Guardado criado com suceeso em:\n\"c:\\your-file-name.xls\".")
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro:" + ex.Message.ToString());
            }
            finally
            {
                if (workbook != null)
                    workbook.Close();

            }
        }

        private void btnCreateVCL_Click(object sender, RoutedEventArgs e)
        {
            Excel.Workbook xlsWorkBook = null;
            StreamWriter strWriter = null;
            try
            {
                Excel.Application xlAPP = new Excel.Application();
                xlsWorkBook = xlAPP.Workbooks.Open(@"C:\Contactos.xls");
                Excel._Worksheet workSheet = xlsWorkBook.Worksheets[1];
                int linha = 1;
                strWriter = new StreamWriter(@"C:\Contactos.vcf");
                while (!string.IsNullOrEmpty(workSheet.Cells[linha, 1].Value))
                {
                    strWriter.WriteLine("BEGIN:VCARD");
                    strWriter.WriteLine("VERSION:2.1");
                    strWriter.WriteLine("N:" + workSheet.Cells[linha, 1].Value.ToString().Trim() + ";;;;");
                    strWriter.WriteLine("FN:" + workSheet.Cells[linha, 1].Value.ToString());
                    int coluna = 2;
                    while (!string.IsNullOrEmpty(workSheet.Cells[linha, coluna].Value))
                    {
                        if (coluna == 2)
                            strWriter.WriteLine("TEL;CELL;PREF:" + workSheet.Cells[linha, coluna].Value.ToString());
                        else
                            strWriter.WriteLine("TEL;CELL:" + workSheet.Cells[linha, coluna].Value.ToString());
                        coluna++;
                    }
                    strWriter.WriteLine("END:VCARD");
                    linha++;
                    MessageBox.Show("Ficheiro criado com Sucesso em:\n\"C:\\Contactos.vcf\"");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro:" + ex.Message.ToString());
            }
            finally
            {
                if (strWriter != null)
                    strWriter.Close();
                if (xlsWorkBook != null)
                    xlsWorkBook.Close();

            }
        }
    }
    public class Contactos
    {
        public string Nome { get; set; }
        public List<string> Numero { get; set; }
        public Contactos()
        {
            Numero = new List<string>();
        }
    }

}
