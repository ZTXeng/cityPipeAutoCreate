using Autodesk.Revit.Creation;
using Org.BouncyCastle.Pqc.Crypto.Lms;
using PipeAutoCreate.ViewModel;
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

namespace PipeAutoCreate.View
{
    /// <summary>
    /// ExcelShowView.xaml 的交互逻辑
    /// </summary>
    public partial class ExcelShowView : Window
    {
        ExcelShowViewModel _viewModel;
        public ExcelShowView(Document doc)
        {
            InitializeComponent();

            DataContext = new ExcelShowViewModel(doc);
        }


        public ExcelShowView()
        {
            InitializeComponent();

            _viewModel = new ExcelShowViewModel();   
            DataContext = _viewModel;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var columsCount = this.ExcelData.Columns.Count;

            this.ThisGrid.Children.Clear();
            for (int i = 0; i < columsCount; i++)
            {
                var columnWith = this.ExcelData.Columns[i].Width;
                var combox = new ComboBox();
                combox.SelectedIndex = -1;
                combox.Width = columnWith.DisplayValue;
                combox.Height = 26;
                combox.ItemsSource = new List<string>() { "编号", "X", "Y", "管高","属性" };
                this.ThisGrid.Children.Add(combox);
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            var columsCount = this.ExcelData.Columns.Count;
            
            for (int i = 0; i < columsCount; i++)
            {
                var comboBox = this.ThisGrid.Children[i] as ComboBox;
                var selectedValue = comboBox.SelectedItem;
            }
        }
    }
}
