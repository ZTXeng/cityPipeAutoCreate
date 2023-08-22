using Autodesk.Revit.DB;
using MathNet.Numerics.LinearAlgebra.Factorization;
using Org.BouncyCastle.Pqc.Crypto.Lms;
using PipeAutoCreate.ExcelControl;
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
        ScrollViewer _dataSVer;
        ScrollViewer _userDefineSVer;
        ExcelShowViewModel _viewModel;
        public ExcelShowView(Document doc)
        {
            InitializeComponent();

            _viewModel = new ExcelShowViewModel(doc);
            DataContext = _viewModel;
        }

        public ExcelShowView()
        {
            InitializeComponent();

            _viewModel = new ExcelShowViewModel();   
            DataContext = _viewModel;

            //ExcelData.ItemsSource = _viewModel.Model.CurrentDataTable.DefaultView;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            AddCobox();
        }

        public void AddCobox()
        {
            var columsCount = this.ExcelData.Columns.Count;

            this.UserDefine.Children.Clear();
            for (int i = 0; i < columsCount; i++)
            {
                var columnWith = this.ExcelData.Columns[i].Width;
                var combox = new ComboBox();
                if (i == 0)
                {
                    combox.SelectedIndex = 0;
                }
                else if (i == 1)
                {
                    combox.SelectedIndex = 1;
                }
                else if (i == 3) { combox.SelectedIndex = 7; }
                else if (i == 4) { combox.SelectedIndex = 2; }
                else if (i == 5) { combox.SelectedIndex = 3; }
                else if (i == 6) { combox.SelectedIndex = 4; }
                else if (i == 7) { combox.SelectedIndex = 5; }
                else if (i == 9) { combox.SelectedIndex = 6; }
                else { combox.SelectedIndex = 8; }


                combox.Width = columnWith.DisplayValue;
                combox.Height = 26;
                combox.ItemsSource = new List<string>() { "起点编号", "终点编号", "X", "Y", "地面高程", "管线高程", "规格", "附属物", "属性" };
                this.UserDefine.Children.Add(combox);
            }
        }

        private void Export(object sender, RoutedEventArgs e)
        {
            var columsCount = this.ExcelData.Columns.Count;
            var userDefinited = new List<string>();

            for (int i = 0; i < columsCount; i++)
            {
                var comboBox = this.UserDefine.Children[i] as ComboBox;
                var selectedValue = comboBox.SelectedItem?.ToString();

                if (selectedValue != null)
                {
                    userDefinited.Add(selectedValue);
                }
                else
                {
                    userDefinited.Add(string.Empty);
                }
            }

            _viewModel.Model.UserDefinited = userDefinited;

            _viewModel.Model.UserChanged=_viewModel.Model.UserChanged==true?false:true; 
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            _dataSVer = VisualTreeHelper.GetChild(VisualTreeHelper.GetChild(this.ExcelData, 0), 0) as ScrollViewer;
            _userDefineSVer = this.UserDifineSVer;

            _dataSVer.ScrollChanged += new ScrollChangedEventHandler(sv1_ScrollChanged);

            AddCobox();
        }

        void sv1_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            _userDefineSVer.ScrollToHorizontalOffset(_dataSVer.HorizontalOffset);
        }

        private void Check(object sender, RoutedEventArgs e)
        {
            //CheckData.ItemsSource = _viewModel.Model.CheckCurrentDataTable.DefaultView;
        }
    }
}
