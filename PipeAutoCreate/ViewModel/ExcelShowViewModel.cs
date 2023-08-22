using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using CommunityToolkit.Mvvm.Input;
using PipeAutoCreate.DataModel;
using PipeAutoCreate.Environment;
using PipeAutoCreate.ExcelControl;
using PipeAutoCreate.Extension;
using PipeAutoCreate.Factory;
using PipeAutoCreate.Model;
using PipeAutoCreate.Request;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;
using static System.Resources.ResXFileRef;

namespace PipeAutoCreate.ViewModel
{
    public class ExcelShowViewModel : ViewModelBase<ExcelShowModel>
    {
        private List<Piping> _pipings;
        private List<BuildingAttachment> _attachments;
        private Document _doc;

        public IRelayCommand FileSelect { get; set; }

        public IRelayCommand<System.Windows.Controls.DataGrid> ProfessionChanged { get; set; }

        public IRelayCommand<System.Windows.Controls.DataGrid> EquipmentChanged { get; set; }

        public IRelayCommand Generate { get; set; }

        public ExcelShowViewModel(Document doc)
        {
            Model = new ExcelShowModel();

            _doc = doc;

            ProfessionChanged = new RelayCommand<System.Windows.Controls.DataGrid>(OnProfessionChanged);
            EquipmentChanged = new RelayCommand<System.Windows.Controls.DataGrid>(OnEquipmentChanged);
            Generate = new RelayCommand(OnGenerate);
            FileSelect = new RelayCommand(OnFileSelect);

            Model.UserChangedEvent += ExportDataCommand;
        }

        public ExcelShowViewModel()
        {
            Model = new ExcelShowModel();

            ProfessionChanged = new RelayCommand<System.Windows.Controls.DataGrid>(OnProfessionChanged);
            EquipmentChanged = new RelayCommand<System.Windows.Controls.DataGrid>(OnEquipmentChanged);
            Generate = new RelayCommand(OnGenerate);
            FileSelect = new RelayCommand(OnFileSelect);

            Model.UserChangedEvent += ExportDataCommand;
        }

        private void OnFileSelect()
        {
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                Model.FilePath = dialog.FileName;

                Model.DataTables = ExcelToData.ReadExcelDataTbales(Model.FilePath, true);
                Model.GetProfessions();
                Model.CurrentDataTable = Model.DataTables.First();
            }
        }

        private void  OnGenerate()
        {
            var tran = new Transaction(_doc);
            tran.Start("SS");
            PipeFactory.CreatePipes(_doc,_pipings);

            tran.Commit();
        }

        private void OnProfessionChanged(System.Windows.Controls.DataGrid par)
        {
            Model.CurrentDataTable = Model.DataTables.FirstOrDefault(x => x.TableName == Model.SelectedProfession);

            par.ItemsSource = Model.CurrentDataTable.DefaultView;

            Model.OriginalText = "合计数量： " + Model.CurrentDataTable.Rows.Count.ToString();
        }

        private void OnEquipmentChanged(System.Windows.Controls.DataGrid grid)
        {
            if (Model.SelectedEquipmentName == "管线")
            {
                Model.CheckCurrentDataTable = Model.CheckDataTables.FirstOrDefault();
            }
            else
            {
                Model.CheckCurrentDataTable = Model.CheckDataTables.LastOrDefault();
            }

            grid.ItemsSource = Model.CheckCurrentDataTable.DefaultView;

            Model.AfterText = "合计数量： " + Model.CheckCurrentDataTable.Rows.Count.ToString();
        }

        private void ExportDataCommand(object sender, EventArgs e)
        {
            var pipes = ExcelToData.ReadPipes(Model.CurrentDataTable, Model.UserDefinited);

            _pipings = pipes;

            var attachments = ExcelToData.ReadAttachments(Model.CurrentDataTable, Model.UserDefinited);

            _attachments = attachments;

            ExcelToData.WriteAttchments(new List<DataModel.ProfessionAttachment>() {
            new DataModel.ProfessionAttachment()
            {
               Profession = Model.SelectedProfession,
                BuildingAttachments = attachments
            }});
            ExcelToData.WritePipes(new List<DataModel.ProfessionPipe>()
            { new DataModel.ProfessionPipe() { PProfession = "dx", Pipings = pipes } });

            var pipedata = DataToExcel.ConvertTo(pipes);
            var attachmentData = DataToExcel.ConvertTo(attachments);

            Model.CheckDataTables.Clear();
            Model.CheckDataTables.Add(pipedata);
            Model.CheckDataTables.Add(attachmentData);
            Model.CheckCurrentDataTable = pipedata;
            Model.AfterText = "合计数量： " + Model.CheckCurrentDataTable.Rows.Count.ToString();
        }
    }
}
