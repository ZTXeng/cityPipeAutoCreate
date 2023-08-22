using CommunityToolkit.Mvvm.ComponentModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PipeAutoCreate.Model
{
    public class ExcelShowModel : ObservableObject
    {
        private string _filePath;
        private string _originalText;
        private string _afterText;

        public List<DataTable> DataTables { get; set; }

        public DataTable CurrentDataTable { get; set; }

        public List<DataTable> CheckDataTables { get; set; }

        public DataTable CheckCurrentDataTable { get; set; }

        public string FilePath { 
            get=>_filePath; 
            set => SetProperty(ref _filePath,value);
        } 

        public string OriginalText
        { 
            get=> _originalText; 
            set => SetProperty(ref _originalText, value);
        } 

        public string AfterText
        { 
            get=> _afterText; 
            set => SetProperty(ref _afterText, value);
        }

        public List<string> Professions { get; set; }

        public string SelectedProfession { get; set; }

        public List<string> EquipmentNames { get; set; }

        public string SelectedEquipmentName { get; set; }

        public List<string> UserDefinited { get; set; }

        public bool _userChanged = false;

        public bool UserChanged
        {
            get { return _userChanged; }
            set
            {
                if (value != _userChanged)
                {
                    _userChanged = value;
                    OnUserChanged();
                }
            }
        }

        public event EventHandler UserChangedEvent;

        public ExcelShowModel()
        {
            DataTables = new List<DataTable>();
            UserDefinited = new List<string>();
            Professions = new List<string>();
            CheckDataTables = new List<DataTable> ();

            EquipmentNames = new List<string>() { "管线","井"};

            FilePath = "";
        }

        public virtual void OnUserChanged()
        {
            UserChangedEvent?.Invoke(this, EventArgs.Empty);
        }

        public void GetProfessions()
        {
            foreach (var item in DataTables)
            {
                Professions.Add(item.TableName);
            }

            SelectedProfession = Professions.First();
        }
    }
}
