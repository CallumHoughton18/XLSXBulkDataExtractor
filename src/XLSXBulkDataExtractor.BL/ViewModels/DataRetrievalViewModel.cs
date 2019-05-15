using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows.Input;
using XLSXBulkDataExtractor.BL.Globals;
using XLSXBulkDataExtractor.BL.Models;
using XLSXBulkDataExtractor.BL.MVVM_Extensions;
using XLSXBulkDataExtractor.Common;
using System.Windows;
using XLSXBulkDataExtractor.BL.Interfaces;
using XLSXDataExtractor;
using System.Linq;
using XLSXDataExtractor.Models;
using System.Collections.Generic;

namespace XLSXBulkDataExtractor.BL.ViewModels
{
    public class DataRetrievalViewModel : NotifyPropertyChangedBase
    {
        private IIOService _iioService;

        DataRetrievalRequest _dataRetrievalRequest;
        DataRetrievalRequest SelectedDataRetrievalRequest
        {
            get
            {
                return _dataRetrievalRequest;
            }
            set
            {
                if (_dataRetrievalRequest != value)
                {
                    _dataRetrievalRequest = value;
                    OnPropertyChanged(nameof(SelectedDataRetrievalRequest));
                }
            }
        }
        private string _directoryForXLSXExtractionFiles;
        public string DirectoryForXLSXExtractionFiles
        {
            get
            {
                return _directoryForXLSXExtractionFiles;
            }
            set
            {
                if (_directoryForXLSXExtractionFiles != value)
                {
                    _directoryForXLSXExtractionFiles = value;
                    OnPropertyChanged(nameof(DirectoryForXLSXExtractionFiles));
                }
            }
        }
        private string _outputDirectory;
        public string OutputDirectory
        {
            get
            {
                return _outputDirectory;
            }
            set
            {
                if (_outputDirectory != value)
                {
                    _outputDirectory = value;
                    OnPropertyChanged(nameof(OutputDirectory));
                }
            }
        }
        public ObservableCollection<DataRetrievalRequest> DataRetrievalRequests { get; private set; } = new ObservableCollection<DataRetrievalRequest>();

        private readonly ICommand AddExtractionRequestCommand;
        private readonly ICommand DeleteExtractionRequestCommand;
        private readonly ICommand BeginExtractionCommand;
        private readonly ICommand SetOutputDirectoryCommand;

        public DataRetrievalViewModel(IIOService iioService)
        {
            _iioService = iioService;

            AddExtractionRequestCommand = new RelayCommand(() => AddNewEmptyExtractionRequest());

            DeleteExtractionRequestCommand = new RelayCommand(() =>
            {
                try
                {
                    DeleteExtractionRequest(SelectedDataRetrievalRequest);
                }
                catch (ArgumentNullException)
                {
                    //fire event to display error messagebox
                }
                catch (CollectionEmptyException)
                {
                    //fire event to display error messagebox
                }
            });

            SetOutputDirectoryCommand = new RelayCommand(() =>
            {
                var chosenPath = SetOutputDirectory();

                if (!string.IsNullOrWhiteSpace(chosenPath)) OutputDirectory = chosenPath;
            });
        }

        public void AddNewEmptyExtractionRequest()
        {
            DataRetrievalRequests.Add(new DataRetrievalRequest());
        }

        private void DeleteExtractionRequest(DataRetrievalRequest dataRetrievalRequest)
        {
            if (dataRetrievalRequest == null) throw new ArgumentNullException("dataRetrievalRequest", "Cannot be null");
            if (DataRetrievalRequests.Count == 0) throw new CollectionEmptyException();

            DataRetrievalRequests.Remove(dataRetrievalRequest); //not concerned about successful removal, no need to interpret return bool.
        }

        private string SetOutputDirectory()
        {
            return _iioService.ChooseFolderDialog();
        }

        private void BeginExtraction(DirectoryInfo documentsDirectory, DataOutputFormat dataOutputFormat)
        {
            //Possibly want this to be async, so the program doesn't lock up on a big extraction
            var validExtensions = new string[] { "xlsx", "xlsm" };
            if (documentsDirectory == null) throw new ArgumentNullException("fileInfo", "Cannot be null");
            IEnumerable<IEnumerable<KeyValuePair<string, object>>> extractedDataCol = null;

            foreach (var document in documentsDirectory.GetFiles())
            {
                if (validExtensions.Contains(Path.GetExtension(document.FullName).ToLower()))
                {
                    var dataExtractor = new DataExtractor(document.FullName);
                    var extractionRequests = DataRetrievalRequestCollectionToExtractionRequestCollection(DataRetrievalRequests);
                    extractedDataCol = dataExtractor.RetrieveDataCollectionFromAllWorksheets<object>(extractionRequests);
                }
            }

            if (extractedDataCol != null)
            {
                switch (dataOutputFormat)
                {
                    case DataOutputFormat.XLSX:
                        var generatedWorksheet = ExtractedDataConverter.ConvertToWorksheet(extractedDataCol);
                        break;
                    case DataOutputFormat.CSV:
                        var generatedCSV = ExtractedDataConverter.ConvertToCSV(extractedDataCol);
                        break;
                    default:
                        break;
                }
            }
        }

        private IEnumerable<ExtractionRequest> DataRetrievalRequestCollectionToExtractionRequestCollection(IEnumerable<DataRetrievalRequest> dataRetrievalRequestsCol)
        {
            foreach (var dataRetrievalRequest in dataRetrievalRequestsCol)
            {
                yield return new ExtractionRequest(dataRetrievalRequest.FieldName, dataRetrievalRequest.Row, dataRetrievalRequest.Column);
            }
        }
    }
}
