using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows.Input;
using XLSXBulkDataExtractor.WPFLogic.Globals;
using XLSXBulkDataExtractor.Common;
using XLSXBulkDataExtractor.WPFLogic.Interfaces;
using XLSXDataExtractor;
using System.Linq;
using XLSXDataExtractor.Models;
using System.Collections.Generic;
using XLSXBulkDataExtractor.MVVMHelpers.MVVM_Extensions;
using ClosedXML.Excel;
using XLSXBulkDataExtractor.WPFLogic.Models;

namespace XLSXBulkDataExtractor.WPFLogic.ViewModels
{
    public class DataRetrievalViewModel : NotifyPropertyChangedBase
    {
        private IIOService _ioService;
        private IXLIOService _xlioService;
        private IUIControlsService _uiControlsService;

        DataRetrievalRequest _dataRetrievalRequest;
        public DataRetrievalRequest SelectedDataRetrievalRequest
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

        public IEnumerable<DataOutputFormat> OutputFormats
        {
            get
            {
                return Enum.GetValues(typeof(DataOutputFormat)).Cast<DataOutputFormat>();
            }
        }

        public DataOutputFormat ChosenOutputFormat { get; set; } = DataOutputFormat.XLSX;

        public ObservableCollection<DataRetrievalRequest> DataRetrievalRequests { get; set; } = new ObservableCollection<DataRetrievalRequest>();

        public ICommand AddExtractionRequestCommand { get; private set; }
        public ICommand DeleteExtractionRequestCommand { get; private set; }
        public ICommand BeginExtractionCommand { get; private set; }
        public ICommand SetOutputDirectoryCommand { get; private set; }

        public DataRetrievalViewModel(IIOService ioService, IXLIOService xlioService, IUIControlsService uiControlsService)
        {
            _ioService = ioService;
            _xlioService = xlioService;
            _uiControlsService = uiControlsService;

            AddExtractionRequestCommand = new RelayCommand(() => AddNewEmptyExtractionRequest());

            DeleteExtractionRequestCommand = new RelayCommand(() =>
            {
                try
                {
                    DeleteExtractionRequest(SelectedDataRetrievalRequest);
                }
                catch (CollectionEmptyException)
                {
                    _uiControlsService.DisplayAlert("No data retrieval requests have been added", "Error!", MessageType.Error);
                }
                catch (ArgumentNullException)
                {
                    _uiControlsService.DisplayAlert("No data retrieval request selected", "Error!", MessageType.Error);
                }
            });

            BeginExtractionCommand = new RelayCommand(() =>
            {
                try
                {
                    BeginExtraction(new DirectoryInfo(OutputDirectory), ChosenOutputFormat);
                }
                catch (ArgumentNullException e)
                {
                    if (e.ParamName.ToLower() == "documentsdirectory") _uiControlsService.DisplayAlert("No output directory set", "Alert!", MessageType.Error);
                }
                catch (NoDataOutputtedException e)
                {
                    _uiControlsService.DisplayAlert(e.Message, e.ExceptionTitle, MessageType.Error);
                }
            });

            SetOutputDirectoryCommand = new RelayCommand(() =>
            {
                var chosenPath = SetOutputDirectory();

                if (!string.IsNullOrWhiteSpace(chosenPath)) OutputDirectory = chosenPath;
            });

        }

        private void AddNewEmptyExtractionRequest()
        {
            DataRetrievalRequests.Add(new DataRetrievalRequest());
        }

        private void DeleteExtractionRequest(DataRetrievalRequest dataRetrievalRequest)
        {
            if (DataRetrievalRequests.Count == 0) throw new CollectionEmptyException();
            if (dataRetrievalRequest == null) throw new ArgumentNullException("dataRetrievalRequest", "Cannot be null");

            DataRetrievalRequests.Remove(dataRetrievalRequest); //not concerned about successful removal, no need to interpret return bool.
        }

        private string SetOutputDirectory()
        {
            return _ioService.ChooseFolderDialog();
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
                SaveExtractedData(dataOutputFormat, extractedDataCol);
            }
            else
            {
                throw new NoDataOutputtedException($"No detected in files at {documentsDirectory.FullName}", "No Data");
            }
        }

        private void SaveExtractedData(DataOutputFormat dataOutputFormat, IEnumerable<IEnumerable<KeyValuePair<string, object>>> extractedDataCol)
        {
            ReturnMessage succesfullySaved;

            switch (dataOutputFormat)
            {
                case DataOutputFormat.XLSX:
                    var generatedWorksheet = ExtractedDataConverter.ConvertToWorksheet(extractedDataCol);

                    var newWorkbook = new XLWorkbook();
                    newWorkbook.AddWorksheet(generatedWorksheet);
                    succesfullySaved = _xlioService.SaveWorkbook(Path.Combine(OutputDirectory, "Output.xlsx"), newWorkbook);
                    DisplaySuccessOrFailMessage(succesfullySaved);

                    break;
                case DataOutputFormat.CSV:
                    var generatedCSV = ExtractedDataConverter.ConvertToCSV(extractedDataCol);
                    succesfullySaved = _ioService.SaveText(generatedCSV, Path.Combine(OutputDirectory, "Output.csv"));
                    DisplaySuccessOrFailMessage(succesfullySaved);

                    break;
                default:
                    break;
            }
        }

        private void DisplaySuccessOrFailMessage(ReturnMessage returnMessage)
        {
            if (returnMessage == null) throw new ArgumentNullException("returnMessage", "cannot be null");

            if (returnMessage.Success) _uiControlsService.DisplayAlert(returnMessage.Message, "Success!", MessageType.Information);
          
            else _uiControlsService.DisplayAlert(returnMessage.Message, "Error!", MessageType.Error);
            
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
