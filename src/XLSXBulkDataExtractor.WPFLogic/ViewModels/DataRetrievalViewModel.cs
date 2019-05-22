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
using System.Threading.Tasks;
using System.Collections.Concurrent;
using MVVMHelpers.MVVM_Extensions;
using MVVMHelpers.Interfaces;
using System.Threading;

namespace XLSXBulkDataExtractor.WPFLogic.ViewModels
{
    public class DataRetrievalViewModel : NotifyPropertyChangedBase
    {
        private IIOService _ioService;
        private IXLIOService _xlioService;
        private IUIControlsService _uiControlsService;

        private CancellationTokenSource _extractionCancellationTokenSource = null;

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

        private double _extractionProgress = 0;
        public double ExtractionProgress
        {
            get
            {
                return _extractionProgress;
            }
            set
            {
                if (_extractionProgress != value)
                {
                    _extractionProgress = value;
                    OnPropertyChanged(nameof(ExtractionProgress));
                }
            }
        }

        private double _totalExtractionCount = 1;
        public double TotalExtractionCount
        {
            get
            {
                return _totalExtractionCount;
            }
            set
            {
                if (_totalExtractionCount != value)
                {
                    _totalExtractionCount = value;
                    OnPropertyChanged(nameof(TotalExtractionCount));
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
        public IAsyncCommand BeginExtractionCommand { get; private set; }
        public ICommand SetOutputDirectoryCommand { get; private set; }
        public Progress<IEnumerable<ReturnMessage>> ExtractionProgressEvent { get; } = new Progress<IEnumerable<ReturnMessage>>();

        public DataRetrievalViewModel(IIOService ioService, IXLIOService xlioService, IUIControlsService uiControlsService)
        {
            _ioService = ioService;
            _xlioService = xlioService;
            _uiControlsService = uiControlsService;

            AddExtractionRequestCommand = new RelayCommand(() => AddNewEmptyExtractionRequest());
            DeleteExtractionRequestCommand = new RelayCommand(() => DeleteExtractionRequest());
            BeginExtractionCommand = new AsyncCommand(async () => await BeginExtraction());
            SetOutputDirectoryCommand = new RelayCommand(() => SetOutputDirectory());

            ExtractionProgressEvent = new Progress<IEnumerable<ReturnMessage>>();
            ExtractionProgressEvent.ProgressChanged += ExtractionProgress_ProgressChanged;
        }

        #region Methods For Commands
        private void AddNewEmptyExtractionRequest()
        {
            DataRetrievalRequests.Add(new DataRetrievalRequest());
        }

        private void DeleteExtractionRequest()
        {
            try
            {
                DeleteExtractionRequestFromCollection(SelectedDataRetrievalRequest);
            }
            catch (CollectionEmptyException)
            {
                _uiControlsService.DisplayAlert("No data retrieval requests have been added", "Error!", MessageType.Error);
            }
            catch (ArgumentNullException)
            {
                _uiControlsService.DisplayAlert("No data retrieval request selected", "Error!", MessageType.Error);
            }
        }

        private async Task BeginExtraction()
        {

            try
            {
                if (string.IsNullOrWhiteSpace(OutputDirectory)) throw new ArgumentNullException(nameof(OutputDirectory),"cannot be null or empty");

                if (_extractionCancellationTokenSource != null)
                {
                    _extractionCancellationTokenSource.Cancel();
                    _extractionCancellationTokenSource = null;
                }
                _extractionCancellationTokenSource = new CancellationTokenSource();

                var successfulExtractions = await ExtractDataFromFiles(new DirectoryInfo(OutputDirectory), ChosenOutputFormat, ExtractionProgressEvent, _extractionCancellationTokenSource.Token);

                if (successfulExtractions.Any(x => x.Success == false))
                {
                    throw new ExtractionFailedException("Some extractions were unsuccessful...",
                                                        "Full Extraction Unsuccessful",
                                                        successfulExtractions.Where(x => x.Success == false).Select(y => y.Message));
                }

                ExtractionProgress = 0;
            }
            catch (ArgumentNullException e)
            {
                if (e.ParamName.ToLower() == "documentsdirectory" || e.ParamName == nameof(OutputDirectory)) _uiControlsService.DisplayAlert("No output directory set", "Alert!", MessageType.Error);
            }
            catch (NoDataOutputtedException e)
            {
                _uiControlsService.DisplayAlert(e.Message, e.ExceptionTitle, MessageType.Error);
            }
            catch (ExtractionFailedException e)
            {
                string alertBody = string.Join(Environment.NewLine, e.FailedExtractionMessages.Select(msg => string.Join(", ", msg)));
                _uiControlsService.DisplayAlert(alertBody, e.ExceptionTitle, MessageType.Error);
            }
        }

        private void ExtractionProgress_ProgressChanged(object sender, IEnumerable<ReturnMessage> e)
        {
            ExtractionProgress += 1;
        }

        public void SetOutputDirectory()
        {
            var chosenPath = SetOutputDirectoryFromIOService();

            if (!string.IsNullOrWhiteSpace(chosenPath)) OutputDirectory = chosenPath;
        }
        #endregion

        private void DeleteExtractionRequestFromCollection(DataRetrievalRequest dataRetrievalRequest)
        {
            if (DataRetrievalRequests.Count == 0) throw new CollectionEmptyException();
            if (dataRetrievalRequest == null) throw new ArgumentNullException("dataRetrievalRequest", "Cannot be null");

            DataRetrievalRequests.Remove(dataRetrievalRequest); //not concerned about successful removal, no need to interpret return bool.
        }

        private string SetOutputDirectoryFromIOService()
        {
            return _ioService.ChooseFolderDialog();
        }

        private async Task<IEnumerable<ReturnMessage>> ExtractDataFromFiles(DirectoryInfo documentsDirectory, DataOutputFormat dataOutputFormat, IProgress<IEnumerable<ReturnMessage>> progress, CancellationToken cancellationToken)
        {
            if (documentsDirectory == null) throw new ArgumentNullException("fileInfo", "Cannot be null");
            if (cancellationToken == null) throw new ArgumentNullException("cancellationToken", "Cannot be null");

            var validExtensions = new string[] { ".xlsx", ".xlsm" };

            var extractedDataCol = new BlockingCollection<IEnumerable<KeyValuePair<string, object>>>();
            var returnMessages = new BlockingCollection<ReturnMessage>();

            await Task.Run(() =>
            {
                var documents = documentsDirectory.GetFiles().Where(x => validExtensions.Contains(Path.GetExtension(x.FullName).ToLower()));
                TotalExtractionCount = documents.Count();

                Parallel.ForEach(documents, (document) =>
                {
                    try
                    {                      
                        if (!cancellationToken.IsCancellationRequested)
                        {
                            var extractionRequests = DataRetrievalRequestCollectionToExtractionRequestCollection(DataRetrievalRequests);
                            var dataExtractor = new DataExtractor(document.FullName);
                            foreach (var extraction in (dataExtractor.RetrieveDataCollectionFromAllWorksheets<object>(extractionRequests)))
                            {
                                extractedDataCol.Add(extraction);
                            }

                            returnMessages.Add(new ReturnMessage(true, $"{document.Name} successfully parsed and extracted from"));
                        }
                    }
                    catch(Exception e)
                    {
                        returnMessages.Add(new ReturnMessage(false, $"Failed to extract from {document.Name}.{Environment.NewLine}Error: {e.Message}"));
                    }

                    progress?.Report(returnMessages);
                });
            },cancellationToken);

            if (extractedDataCol != null)
            {
                SaveExtractedData(dataOutputFormat, extractedDataCol);
            }
            else
            {
                throw new NoDataOutputtedException($"No detected in files at {documentsDirectory.FullName}", "No Data");
            }

            return returnMessages;
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
                    succesfullySaved = _xlioService.SaveWorkbook(Path.Combine(OutputDirectory,$"{DateTime.Now.Ticks} Output.xlsx"), newWorkbook);
                    DisplaySuccessOrFailMessage(succesfullySaved);

                    break;
                case DataOutputFormat.CSV:
                    var generatedCSV = ExtractedDataConverter.ConvertToCSV(extractedDataCol);
                    succesfullySaved = _ioService.SaveText(generatedCSV, Path.Combine(OutputDirectory, $"{DateTime.Now.Ticks} Output.csv"));
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
