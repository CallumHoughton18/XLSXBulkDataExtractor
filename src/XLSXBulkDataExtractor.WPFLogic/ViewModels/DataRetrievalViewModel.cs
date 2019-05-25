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
using System.Data;

namespace XLSXBulkDataExtractor.WPFLogic.ViewModels
{
    /// <summary>
    /// Class DataRetrievalViewModel.
    /// Implements the <see cref="NotifyPropertyChangedBase" />
    /// </summary>
    /// <seealso cref="NotifyPropertyChangedBase" />
    public class DataRetrievalViewModel : NotifyPropertyChangedBase
    {
        private IIOService _ioService;
        private IXLIOService _xlioService;
        private IUIControlsService _uiControlsService;

        private CancellationTokenSource _extractionCancellationTokenSource = null;

        private DataTable _extractedDataTable;
        public DataTable ExtractedDataTable
        {
            get
            {
                return _extractedDataTable;
            }
            set
            {
                if (_extractedDataTable != value)
                {
                    _extractedDataTable = value;
                    OnPropertyChanged(nameof(ExtractedDataTable));
                }
            }
        }

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
        public string ExtractionDirectory
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
                    OnPropertyChanged(nameof(ExtractionDirectory));
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

        /// <summary>
        /// Gets the<seealso cref="DataOutputFormat"/> enum as a collection.
        /// </summary>
        /// <value>The output formats.</value>
        public IEnumerable<DataOutputFormat> OutputFormats
        {
            get
            {
                return Enum.GetValues(typeof(DataOutputFormat)).Cast<DataOutputFormat>();
            }
        }

        public DataOutputFormat ChosenOutputFormat { get; set; } = DataOutputFormat.XLSX;

        public ObservableCollection<DataRetrievalRequest> DataRetrievalRequests { get; set; } = new ObservableCollection<DataRetrievalRequest>();

        /// <summary>
        /// Gets the add extraction request command. Adds an empty <seealso cref="DataRetrievalRequest"/> object to the <see cref="DataRetrievalRequests"/> collection.
        /// </summary>
        /// <value>The add extraction request command.</value>
        public ICommand AddExtractionRequestCommand { get; private set; }
        /// <summary>
        /// Gets the delete extraction request command.Removes the <seealso cref="SelectedDataRetrievalRequest"/> object from the<see cref="DataRetrievalRequests"/> collection.
        /// </summary>
        /// <value>The delete extraction request command.</value>
        public ICommand DeleteExtractionRequestCommand { get; private set; }
        /// <summary>
        /// Gets the begin extraction command. Begins extraction of Excel files in <see cref="ExtractionDirectory"/> using the <see cref="DataRetrievalRequests"/>
        /// </summary>
        /// <value>The begin extraction command.</value>
        public IAsyncCommand BeginExtractionCommand { get; private set; }
        /// <summary>
        /// Gets the set output directory command. Sets the <see cref="ExtractionDirectory"/> for Excel files. This directory is also used as the output directory for the extracted data.
        /// </summary>
        /// <value>The set output directory command.</value>
        public ICommand SetDirectoryCommand { get; private set; }
        public Progress<IEnumerable<ReturnMessage>> ExtractionProgressEvent { get; } = new Progress<IEnumerable<ReturnMessage>>();

        public DataRetrievalViewModel(IIOService ioService, IXLIOService xlioService, IUIControlsService uiControlsService)
        {
            _ioService = ioService;
            _xlioService = xlioService;
            _uiControlsService = uiControlsService;

            AddExtractionRequestCommand = new RelayCommand(() => AddNewEmptyExtractionRequest());
            DeleteExtractionRequestCommand = new RelayCommand(() => DeleteExtractionRequest());
            BeginExtractionCommand = new AsyncCommand(async () => await BeginExtraction());
            SetDirectoryCommand = new RelayCommand(() => SetOutputDirectory());

            ExtractionProgressEvent = new Progress<IEnumerable<ReturnMessage>>();
            ExtractionProgressEvent.ProgressChanged += ExtractionProgress_ProgressChanged;
        }

        #region Methods For Commands
        protected void AddNewEmptyExtractionRequest()
        {
            DataRetrievalRequests.Add(new DataRetrievalRequest());
        }

        protected void DeleteExtractionRequest()
        {
            try
            {
                DeleteExtractionRequestFromCollection(SelectedDataRetrievalRequest);
            }
            catch (CollectionEmptyException)
            {
                _uiControlsService.DisplayAlert("No data retrieval requests have been added", MessageType.Error);
            }
            catch (ArgumentNullException)
            {
                _uiControlsService.DisplayAlert("No data retrieval request selected", MessageType.Error);
            }
        }

        protected async Task BeginExtraction()
        {

            try
            {
                if (string.IsNullOrWhiteSpace(ExtractionDirectory)) throw new ArgumentNullException(nameof(ExtractionDirectory),"cannot be null or empty");

                if (_extractionCancellationTokenSource != null)
                {
                    _extractionCancellationTokenSource.Cancel();
                    _extractionCancellationTokenSource = null;
                }
                _extractionCancellationTokenSource = new CancellationTokenSource();

                var successfulExtractions = await ExtractDataFromFiles(new DirectoryInfo(ExtractionDirectory), ChosenOutputFormat, ExtractionProgressEvent, _extractionCancellationTokenSource.Token);

                ExtractionProgress = 0;

                if (successfulExtractions.Any(x => x.Success == false))
                {
                    throw new ExtractionFailedException("Some extractions were unsuccessful...",
                                                        "Full Extraction Unsuccessful",
                                                        successfulExtractions.Where(x => x.Success == false).Select(y => y.Message));
                }
            }
            catch (ArgumentNullException e)
            {
                if (e.ParamName.ToLower() == "documentsdirectory" || e.ParamName == nameof(ExtractionDirectory)) _uiControlsService.DisplayAlert("No output directory set", MessageType.Error);
            }
            catch (NoDataOutputtedException e)
            {
                _uiControlsService.DisplayAlert(e.Message, MessageType.Error);
            }
            catch (ExtractionFailedException e)
            {
                foreach (var failedExtractionMessage in e.FailedExtractionMessages)
                {
                    _uiControlsService.DisplayAlert(failedExtractionMessage, MessageType.Error);
                }
            }
        }

        protected void ExtractionProgress_ProgressChanged(object sender, IEnumerable<ReturnMessage> e)
        {
            ExtractionProgress += 1;
        }

        protected void SetOutputDirectory()
        {
            var chosenPath = SetOutputDirectoryFromIOService();

            if (!string.IsNullOrWhiteSpace(chosenPath)) ExtractionDirectory = chosenPath;
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
                //filter files to valid extensions, and ignore temporary Excel files if excel file is open.
                var documents = documentsDirectory.GetFiles().Where(x => validExtensions.Contains(Path.GetExtension(x.FullName).ToLower()) && !x.Name.StartsWith(@"~$"));
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
            DataTable extractedDataTable = null;

            switch (dataOutputFormat)
            {
                case DataOutputFormat.XLSX:
                    extractedDataTable = ExtractedDataConverter.GenerateDataTable(extractedDataCol);
                    extractedDataTable.TableName = "ExtractedData";
                    var newWorkbook = new XLWorkbook();
                    newWorkbook.AddWorksheet(extractedDataTable);
                    succesfullySaved = _xlioService.SaveWorkbook(Path.Combine(ExtractionDirectory,$"{DateTime.Now.Ticks} Output.xlsx"), newWorkbook);
                    DisplaySuccessOrFailMessage(succesfullySaved);

                    break;
                case DataOutputFormat.CSV:
                    extractedDataTable = ExtractedDataConverter.GenerateDataTable(extractedDataCol);
                    var generatedCSV = ExtractedDataConverter.ConvertToCSV(extractedDataTable);
                    succesfullySaved = _ioService.SaveText(generatedCSV, Path.Combine(ExtractionDirectory, $"{DateTime.Now.Ticks} Output.csv"));
                    DisplaySuccessOrFailMessage(succesfullySaved);

                    break;
                default:
                    break;
            }

            if (extractedDataTable != null) ExtractedDataTable = extractedDataTable;
        }

        private void DisplaySuccessOrFailMessage(ReturnMessage returnMessage)
        {
            if (returnMessage == null) throw new ArgumentNullException("returnMessage", "cannot be null");

            if (returnMessage.Success) _uiControlsService.DisplayAlert(returnMessage.Message, MessageType.Success);

            else _uiControlsService.DisplayAlert(returnMessage.Message, MessageType.Error);

        }

        private IEnumerable<ExtractionRequest> DataRetrievalRequestCollectionToExtractionRequestCollection(IEnumerable<DataRetrievalRequest> dataRetrievalRequestsCol)
        {
            foreach (var dataRetrievalRequest in dataRetrievalRequestsCol)
            {
                yield return new ExtractionRequest(dataRetrievalRequest.FieldName, dataRetrievalRequest.Row, dataRetrievalRequest.ColumnNumber);
            }
        }
    }
}
