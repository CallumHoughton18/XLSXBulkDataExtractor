using Moq;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Threading.Tasks;
using XLSXBulkDataExtractor.WPFLogic.Globals;
using XLSXBulkDataExtractor.WPFLogic.Models;
using XLSXBulkDataExtractor.WPFLogic.ViewModels;

namespace XLSXBulkDataExtractor.WPFLogicTests
{
    public class DataRetrievalViewModelTests : BaseTest
    {
        public override void InitialOneTimeSetup()
        {
            base.InitialOneTimeSetup();
        }

        [Test]
        public void AddExtractionRequestTest()
        {
            DataRetrievalViewModel sut = new DataRetrievalViewModel(ioServiceMock.Object, xlIOServiceMock.Object, uiControlsServiceMock.Object);

            sut.AddExtractionRequestCommand.Execute(null);

            Assert.That(sut.DataRetrievalRequests.Count, Is.EqualTo(1));

            Assert.That(sut.DataRetrievalRequests[0].ColumnNumber, Is.EqualTo(1));
            Assert.That(sut.DataRetrievalRequests[0].ColumnNumber, Is.EqualTo(1));
            Assert.That(sut.DataRetrievalRequests[0].FieldName, Is.EqualTo(null));
        }

        [Test]
        public void DeleteExtractionRequestValidTest()
        {
            DataRetrievalViewModel sut = new DataRetrievalViewModel(ioServiceMock.Object, xlIOServiceMock.Object, uiControlsServiceMock.Object);
            sut.DataRetrievalRequests = new ObservableCollection<DataRetrievalRequest>(WPFLogicTestsHelper.GenerateMockRequests());
            sut.SelectedDataRetrievalRequest = sut.DataRetrievalRequests[1];

            sut.DeleteExtractionRequestCommand.Execute(null);

            Assert.That(sut.DataRetrievalRequests.Count, Is.EqualTo(2));
            Assert.That(sut.DataRetrievalRequests[0].FieldName, Is.EqualTo("MockFieldName0"));
            Assert.That(sut.DataRetrievalRequests[1].FieldName, Is.EqualTo("MockFieldName2"));
        }

        [Test]
        public void DeleteExtractionRequestNoRequestSelected()
        {
            DataRetrievalViewModel sut = new DataRetrievalViewModel(ioServiceMock.Object, xlIOServiceMock.Object, uiControlsServiceMock.Object);
            sut.DataRetrievalRequests = new ObservableCollection<DataRetrievalRequest>(WPFLogicTestsHelper.GenerateMockRequests());

            sut.DeleteExtractionRequestCommand.Execute(null);

            uiControlsServiceMock.Verify(x => x.DisplayAlert("No data retrieval request selected", MessageType.Error), Times.Once());
            Assert.That(sut.DataRetrievalRequests.Count, Is.EqualTo(3));
        }

        [Test]
        public void DeleteExtractionRequestNoRequestsInCollection()
        {
            DataRetrievalViewModel sut = new DataRetrievalViewModel(ioServiceMock.Object, xlIOServiceMock.Object, uiControlsServiceMock.Object);

            sut.DeleteExtractionRequestCommand.Execute(null);

            uiControlsServiceMock.Verify(x => x.DisplayAlert("No data retrieval requests have been added", MessageType.Error), Times.Once());
        }

        [Test]
        public void SetOutDirectoryValid()
        {
            DataRetrievalViewModel sut = new DataRetrievalViewModel(ioServiceMock.Object, xlIOServiceMock.Object, uiControlsServiceMock.Object);
            ioServiceMock.Setup(x => x.ChooseFolderDialog()).Returns("test/path");

            sut.SetOutputDirectoryCommand.Execute(null);

            Assert.That(sut.OutputDirectory, Is.EqualTo("test/path"));
        }

        [Test]
        public void SetOutDirectoryEmptyStringReturned()
        {
            DataRetrievalViewModel sut = new DataRetrievalViewModel(ioServiceMock.Object, xlIOServiceMock.Object, uiControlsServiceMock.Object);
            ioServiceMock.Setup(x => x.ChooseFolderDialog()).Returns("");

            sut.SetOutputDirectoryCommand.Execute(null);

            Assert.That(sut.OutputDirectory, Is.Null);
        }

        [Test]
        public async Task ExtractDataValid()
        {
            ClosedXML.Excel.XLWorkbook generatedWorkbook = null;
            DataRetrievalViewModel sut = new DataRetrievalViewModel(ioServiceMock.Object, xlIOServiceMock.Object, uiControlsServiceMock.Object);
            sut.OutputDirectory = Path.Combine(debugDirectory, "XLSXFiles");
            sut.DirectoryForXLSXExtractionFiles = Path.Combine(debugDirectory, "XLSXFiles");
            sut.DataRetrievalRequests = new ObservableCollection<DataRetrievalRequest>();
            sut.DataRetrievalRequests.Add(new DataRetrievalRequest { FieldName = "Test", ColumnName = 1.ToString(), Row = 1 });
            sut.DataRetrievalRequests.Add(new DataRetrievalRequest { FieldName = "Test2", ColumnName = 2.ToString(), Row = 2 });
            xlIOServiceMock.Setup(x => x.SaveWorkbook(It.IsAny<string>(), It.IsAny<ClosedXML.Excel.XLWorkbook>()))
                .Callback<string, ClosedXML.Excel.XLWorkbook>((path, workbook) => generatedWorkbook = workbook)
                .Returns(new Common.ReturnMessage(true, "workbook generated"));

            await sut.BeginExtractionCommand.ExecuteAsync();

            AssertCellValueForWorksheet(2, 1, "Postcode", generatedWorkbook);
            AssertCellValueForWorksheet(2, 2, "456", generatedWorkbook);


        }

        //TODO: Add in unit tests for task progress tracking

        [Test]
        public async Task ExtractDataNoOutputDirectory()
        {
            ClosedXML.Excel.XLWorkbook generatedWorkbook = null;
            DataRetrievalViewModel sut = new DataRetrievalViewModel(ioServiceMock.Object, xlIOServiceMock.Object, uiControlsServiceMock.Object);
            sut.DirectoryForXLSXExtractionFiles = Path.Combine(debugDirectory, "XLSXFiles");
            sut.DataRetrievalRequests = new ObservableCollection<DataRetrievalRequest>();
            sut.DataRetrievalRequests.Add(new DataRetrievalRequest { FieldName = "Test", ColumnName = 1.ToString(), Row = 1 });
            sut.DataRetrievalRequests.Add(new DataRetrievalRequest { FieldName = "Test2", ColumnName = 2.ToString(), Row = 2 });
            sut.DataRetrievalRequests.Add(new DataRetrievalRequest { FieldName = "Test3", ColumnName = 3.ToString(), Row = 3 });
            sut.DataRetrievalRequests.Add(new DataRetrievalRequest { FieldName = "Test4", ColumnName = 4.ToString(), Row = 4 });
            xlIOServiceMock.Setup(x => x.SaveWorkbook(It.IsAny<string>(), It.IsAny<ClosedXML.Excel.XLWorkbook>()))
                .Callback<string, ClosedXML.Excel.XLWorkbook>((path, workbook) => generatedWorkbook = workbook)
                .Returns(new Common.ReturnMessage(true, "workbook generated"));

            await sut.BeginExtractionCommand.ExecuteAsync();

            uiControlsServiceMock.Verify(x => x.DisplayAlert("No output directory set", MessageType.Error), Times.Once);
        }

        [Test]
        public async Task ExtractDataTaskProgressTest()
        {
            double timesExtractionProgressEventFired = 0;

            DataRetrievalViewModel sut = new DataRetrievalViewModel(ioServiceMock.Object, xlIOServiceMock.Object, uiControlsServiceMock.Object);

            sut.ExtractionProgressEvent.ProgressChanged += (sender, e) =>
            {
                timesExtractionProgressEventFired += 1;
            };

            sut.OutputDirectory = Path.Combine(debugDirectory, "XLSXFiles");
            sut.DirectoryForXLSXExtractionFiles = Path.Combine(debugDirectory, "XLSXFiles");
            sut.DataRetrievalRequests = new ObservableCollection<DataRetrievalRequest>();
            sut.DataRetrievalRequests.Add(new DataRetrievalRequest { FieldName = "Test", ColumnName = 1.ToString(), Row = 1 });
            sut.DataRetrievalRequests.Add(new DataRetrievalRequest { FieldName = "Test2", ColumnName = 2.ToString(), Row = 2 });
            sut.DataRetrievalRequests.Add(new DataRetrievalRequest { FieldName = "Test3", ColumnName = 3.ToString(), Row = 3 });
            sut.DataRetrievalRequests.Add(new DataRetrievalRequest { FieldName = "Test4", ColumnName = 4.ToString(), Row = 4 });
            xlIOServiceMock.Setup(x => x.SaveWorkbook(It.IsAny<string>(), It.IsAny<ClosedXML.Excel.XLWorkbook>()))
                .Returns(new Common.ReturnMessage(true, "workbook generated"));

            await sut.BeginExtractionCommand.ExecuteAsync();

            Assert.That(timesExtractionProgressEventFired, Is.EqualTo(5D));
            Assert.That(sut.TotalExtractionCount, Is.EqualTo(5D));
            Assert.That(sut.ExtractionProgress, Is.EqualTo(0D)); //check that extraction progress reset when complete.
        }

        [Test]
        public async Task ExtractDataTaskDataTableTest()
        {

            DataRetrievalViewModel sut = new DataRetrievalViewModel(ioServiceMock.Object, xlIOServiceMock.Object, uiControlsServiceMock.Object);

            sut.OutputDirectory = Path.Combine(debugDirectory, "XLSXFiles");
            sut.DirectoryForXLSXExtractionFiles = Path.Combine(debugDirectory, "XLSXFiles");
            sut.DataRetrievalRequests = new ObservableCollection<DataRetrievalRequest>();
            sut.DataRetrievalRequests.Add(new DataRetrievalRequest { FieldName = "Test", ColumnName = 1.ToString(), Row = 1 });
            sut.DataRetrievalRequests.Add(new DataRetrievalRequest { FieldName = "Test2", ColumnName = 2.ToString(), Row = 2 });
            sut.DataRetrievalRequests.Add(new DataRetrievalRequest { FieldName = "Test3", ColumnName = 3.ToString(), Row = 3 });
            sut.DataRetrievalRequests.Add(new DataRetrievalRequest { FieldName = "Test4", ColumnName = 4.ToString(), Row = 4 });
            xlIOServiceMock.Setup(x => x.SaveWorkbook(It.IsAny<string>(), It.IsAny<ClosedXML.Excel.XLWorkbook>()))
                .Returns(new Common.ReturnMessage(true, "workbook generated"));

            Assert.That(sut.ExtractedDataTable, Is.Null);

            await sut.BeginExtractionCommand.ExecuteAsync();

            Assert.That(sut.ExtractedDataTable, Is.Not.Null);
        }
        public void AssertCellValueForWorksheet(int row, int column,string expectedValue, ClosedXML.Excel.XLWorkbook generatedWorkbook)
        {
            Assert.That(generatedWorkbook.Worksheet(1).Cell(row, column).Value.ToString(), Is.EqualTo(expectedValue));
        }
        
        //TODO: Add DataTable unit test
    }
}
