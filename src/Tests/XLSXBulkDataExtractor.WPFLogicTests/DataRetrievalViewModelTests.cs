using Moq;
using NUnit.Framework;
using System;
using System.Collections.ObjectModel;
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

            Assert.That(sut.DataRetrievalRequests[0].Column, Is.EqualTo(0));
            Assert.That(sut.DataRetrievalRequests[0].Column, Is.EqualTo(0));
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

            uiControlsServiceMock.Verify(x => x.DisplayAlert("No data retrieval request selected", It.IsAny<string>(), MessageType.Error), Times.Once());
            Assert.That(sut.DataRetrievalRequests.Count, Is.EqualTo(3));
        }

        [Test]
        public void DeleteExtractionRequestNoRequestsInCollection()
        {
            DataRetrievalViewModel sut = new DataRetrievalViewModel(ioServiceMock.Object, xlIOServiceMock.Object, uiControlsServiceMock.Object);

            sut.DeleteExtractionRequestCommand.Execute(null);

            uiControlsServiceMock.Verify(x => x.DisplayAlert("No data retrieval requests have been added", It.IsAny<string>(), MessageType.Error), Times.Once());
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
        [Ignore("Not completed yet")]
        public void ExtractDataValid()
        {
        }
    }
}
