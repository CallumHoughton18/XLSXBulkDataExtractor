using NUnit.Framework;
using System;
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
    }
}
