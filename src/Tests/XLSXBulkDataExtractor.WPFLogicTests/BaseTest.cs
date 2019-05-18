﻿using Moq;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLSXBulkDataExtractor.WPFLogic.Interfaces;

namespace XLSXBulkDataExtractor.WPFLogicTests
{
    public abstract class BaseTest
    {
        public Mock<IIOService> ioServiceMock = new Mock<IIOService>();
        public Mock<IUIControlsService> uiControlsServiceMock = new Mock<IUIControlsService>();
        public Mock<IXLIOService> xlIOServiceMock = new Mock<IXLIOService>();

        [OneTimeSetUp]
        public virtual void InitialOneTimeSetup()
        {
        }
    }
}