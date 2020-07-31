using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using AventStack.ExtentReports;
using System.Configuration;
using System.Reflection;
using AventStack.ExtentReports.Reporter;

namespace FeatureVision8.FeatureVision8
{
    [SetUpFixture]
    public class SetUpFixture : Base
    {
        [OneTimeSetUp]
        public void SetUp()
        {
            extent = ExtentManager.Instance;
            extent.AttachReporter(ExtentManager.htmlReporter);
        }

        [OneTimeTearDown]
        public void TearDown()
        {
            extent.Flush();
        }
    }
}
