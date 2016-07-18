using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using FrameworkChrome;

namespace chrometest
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            {
               //// Class1 test = new Class1();
                //test.Go();
                Class1 test2 = new Class1();

                test2.InvokeReporting();
               // test2.GetTestCaseToRun("Test-Oracle");
               // test2.GetTestCaseRun("Test-Oracle");
                //test2.CreateFolders();
                //test2.ResultReport("hello", "3");
                // Class1 test3 = new Class1();
                //test2.writeResult();


            }

        }
    }
}
