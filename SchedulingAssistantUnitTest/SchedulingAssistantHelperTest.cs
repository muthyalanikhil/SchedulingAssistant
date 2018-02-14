using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SchedulingAssistant;

namespace SchedulingAssistantUnitTest
{
    [TestClass]
    public class SchedulingAssistantHelperTest
    {
        [TestMethod]
        public void IsTimeOverLapping_Test()
        {
            // arrange  
            string time1 = "1600-1650";
            string time2 = "1645-1745";
            bool expected = true;
            SchedulingAssistantHelper helper = new SchedulingAssistantHelper();

            // assert  
            bool actual = helper.IsTimeOverLapping(time1, time2);
            Assert.AreEqual(expected, actual, null, "The overlapping method fails");
        }

        [TestMethod]
        public void matchString_Test()
        {
            SchedulingAssistantHelper helper = new SchedulingAssistantHelper();
            // arrange  
            string regexExpression = @"^\d{4}-\d{4}$";
            string time1 = "1645-1745";
            bool expected1 = true;
            // assert  
            bool actual1 = helper.matchString(regexExpression, time1);
            Assert.AreEqual(expected1, actual1, null, "The match string method fails");

            string time2 = "16:45-17-45";
            bool expected2 = false;
            // assert  
            bool actual2 = helper.matchString(regexExpression, time2);
            Assert.AreEqual(expected2, actual2, null, "The match string method fails");

        }
    }
}
