using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    public static class WebDriverExtension
    {
        public static IWebElement? TryGetElement(this EdgeDriver edgeDriver,By by)
        {
            var elements = edgeDriver.FindElements(by);
            if (elements.Count() == 0) return null;
            return elements.First();
        }

        public static IWebElement? TryGetElement(this ChromeDriver edgeDriver, By by)
        {
            var elements = edgeDriver.FindElements(by);
            if (elements.Count() == 0) return null;
            return elements.First();
        }
    }
}
