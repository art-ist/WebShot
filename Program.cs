using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace WebShot {
    class Program {
        static int Main(string[] args) {
            if (args.Length == 0 || args.Contains("-?") || args.Contains("/?")) {
                Help();
            }
            else {
                string url = args[0];
                string outPath = string.Format(@".\WebShot_{0:yyyyMMdd_hhmmss_nnn}.png", DateTime.Now);
                string browser = "IE";
                int windowWidth = 1024;
                int imageWidth = 64;
                int imageHeight = imageWidth;
                int wait = 0;
                int timeout = 20;

                try {
                    //process arguments
                    char[] removeQuotes = new char[] { '"' };
                    foreach (var arg in args) {
                        if (arg.ToLower().StartsWith("-url=")) url = arg.Split('=')[1].Trim(removeQuotes);
                        if (arg.ToLower().StartsWith("-out=")) outPath = arg.Split('=')[1].Trim(removeQuotes);
                        if (arg.ToLower().StartsWith("-wait=")) wait = int.Parse(arg.Split('=')[1]);
                        if (arg.ToLower().StartsWith("-timeout=")) timeout = int.Parse(arg.Split('=')[1]);
                        if (arg.ToLower().StartsWith("-windowwidth=")) windowWidth = int.Parse(arg.Split('=')[1]);
                        if (arg.ToLower().StartsWith("-imagewidth=")) imageWidth = int.Parse(arg.Split('=')[1]);
                        imageHeight = imageWidth;
                        if (arg.ToLower().StartsWith("-imageheight=")) imageHeight = int.Parse(arg.Split('=')[1]);
                        if (arg.ToLower().StartsWith("-browser=")) browser = arg.Split('=')[1];
                    }
                    //create image
                    CreateImage(url, outPath, windowWidth, imageWidth, imageHeight, wait, timeout, browser);
                }
                catch (Exception ex) {
                    Console.Error.WriteLine(ex.Message);
                    return 1;
                }
            }

#if DEBUG
            Console.ReadKey();
#endif
            return 0;
        }

        static void Help() {
            Console.WriteLine(@"WebShot (Alpha)

Usage: WebShot URL  [Options]

Options: 
    -ImageWidth=64                          ... image width in pixels
    -ImageHeight=64                         ... image height in pixels
    -WindowWidth=1024                       ... width of the browser window in pixels
    -Out=.\WebShot_20160831_153000_000.png  ... path of output file
    -Timeout=20                             ... wait max. 20 seconds for page to load
    -Wait=0                                 ... after loading page wait 0 seconds before taking screenshot
    -Browser=IE                             ... browser (needs to be installed)
");
        }

        static void CreateImage(string Url, string OutPath, int WindowWidth, int ImageWidth, int ImageHeight, int Wait, int TimeOut, string Browser) {
            float ratio = (float)ImageWidth / (float)ImageHeight;

            IWebDriver browser = null;
            try {
                browser = GetDriver(Browser, new Size(WindowWidth, (int)(WindowWidth / ratio)), TimeOut);
                //load url
                browser.Navigate().GoToUrl(Url);
                if (Wait > 0) {
                    Thread.Sleep(Wait * 1000);
                }
                //take screenshot
                Screenshot ss = ((ITakesScreenshot)browser).GetScreenshot();
                //proces image
                using (MemoryStream ms = new MemoryStream(ss.AsByteArray)) {
                    using (Image image = Image.FromStream(ms)) {
                        SaveImage(image, OutPath, new Size(ImageWidth, ImageHeight));
                    }
                }
            }
            catch (Exception) {
                throw;
            }
            finally {
                if (browser != null) {
                    browser.Quit();
                    browser.Dispose();
                }
            }
        }

        static IWebDriver GetDriver(string Name, Size WindowsSize, int TimeOut) {
            IWebDriver driver = null;
            switch (Name.ToLower()) {
                case "ff":
                case "firefox":
                    string ffPath = ConfigurationManager.AppSettings["FireFoxPath"];
                    if (!string.IsNullOrEmpty(ffPath)) {

                    }
                    driver = new FirefoxDriver(new FirefoxBinary(@"C:\Program Files\Mozilla Firefox\firefox.exe"), new FirefoxProfile());
                    break;
                case "crome":
                    driver = new ChromeDriver();
                    break;
                case "ie":
                case "internetexplorer":
                default:
                    driver = new InternetExplorerDriver();
                    break;
            }
            driver.Manage().Window.Size = WindowsSize;
            driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(TimeOut));
            return driver;
        }

        static void SaveImage(Image image, String FilePath) {
            image.Save(FilePath, GetImageFormat(FilePath));
        }

        static void SaveImage(Image image, String FilePath, Size size) {
            Image resized = new Bitmap(size.Width, size.Height);
            Graphics g = Graphics.FromImage(resized);
            g.CompositingQuality = CompositingQuality.HighQuality;
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.DrawImage(image, 0, 0, size.Width, size.Height);
            resized.Save(FilePath, GetImageFormat(FilePath));
        }

        static ImageFormat GetImageFormat(string FilePath) {
            switch (Path.GetExtension(FilePath).ToLower()) {
                case "png":
                    return ImageFormat.Png;
                case "gif":
                    return ImageFormat.Gif;
                default:
                    return ImageFormat.Jpeg;
            }
        }

    }
}
