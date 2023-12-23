using Microsoft.Office.Interop.Word;
using MyCOMLib;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace COMTesting
{
    internal class Program
    {
        [Obsolete]
        public static void TestCom()
        {
            var type = typeof(MyClass);
            var assemblyName = type.Assembly.GetName();
            var assemblyFileName = Path.GetFileName(assemblyName.CodeBase);
            var comObject = Activator.CreateComInstanceFrom(assemblyFileName, type.FullName);
            var unwrapped = comObject.Unwrap() as MyClass;

            unwrapped.MyProperty = "hello";
            Console.WriteLine(unwrapped.MyProperty);
            var cookieValue = ROTHelpers.RegisterInRunningComObjectTable(unwrapped);

            Console.WriteLine("Added to ROT!");

            // alternative: Marshal.GetActiveObject()
            var object2 = ROTHelpers.GetRunningComObjectByDisplayName<MyClass>(assemblyName.Name);
            object2.MyProperty = "hello2";
            Console.WriteLine(object2.MyProperty);

            TestGetRunningComObjectNames();
            ROTHelpers.UnregisterRunningComObject(cookieValue);

            Console.WriteLine(Environment.NewLine);

            TestGetRunningComObjectNames();
        }

        public static void TestExe()
        {
            // 1. start process
            // 2. find out what the display name is!!! -> Read out rot!!!
            // 3. ->

            var wordDocPath = @"H:\testdata\my-word-doc.docx";
            var procArguments = $"/q"; ///t {wordDocPath}";
            var procFileName = "winword.exe";

            Document wordDoc = null;
            try
            {
                wordDoc = ROTHelpers.LaunchProcessAndWaitForComInstance<Document>(wordDocPath); // procArguments
                wordDoc.Application.Visible = false;
                wordDoc.Application.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                Console.WriteLine("File name: " + wordDoc.Name);
                object missing = System.Reflection.Missing.Value;
                //wordDoc.Content.Text += DateTime.Now.ToString() + "\tfoobar";
                //wordDoc.Save();

                wordDoc.Application.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                if (wordDoc != null)
                {
                    Marshal.ReleaseComObject(wordDoc);
                }
            }



            //rr app = new rr.Application();

            //// Optionally, set additional properties or configurations

            //// Open an existing project or create a new one
            //IRPProject project = rr.OpenProject("C:\\Path\\To\\Your\\Project.rpy");

            //RPApplication application = (RPApplication)Marshal.GetActiveObject("rr.Application");
            //RPProject project = application.activeProject();
            //RPCollection allElements = project.getNestedElementsRecursive();

            //foreach (RPModelElement element in allElements)
            //{
            //    //do something
            //}

            // https://www.ibm.com/support/pages/system/files/support/swg/swgdocs.nsf/0/cb4076a85185512385257de2005179e5/$FILE/rational_rr_api_getting_started.pdf


            // https://www.ibm.com/docs/en/engineering-lifecycle-management-suite/design-rr/9.0.2?topic=interface-command-line-syntax
            // https://www.ibm.com/docs/en/engineering-lifecycle-management-suite/design-rr/8.2.1?topic=api-using-rpapplicationlistener-respond-events
            //var app = GetRunningCOMObjectByName<RPApplication>("rr.Application");
            //if (app == null)
            //    continue;

            //// Do your stuff app (com object) in here..
            //Console.WriteLine("rr" + app.version() + " (" + app.BuildNo + ")");
            //// send a message to the rr instance
            //app.writeToOutputWindow("Log", "External script connected via API");
            //IRPProject proj = app.activeProject();
            //Console.WriteLine(proj.getLanguage() + " Project " + proj.name + "loaded.\n");
        }

        /// <summary>
        /// No need to use Interop here!!!!
        /// </summary>
        public static void TestWord()
        {
            var fileName = @"H:\testdata\my-word-doc.docx";
            Application word = null;
            word = new Application
            {
                Visible = false,
                DisplayAlerts = WdAlertLevel.wdAlertsNone
            };
            var doc = word.Documents.Open(fileName);
            doc.Content.Text += "barfoo";
            doc.Save();
            word.Quit();
        }

        public static void TestGetRunningComObjectNames()
        {
            var names = ROTHelpers.GetRunningComObjectNames();
            foreach (var name in names)
            {
                Console.WriteLine(name);
            }
        }


        [Obsolete]
        public static void Main(string[] args)
        {
            TestExe();
            Console.WriteLine("Done.");
            Console.ReadKey();
        }
    }
}
