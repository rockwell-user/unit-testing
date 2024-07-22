// ---------------------------------------------------------------------------------------------------------------------------------------------------------------
//
// FileName: Logix_Program.cs
// FileType: Visual C# Source file
// Author : Rockwell Automation
// Created : 2024
// Description : This script provides an example test in a CI/CD pipeline utilizing Studio 5000 Logix Designer SDK and Factory Talk Logix Echo SDK.
//
// ---------------------------------------------------------------------------------------------------------------------------------------------------------------

using Google.Protobuf;
using LogixEcho_ClassLibrary;
using OfficeOpenXml;
using RockwellAutomation.LogixDesigner;
using System.Collections;
using System.Text;
using System.Xml.Linq;
using static RockwellAutomation.LogixDesigner.LogixProject;
using DataType = RockwellAutomation.LogixDesigner.LogixProject.DataType;

namespace UnitTesting
{
    internal class UnitTest
    {
        struct AOIParameters
        {
            public string? Name { get; set; }
            public string? DataType { get; set; }
            public string? Usage { get; set; }
            public bool? Required { get; set; }
            public bool? Visible { get; set; }
            public string? Value { get; set; }
            public int BytePosition { get; set; }
            public int BoolPosition { get; set; }
        }

        static async Task Main()
        {
            // XML FILE MANIPULATIONS
            string rungXML = CopyXmlFile(@"C:\Users\ASYost\Desktop\UnitTesting\AOI_L5Xs\WetBulbTemperature_AOI.L5X");
            Console.WriteLine("rungXML filepath: " + rungXML);

            string aoiName = Get_AttributeValue(rungXML, "AddOnInstructionDefinition", "Name");

            //Modify top half
            DeleteAttributeFromRoot(rungXML, "TargetName");
            DeleteAttributeFromRoot(rungXML, "TargetRevision");
            DeleteAttributeFromRoot(rungXML, "TargetLastEdited");
            DeleteAttributeFromComplexElement(rungXML, "AddOnInstructionDefinition", "Use");
            ChangeComplexElementAttribute(rungXML, "RSLogix5000Content", "TargetCount", "1");
            ChangeComplexElementAttribute(rungXML, "RSLogix5000Content", "TargetType", "Rung");

            // Create bottom half
            AddElementToComplexElement(rungXML, "Controller", "Tags");
            AddAttributeToComplexElement(rungXML, "Tags", "Use", "Context");

            AddElementToComplexElement(rungXML, "Tags", "Tag");
            AddAttributeToComplexElement(rungXML, "Tag", "Name", "AOI_" + aoiName);
            AddAttributeToComplexElement(rungXML, "Tag", "TagType", "Base");
            AddAttributeToComplexElement(rungXML, "Tag", "DataType", aoiName);
            AddAttributeToComplexElement(rungXML, "Tag", "Constant", "false");
            AddAttributeToComplexElement(rungXML, "Tag", "ExternalAccess", "Read/Write");
            AddAttributeToComplexElement(rungXML, "Tag", "OpcUaAccess", "None");

            AddElementToComplexElement(rungXML, "Tag", "Data");
            AddAttributeToComplexElement(rungXML, "Data", "Format", "L5K");

            string cdataInfo_forData = Get_CDATAfromXML_forData(rungXML);
            CreateCData(rungXML, "Data", cdataInfo_forData);

            AddElementToComplexElement(rungXML, "Tag", "Data");
            AddAttributeToComplexElement(rungXML, "Data", "Format", "Decorated");

            AddElementToComplexElement(rungXML, "Data", "Structure");
            AddAttributeToComplexElement(rungXML, "Structure", "DataType", aoiName);

            //AddAttributeToComplexElement(rungXML, "DataValueMember", "Name", aoiName);
            List<Dictionary<string, string>> attributesList = Get_DataValueMemberInfofromXML(rungXML);

            //foreach (var attributes in attributesList)
            //{
            //    Console.WriteLine($"Name: {attributes["Name"]}, DataType: {attributes["DataType"]}, Radix: {attributes["Radix"]}");
            //}

            AddComplexElementsToXml(rungXML, attributesList);

            AddElementToComplexElement(rungXML, "Controller", "Programs");
            AddAttributeToComplexElement(rungXML, "Programs", "Use", "Context");

            AddElementToComplexElement(rungXML, "Programs", "Program");
            AddAttributeToComplexElement(rungXML, "Program", "Use", "Context");
            AddAttributeToComplexElement(rungXML, "Program", "Name", "P00_AOI_Testing");

            AddElementToComplexElement(rungXML, "Program", "Routines");
            AddAttributeToComplexElement(rungXML, "Routines", "Use", "Context");

            AddElementToComplexElement(rungXML, "Routines", "Routine");
            AddAttributeToComplexElement(rungXML, "Routine", "Use", "Context");
            AddAttributeToComplexElement(rungXML, "Routine", "Name", "R00_AOI_Testing");

            AddElementToComplexElement(rungXML, "Routine", "RLLContent");
            AddAttributeToComplexElement(rungXML, "RLLContent", "Use", "Context");

            AddElementToComplexElement(rungXML, "RLLContent", "Rung");
            AddAttributeToComplexElement(rungXML, "Rung", "Use", "Target");
            AddAttributeToComplexElement(rungXML, "Rung", "Number", "0");
            AddAttributeToComplexElement(rungXML, "Rung", "Type", "N");

            AddElementToComplexElement(rungXML, "Rung", "Text");

            string cdataInfo_forText = Get_CDATAfromXML_forText(rungXML);
            CreateCData(rungXML, "Text", cdataInfo_forText);


            // THE ONLY PARAMETER THAT WILL NEED TO BE MODIFIED (MAYBE PASS THIS IN FROM JENKINS)
            // Parameter to be specified in jenkins
            string unitTestExcelWorkbooks_folderPath = @"C:\Users\ASYost\Desktop\UnitTesting\AOIs_toTest";
            string exampleTestReportsFolder_filePath = @"C:\Users\ASYost\Desktop\UnitTesting\exampleTestReports";
            AOIParameters[] testParams = Get_AOIParameters_FromL5X(@"C:\Users\ASYost\Desktop\UnitTesting\AOI_L5Xs\WetBulbTemperature_AOI.L5X");
            Print_AOIParameters(testParams);


            // Title Banner 
            Console.WriteLine("\n  ====================================================================================================================");
            Console.WriteLine("========================================================================================================================");
            Console.WriteLine("                      UNIT TESTING | " + DateTime.Now + " " + TimeZoneInfo.Local);
            Console.WriteLine("========================================================================================================================");
            Console.WriteLine("  ====================================================================================================================");

            // From a string array to a list, store the name (including their path) for each excel workbook.
            // With the current implementation, each Excel Workbook tests a single Add-On Instruction.
            string[] excelFiles = Directory.GetFiles(unitTestExcelWorkbooks_folderPath);
            List<FileInfo> orderedExcelFiles = [.. excelFiles.Select(f => new FileInfo(f)).OrderBy(f => f.CreationTime)];

            // Parameters from the excel sheet that determine the test to be run. 
            string controllerName = "";
            string acdFilePath = "";
            string aoiTagName = "";
            string aoiTagScope = "";

            // Increment through each Excel Workbook in the specified folder.
            for (int i = 0; i < (orderedExcelFiles.Count); i++)
            {
                var currentExcelUnitTest_filePath = orderedExcelFiles[i].FullName;
                FileInfo existingFile = new FileInfo(currentExcelUnitTest_filePath);
                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    //get the first worksheet in the workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    controllerName = worksheet.Cells[11, 2].Value?.ToString()!.Trim()!;
                    acdFilePath = worksheet.Cells[11, 3].Value?.ToString()!.Trim()!;
                    aoiTagName = worksheet.Cells[11, 11].Value?.ToString()!.Trim()!;
                    aoiTagScope = worksheet.Cells[11, 15].Value?.ToString()!.Trim()!;
                }


                //string githubPath = @"C:\examplefolder";                                           // 1st incoming argument = GitHub folder path
                //string acdFilename = "CICD_test.ACD";                                          // 2nd incoming argument = Logix Designer ACD filename
                string name_mostRecentCommit = "example name";                                // 3rd incoming argument = name of person assocatied with most recent git commit
                string email_mostRecentCommit = "example email";                               // 4th incoming argument = email of person associated with most recent git commit
                string message_mostRecentCommit = "example commmit message";                             // 5th incoming argument = message provided in the most recent git commit
                string hash_mostRecentCommit = "example commit hash";                                // 6th incoming argument = hash ID from most recent git commit
                string jenkinsJobName = "example jenkins job";                                       // 7th incoming argument = the Jenkins job name
                string jenkinsBuildNumber = "example jenkins build number";                                   // 8th incoming argument = the Jenkins job build number
                                                                                                              //string acdFilePath = @"C:\Users\ASYost\Desktop\UnitTesting\ACD_testFiles\CICD_test.ACD"; // file path to ACD project
                                                                                                              // string textFileReportDirectory = githubPath + @"test-reports\textFiles\";   // folder path to text test reports
                                                                                                              // string excelFileReportDirectory = githubPath + @"test-reports\excelFiles\"; // folder path to excel test reports
                                                                                                              // string textFileReportPath = Path.Combine(textFileReportDirectory, DateTime.Now.ToString("yyyyMMddHHmmss") + "_testfile.txt");    // new text test report filename
                string excelFileReportPath = Path.Combine(exampleTestReportsFolder_filePath, DateTime.Now.ToString("yyyyMMddHHmmss") + "_testfile.xlsx"); // new excel test report filename

                Console.WriteLine("\n\n");

                // Executed only once on the first AOI tested.
                if (i == 0)
                {
                    // Create an excel test report to be filled out during testing.
                    Console.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss")}] START setting up excel test report workbook...");
                    CreateFormattedExcelFile(excelFileReportPath, acdFilePath, name_mostRecentCommit, email_mostRecentCommit,
                        jenkinsBuildNumber, jenkinsJobName, hash_mostRecentCommit, message_mostRecentCommit);
                    Console.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss")}] DONE setting up excel test report workbook...\n---");

                    // Check the test-reports folder and if over the specified file number limit, delete the oldest test files.
                    Console.WriteLine($"[{DateTime.Now.ToString("T")}] START checking test-reports folder...");
                    CleanTestReportsFolder(exampleTestReportsFolder_filePath, 5);
                    Console.WriteLine($"[{DateTime.Now.ToString("T")}] DONE checking test-reports folder...\n---");
                }

                // Set up emulated controller (based on the specified ACD file path) if one does not yet exist. If not, continue.
                Console.WriteLine($"[{DateTime.Now.ToString("T")}] START setting up Factory Talk Logix Echo emulated controller...");
                string commPath = SetUpEmulatedController_Sync(acdFilePath, "UnitTest_Chassis", controllerName);
                Console.WriteLine($"[{DateTime.Now.ToString("T")}] DONE setting up Factory Talk Logix Echo emulated controller\n---");

                // Create a new ACD project file.
                Console.WriteLine($"[{DateTime.Now.ToString("T")}] START creating & opening ACD file...");
                string acdPath = Path.Combine(@"C:\Users\ASYost\Desktop\UnitTesting\ACD_testFiles_generated\", DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + aoiTagName + "_UnitTest.ACD");
                //string acdPath = @"C:\Users\ASYost\Desktop\UnitTesting\ACD_testFiles_generated\20240716141455_AOIunittest.ACD";
                uint majorRevision = 36;
                string processorTypeName = "1756-L85E";
                string controllerName2 = "UnitTest_Controller";
                LogixProject project = await CreateNewProjectAsync(acdPath, majorRevision, processorTypeName, controllerName2);
                Console.WriteLine($"SUCCESS: file created at {acdPath}");
                Console.WriteLine($"[{DateTime.Now.ToString("T")}] DONE creating & opening ACD file\n---");

                //Console.WriteLine($"[{DateTime.Now.ToString("T")}] START uploading ACD file...");
                //string acdPath = Path.Combine(@"C:\Users\ASYost\Desktop\UnitTesting\ACD_testFiles_generated\", DateTime.Now.ToString("yyyyMMddHHmmss") + "_AOI_UnitTest.ACD");
                //await LogixProject.UploadToNewProjectAsync(acdPath, commPath);
                //Console.WriteLine($"[{DateTime.Now.ToString("T")}] SUCCESS uploading ACD file\n---");

                //string export_filePath = @"C:\Users\ASYost\Desktop\UnitTesting\ACD_testFiles_generated\AOI_UnitTestTemplate.ACD";
                //Console.WriteLine($"[{DateTime.Now.ToString("T")}] START exporting program file...");
                //string xPath2 = @"Controller/Programs";
                ////string export_filePath = @"C:\Users\ASYost\Desktop\UnitTesting\AOI_L5Xs\NEW_P00_AOI_Testing_Program.L5X";
                //await project.PartialExportToXmlFileAsync(xPath2, export_filePath);
                //Console.WriteLine($"[{DateTime.Now.ToString("T")}] DONE exporting program file\n---");



                //string acdPath = @"C:\Users\ASYost\Desktop\UnitTesting\ACD_testFiles\UnitTest_AOI.ACD"; // USED DURING DEV



                //// =========================================================================
                //// LOGGER INFO
                //var logger = new StdOutEventLogger();
                //project.AddEventHandler(logger);
                //// =========================================================================

                //Console.WriteLine("\n\n\n\n");

                Console.WriteLine($"[{DateTime.Now.ToString("T")}] START preparing programmatically created ACD...");
                //Console.WriteLine("\n\n\n\n");

                //string filePath = @"C:\Users\ASYost\Desktop\UnitTesting\AOI_L5Xs\NEW_P00_AOI_Testing_Program.L5X";
                string xPath = @"Controller/Programs";  // THIS WORKS BUT GOES TO UNSCHEDULED FOLDER IN ACD
                                                        //string filePath = @"C:\Users\ASYost\Desktop\UnitTesting\AOI_L5Xs\P00_AOI_Testing_Program.L5X";
                string filePath1 = @"C:\Users\ASYost\Desktop\UnitTesting\AOI_L5Xs\PROGRAMTARGET_P00_AOI_Testing_Program.L5X";
                await project.PartialImportFromXmlFileAsync(xPath, filePath1, LogixProject.ImportCollisionOptions.OverwriteOnColl);
                //Console.WriteLine("\n\n\n\n");

                string filePath2 = @"C:\Users\ASYost\Desktop\UnitTesting\AOI_L5Xs\MyTaskExp.L5X";
                string xPath2 = @"Controller/Tasks";  // include program in task
                await project.PartialImportFromXmlFileAsync(xPath2, filePath2, LogixProject.ImportCollisionOptions.OverwriteOnColl);
                //Console.WriteLine("\n\n\n\n");

                string filePath3 = @"C:\Users\ASYost\Desktop\UnitTesting\AOI_L5Xs\WetBulbTemperature_AOI.L5X";
                string xPath3 = @"Controller/AddOnInstructionDefinitions";
                await project.PartialImportFromXmlFileAsync(xPath3, filePath3, LogixProject.ImportCollisionOptions.OverwriteOnColl);
                //Console.WriteLine("\n\n\n\n");

                string filePath4 = @"C:\Users\ASYost\Desktop\UnitTesting\AOI_L5Xs\Rung0_from_R00_AOI_Testing.L5X";
                //string xPath4 = @"Controller/Programs/Program[@Name='P00_AOI_Testing']/Routines/Routine[@Name='R00_AOI_Testing']/RLLContent/Rung[@Number>='0]";
                //string xPath4 = @"Controller/Programs/Program[@Name='P00_AOI_Testing']/Routines/Routine[@Name='R00_AOI_Testing']/RLLContent/Rung[@Number>='1]";
                //string xPath4 = @"Controller/Programs/Program[@Name='P00_AOI_Testing']/Routines/Routine[@Name='R00_AOI_Testing']";
                //string xPath4 = @"Controller/Programs/Program[@Name='P00_AOI_Testing']/Routines/Routine[@Name='R00_AOI_Testing']/RLLContent/Rung[@Number='0']";
                string xPath4 = @"Controller/Programs/Program[@Name='P00_AOI_Testing']/Routines";
                await project.PartialImportFromXmlFileAsync(xPath4, filePath4, LogixProject.ImportCollisionOptions.OverwriteOnColl);
                await project.SaveAsync();
                //Console.WriteLine("\n\n\n\n");

                Console.WriteLine($"[{DateTime.Now.ToString("T")}] DONE preparing programmatically created ACD\n---");
                //Console.WriteLine("\n\n\n\n");







                //XmlDocument doc1 = new XmlDocument();
                //doc1.Load(filePath);
                //XmlNode node1 = doc1.DocumentElement.SelectSingleNode("/")

                //XmlDocument doc2 = new XmlDocument();
                //doc2.Load(filePath2);



                //Console.WriteLine($"[{DateTime.Now.ToString("T")}] START importing AOI.L5X...");
                ////string xPath = @"Controller/Programs/Program[@Name='P00_AOI_Testing']";
                ////string filePath = @"C:\Users\ASYost\Desktop\UnitTesting\AOI_L5Xs\WetBulbTemperature_AOI.L5X";
                //string xPath = @"Controller"; // didn't work with P00_AOI_Testing_Program.L5X
                // string xPath = @"Controller/Programs";  // THIS WORKS BUT GOES TO UNSCHEDULED FOLDER IN ACD
                //string xPath = @"Controller/Programs"; 
                // string xPath = @"/RSLogix5000Content/Controller/Tasks/Task[@Name='T00_AOI_Testing']";
                //string xPath = @"/RSLogix5000Content/Controller/Tasks";
                ////string xPath = @"Controller/Tasks";
                //string xPath = @"RSLogix5000Content/Controller/Tasks"; // didn't work with P00_AOI_Testing_Program.L5X
                // string xPath = @"RSLogix5000Content/Controller"; // didn't work with P00_AOI_Testing_Program.L5X
                //string filePath = @"C:\Users\ASYost\Desktop\UnitTesting\AOI_L5Xs\P00_AOI_Testing_Program.L5X";
                ////string filePath = @"C:\Users\ASYost\Desktop\UnitTesting\AOI_L5Xs\20240716140136_AOIunittest.L5X";
                //string filePath = @"C:\Users\ASYost\Desktop\UnitTesting\AOI_L5Xs\WetBulbTemperature_AOI.L5X";
                //string targetName = @"/RSLogix5000Content/Controller/Programs/Program"; // PartialImportWithTargetFromXmlFile does not allow to import XML xPath: /RSLogix5000Content/Controller/Programs
                //string targetName = @"/RSLogix5000Content/Controller/Programs"; //PartialImportWithTargetFromXmlFile does not allow to import XML xPath: /RSLogix5000Content/Controller/Programs
                //string xPath = @"Controller/Tasks/Task[@Name='T00_AOI_Testing']"; // Invalid import target for the XML target type. (C:\Users\ASYost\Desktop\UnitTesting\ACD_testFiles\UnitTest_AOI.ACD)
                //string targetName = @"/RSLogix5000Content"; //PartialImportWithTargetFromXmlFile does not allow to import XML xPath: /RSLogix5000Content/Controller/Programs
                //await project.PartialImportWithTargetFromXmlFileAsync(xPath, targetName, filePath, LogixProject.PartialImportOption.FinalizeEdits);
                //Console.WriteLine($"[{DateTime.Now.ToString("T")}] SUCCESS importing AOI.L5X\n---");

                // Change controller mode to program & verify.
                Console.WriteLine($"[{DateTime.Now.ToString("T")}] START changing controller to PROGRAM...");
                ChangeControllerMode_Async(commPath, "Program", project).GetAwaiter().GetResult();
                if (ReadControllerMode_Async(commPath, project).GetAwaiter().GetResult() == "PROGRAM")
                    Console.WriteLine($"[{DateTime.Now.ToString("T")}] SUCCESS changing controller to PROGRAM\n---");
                else
                    Console.WriteLine($"[{DateTime.Now.ToString("T")}] FAILURE changing controller to PROGRAM\n---");

                // Download project.
                Console.WriteLine($"[{DateTime.Now.ToString("T")}] START downloading ACD file...");
                DownloadProject_Async(commPath, project).GetAwaiter().GetResult();
                Console.WriteLine($"[{DateTime.Now.ToString("T")}] SUCCESS downloading ACD file\n---");

                // Change controller mode to run.
                Console.WriteLine($"[{DateTime.Now.ToString("T")}] START Changing controller to RUN...");
                ChangeControllerMode_Async(commPath, "Run", project).GetAwaiter().GetResult();
                if (ReadControllerMode_Async(commPath, project).GetAwaiter().GetResult() == "RUN")
                    Console.WriteLine($"[{DateTime.Now.ToString("T")}] SUCCESS changing controller to RUN\n---");
                else
                    Console.WriteLine($"[{DateTime.Now.ToString("T")}] FAILURE changing controller to RUN\n---");



                //// Open the ACD project file and store the reference as myProject.
                //Console.WriteLine($"[{DateTime.Now.ToString("T")}] START opening ACD file...");
                //LogixProject myProject = await LogixProject.OpenLogixProjectAsync(acdFilePath);
                //Console.WriteLine($"[{DateTime.Now.ToString("T")}] SUCCESS opening ACD file\n---");

                //// Change controller mode to program & verify.
                //Console.WriteLine($"[{DateTime.Now.ToString("T")}] START changing controller to PROGRAM...");
                //ChangeControllerMode_Async(commPath, "Program", myProject).GetAwaiter().GetResult();
                //if (ReadControllerMode_Async(commPath, myProject).GetAwaiter().GetResult() == "PROGRAM")
                //    Console.WriteLine($"[{DateTime.Now.ToString("T")}] SUCCESS changing controller to PROGRAM\n---");
                //else
                //    Console.WriteLine($"[{DateTime.Now.ToString("T")}] FAILURE changing controller to PROGRAM\n---");

                //// Download project.
                //Console.WriteLine($"[{DateTime.Now.ToString("T")}] START downloading ACD file...");
                //DownloadProject_Async(commPath, myProject).GetAwaiter().GetResult();
                //Console.WriteLine($"[{DateTime.Now.ToString("T")}] SUCCESS downloading ACD file\n---");

                //// Change controller mode to run.
                //Console.WriteLine($"[{DateTime.Now.ToString("T")}] START Changing controller to RUN...");
                //ChangeControllerMode_Async(commPath, "Run", myProject).GetAwaiter().GetResult();
                //if (ReadControllerMode_Async(commPath, myProject).GetAwaiter().GetResult() == "RUN")
                //    Console.WriteLine($"[{DateTime.Now.ToString("T")}] SUCCESS changing controller to RUN\n---");
                //else
                //    Console.WriteLine($"[{DateTime.Now.ToString("T")}] FAILURE changing controller to RUN\n---");


                // ---------------------------------------------------------------------------------------------------------------------------

                //TagData[] testDataPoint = GetAOIParameters(currentExcelUnitTest_filePath);

                //Console.WriteLine("current fullTagPath: " + aoiTagScope);
                //ByteString udtoraoi_byteString = Get_UDTorAOI_ByteString_Sync(aoiTagScope, project, OperationMode.Online);

                //TagData[] tagdata_UDTorAOI = Get_UDTorAOI(testDataPoint, udtoraoi_byteString, true);
                //ShowDataPoints(tagdata_UDTorAOI);

                //int testcases = GetPopulatedColumnCount(currentExcelUnitTest_filePath, 20) - 3;
                //Console.WriteLine("testcases: " + testcases);
                ////await SetSingleValue_UDTorAOI("40", aoiTagScope, "Temperature", OperationMode.Online, tagdata_UDTorAOI, myProject);
                ////await SetSingleValue_UDTorAOI("0", aoiTagScope, "isFahrenheit", OperationMode.Online, tagdata_UDTorAOI, myProject);

                //ShowDataPoints(Get_UDTorAOI(testDataPoint, Get_UDTorAOI_ByteString_Sync(aoiTagScope, project, OperationMode.Online), true));


                //await myProject.GoOfflineAsync();
            }


        }






























        #region METHODS: manipulate L5X
        public static string CopyXmlFile(string sourceFilePath)
        {
            // Check if the source file exists
            if (!File.Exists(sourceFilePath))
            {
                Console.WriteLine($"Source file '{sourceFilePath}' does not exist.");
            }

            // Get the directory and file name from the source file path
            string? directory = Path.GetDirectoryName(sourceFilePath);
            string fileName = Path.GetFileNameWithoutExtension(sourceFilePath);
            string extension = Path.GetExtension(sourceFilePath);

            // Construct the new file path for the copied file
            string newFileName = $"COPY_{fileName}{extension}";
            string newFilePath = Path.Combine(directory, newFileName);

            // Copy the file
            File.Copy(sourceFilePath, newFilePath, overwrite: true);

            return newFilePath;
        }

        public static string? Get_AttributeValue(string xmlFilePath, string complexElementName, string attributeName)
        {
            // Load the XML document
            XDocument xdoc = XDocument.Load(xmlFilePath);

            // Find the complex element by name
            XElement? complexElement = xdoc.Descendants(complexElementName).FirstOrDefault();

            if (complexElement != null)
            {
                // Find the attribute within the complex element
                XAttribute? attribute = complexElement.Attribute(attributeName);
                if (attribute != null)
                {
                    // Return the attribute value
                    return attribute.Value;
                }
                else
                {
                    Console.WriteLine($"Attribute '{attributeName}' not found in element '{complexElementName}'.");
                }
            }
            else
            {
                Console.WriteLine($"The complex element '{complexElementName}' was not found in the XML file.");
            }

            return null; // Return null if attribute value is not found
        }


        public Dictionary<string, string> CopyAttributes(string xmlFilePath, string elementName)
        {
            // Load the XML document
            XDocument xdoc = XDocument.Load(xmlFilePath);

            // Find the specified element. Assuming there is only one unique element with this name
            XElement element = xdoc.Descendants(elementName).FirstOrDefault();

            if (element == null)
            {
                throw new InvalidOperationException($"The element '{elementName}' was not found in the XML file.");
            }

            // Create a dictionary to hold the attribute names and values
            Dictionary<string, string> attributesDictionary = element.Attributes().ToDictionary(attr => attr.Name.LocalName, attr => attr.Value);

            return attributesDictionary;
        }

        public static void DeleteAttributeFromComplexElement(string xmlFilePath, string complexElementName, string attributeToDelete)
        {
            try
            {
                // Load the XML document
                XDocument xdoc = XDocument.Load(xmlFilePath);

                // Find the complex element by name
                XElement complexElement = xdoc.Descendants(complexElementName).FirstOrDefault();

                if (complexElement != null)
                {
                    // Find the attribute within the complex element
                    XAttribute attribute = complexElement.Attribute(attributeToDelete);
                    if (attribute != null)
                    {
                        // Remove the attribute
                        attribute.Remove();
                        Console.WriteLine($"Attribute '{attributeToDelete}' has been removed from the element '{complexElementName}'.");

                        // Save the changes back to the file
                        xdoc.Save(xmlFilePath);
                    }
                    else
                    {
                        Console.WriteLine($"Attribute '{attributeToDelete}' not found in element '{complexElementName}'.");
                    }
                }
                else
                {
                    Console.WriteLine($"The complex element '{complexElementName}' was not found in the XML file.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public static void DeleteAttributeFromRoot(string xmlFilePath, string attributeToDelete)
        {
            // Load the XML document
            XDocument xdoc = XDocument.Load(xmlFilePath);

            // Access the root element
            XElement root = xdoc.Root;

            // Name of root element in L5Xs
            string complexElementName = "RSLogix5000Content";

            // Find the attribute and remove it
            XAttribute attribute = root.Attribute(attributeToDelete);
            if (attribute != null)
            {
                attribute.Remove();
                Console.WriteLine($"Attribute '{attributeToDelete}' has been removed from the root complex element '{complexElementName}'.");
                // Save the changes back to the file
                xdoc.Save(xmlFilePath);
            }
            else
            {
                Console.WriteLine($"Attribute '{attributeToDelete}' not found in the root complex element '{complexElementName}'.");
            }
        }


        public static void ChangeComplexElementAttribute(string xmlFilePath, string complexElementName, string attributeName, string attributeValue)
        {
            // Load the XML document
            XDocument xdoc = XDocument.Load(xmlFilePath);

            // Find the complex element by name
            XElement complexElement = xdoc.Descendants(complexElementName).FirstOrDefault();

            if (complexElement != null)
            {
                // Add the attribute to the complex element
                complexElement.SetAttributeValue(attributeName, attributeValue);
                Console.WriteLine($"Attribute '{attributeName}' with value '{attributeValue}' has been added to the element '{complexElementName}'.");

                // Save the changes back to the file
                xdoc.Save(xmlFilePath);
            }
            else
            {
                Console.WriteLine($"The complex element '{complexElementName}' was not found in the XML file.");
            }
        }

        public static void AddAttributeToComplexElement(string xmlFilePath, string complexElementName, string attributeName, string attributeValue)
        {
            try
            {
                // Load the XML document
                XDocument xdoc = XDocument.Load(xmlFilePath);

                // Find the complex element by name
                XElement complexElement = xdoc.Descendants(complexElementName).LastOrDefault();

                if (complexElement != null)
                {
                    // Add the attribute to the complex element
                    complexElement.SetAttributeValue(attributeName, attributeValue);
                    Console.WriteLine($"Attribute '{attributeName}' with value '{attributeValue}' has been added to the element '{complexElementName}'.");

                    // Save the changes back to the file
                    xdoc.Save(xmlFilePath);
                }
                else
                {
                    Console.WriteLine($"The complex element '{complexElementName}' was not found in the XML file.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public static void AddElementToComplexElement(string xmlFilePath, string complexElementName, string newElementName)
        {
            try
            {
                // Load the XML document
                XDocument xdoc = XDocument.Load(xmlFilePath);

                // Find the complex element by name
                XElement complexElement = xdoc.Descendants(complexElementName).LastOrDefault();

                if (complexElement != null)
                {
                    // Create the new element
                    XElement newElement = new XElement(newElementName);
                    complexElement.Add(newElement);
                    Console.WriteLine($"Element '{newElementName}' has been added to the complex element '{complexElementName}'.");

                    // Save the changes back to the file
                    xdoc.Save(xmlFilePath);
                }
                else
                {
                    Console.WriteLine($"The complex element '{complexElementName}' was not found in the XML file.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }


        public static void CreateCData(string xmlFilePath, string complexElementName, string cdataContent)
        {
            try
            {
                // Load the XML document
                XDocument xdoc = XDocument.Load(xmlFilePath);

                // Find the complex element by name
                XElement complexElement = xdoc.Descendants(complexElementName).LastOrDefault();

                if (complexElement != null)
                {
                    // Create a new CDATA section and add it to the complex element
                    XCData cdataSection = new XCData(cdataContent);
                    complexElement.Add(cdataSection);
                    Console.WriteLine($"A new CDATA section has been created and added to the element '{complexElementName}'.");

                    // Save the changes back to the file
                    xdoc.Save(xmlFilePath);
                }
                else
                {
                    Console.WriteLine($"The complex element '{complexElementName}' was not found in the XML file.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public static string Get_CDATAfromXML_forData(string xmlFilePath)
        {
            try
            {
                // Load the XML document
                XDocument doc = XDocument.Load(xmlFilePath);

                // Find all "Parameter" elements
                var parameterElements = doc
                    .Descendants("Parameters")
                    .Elements("Parameter")
                    .Where(param => param.Attribute("DataType")?.Value != "BOOL")
                    .Descendants("DefaultData")
                    .Where(defaultData => defaultData.FirstNode is XCData)
                    .Select(defaultData => ((XCData)defaultData.FirstNode).Value.Trim())
                    .ToList();

                string joined_pCDATA = string.Join(",", parameterElements);

                // Find all "LocalTag" elements
                var localtagElements = doc
                    .Descendants("LocalTags")
                    .Elements("LocalTag")
                    .Where(param => param.Attribute("DataType")?.Value != "BOOL")
                    .Descendants("DefaultData")
                    .Where(defaultData => defaultData.FirstNode is XCData)
                    .Select(defaultData => ((XCData)defaultData.FirstNode).Value.Trim())
                    .ToList();

                string joined_ltCDATA = string.Join(",", localtagElements);

                string returnString = "[1," + joined_pCDATA + "," + joined_ltCDATA + "]";
                Console.WriteLine("\n\n\n\nreturnString: " + returnString);
                return returnString;
            }
            catch (Exception e)
            {
                Console.WriteLine("ERROR: " + e.Message);
                return e.Message;
            }
        }

        public static string Get_CDATAfromXML_forText(string xmlFilePath)
        {
            string? aoiName = Get_AttributeValue(xmlFilePath, "AddOnInstructionDefinition", "Name");
            StringBuilder sb = new StringBuilder();
            try
            {
                // Load the XML document
                XDocument doc = XDocument.Load(xmlFilePath);

                // Find all "Parameter" elements
                var parameterElements = doc.Descendants("Parameter");

                foreach (var param in parameterElements)
                {
                    XAttribute nameAttribute = param.Attribute("Name");

                    if ((nameAttribute != null) && (param.Attribute("Required").Value == "true"))
                        sb.Append($",AOI_{aoiName}.{nameAttribute.Value}");
                }

                string returnString = $"{aoiName}(AOI_{aoiName}{sb});";
                Console.WriteLine("\n\n\n\nreturnString: " + returnString);
                return returnString;
            }
            catch (Exception e)
            {
                Console.WriteLine("ERROR: " + e.Message);
                return e.Message;
            }
        }

        public static List<Dictionary<string, string>> Get_DataValueMemberInfofromXML(string xmlFilePath)
        {
            List<Dictionary<string, string>> attributeList = new List<Dictionary<string, string>>();

            try
            {
                XDocument doc = XDocument.Load(xmlFilePath);

                foreach (var paramElem in doc.Descendants("Parameter"))
                {
                    Dictionary<string, string> attributes = new Dictionary<string, string>
                {
                    { "Name", paramElem.Attribute("Name")?.Value },
                    { "DataType", paramElem.Attribute("DataType")?.Value },
                    { "Radix", paramElem.Attribute("Radix")?.Value }
                };

                    attributeList.Add(attributes);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error: {e.Message}");
            }

            return attributeList;
        }

        public static string AddComplexElementsToXml(string xmlFilePath, List<Dictionary<string, string>> attributesList)
        {
            try
            {
                // Add the new attributes
                foreach (var attributes in attributesList)
                {
                    AddElementToComplexElement(xmlFilePath, "Structure", "DataValueMember");

                    AddAttributeToComplexElement(xmlFilePath, "DataValueMember", "Name", attributes["Name"]);

                    AddAttributeToComplexElement(xmlFilePath, "DataValueMember", "DataType", attributes["DataType"]);

                    if (attributes["DataType"] != "BOOL")
                        AddAttributeToComplexElement(xmlFilePath, "DataValueMember", "Radix", attributes["Radix"]);

                    if (attributes["Name"] == "EnableIn")
                        AddAttributeToComplexElement(xmlFilePath, "DataValueMember", "Value", "1");
                    else if (attributes["DataType"] == "REAL")
                        AddAttributeToComplexElement(xmlFilePath, "DataValueMember", "Value", "0.0");
                    else
                        AddAttributeToComplexElement(xmlFilePath, "DataValueMember", "Value", "0");
                }

                return "Complex elements added successfully.";
            }
            catch (Exception e)
            {
                return $"Error: {e.Message}";
            }
        }
        #endregion

        #region METHODS: reading excel file
        private static int GetPopulatedColumnCount(string filePath, int rowNumber)
        {
            int return_populatedColumnCount = 0;
            FileInfo existingFile = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int maxColumnNum = worksheet.Dimension.End.Column;

                for (int col = 1; col <= maxColumnNum; col++)
                {
                    var cellValue = worksheet.Cells[rowNumber, col].Value;

                    if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                        return_populatedColumnCount++;
                }
            }
            return return_populatedColumnCount;
        }


        private static int GetPopulatedRowCount(string filePath, int columnNumber)
        {
            int return_populatedRowCount = 0;
            FileInfo existingFile = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int maxRowNum = worksheet.Dimension.End.Row;

                for (int row = 1; row <= maxRowNum; row++)
                {
                    var cellValue = worksheet.Cells[row, columnNumber].Value;

                    if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                        return_populatedRowCount++;
                }
            }
            return return_populatedRowCount;
        }


        private static void ReadXLS(string filePath)
        {
            FileInfo existingFile = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count
                Console.WriteLine("colCount: " + colCount);
                Console.WriteLine("rowCount: " + rowCount);
                for (int row = 1; row <= rowCount; row++)
                {
                    for (int col = 1; col <= colCount; col++)
                    {
                        Console.WriteLine(" Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value?.ToString()!.Trim()!);
                    }
                }
            }
        }

        private static void Print_AOIParameters(AOIParameters[] dataPointsArray)
        {
            int arraySize = dataPointsArray.Length;
            Console.WriteLine("arraySize: " + arraySize);

            for (int i = 0; i < arraySize; i++)
            {
                Console.WriteLine($"Name: {dataPointsArray[i].Name,-20} | Data Type: {dataPointsArray[i].DataType,-9} | " +
                    $"Scope: {dataPointsArray[i].Usage,-7} | Required: {dataPointsArray[i].Required,-5} | " +
                    $"Visible: {dataPointsArray[i].Visible,-5} |  Value: {dataPointsArray[i].Value,-20} | " +
                    $"Byte Position: {dataPointsArray[i].BytePosition,-3} | Bool Position: {dataPointsArray[i].BoolPosition}");
            }
        }

        private static AOIParameters[] Get_AOIParameters(string filePath)
        {
            int parameterCount;
            AOIParameters[] returnDataPoints;

            FileInfo existingFile = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                parameterCount = GetPopulatedRowCount(filePath, 2) - 6;
                returnDataPoints = new AOIParameters[parameterCount];

                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                for (int row = 0; row < parameterCount; row++)
                {
                    var paramName = worksheet.Cells[row + 20, 2].Value.ToString()!.Trim();
                    var paramDataType = worksheet.Cells[row + 20, 3].Value.ToString()!.Trim();
                    var paramScope = worksheet.Cells[row + 20, 4].Value.ToString()!.Trim();

                    AOIParameters dataPoint = new AOIParameters();
                    dataPoint.Name = paramName;
                    dataPoint.DataType = paramDataType;
                    dataPoint.Usage = paramScope;

                    returnDataPoints[row] = dataPoint;
                }
            }
            return returnDataPoints;
        }


        private static AOIParameters[] Get_AOIParameters_FromL5X(string l5xPath)
        {
            XDocument xDoc = XDocument.Load(l5xPath);
            int parameterCount = xDoc.Descendants("Parameters").FirstOrDefault().Elements().Count();
            AOIParameters[] returnDataPoints = new AOIParameters[parameterCount];
            int paramIndex = 0;

            foreach (var p in xDoc.Descendants("Parameter"))
            {
                returnDataPoints[paramIndex].Name = p.Attribute("Name").Value;
                returnDataPoints[paramIndex].DataType = p.Attribute("DataType").Value;
                returnDataPoints[paramIndex].Usage = p.Attribute("Usage").Value;
                returnDataPoints[paramIndex].Required = ToBoolean(p.Attribute("Required").Value);
                returnDataPoints[paramIndex].Visible = ToBoolean(p.Attribute("Visible").Value);
                paramIndex++;
            }

            return returnDataPoints;
        }
        #endregion

        #region METHODS: formatting text file
        /// <summary>
        /// Modify the input string to wrap the text to the next line after a certain length.<br/>
        /// The input string is seperated per word and then each line is incrementally added to per word.<br/>
        /// Start a new line when the character count of a line exceeds 125.
        /// </summary>
        /// <param name="inputString">The input string to be wrapped.</param>
        /// <param name="indentLength">An integer that defines the length of the characters in the indent starting each new line.</param>
        /// <param name="lineLimit">An integer that defines the maximum number of characters per line before a new line is created.</param>
        /// <returns>A modified string that wraps every 125 characters.</returns>
        private static string WrapText(string inputString, int indentLength, int lineLimit)
        {
            string[] words = inputString.Split(' ');
            string indent = new(' ', indentLength);
            StringBuilder newSentence = new();
            string line = "";
            int numberOfNewLines = 0;
            foreach (string word in words)
            {
                word.Trim();
                if ((line + word).Length > lineLimit)
                {
                    if (numberOfNewLines == 0)
                        newSentence.AppendLine(line);
                    else
                        newSentence.AppendLine(indent + line);
                    line = "";
                    numberOfNewLines++;
                }
                line += string.Format($"{word} ");
            }
            if (line.Length > 0)
            {
                if (numberOfNewLines > 0)
                    newSentence.AppendLine(indent + line);
                else
                    newSentence.AppendLine(line);
            }
            return newSentence.ToString();
        }

        /// <summary>
        /// Create a banner used to identify the portion of the test being executed and write it to console.
        /// </summary>
        /// <param name="bannerName">The name displayed in the console banner.</param>
        private static void CreateBanner(string bannerName)
        {
            string final_banner = "-=[" + bannerName + "]=---";
            final_banner = final_banner.PadLeft(125, '-');
            Console.WriteLine(final_banner);
        }

        /// <summary>
        /// Delete all the oldest files in a specified folder that exceeds the chosen number of files to keep.
        /// </summary>
        /// <param name="folderPath">The full path to the folder that will be cleaned.</param>
        /// <param name="keepCount">The number of files in a folder to be kept.</param>
        private static void CleanTestReportsFolder(string folderPath, int keepCount)
        {
            Console.WriteLine($"STATUS:  {folderPath} set to retain {keepCount} test files");
            string[] all_files = Directory.GetFiles(folderPath);
            var orderedFiles = all_files.Select(f => new FileInfo(f)).OrderBy(f => f.CreationTime).ToList();
            if (orderedFiles.Count > keepCount)
            {
                for (int i = 0; i < (orderedFiles.Count - keepCount); i++)
                {
                    FileInfo deleteThisFile = orderedFiles[i];
                    deleteThisFile.Delete();
                    Console.WriteLine($"SUCCESS: deleted {deleteThisFile.FullName}");
                }
            }
            else
                Console.WriteLine($"SUCCESS: no files needed to be deleted (currently {orderedFiles.Count} test files)");
        }
        #endregion

        #region METHODS: formatting excel file
        /// <summary>
        /// Convert the integer number of a column to the Microsoft Excel letter formatting.
        /// </summary>
        /// <param name="columnNumber">Integer input to convert to excel letter formatting.</param>
        /// <returns>A string in excel letter formatting.</returns>
        private static string ConvertToExcelColumn(int columnNumber)
        {
            string columnName = "";

            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }

            return columnName;
        }

        /// <summary>
        /// Initialize the Mircosoft Excel test report with one sheet containing test information.
        /// </summary>
        /// <param name="excelFilePath">The file path to the excel workbook containing the test results.</param>
        /// <param name="acdFilePath">The file path to the Studio 5000 Logix Designer ACD file being tested.</param>
        /// <param name="name">The name of the person starting the test.</param>
        /// <param name="email">The email of the person starting the test.</param>
        /// <param name="buildNumber">The Jenkins job build number.</param>
        /// <param name="jobName">The Jenkins job name.</param>
        /// <param name="gitHash">The git hash of the commit being tested.</param>
        /// <param name="gitMessage">The git message of the commit being tested.</param>
        private static void CreateFormattedExcelFile(string excelFilePath, string acdFilePath, string name, string email, string buildNumber, string jobName, string gitHash, string gitMessage)
        {
            ExcelPackage excelPackage = new ExcelPackage();
            var returnWorkBook = excelPackage.Workbook;

            var TestSummarySheet = returnWorkBook.Worksheets.Add("TestSummary");
            TestSummarySheet.Cells["B2:C2"].Merge = true;
            TestSummarySheet.Cells["B2"].Value = "CI/CD Test Stage Results";
            TestSummarySheet.Cells["B2"].Style.Font.Bold = true;
            TestSummarySheet.Cells["B3"].Value = "Jenkins job name:";
            TestSummarySheet.Cells["C3"].Value = jobName;
            TestSummarySheet.Cells["B4"].Value = "Jenkins job build number:";
            TestSummarySheet.Cells["C4"].Value = buildNumber;
            TestSummarySheet.Cells["B5"].Value = "Tester name:";
            TestSummarySheet.Cells["C5"].Value = name;
            TestSummarySheet.Cells["B6"].Value = "Tester contact:";
            TestSummarySheet.Cells["C6"].Value = email;
            TestSummarySheet.Cells["B7"].Value = "ACD file specified:";
            TestSummarySheet.Cells["C7"].Value = acdFilePath;
            TestSummarySheet.Cells["B8"].Value = "Git commit hash to be verified:";
            TestSummarySheet.Cells["C8"].Value = gitHash;
            TestSummarySheet.Cells["B9"].Value = "Git commit message to be verified:";
            TestSummarySheet.Cells["C9"].Value = gitMessage;

            TestSummarySheet.Column(2).Style.Font.Size = 14;
            TestSummarySheet.Column(2).AutoFit();
            TestSummarySheet.Column(2).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            TestSummarySheet.Column(2).Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
            TestSummarySheet.Cells["B2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            TestSummarySheet.Cells["B2"].Style.Font.Size = 20;

            TestSummarySheet.Column(3).Style.Font.Size = 14;
            TestSummarySheet.Column(3).Width = 95;
            TestSummarySheet.Column(3).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            TestSummarySheet.Cells["C9"].Style.WrapText = true;
            TestSummarySheet.Cells["C9"].Style.ShrinkToFit = true;

            excelPackage.SaveAs(new System.IO.FileInfo(excelFilePath));
        }

        /// <summary>
        /// Create the InitialTagValues sheet in the existing Microsoft Excel workbook.
        /// </summary>
        /// <param name="excelFilePath">The file path to the excel workbook containing the test results.</param>
        private static void CreateInitialTagValuesSheet(string excelFilePath)
        {
            FileInfo existingFile = new FileInfo(excelFilePath);
            using (ExcelPackage excelPackage = new ExcelPackage(existingFile))
            {
                var InitialTagValues = excelPackage.Workbook.Worksheets.Add("InitialTagValues");

                InitialTagValues.Cells["A1"].Value = "Tag Name";
                InitialTagValues.Cells["B1"].Value = "Tag Type";
                InitialTagValues.Cells["C1"].Value = "Online Value";
                InitialTagValues.Cells["D1"].Value = "Offline Value";
                InitialTagValues.Cells["A1:D1"].Style.Font.Bold = true;

                for (int i = 1; i <= 4; i++)
                    InitialTagValues.Column(i).AutoFit();

                InitialTagValues.View.FreezePanes(2, 1);
                excelPackage.Save();
            }
        }

        /// <summary>
        /// Create the TurthTable_Example sheet and add the results of the MainRoutine Rung 3 test.
        /// </summary>
        /// <param name="excelFilePath">The file path to the excel workbook containing the test results.</param>
        /// <param name="truthTableResults">The string array containing every boolean combination result from MainRoutine Rung 3 testing.</param>
        private static void CreateTruthTableSheet(string excelFilePath, string[] truthTableResults)
        {
            FileInfo existingFile = new FileInfo(excelFilePath);
            using (ExcelPackage excelPackage = new ExcelPackage(existingFile))
            {
                var TruthTable_ExampleSheet = excelPackage.Workbook.Worksheets.Add("TruthTable_Example");
                TruthTable_ExampleSheet.View.FreezePanes(2, 2);
                TruthTable_ExampleSheet.Cells["A1"].Value = "Rung 3 of Main Program Test Cases";
                TruthTable_ExampleSheet.Cells["A1"].Style.Font.Bold = true;
                TruthTable_ExampleSheet.Cells["A2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                TruthTable_ExampleSheet.Cells["A3"].Value = "UDT_AllAtomicDataTypes.ex_BOOL1";
                TruthTable_ExampleSheet.Cells["A4"].Value = "UDT_AllAtomicDataTypes.ex_BOOL2";
                TruthTable_ExampleSheet.Cells["A5"].Value = "UDT_AllAtomicDataTypes.ex_BOOL3";
                TruthTable_ExampleSheet.Cells["A6"].Value = "UDT_AllAtomicDataTypes.ex_BOOL4";
                TruthTable_ExampleSheet.Cells["A7"].Value = "UDT_AllAtomicDataTypes.ex_BOOL5";
                TruthTable_ExampleSheet.Cells["A8"].Value = "UDT_AllAtomicDataTypes.ex_BOOL8";
                int numberOfTests = truthTableResults.Length;
                string columnLetter = "";
                for (int i = 1; i < numberOfTests; i++)
                {
                    columnLetter = ConvertToExcelColumn(i + 1);
                    TruthTable_ExampleSheet.Cells[$"{columnLetter}1"].Value = i;

                    TruthTable_ExampleSheet.Cells[$"{columnLetter}3"].Value = Char.GetNumericValue(truthTableResults[i][7]);
                    TruthTable_ExampleSheet.Cells[$"{columnLetter}4"].Value = Char.GetNumericValue(truthTableResults[i][6]);
                    TruthTable_ExampleSheet.Cells[$"{columnLetter}5"].Value = Char.GetNumericValue(truthTableResults[i][5]);
                    TruthTable_ExampleSheet.Cells[$"{columnLetter}6"].Value = Char.GetNumericValue(truthTableResults[i][4]);
                    TruthTable_ExampleSheet.Cells[$"{columnLetter}7"].Value = Char.GetNumericValue(truthTableResults[i][3]);
                    TruthTable_ExampleSheet.Cells[$"{columnLetter}8"].Value = Char.GetNumericValue(truthTableResults[i][0]);

                    TruthTable_ExampleSheet.Column(i + 1).Width = 4;
                    TruthTable_ExampleSheet.Column(i + 1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                }

                TruthTable_ExampleSheet.Column(1).AutoFit();
                TruthTable_ExampleSheet.Column(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                excelPackage.Save();
            }
        }

        /// <summary>
        /// Create the FinalTagValues sheet in the existing Microsoft Excel workbook.
        /// </summary>
        /// <param name="excelFilePath">The file path to the excel workbook containing the test results.</param>
        private static void CreateFinalTagValuesSheet(string excelFilePath)
        {
            FileInfo existingFile = new FileInfo(excelFilePath);
            using (ExcelPackage excelPackage = new ExcelPackage(existingFile))
            {
                var InitialTagValues = excelPackage.Workbook.Worksheets.Add("FinalTagValues");

                InitialTagValues.Cells["A1"].Value = "Tag Name";
                InitialTagValues.Cells["B1"].Value = "Tag Type";
                InitialTagValues.Cells["C1"].Value = "Online Value";
                InitialTagValues.Cells["D1"].Value = "Offline Value";
                InitialTagValues.Cells["A1:D1"].Style.Font.Bold = true;

                for (int i = 1; i <= 4; i++)
                    InitialTagValues.Column(i).AutoFit();

                InitialTagValues.View.FreezePanes(2, 1);
                excelPackage.Save();
            }
        }

        /// <summary>
        /// Add a row containing tag information to a specific worksheet in an existing excel workbook.
        /// </summary>
        /// <param name="excelFilePath">The file path to the excel workbook containing the test results.</param>
        /// <param name="sheetName">The name of the worksheet being modified.</param>
        /// <param name="rowNumber">The row number of the worksheet to be modified.</param>
        /// <param name="tagName">The name of the tag for which further information is provided.</param>
        /// <param name="dataType">The datatype of the tag.</param>
        /// <param name="onlineValue">The online value of the tag.</param>
        /// <param name="offlineValue">The offline value of the tag.</param>
        private static void AddRowToSheet(string excelFilePath, string sheetName, int rowNumber, string tagName, string dataType, string onlineValue, string offlineValue)
        {
            FileInfo existingFile = new FileInfo(excelFilePath);
            using (ExcelPackage excelPackage = new ExcelPackage(existingFile))
            {
                ExcelWorksheet modifiedSheet = excelPackage.Workbook.Worksheets[sheetName];

                modifiedSheet.Cells[$"A{rowNumber}"].Value = tagName;
                modifiedSheet.Cells[$"B{rowNumber}"].Value = dataType;
                modifiedSheet.Cells[$"C{rowNumber}"].Value = onlineValue;
                modifiedSheet.Cells[$"D{rowNumber}"].Value = offlineValue;

                for (int i = 1; i <= 4; i++)
                    modifiedSheet.Column(i).AutoFit();

                excelPackage.Save();
            }
        }
        #endregion

        #region METHODS: get/set basic data type tags
        /// <summary>
        /// Asynchronously get the online and offline value of a basic data type tag.<br/>
        /// (basic data types handled: boolean, single integer, integer, double integer, long integer, real, string)
        /// </summary>
        /// <param name="tagName">The name of the tag in Studio 5000 Logix Designer whose value will be returned.</param>
        /// <param name="type">The data type of the tag whose value will be returned.</param>
        /// <param name="tagPath">
        /// The tag path specifying the tag's scope and location in the Studio 5000 Logix Designer project.<br/>
        /// The tag path is based on the XML filetype (L5X) encapsulation of elements.
        /// </param>
        /// <param name="project">An instance of the LogixProject class.</param>
        /// <param name="printout">A boolean that, if True, prints the online and offline values to the console.</param>
        /// <returns>
        /// A Task that results in a string array containing tag information:<br/>
        /// return_array[0] = tag name<br/>
        /// return_array[1] = online tag value<br/>
        /// return_array[2] = offline tag value
        /// </returns>
        private static async Task<string[]> Get_TagValue_Async(string tagName, DataType type, string tagPath, LogixProject project, bool printout)
        {
            string[] return_array = new string[3];
            tagPath = tagPath + $"[@Name='{tagName}']";
            return_array[0] = tagName;
            try
            {
                if (type == DataType.BOOL)
                {
                    var tagValue_online = await project.GetTagValueBOOLAsync(tagPath, OperationMode.Online);
                    return_array[1] = $"{tagValue_online}";
                    var tagValue_offline = await project.GetTagValueBOOLAsync(tagPath, OperationMode.Offline);
                    return_array[2] = $"{tagValue_offline}";

                }
                else if (type == DataType.SINT)
                {
                    var tagValue_online = await project.GetTagValueSINTAsync(tagPath, OperationMode.Online);
                    return_array[1] = $"{tagValue_online}";
                    var tagValue_offline = await project.GetTagValueSINTAsync(tagPath, OperationMode.Offline);
                    return_array[2] = $"{tagValue_offline}";
                }
                else if (type == DataType.INT)
                {
                    var tagValue_online = await project.GetTagValueINTAsync(tagPath, OperationMode.Online);
                    return_array[1] = $"{tagValue_online}";
                    var tagValue_offline = await project.GetTagValueINTAsync(tagPath, OperationMode.Offline);
                    return_array[2] = $"{tagValue_offline}";
                }
                else if (type == DataType.DINT)
                {
                    var tagValue_online = await project.GetTagValueDINTAsync(tagPath, OperationMode.Online);
                    return_array[1] = $"{tagValue_online}";
                    var tagValue_offline = await project.GetTagValueDINTAsync(tagPath, OperationMode.Offline);
                    return_array[2] = $"{tagValue_offline}";
                }
                else if (type == DataType.LINT)
                {
                    var tagValue_online = await project.GetTagValueLINTAsync(tagPath, OperationMode.Online);
                    return_array[1] = $"{tagValue_online}";
                    var tagValue_offline = await project.GetTagValueLINTAsync(tagPath, OperationMode.Offline);
                    return_array[2] = $"{tagValue_offline}";
                }
                else if (type == DataType.REAL)
                {
                    var tagValue_online = await project.GetTagValueREALAsync(tagPath, OperationMode.Online);
                    return_array[1] = $"{tagValue_online}";
                    var tagValue_offline = await project.GetTagValueREALAsync(tagPath, OperationMode.Offline);
                    return_array[2] = $"{tagValue_offline}";
                }
                else if (type == DataType.STRING)
                {
                    var tagValue_online = await project.GetTagValueSTRINGAsync(tagPath, OperationMode.Online);
                    return_array[1] = (tagValue_online == "") ? "<empty_string>" : $"{tagValue_online}";
                    var tagValue_offline = await project.GetTagValueSTRINGAsync(tagPath, OperationMode.Offline);
                    return_array[2] = (tagValue_offline == "") ? "<empty_string>" : $"{tagValue_offline}";
                }
                else
                    Console.WriteLine(WrapText($"ERROR executing command: The tag {tagName} cannot be handled. Select either DINT, BOOL, or REAL.", 9, 125));
            }
            catch (LogixSdkException ex)
            {
                Console.WriteLine($"ERROR getting tag {tagName}");
                Console.WriteLine(ex.Message);
            }

            if (printout)
            {
                string online_message = $"online value: {return_array[1]}";
                string offline_message = $"offline value: {return_array[2]}";
                Console.WriteLine($"SUCCESS: " + tagName.PadRight(40, ' ') + online_message.PadRight(35, ' ') + offline_message.PadRight(35, ' '));
            }

            return return_array;
        }

        /// <summary>
        /// Run the GetTagValueAsync method synchronously.<br/>
        /// Get the online and offline value of a basic data type tag.<br/>
        /// (basic data types handled: boolean, single integer, integer, double integer, long integer, real, string)
        /// </summary>
        /// <param name="tagName">The name of the tag whose value will be returned.</param>
        /// <param name="type">The data type of the tag whose value will be returned.</param>
        /// <param name="tagPath">
        /// The tag path specifying the tag's scope and location in the Studio 5000 Logix Designer project.<br/>
        /// The tag path is based on the XML filetype (L5X) encapsulation of elements.
        /// </param>
        /// <param name="project">An instance of the LogixProject class.</param>
        /// <param name="printout">A boolean that, if True, prints the online and offline values to the console.</param>
        /// <returns>
        /// A Task that results in a string array containing tag information:<br/>
        /// return_array[0] = tag name<br/>
        /// return_array[1] = online tag value<br/>
        /// return_array[2] = offline tag value
        /// </returns>
        private static string[] Get_TagValue_Sync(string tagName, DataType type, string tagPath, LogixProject project, bool printout)
        {
            var task = Get_TagValue_Async(tagName, type, tagPath, project, printout);
            task.Wait();
            return task.Result;
        }

        /// <summary>
        /// Asynchronously set either the online or offline value of a basic data type tag.
        /// (basic data types handled: boolean, single integer, integer, double integer, long integer, real, string)
        /// </summary>
        /// <param name="tagName">The name of the tag whose value will be set.</param>
        /// <param name="newTagValue">The value of the tag that will be set.</param>
        /// <param name="mode">This specifies whether the 'Online' or 'Offline' value of the tag is the one to set.</param>
        /// <param name="type">The data type of the tag whose value will be set.</param>
        /// <param name="tagPath">
        /// The tag path specifying the tag's scope and location in the Studio 5000 Logix Designer project.
        /// The tag path is based on the XML filetype (L5X) encapsulation of elements.
        /// </param>
        /// <param name="project">An instance of the LogixProject class.</param>
        /// <param name="printout">A boolean that, if True, prints the online and offline values to the console.</param>
        /// <returns>A Task that will set the online or offline value of a basic data type tag.</returns>
        private static async Task Set_TagValue_Async(string tagName, string newTagValue, OperationMode mode, DataType type, string tagPath, LogixProject project, bool printout)
        {
            tagPath = tagPath + $"[@Name='{tagName}']";
            string[] old_tag_values = await Get_TagValue_Async(tagName, type, tagPath, project, false);
            string old_tag_value = "";
            try
            {
                if (mode == OperationMode.Online)
                {

                    if (type == DataType.BOOL)
                        await project.SetTagValueBOOLAsync(tagPath, OperationMode.Online, bool.Parse(newTagValue));
                    else if (type == DataType.SINT)
                        await project.SetTagValueSINTAsync(tagPath, OperationMode.Online, sbyte.Parse(newTagValue));
                    else if (type == DataType.INT)
                        await project.SetTagValueINTAsync(tagPath, OperationMode.Online, short.Parse(newTagValue));
                    else if (type == DataType.DINT)
                        await project.SetTagValueDINTAsync(tagPath, OperationMode.Online, int.Parse(newTagValue));
                    else if (type == DataType.LINT)
                        await project.SetTagValueLINTAsync(tagPath, OperationMode.Online, long.Parse(newTagValue));
                    else if (type == DataType.REAL)
                        await project.SetTagValueREALAsync(tagPath, OperationMode.Online, float.Parse(newTagValue));
                    else if (type == DataType.STRING)
                        await project.SetTagValueSTRINGAsync(tagPath, OperationMode.Online, newTagValue);
                    else
                        Console.WriteLine($"ERROR executing command: The data type cannot be handled. Select either DINT, BOOL, or REAL.");
                    old_tag_value = old_tag_values[1];
                }
                else if (mode == OperationMode.Offline)
                {
                    if (type == DataType.BOOL)
                        await project.SetTagValueBOOLAsync(tagPath, OperationMode.Offline, bool.Parse(newTagValue));
                    else if (type == DataType.SINT)
                        await project.SetTagValueSINTAsync(tagPath, OperationMode.Offline, sbyte.Parse(newTagValue));
                    else if (type == DataType.INT)
                        await project.SetTagValueINTAsync(tagPath, OperationMode.Offline, short.Parse(newTagValue));
                    else if (type == DataType.DINT)
                        await project.SetTagValueDINTAsync(tagPath, OperationMode.Offline, int.Parse(newTagValue));
                    else if (type == DataType.LINT)
                        await project.SetTagValueLINTAsync(tagPath, OperationMode.Offline, long.Parse(newTagValue));
                    else if (type == DataType.REAL)
                        await project.SetTagValueREALAsync(tagPath, OperationMode.Offline, float.Parse(newTagValue));
                    else if (type == DataType.STRING)
                        await project.SetTagValueSTRINGAsync(tagPath, OperationMode.Offline, newTagValue);
                    else
                        Console.WriteLine($"ERROR executing command: The data type cannot be handled. Select either DINT, BOOL, or REAL.");
                    old_tag_value = old_tag_values[2];
                }
            }
            catch (LogixSdkException ex)
            {
                Console.WriteLine("Unable to set tag value.");
                Console.WriteLine(ex.Message);
            }

            try
            {
                await project.SaveAsync();
            }
            catch (LogixSdkException ex)
            {
                Console.WriteLine("Unable to save project");
                Console.WriteLine(ex.Message);
            }

            if (printout)
            {
                string new_tag_value_string = Convert.ToString(newTagValue);
                if ((new_tag_value_string == "1") && (type == DataType.BOOL)) { new_tag_value_string = "True"; }
                if ((new_tag_value_string == "0") && (type == DataType.BOOL)) { new_tag_value_string = "False"; }
                Console.WriteLine("SUCCESS: " + mode + " " + old_tag_values[0].PadRight(40, ' ') + old_tag_value.PadLeft(20, ' ') + "  -->  " + new_tag_value_string);
            }
        }

        /// <summary>
        /// Run the SetTagValueAsync method synchronously.
        /// Set either the online or offline value of a basic data type tag.
        /// (basic data types handled: boolean, single integer, integer, double integer, long integer, real, string)
        /// </summary>
        /// <param name="tagName">The name of the tag whose value will be set.</param>
        /// <param name="newTagValue">The value of the tag that will be set.</param>
        /// <param name="mode">This specifies whether the 'Online' or 'Offline' value of the tag is the one to set.</param>
        /// <param name="type">The data type of the tag whose value will be set.</param>
        /// <param name="tagPath">
        /// The tag path specifying the tag's scope and location in the Studio 5000 Logix Designer project.
        /// The tag path is based on the XML filetype (L5X) encapsulation of elements.
        /// </param>
        /// <param name="project">An instance of the LogixProject class.</param>
        /// <param name="printout">A boolean that, if True, prints the online and offline values to the console.</param>
        private static void Set_TagValue_Sync(string tagName, string newTagValue, OperationMode mode, DataType type, string tagPath, LogixProject project, bool printout)
        {
            var task = Set_TagValue_Async(tagName, newTagValue, mode, type, tagPath, project, printout);
            task.Wait();
        }
        #endregion

        #region METHODS: get/set complex data type tags
        /// <summary>
        /// Asynchronously get the online and offline data in ByteString form for the complex tag UDT_AllAtomicDataTypes.
        /// </summary>
        /// <param name="fullTagPath">
        /// The tag path specifying the tag's scope and location in the Studio 5000 Logix Designer project.<br/>
        /// The tag path is based on the XML filetype (L5X) encapsulation of elements.</param>
        /// <param name="project">An instance of the LogixProject class.</param>
        /// <returns>
        /// A Task that results in a ByteString array containing UDT_AllAtomicDataTypes information:<br/>
        /// returnByteStringArray[0] = online tag values<br/>
        /// returnByteStringArray[1] = offline tag values
        /// </returns>
        private static async Task<ByteString> Get_UDTorAOI_ByteString_Async(string fullTagPath, LogixProject project, OperationMode online_or_offline)
        {
            ByteString returnByteStringArray = ByteString.Empty;
            if (online_or_offline == OperationMode.Online)
                returnByteStringArray = await project.GetTagValueAsync(fullTagPath, OperationMode.Online, DataType.BYTE_ARRAY);
            else if (online_or_offline == OperationMode.Offline)
                returnByteStringArray = await project.GetTagValueAsync(fullTagPath, OperationMode.Offline, DataType.BYTE_ARRAY);
            else
                Console.WriteLine("FAILURE: The input " + online_or_offline + " is not a valid selection. Input either OperationMode.Online or OperationMode.Offline");

            return returnByteStringArray;
        }

        /// <summary>
        /// Run the GetUDT_AllAtomicDataTypesAsync Method synchronously.<br/>
        /// Get the online and offline data in ByteString form for the complex tag UDT_AllAtomicDataTypes.
        /// </summary>
        /// <param name="tagPath">
        /// The tag path specifying the tag's scope and location in the Studio 5000 Logix Designer project.<br/>
        /// The tag path is based on the XML filetype (L5X) encapsulation of elements.</param>
        /// <param name="project">An instance of the LogixProject class.</param>
        /// <returns>
        /// A ByteString array containing UDT_AllAtomicDataTypes information:<br/>
        /// returnByteStringArray[0] = online tag values<br/>
        /// returnByteStringArray[1] = offline tag values
        /// </returns>
        private static ByteString Get_UDTorAOI_ByteString_Sync(string tagPath, LogixProject project, OperationMode online_or_offline)
        {
            var task = Get_UDTorAOI_ByteString_Async(tagPath, project, online_or_offline);
            task.Wait();
            return task.Result;
        }

        private static string ReverseByteArrayToString(byte[] byteArray)
        {
            StringBuilder sb = new StringBuilder();

            for (int i = byteArray.Length - 1; i >= 0; i--)
            {
                sb.Append(Convert.ToString(byteArray[i], 2).PadLeft(8, '0'));
            }

            return sb.ToString();
        }

        /// <summary>
        /// Custom conversion of a string to its boolean equivalent.
        /// </summary>
        /// <param name="val">String input to be converted to a boolean equivalent.</param>
        /// <returns>Any uppercase/lowercase combination of "TRUE" or "YES" or "1" returns boolean true. All other inputs return boolean false.</returns>
        public static bool ToBoolean(string val)
        {
            switch (val.Trim().ToUpper())
            {
                case "TRUE":
                case "YES":
                case "1":
                    return true;
                default:
                    return false;
            }
        }

        private static async Task SetSingleValue_UDTorAOI(string newParameterValue, string aoiTagPath, string parameterName, OperationMode mode, AOIParameters[] input_TagDataArray, LogixProject project)
        {
            ByteString input_ByteString = Get_UDTorAOI_ByteString_Sync(aoiTagPath, project, mode);
            byte[] new_byteArray = input_ByteString.ToByteArray();
            int arraySize = input_TagDataArray.Length;
            string oldParameterValue = "";

            for (int j = 0; j < arraySize; j++)
            {
                // Search the TagData[] array to get the associated newTagValue data needed.
                if (input_TagDataArray[j].Name == parameterName)
                {
                    DataType dataType = Get_DataType(input_TagDataArray[j].DataType);
                    int bytePosition = input_TagDataArray[j].BytePosition;
                    oldParameterValue = input_TagDataArray[j].Value;

                    if (dataType == DataType.BOOL)
                    {
                        byte[] bools_byteArray = new byte[4];
                        Array.ConstrainedCopy(new_byteArray, bytePosition, bools_byteArray, 0, 4);
                        var bitArray = new BitArray(bools_byteArray);

                        int boolPosition = 31 - input_TagDataArray[j].BoolPosition;
                        bool bool_newTagValue = ToBoolean(newParameterValue);
                        bitArray[boolPosition] = bool_newTagValue;
                        bitArray.CopyTo(bools_byteArray, 0);


                        for (int i = 0; i < 4; ++i)
                            new_byteArray[i + bytePosition] = bools_byteArray[i];
                    }

                    else if (dataType == DataType.SINT)
                    {
                        string sint_string = Convert.ToString(long.Parse(newParameterValue), 2);
                        sint_string = sint_string.Substring(sint_string.Length - 8);
                        new_byteArray[bytePosition] = Convert.ToByte(sint_string, 2);
                    }

                    else if (dataType == DataType.INT)
                    {
                        byte[] int_byteArray = BitConverter.GetBytes(int.Parse(newParameterValue));
                        for (int i = 0; i < 2; ++i)
                            new_byteArray[i + bytePosition] = int_byteArray[i];
                    }

                    else if (dataType == DataType.DINT)
                    {
                        byte[] dint_byteArray = BitConverter.GetBytes(long.Parse(newParameterValue));
                        for (int i = 0; i < 4; ++i)
                            new_byteArray[i + bytePosition] = dint_byteArray[i];
                    }

                    else if (dataType == DataType.LINT)
                    {
                        byte[] lint_byteArray = BitConverter.GetBytes(long.Parse(newParameterValue));
                        for (int i = 0; i < 8; ++i)
                            new_byteArray[i + bytePosition] = lint_byteArray[i];
                    }

                    else if (dataType == DataType.REAL)
                    {
                        byte[] real_byteArray = BitConverter.GetBytes(float.Parse(newParameterValue));
                        for (int i = 0; i < 4; ++i)
                            new_byteArray[i + bytePosition] = real_byteArray[i];
                    }
                    else
                    {
                        Console.WriteLine("ERROR: Data type not supported.");
                    }
                }
            }

            await project.SetTagValueAsync(aoiTagPath, mode, new_byteArray, DataType.BYTE_ARRAY);
            Console.WriteLine($"SUCCESS: {parameterName,-14} | {oldParameterValue,20} -> {newParameterValue,-20}");
        }

        private static DataType Get_DataType(string dataType)
        {
            DataType type;
            switch (dataType)
            {
                case "BOOL":
                    type = DataType.BOOL;
                    break;
                case "SINT":
                    type = DataType.SINT;
                    break;
                case "INT":
                    type = DataType.INT;
                    break;
                case "DINT":
                    type = DataType.DINT;
                    break;
                case "REAL":
                    type = DataType.REAL;
                    break;
                case "LINT":
                    type = DataType.LINT;
                    break;
                default:
                    Console.WriteLine("Error in Get_DataType method: data type not recognized!");
                    throw new ArgumentException();
            }
            return type;
        }

        private static AOIParameters[] Get_AOIParameterValues(AOIParameters[] input_TagDataArray, ByteString input_AOIorUDT_ByteString, bool printout)
        {
            // initialize values needed for this method
            AOIParameters[] output_TagDataArray = input_TagDataArray;
            byte[] input_bytearray = input_AOIorUDT_ByteString.ToByteArray();
            int byteStartPosition_InputByteArray = 0;
            int boolStartPosition_InputByteArray = 0;
            int boolCount = 0;

            // loop through each of the data types provided in the excel sheet
            int arraySize = input_TagDataArray.Length;
            for (int i = 0; i < arraySize; i++)
            {
                string datatype_TagData = input_TagDataArray[i].DataType;

                // Console.WriteLine("START WITH CURRENT DATATYPE: " + name_TagData + " | inputposition_inbytearray: " + inputposition_inbytearray);


                if (datatype_TagData == "BOOL")
                {
                    // Update the "boolean host member" location of the input byte array that is being checked every 32 booleans.
                    if (((boolCount % 32 == 1) && (boolCount > 1)) || (boolCount == 0))
                    {
                        boolStartPosition_InputByteArray = byteStartPosition_InputByteArray;
                        byteStartPosition_InputByteArray += 4;
                    }

                    byte[] bools_bytearray = new byte[4];
                    Array.ConstrainedCopy(input_bytearray, boolStartPosition_InputByteArray, bools_bytearray, 0, 4);

                    string bools_string = ReverseByteArrayToString(bools_bytearray);

                    output_TagDataArray[i].Value = bools_string[31 - boolCount].ToString();
                    output_TagDataArray[i].BytePosition = boolStartPosition_InputByteArray;
                    output_TagDataArray[i].BoolPosition = 31 - boolCount;

                    boolCount++;
                }

                else if (datatype_TagData == "SINT")
                {
                    byte[] sint_bytearray = new byte[1];
                    Array.ConstrainedCopy(input_bytearray, byteStartPosition_InputByteArray, sint_bytearray, 0, 1);
                    string sint_string = Convert.ToString(unchecked((sbyte)sint_bytearray[0]));
                    output_TagDataArray[i].Value = sint_string;
                    output_TagDataArray[i].BytePosition = byteStartPosition_InputByteArray;
                    byteStartPosition_InputByteArray += 1;
                }

                else if (datatype_TagData == "INT")
                {
                    if ((byteStartPosition_InputByteArray % 2) > 0)
                        byteStartPosition_InputByteArray += 2 - (byteStartPosition_InputByteArray % 2);

                    byte[] int_bytearray = new byte[2];
                    Array.ConstrainedCopy(input_bytearray, byteStartPosition_InputByteArray, int_bytearray, 0, 2);
                    string int_string = Convert.ToString(BitConverter.ToInt16(int_bytearray));
                    output_TagDataArray[i].Value = int_string;
                    output_TagDataArray[i].BytePosition = byteStartPosition_InputByteArray;
                    byteStartPosition_InputByteArray += 2;
                }

                else if (datatype_TagData == "DINT")
                {
                    if ((byteStartPosition_InputByteArray % 4) > 0)
                        byteStartPosition_InputByteArray += 4 - (byteStartPosition_InputByteArray % 4);

                    byte[] dint_bytearray = new byte[4];
                    Array.ConstrainedCopy(input_bytearray, byteStartPosition_InputByteArray, dint_bytearray, 0, 4);
                    string dint_string = Convert.ToString(BitConverter.ToInt32(dint_bytearray));
                    output_TagDataArray[i].Value = dint_string;
                    output_TagDataArray[i].BytePosition = byteStartPosition_InputByteArray;
                    byteStartPosition_InputByteArray += 4;
                }

                else if (datatype_TagData == "LINT")
                {
                    if ((byteStartPosition_InputByteArray % 8) > 0)
                        byteStartPosition_InputByteArray += 8 - (byteStartPosition_InputByteArray % 8);

                    byte[] lint_bytearray = new byte[8];
                    Array.ConstrainedCopy(input_bytearray, byteStartPosition_InputByteArray, lint_bytearray, 0, 8);
                    string lint_string = Convert.ToString(BitConverter.ToInt64(lint_bytearray));
                    output_TagDataArray[i].Value = lint_string;
                    output_TagDataArray[i].BytePosition = byteStartPosition_InputByteArray;
                    byteStartPosition_InputByteArray += 8;
                }

                else if (datatype_TagData == "REAL")
                {
                    if ((byteStartPosition_InputByteArray % 4) > 0)
                        byteStartPosition_InputByteArray += 4 - (byteStartPosition_InputByteArray % 4);

                    byte[] real_bytearray = new byte[4];
                    Array.ConstrainedCopy(input_bytearray, byteStartPosition_InputByteArray, real_bytearray, 0, 4);
                    string real_string = Convert.ToString(BitConverter.ToSingle(real_bytearray));
                    output_TagDataArray[i].Value = real_string;
                    output_TagDataArray[i].BytePosition = byteStartPosition_InputByteArray;
                    byteStartPosition_InputByteArray += 4;
                }

                else
                {
                    //byte[] real_bytearray = new byte[24];
                    //Array.ConstrainedCopy(input_bytearray, inputposition_inbytearray, real_bytearray, 0, 4);
                    //string real_string = Convert.ToString(BitConverter.ToSingle(real_bytearray));
                    //outputDataPointArray[i].Value = real_string;
                    //inputposition_inbytearray += 4;
                    output_TagDataArray[i].BytePosition = byteStartPosition_InputByteArray;
                }
            }

            return output_TagDataArray;
        }
        #endregion

        #region METHODS: setting up Logix Echo emulated controller
        /// <summary>
        /// Run the Echo_Program script synchronously.<br/>
        /// Script that sets up an emulated controller for CI/CD software in the loop (SIL) testing.<br/>
        /// If no emulated controller based on the ACD file path yet exists, create one, and then return the communication path.<br/>
        /// If an emulated controller based on the ACD file path exists, only return the communication path.
        /// </summary>
        /// <param name="acdFilePath">The file path to the Studio 5000 Logix Designer ACD file being tested.</param>
        /// <returns>A string containing the communication path of the emulated controller that the ACD project file will go online with during testing.</returns>
        private static string SetUpEmulatedController_Sync(string acdFilePath, string chassis_name, string controller_name)
        {
            var task = LogixEchoMethods.Main(acdFilePath, chassis_name, controller_name);
            task.Wait();
            return task.Result;
        }
        #endregion

        #region METHODS: changing controller mode / download
        /// <summary>
        /// Asynchronously change the controller mode to either Program, Run, or Test mode.
        /// </summary>
        /// <param name="commPath">The controller communication path.</param>
        /// <param name="mode">The controller mode to switch to.</param>
        /// <param name="project">An instance of the LogixProject class.</param>
        /// <returns>A Task that changes the controller mode.</returns>
        private static async Task ChangeControllerMode_Async(string commPath, string mode, LogixProject project)
        {
            var requestedControllerMode = default(LogixProject.RequestedControllerMode);
            if (mode == "Program")
                requestedControllerMode = LogixProject.RequestedControllerMode.Program;
            else if (mode == "Run")
                requestedControllerMode = LogixProject.RequestedControllerMode.Run;
            else if (mode == "Test")
                requestedControllerMode = LogixProject.RequestedControllerMode.Test;
            else
                Console.WriteLine($"ERROR: {mode} is not supported.");

            try
            {
                await project.SetCommunicationsPathAsync(commPath);
            }
            catch (LogixSdkException ex)
            {
                Console.WriteLine($"Unable to set commpath to {commPath}");
                Console.WriteLine(ex.Message);
            }

            try
            {
                await project.ChangeControllerModeAsync(requestedControllerMode);
            }
            catch (LogixSdkException ex)
            {
                Console.WriteLine($"Unable to set mode. Requested mode was {mode}");
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Asynchronously download to the specified controller.
        /// </summary>
        /// <param name="commPath">The controller communication path.</param>
        /// <param name="project">An instance of the LogixProject class.</param>
        /// <returns>An Task that downloads to the specified controller.</returns>
        private static async Task DownloadProject_Async(string commPath, LogixProject project)
        {
            try
            {
                await project.SetCommunicationsPathAsync(commPath);
            }
            catch (LogixSdkException ex)
            {
                Console.WriteLine($"Unable to set commpath to {commPath}");
                Console.WriteLine(ex.Message);
            }

            try
            {
                LogixProject.ControllerMode controllerMode = await project.ReadControllerModeAsync();
                if (controllerMode != LogixProject.ControllerMode.Program)
                {
                    Console.WriteLine($"Controller mode is {controllerMode}. Downloading is possible only if the controller is in 'Program' mode");
                }
            }
            catch (LogixSdkException ex)
            {
                Console.WriteLine($"Unable to read ControllerMode");
                Console.WriteLine(ex.Message);
            }

            try
            {
                await project.DownloadAsync();
            }
            catch (LogixSdkException ex)
            {
                Console.WriteLine($"Unable to download");
                Console.WriteLine(ex.Message);
            }

            // Download modifies the project.
            // Without saving, if used file will be opened again, commands which need correlation
            // between program in the controller and opened project like LoadImageFromSDCard or StoreImageOnSDCard
            // may not be able to succeed because project in the controller won't match opened project.
            try
            {
                await project.SaveAsync();
            }
            catch (LogixSdkException ex)
            {
                Console.WriteLine($"Unable to save project");
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Asynchronously get the current controller mode (FAULTED, PROGRAM, RUN, or TEST).
        /// </summary>
        /// <param name="commPath">The controller communication path.</param>
        /// <param name="project">An instance of the LogixProject class.</param>
        /// <returns>A Task that returns a string of the current controller mode.</returns>
        /// <exception cref="ArgumentOutOfRangeException">Thrown if the returned controller mode is not FAULTED, PROGRAM, RUN, or TEST.</exception>
        private static async Task<string> ReadControllerMode_Async(string commPath, LogixProject project)
        {
            try
            {
                await project.SetCommunicationsPathAsync(commPath);
            }
            catch (LogixSdkException ex)
            {
                Console.WriteLine($"Unable to set commpath to {commPath}");
                Console.WriteLine(ex.Message);
            }

            try
            {
                LogixProject.ControllerMode result = await project.ReadControllerModeAsync();
                switch (result)
                {
                    case LogixProject.ControllerMode.Faulted:
                        return "FAULTED";
                    case LogixProject.ControllerMode.Program:
                        return "PROGRAM";
                    case LogixProject.ControllerMode.Run:
                        return "RUN";
                    case LogixProject.ControllerMode.Test:
                        return "TEST";
                    default:
                        throw new ArgumentOutOfRangeException("Controller mode is unrecognized");
                }
            }
            catch (LogixSdkException ex)
            {
                Console.WriteLine($"Unable to read controller mode");
                Console.WriteLine(ex.Message);
            }

            return "";
        }
        #endregion
    }
}
