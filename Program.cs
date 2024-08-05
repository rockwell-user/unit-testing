// ------------------------------------------------------------------------------------------------------------------------------------------------------------
//
// FileName: UnitTest_Program.cs
// FileType: Visual C# Source file
// Author : Rockwell Automation Engineering
// Created : 2024
// Description : This script conducts Add-On Instruction (AOI) unit testing, utilizing Studio 5000 Logix Designer SDK and Factory Talk Logix Echo SDK.
//
// ------------------------------------------------------------------------------------------------------------------------------------------------------------

using Google.Protobuf;
using L5Xfiles;
using LogixEcho;
using OfficeOpenXml;
using RockwellAutomation.LogixDesigner;
using System.Collections;
using System.Text;
using System.Xml.Linq;
using static RockwellAutomation.LogixDesigner.LogixProject;

namespace AOIUnitTest
{
    /// <summary>
    /// This class contains the methods and logic to programmatically conduct unit testing for Studio 5000 Logix Designer Add-On Instructions.
    /// </summary>
    public class AOIUnitTestMethods
    {
        /// <summary>
        /// The "AOI Parameter" structure houses all the information required to read & use a single parameter of an AOI.<br/>
        /// Note that this structure will always be used in a list, wherein each element pertains to an AOI parameter. 
        /// </summary>
        public struct AOIParameter
        {
            public string? Name { get; set; }     // the AOI parameter's name
            public string? DataType { get; set; } // BOOL/SINT/INT/DINT/LINT/REAL
            public string? Usage { get; set; }    // Input/Output/InOut
            public bool? Required { get; set; }   // is the parameter required in an instruction
            public bool? Visible { get; set; }    // is the parameter visible in an instruction
            public string? Value { get; set; }    // the value of the AOI parameter
            public int BytePosition { get; set; } // used during tag conversion
            public int BoolPosition { get; set; } // used during tag conversion
        }

        static async Task Main()
        {
            // Parse the incoming variables into the main method & set up global variables
            #region PARSE & INITIALIZE VARIABLES
            // static async Task Main(string[] args)
            // {
            //      string inputExcel_UnitTestSetup_filePath = args.Length > 0 ? args[1] : "default value";
            //      string tempFolder = args.Length > 0 ? args[1] : "default value";
            //      string outputExcel_UnitTestResults_filepath = args.Length > 0 ? args[1] : "default value";
            // }
            string inputExcel_UnitTestSetup_filePath = @"C:\Users\ASYost\Desktop\UnitTesting\AOIs_toTest\WetBulbTemperature_FaultCase.xlsx";
            string tempFolder = @"C:\Users\ASYost\Desktop\UnitTesting\ACD_testFiles_generated\";// ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------[MODIFY FOR GITHUB DIRECTORY]
            string outputExcel_UnitTestResults_filepath = @"C:\Users\ASYost\Desktop\UnitTesting\exampleTestReports" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_testfile.xlsx";

            // The below parameters are used to create the ACD application that will be testing the Add-On Instruction specified in the input excel sheet.
            string echoChassisName = "UnitTest_Chassis";
            string controllerName = "UnitTest_Controller";
            string taskName = "T00_AOI_Testing";
            string programName = "P00_AOI_Testing";
            string routineName = "R00_AOI_Testing";
            string programName_FaultHandler = "PXX_FaultHandler";
            string routineName_FaultHandler = "RXX_FaultHandler";
            string processorType = "1756-L85E";
            bool keepACD;
            bool keepL5Xs;
            string aoiFileName = "";
            //string commPath = "";

            FileInfo existingFile = new FileInfo(inputExcel_UnitTestSetup_filePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                aoiFileName = worksheet.Cells[9, 2].Value?.ToString()!.Trim()!;
                keepACD = ToBoolean(worksheet.Cells[9, 4].Value?.ToString()!.Trim()!);
                keepL5Xs = ToBoolean(worksheet.Cells[9, 14].Value?.ToString()!.Trim()!);
            }

            // INCLUDE NEAR TOP
            string aoi_L5Xfilepath = tempFolder + aoiFileName; // ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------[MODIFY FOR GITHUB DIRECTORY]
            string convertedAOIrung_L5Xfilepath = CopyXmlFile(aoi_L5Xfilepath, false);// ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------[MODIFY FOR GITHUB DIRECTORY]
            string? aoiName = GetAttributeValue(convertedAOIrung_L5Xfilepath, "AddOnInstructionDefinition", "Name", false); // The name of the AOI being testing.
            string aoiTagScope = $"Controller/Tags/Tag[@Name='AOI_{aoiName}']";
            #endregion

            #region STAGING TEST: create new ACD -> Logix Echo emulation FINISH THIS LATER
            // Create a new ACD project file.
            ConsoleMessage("START creating new ACD unit test application file...", "NEWSECTION", false);
            string? softwareRevision = GetAttributeValue(convertedAOIrung_L5Xfilepath, "RSLogix5000Content", "SoftwareRevision", false);
            uint softwareRevision_uint = ConvertStringToUint(softwareRevision);

            string faultHandlingApplication_L5Xcontents = L5XfileMethods.GetFaultHandlingApplicationL5XContents(routineName, programName, taskName, routineName_FaultHandler, programName_FaultHandler, controllerName, processorType, softwareRevision); ;
            string L5Xpath = tempFolder + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + aoiName + "_UnitTest.L5X";
            File.WriteAllText(L5Xpath, faultHandlingApplication_L5Xcontents);
            LogixProject projectL5X = await LogixProject.ConvertAsync(L5Xpath, (int)softwareRevision_uint);
            ConsoleMessage($"L5X application file created at '{L5Xpath}'.", "STATUS");
            string acdPath = tempFolder + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + aoiName + "_UnitTest.ACD"; // ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------[MODIFY FOR GITHUB DIRECTORY]
            await projectL5X.SaveAsAsync(acdPath, true);
            ConsoleMessage($"ACD application file created at '{acdPath}'.", "STATUS");

            ConsoleMessage("START opening new ACD unit test application file...", "NEWSECTION", false);
            LogixProject projectACD = await LogixProject.OpenLogixProjectAsync(acdPath);
            ConsoleMessage($"'{acdPath}' application file opened.", "STATUS");

            //// =========================================================================
            //// LOGGER INFO (UNCOMMENT IF TROUBLESHOOTING)
            //// Note: may need to add using RockwellAutomation.LogixDesigner.Logging;
            //var logger = new StdOutEventLogger();
            //project.AddEventHandler(logger);
            //// =========================================================================

            // Execute only if the specified output file does not exist
            if (!File.Exists(outputExcel_UnitTestResults_filepath))
            {
                // Create an excel test report to be filled out during testing.
                ConsoleMessage("START setting up excel test report workbook...", "NEWSECTION");
                //CreateFormattedExcelFile(outputExcel_UnitTestResults_filepath, acdPath, name_mostRecentCommit, email_mostRecentCommit, jenkinsBuildNumber, jenkinsJobName, hash_mostRecentCommit, message_mostRecentCommit);
            }

            // Set up emulated controller (based on the specified ACD file path) if one does not yet exist. If not, continue.
            ConsoleMessage("START setting up Factory Talk Logix Echo emulated controller...", "NEWSECTION");
            string commPath = LogixEchoMethods.Main(acdPath, echoChassisName, controllerName).GetAwaiter().GetResult();
            ConsoleMessage($"Project communication path specified is '{commPath}'.", "STATUS");

            ConsoleMessage("START preparing ACD application for test...", "NEWSECTION");
            // Import the AOI.L5X being tested
            string xPath_aoiDef = @"Controller/AddOnInstructionDefinitions";
            await projectACD.PartialImportFromXmlFileAsync(xPath_aoiDef, aoi_L5Xfilepath, LogixProject.ImportCollisionOptions.OverwriteOnColl);
            await projectACD.SaveAsync();

            // Add custom AOI rung to rung 1
            bool conversionPrintOut = false;
            ConsoleMessage($"Print STATUS messages for AOI to rung conversion? Currently set to '{conversionPrintOut}'.", "STATUS");
            ConvertAOItoRUNGxml(convertedAOIrung_L5Xfilepath, routineName, programName, conversionPrintOut);
            string xPath_convertedAOIrung = @"Controller/Programs/Program[@Name='" + programName + @"']/Routines";
            await projectACD.PartialImportFromXmlFileAsync(xPath_convertedAOIrung, convertedAOIrung_L5Xfilepath, LogixProject.ImportCollisionOptions.OverwriteOnColl);
            await projectACD.SaveAsync();
            ConsoleMessage($"Imported '{convertedAOIrung_L5Xfilepath}' to '{acdPath}'.", "STATUS");

            // Change controller mode to program & verify.
            ConsoleMessage("START changing controller to PROGRAM...", "NEWSECTION");
            ChangeControllerMode_Async(commPath, "PROGRAM", projectACD).GetAwaiter().GetResult();
            if (ReadControllerMode_Async(commPath, projectACD).GetAwaiter().GetResult() == "PROGRAM")
                ConsoleMessage("SUCCESS changing controller to PROGRAM.", "STATUS", false);
            else
                ConsoleMessage("FAILURE changing controller to PROGRAM.", "ERROR", false);

            // Download project.
            ConsoleMessage("START downloading ACD file...", "NEWSECTION");
            DownloadProject_Async(commPath, projectACD).GetAwaiter().GetResult();
            ConsoleMessage("SUCCESS downloading ACD file.", "STATUS", false);

            // Change controller mode to run.
            ConsoleMessage("START changing controller to RUN...", "NEWSECTION");
            ChangeControllerMode_Async(commPath, "RUN", projectACD).GetAwaiter().GetResult();
            if (ReadControllerMode_Async(commPath, projectACD).GetAwaiter().GetResult() == "RUN")
                ConsoleMessage("SUCCESS changing controller to RUN.", "STATUS", false);
            else
                ConsoleMessage("FAILURE changing controller to RUN.", "ERROR", false);
            #endregion

            #region COMMENCE TEST: Iterate through each test case from the excel sheet. Set & check parameters per test case.
            ConsoleMessage($"START {aoiName} unit testing...", "NEWSECTION");

            int failureCondition = 0;  // This variable tracks the number of failed test cases or controller faults.

            //Console.WriteLine("current fullTagPath: " + aoiTagScope);
            ByteString udtoraoi_byteString = GetAOIbytestring_Sync(aoiTagScope, projectACD, OperationMode.Online);

            AOIParameter[] testParams = GetAOIParameters_FromL5X(convertedAOIrung_L5Xfilepath); // ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------[MODIFY FOR GITHUB DIRECTORY]
            testParams = GetAOIParameterValues(testParams, GetAOIbytestring_Sync(aoiTagScope, projectACD, OperationMode.Online), true);
            Print_AOIParameters(testParams, aoiName, false);

            // Test variables
            string[] AT_FaultType_TagValue;
            string[] AT_FaultCode_TagValue;
            bool breakUnitTestLoop = false;
            bool faultedState;
            bool breakOutputParameterLoop;

            // Iterate through and verify each test case (each column in the input excel sheet).
            int testCases = GetPopulatedColumnCount(inputExcel_UnitTestSetup_filePath, 18) - 1;
            for (int columnNumber = 4; columnNumber < (testCases + 4); columnNumber++)
            {
                int testNumber = columnNumber - 3;
                string testBannerContents = $"test {testNumber}/{testCases}";
                ConsoleMessage($"START {testBannerContents}...", "NEWSECTION", false);
                breakOutputParameterLoop = false;

                // Rotate through inputs
                Dictionary<string, string> currentColumn = GetExcelTestValues(inputExcel_UnitTestSetup_filePath, columnNumber);
                foreach (var kvp in currentColumn)
                {
                    if (GetAOIParameter(kvp.Key, "Usage", testParams) != "Output")
                    {
                        //SetSingleValue_UDTorAOI(kvp.Value, aoiTagScope, kvp.Key, OperationMode.Online, testParams, project).GetAwaiter().GetResult();
                        await SetSingleValue_UDTorAOI(kvp.Value, aoiTagScope, kvp.Key, OperationMode.Online, testParams, projectACD, true);
                    }

                    // Check if faulted
                    AT_FaultType_TagValue = GetTagValue_Sync("AT_FaultType", DataType.DINT, "Controller/Tags/Tag", projectACD);
                    AT_FaultCode_TagValue = GetTagValue_Sync("AT_FaultCode", DataType.DINT, "Controller/Tags/Tag", projectACD);
                    faultedState = (AT_FaultType_TagValue[1] != "0") || (AT_FaultCode_TagValue[1] != "0");

                    if (faultedState)
                    {
                        failureCondition++;

                        // Do faulted stuff
                        ConsoleMessage($"Controller faulted upon setting '{kvp.Key}' to '{kvp.Value}'. | Fault Type: '{AT_FaultType_TagValue[1]}' & Fault Code: '{AT_FaultCode_TagValue[1]}'.", "ERROR");

                        ConsoleMessage($"Attempting to clear fault. Setting '{kvp.Key}' to '0' & test if controller still faulted.", "ERROR");

                        await SetSingleValue_UDTorAOI("0", aoiTagScope, kvp.Key, OperationMode.Online, testParams, projectACD);

                        // Toggle reset to clear fault
                        SetTagValue_Sync("AT_ClearFault", "true", OperationMode.Online, DataType.BOOL, "Controller/Tags/Tag", projectACD);
                        SetTagValue_Sync("AT_ClearFault", "false", OperationMode.Online, DataType.BOOL, "Controller/Tags/Tag", projectACD);

                        AT_FaultType_TagValue = GetTagValue_Sync("AT_FaultType", DataType.DINT, "Controller/Tags/Tag", projectACD);
                        AT_FaultCode_TagValue = GetTagValue_Sync("AT_FaultCode", DataType.DINT, "Controller/Tags/Tag", projectACD);
                        faultedState = (AT_FaultType_TagValue[1] != "0") || (AT_FaultCode_TagValue[1] != "0");

                        if (faultedState)
                        {
                            ConsoleMessage("Controller still faulted. Ending Test.", "ERROR");
                            await LogixEchoMethods.DeleteChassis_Async(echoChassisName);
                            ConsoleMessage($"Deleted chassis '{echoChassisName}'.", "STATUS");
                            breakUnitTestLoop = true;
                            break;
                        }
                        else if (testNumber < testCases)
                        {
                            ConsoleMessage($"Fault cleared. Moving to next test case...", "SUCCESS");
                            breakOutputParameterLoop = true;
                            break;
                        }
                    }
                }

                if (breakUnitTestLoop)
                    break;


                // Rotate through outputs
                foreach (var kvp in currentColumn)
                {
                    if (breakOutputParameterLoop)
                        break;

                    if (GetAOIParameter(kvp.Key, "Usage", testParams) != "Input")
                    {
                        AOIParameter[] newTestParameters = GetAOIParameterValues(testParams, GetAOIbytestring_Sync(aoiTagScope, projectACD, OperationMode.Online), true);
                        string outputValue = GetAOIParameter(kvp.Key, "Value", newTestParameters);

                        failureCondition += TEST_CompareForExpectedValue(kvp.Key, kvp.Value, outputValue, true);
                    }

                }
            }
            #endregion

            #region END TEST: Print final test results & retain/delete generated files as specified in input excel sheet.
            // Based on the AOI unit test result, print a final result message in red or green.
            if (failureCondition > 0)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                ConsoleMessage($"{aoiName} Unit Test Final Result: FAIL | {failureCondition} Issues Encountered", "NEWSECTION", false);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Green;
                ConsoleMessage($"{aoiName} Unit Test Final Result: PASS", "NEWSECTION", false);
                Console.ForegroundColor = ConsoleColor.Gray;
            }

            // Based on the AOI Excel Worksheet for this AOI, keep or delete generated L5X files.
            ConsoleMessage("START retaining or deleting programmatically generated L5X files...", "NEWSECTION");
            if (!keepL5Xs)
            {
                File.Delete(L5Xpath);
                File.Delete(convertedAOIrung_L5Xfilepath);
                ConsoleMessage($"Deleted '{L5Xpath}'.", "STATUS");
                ConsoleMessage($"Deleted '{convertedAOIrung_L5Xfilepath}'.", "STATUS");
            }
            else
            {
                ConsoleMessage($"Retained '{L5Xpath}'.", "STATUS");
                ConsoleMessage($"Retained '{convertedAOIrung_L5Xfilepath}'.", "STATUS");
            }

            // Based on the AOI Excel Worksheet for this AOI, keep or delete the generated ACD file.
            ConsoleMessage("START retaining or deleting programmatically generated ACD file...", "NEWSECTION");
            if (!keepACD)
            {
                File.Delete(acdPath);
                ConsoleMessage($"Deleted '{acdPath}'.", "STATUS");
            }
            else
            {
                ConsoleMessage($"Retained '{acdPath}'.", "STATUS");
            }

            await projectACD.GoOfflineAsync(); // Testing is complete. Go offline with the emulated controller.
            #endregion
        }

        #region METHODS: L5X Manipulation
        private static AOIParameter[] GetAOIParameters_FromL5X(string l5xPath)
        {
            XDocument xDoc = XDocument.Load(l5xPath);
            int parameterCount = xDoc.Descendants("Parameters").FirstOrDefault().Elements().Count();
            AOIParameter[] returnDataPoints = new AOIParameter[parameterCount];
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

        public static string GetAOIParameter(string parameterName, string AOIParameterField, AOIParameter[] AOIParameter)
        {
            AOIParameterField = AOIParameterField.Trim().ToUpper();
            string returnString = "";
            for (int i = 0; i < AOIParameter.Length; i++)
            {
                if (AOIParameter[i].Name == parameterName)
                {
                    if (AOIParameterField == "NAME")
                    {
                        returnString = AOIParameter[i].Name;
                    }
                    if (AOIParameterField == "DATATYPE")
                    {
                        returnString = AOIParameter[i].DataType;
                    }
                    if (AOIParameterField == "USAGE")
                    {
                        returnString = AOIParameter[i].Usage;
                    }
                    if (AOIParameterField == "REQUIRED")
                    {
                        returnString = AOIParameter[i].Required.ToString();
                    }
                    if (AOIParameterField == "VISIBLE")
                    {
                        returnString = AOIParameter[i].Visible.ToString();
                    }
                    if (AOIParameterField == "VALUE")
                    {
                        returnString = AOIParameter[i].Value;
                    }
                    if (AOIParameterField == "BYTEPOSITION")
                    {
                        returnString = AOIParameter[i].BytePosition.ToString();
                    }
                    if (AOIParameterField == "BOOLPOSITION")
                    {
                        returnString = AOIParameter[i].BoolPosition.ToString();
                    }
                }
            }
            return returnString;
        }

        public static Dictionary<string, string> GetExcelTestValues(string filePath, int columnNumber)
        {
            Dictionary<string, string> returnDictionary = [];

            FileInfo existingFile = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int numberOfParameters = GetPopulatedRowCount(filePath, 2) - 6;
                for (int rowNumber = 18; rowNumber < (numberOfParameters + 18); rowNumber++)
                {
                    returnDictionary[worksheet.Cells[rowNumber, 2].Value?.ToString()!.Trim()!] = worksheet.Cells[rowNumber, columnNumber].Value?.ToString()!.Trim()!;
                }
            }

            return returnDictionary;
        }

        public static void ConvertAOItoRUNGxml(string xmlFilePath, string routineName, string programName, bool printOut)
        {
            string aoiName = GetAttributeValue(xmlFilePath, "AddOnInstructionDefinition", "Name", printOut);

            //Modify top half
            DeleteAttributeFromRoot(xmlFilePath, "TargetName", printOut);
            DeleteAttributeFromRoot(xmlFilePath, "TargetRevision", printOut);
            DeleteAttributeFromRoot(xmlFilePath, "TargetLastEdited", printOut);
            DeleteAttributeFromComplexElement(xmlFilePath, "AddOnInstructionDefinition", "Use", printOut);
            ChangeComplexElementAttribute(xmlFilePath, "RSLogix5000Content", "TargetCount", "1", printOut);
            ChangeComplexElementAttribute(xmlFilePath, "RSLogix5000Content", "TargetType", "Rung", printOut);

            // Create bottom half
            AddElementToComplexElement(xmlFilePath, "Controller", "Tags", printOut);
            AddAttributeToComplexElement(xmlFilePath, "Tags", "Use", "Context", printOut);

            AddElementToComplexElement(xmlFilePath, "Tags", "Tag", printOut);
            AddAttributeToComplexElement(xmlFilePath, "Tag", "Name", "AOI_" + aoiName, printOut);
            AddAttributeToComplexElement(xmlFilePath, "Tag", "TagType", "Base", printOut);
            AddAttributeToComplexElement(xmlFilePath, "Tag", "DataType", aoiName, printOut);
            AddAttributeToComplexElement(xmlFilePath, "Tag", "Constant", "false", printOut);
            AddAttributeToComplexElement(xmlFilePath, "Tag", "ExternalAccess", "Read/Write", printOut);
            AddAttributeToComplexElement(xmlFilePath, "Tag", "OpcUaAccess", "None", printOut);

            AddElementToComplexElement(xmlFilePath, "Tag", "Data", printOut);
            AddAttributeToComplexElement(xmlFilePath, "Data", "Format", "L5K", printOut);

            string cdataInfo_forData = GetCDATAfromXML_forData(xmlFilePath, printOut);
            AddCDATA(xmlFilePath, "Data", cdataInfo_forData, printOut);

            AddElementToComplexElement(xmlFilePath, "Tag", "Data", printOut);
            AddAttributeToComplexElement(xmlFilePath, "Data", "Format", "Decorated", printOut);

            AddElementToComplexElement(xmlFilePath, "Data", "Structure", printOut);
            AddAttributeToComplexElement(xmlFilePath, "Structure", "DataType", aoiName, printOut);

            List<Dictionary<string, string>> attributesList = GetDataValueMemberInfofromXML(xmlFilePath, printOut);
            AddComplexElementsWithAttributesToXml(xmlFilePath, attributesList, printOut);

            AddElementToComplexElement(xmlFilePath, "Controller", "Programs", printOut);
            AddAttributeToComplexElement(xmlFilePath, "Programs", "Use", "Context", printOut);

            AddElementToComplexElement(xmlFilePath, "Programs", "Program", printOut);
            AddAttributeToComplexElement(xmlFilePath, "Program", "Use", "Context", printOut);
            AddAttributeToComplexElement(xmlFilePath, "Program", "Name", programName, printOut);

            AddElementToComplexElement(xmlFilePath, "Program", "Routines", printOut);
            AddAttributeToComplexElement(xmlFilePath, "Routines", "Use", "Context", printOut);

            AddElementToComplexElement(xmlFilePath, "Routines", "Routine", printOut);
            AddAttributeToComplexElement(xmlFilePath, "Routine", "Use", "Context", printOut);
            AddAttributeToComplexElement(xmlFilePath, "Routine", "Name", routineName, printOut);

            AddElementToComplexElement(xmlFilePath, "Routine", "RLLContent", printOut);
            AddAttributeToComplexElement(xmlFilePath, "RLLContent", "Use", "Context", printOut);

            AddElementToComplexElement(xmlFilePath, "RLLContent", "Rung", printOut);
            AddAttributeToComplexElement(xmlFilePath, "Rung", "Use", "Target", printOut);
            AddAttributeToComplexElement(xmlFilePath, "Rung", "Number", "1", printOut);
            AddAttributeToComplexElement(xmlFilePath, "Rung", "Type", "N", printOut);

            AddElementToComplexElement(xmlFilePath, "Rung", "Comment", printOut);
            string cdataInfo_forComment = @"AUTOMATED TESTING | " + aoiName + @" AOI Unit Test
                    - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                    This a programmatically created rung with a populated instance of the AOI instruction added using the Logix Designer SDK.";
            AddCDATA(xmlFilePath, "Comment", cdataInfo_forComment, printOut);

            AddElementToComplexElement(xmlFilePath, "Rung", "Text", printOut);
            string cdataInfo_forText = GetCDATAfromXML_forText(xmlFilePath, printOut);
            AddCDATA(xmlFilePath, "Text", cdataInfo_forText, printOut);
        }

        public static string CopyXmlFile(string sourceFilePath, bool printOut)
        {
            // Check if the source file exists
            if (!File.Exists(sourceFilePath) && (printOut))
            {
                ConsoleMessage($"Source file '{sourceFilePath}' does not exist.", "ERROR");
            }

            // Get the directory and file name from the source file path
            string? directory = Path.GetDirectoryName(sourceFilePath);
            string fileName = Path.GetFileNameWithoutExtension(sourceFilePath);
            string extension = Path.GetExtension(sourceFilePath);

            // Construct the new file path for the copied file
            string newFileName = $"generated_{fileName}{extension}";
            string newFilePath = Path.Combine(directory, newFileName);

            // Copy the file
            File.Copy(sourceFilePath, newFilePath, overwrite: true);

            return newFilePath;
        }

        public static string? GetAttributeValue(string xmlFilePath, string complexElementName, string attributeName, bool printOut)
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
                else if (printOut)
                {
                    ConsoleMessage($"Attribute '{attributeName}' not found in element '{complexElementName}'.", "ERROR");
                }
            }
            else if (printOut)
            {
                ConsoleMessage($"The complex element '{complexElementName}' was not found in the XML file.", "ERROR");
            }

            return null; // Return null if attribute value is not found
        }

        public static void DeleteAttributeFromComplexElement(string xmlFilePath, string complexElementName, string attributeToDelete, bool printOut)
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

                        if (printOut)
                        {
                            ConsoleMessage($"Attribute '{attributeToDelete}' has been removed from the element '{complexElementName}'.", "STATUS");
                        }

                        // Save the changes back to the file
                        xdoc.Save(xmlFilePath);
                    }
                    else if (printOut)
                    {
                        ConsoleMessage($"Attribute '{attributeToDelete}' not found in element '{complexElementName}'.", "ERROR");
                    }
                }
                else if (printOut)
                {
                    ConsoleMessage($"The complex element '{complexElementName}' was not found in the XML file.", "ERROR");
                }
            }
            catch (Exception e)
            {
                ConsoleMessage(e.Message, "ERROR");
            }
        }

        public static void DeleteAttributeFromRoot(string xmlFilePath, string attributeToDelete, bool printOut)
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

                if (printOut)
                {
                    ConsoleMessage($"Attribute '{attributeToDelete}' has been removed from the root complex element '{complexElementName}'.", "STATUS");
                }

                // Save the changes back to the file
                xdoc.Save(xmlFilePath);
            }
            else if (printOut)
            {
                ConsoleMessage($"Attribute '{attributeToDelete}' not found in the root complex element '{complexElementName}'.", "ERROR");
            }
        }

        public static void ChangeComplexElementAttribute(string xmlFilePath, string complexElementName, string attributeName, string attributeValue, bool printOut)
        {
            // Load the XML document
            XDocument xdoc = XDocument.Load(xmlFilePath);

            // Find the complex element by name
            XElement complexElement = xdoc.Descendants(complexElementName).FirstOrDefault();

            if (complexElement != null)
            {
                // Add the attribute to the complex element
                complexElement.SetAttributeValue(attributeName, attributeValue);

                if (printOut)
                {
                    ConsoleMessage($"Attribute '{attributeName}' with value '{attributeValue}' has been added to the element '{complexElementName}'.", "STATUS");
                }

                // Save the changes back to the file
                xdoc.Save(xmlFilePath);
            }
            else if (printOut)
            {
                ConsoleMessage($"The complex element '{complexElementName}' was not found in the XML file.", "ERROR");
            }
        }

        public static void AddAttributeToComplexElement(string xmlFilePath, string complexElementName, string attributeName, string attributeValue, bool printOut)
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

                    if (printOut)
                    {
                        ConsoleMessage($"Attribute '{attributeName}' with value '{attributeValue}' has been added to the element '{complexElementName}'.", "STATUS");
                    }

                    // Save the changes back to the file
                    xdoc.Save(xmlFilePath);
                }
                else if (printOut)
                {
                    ConsoleMessage($"The complex element '{complexElementName}' was not found in the XML file.", "ERROR");
                }
            }
            catch (Exception e)
            {
                ConsoleMessage(e.Message, "ERROR");
            }
        }

        public static void AddElementToComplexElement(string xmlFilePath, string complexElementName, string newElementName, bool printOut)
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

                    if (printOut)
                    {
                        ConsoleMessage($"Element '{newElementName}' has been added to the complex element '{complexElementName}'.", "STATUS");
                    }

                    // Save the changes back to the file
                    xdoc.Save(xmlFilePath);
                }
                else if (printOut)
                {
                    ConsoleMessage($"The complex element '{complexElementName}' was not found in the XML file.", "ERROR");
                }
            }
            catch (Exception e)
            {
                ConsoleMessage(e.Message, "ERROR");
            }
        }

        /// <summary>
        /// Create a new CDATA element to the last or default instance of a specified complex element.
        /// </summary>
        /// <param name="xmlFilePath">The AOI L5X file path.</param>
        /// <param name="complexElementName">The name of the complex element to which the CDATA element will be added.</param>
        /// <param name="cdataContent">The contents of the CDATA element.</param>
        public static void AddCDATA(string xmlFilePath, string complexElementName, string cdataContent, bool printOut)
        {
            try
            {
                // Load the XML document.
                XDocument xdoc = XDocument.Load(xmlFilePath);

                // Find the complex element by name.
                XElement complexElement = xdoc.Descendants(complexElementName).LastOrDefault();

                if (complexElement != null)
                {
                    // Create a new CDATA section and add it to the complex element.
                    XCData cdataSection = new XCData(cdataContent);
                    complexElement.Add(cdataSection);

                    if (printOut)
                    {
                        ConsoleMessage($"A new CDATA section has been created and added to the element '{complexElementName}'.", "STATUS");
                    }

                    // Save the changes back to the file.
                    xdoc.Save(xmlFilePath);
                }
                else if (printOut)
                {
                    ConsoleMessage($"The complex element '{complexElementName}' was not found in the XML file.", "ERROR");
                }
            }
            catch (Exception e)
            {
                ConsoleMessage(e.Message, "ERROR");
            }
        }

        /// <summary>
        /// Programmatically get the CDATA contents for the Data complex element.
        /// </summary>
        /// <param name="xmlFilePath">The AOI L5X file path.</param>
        /// <returns>A string of formatted CDATA contents.</returns>
        public static string GetCDATAfromXML_forData(string xmlFilePath, bool printOut)
        {
            try
            {
                // Load the XML document
                XDocument doc = XDocument.Load(xmlFilePath);

                // Get a list filtered to contain only CDATA information from nonboolean "Parameter" elements
                var parameterElements = doc
                    .Descendants("Parameters")
                    .Elements("Parameter")
                    .Where(param => param.Attribute("DataType")?.Value != "BOOL")
                    .Descendants("DefaultData")
                    .Where(defaultData => defaultData.FirstNode is XCData)
                    .Select(defaultData => ((XCData)defaultData.FirstNode).Value.Trim())
                    .ToList();

                // Join all parameterElements list elements into a single string, with each element separated by a comma without spaces.
                string joined_pCDATA = string.Join(",", parameterElements);

                // Get a list filtered to contain only CDATA information from nonboolean "LocalTag" elements
                var localtagElements = doc
                    .Descendants("LocalTags")
                    .Elements("LocalTag")
                    .Where(param => param.Attribute("DataType")?.Value != "BOOL")
                    .Descendants("DefaultData")
                    .Where(defaultData => defaultData.FirstNode is XCData)
                    .Select(defaultData => ((XCData)defaultData.FirstNode).Value.Trim())
                    .ToList();

                // Join all localtagElements list elements into a single string, with each element separated by a comma without spaces.
                string joined_ltCDATA = string.Join(",", localtagElements);

                // Create the final formatted string to be used as CDATA content information (in the Data complex element of L5X).
                string returnString = "[1," + joined_pCDATA + "," + joined_ltCDATA + "]";

                if (printOut)
                {
                    ConsoleMessage($"CDATA contents: {returnString}", "STATUS");
                }

                return returnString;
            }
            catch (Exception e)
            {
                ConsoleMessage(e.Message, "ERROR");
                return e.Message;
            }
        }

        /// <summary>
        /// Programmatically get the CDATA contents for the Text complex element.<br/>
        /// This method is where the information needed for a new AOI tag is programmatically gathered and formatted.
        /// </summary>
        /// <param name="xmlFilePath">The AOI L5X file path.</param>
        /// <param name="printOut">A boolean that, if true, prints updates to the console.</param>
        /// <returns>A string of formatted CDATA contents.</returns>
        public static string GetCDATAfromXML_forText(string xmlFilePath, bool printOut)
        {
            // The name of the AOI being testing.
            string? aoiName = GetAttributeValue(xmlFilePath, "AddOnInstructionDefinition", "Name", printOut);

            // Initialize the StringBuilder that will contain the AOI parameter tag names.
            StringBuilder aoiTagParameterNames = new();

            try
            {
                // Load the XML document.
                XDocument doc = XDocument.Load(xmlFilePath);

                // Find all "Parameter" elements.
                var parameterElements = doc.Descendants("Parameter");

                // Cycle through each AOI parameter and add it to the list if it is a required parameter.
                foreach (var param in parameterElements)
                {
                    XAttribute? nameAttribute = param.Attribute("Name");
                    string? requiredAttributeValue = param.Attribute("Required")?.Value;
                    if ((nameAttribute != null) && (requiredAttributeValue == "true"))
                    {
                        aoiTagParameterNames.Append($",AOI_{aoiName}.{nameAttribute.Value}");
                    }
                }

                string returnString = $"{aoiName}(AOI_{aoiName}{aoiTagParameterNames});";

                if (printOut)
                {
                    ConsoleMessage($"CDATA contents: {returnString}", "STATUS");
                }

                return returnString;
            }
            catch (Exception e)
            {
                ConsoleMessage(e.Message, "ERROR");
                return e.Message;
            }
        }

        /// <summary>
        /// Get all the attribute names and values for each parameter in an AOI L5X file.
        /// </summary>
        /// <param name="xmlFilePath">The AOI L5X file path.</param>
        /// <param name="printOut">A boolean that, if true, prints updates to the console.</param>
        /// <returns>A list of dictionaries for each AOI parameter's attributes.</returns>
        public static List<Dictionary<string, string>> GetDataValueMemberInfofromXML(string xmlFilePath, bool printOut)
        {
            List<Dictionary<string, string>> return_attributeList = new List<Dictionary<string, string>>();

            try
            {
                // Load the XML document
                XDocument doc = XDocument.Load(xmlFilePath);

                // Cycle through each "Parameter" element in the L5X file.
                foreach (var parameterElement in doc.Descendants("Parameter"))
                {
                    Dictionary<string, string> attributes = new Dictionary<string, string>
                    {
                        { "Name", parameterElement.Attribute("Name").Value },
                        { "DataType", parameterElement.Attribute("DataType").Value },
                        { "Radix", parameterElement.Attribute("Radix").Value }
                    };

                    // Store the new dictionary containing attributes for a single AOI parameter.
                    return_attributeList.Add(attributes);
                }
            }
            catch (Exception e)
            {
                ConsoleMessage(e.Message, "ERROR");
            }

            return return_attributeList;
        }

        /// <summary>
        /// For each AOI parameter, add the element "DataValueMember" with its attributes to the L5X complex element "Structure".<br/>
        /// This method creates XML children needed to create an AOI tag in the L5X file.
        /// </summary>
        /// <param name="xmlFilePath">The AOI L5X file path.</param>
        /// <param name="attributesList">A list of dictionaries for each AOI parameter's attributes.</param>
        /// <param name="printOut">A boolean that, if true, prints updates to the console.</param>
        public static void AddComplexElementsWithAttributesToXml(string xmlFilePath, List<Dictionary<string, string>> attributesList, bool printOut)
        {
            try
            {
                // Add the new attributes
                foreach (var attributes in attributesList)
                {
                    // Add a new element "DataValueMember" to complex element "Structure" for each AOI parameter.
                    AddElementToComplexElement(xmlFilePath, "Structure", "DataValueMember", printOut);

                    // Add the "Name" attribute and its value for the current AOI parameter.
                    AddAttributeToComplexElement(xmlFilePath, "DataValueMember", "Name", attributes["Name"], printOut);

                    // Add the "DataType" attribute and its value for the current AOI parameter.
                    AddAttributeToComplexElement(xmlFilePath, "DataValueMember", "DataType", attributes["DataType"], printOut);

                    // Add the "Radix" attribute and its value for the current AOI parameter.
                    // Note: BOOL datatype parameters don't have a "Radix" attribute and are therefore skipped.
                    if (attributes["DataType"] != "BOOL")
                    {
                        AddAttributeToComplexElement(xmlFilePath, "DataValueMember", "Radix", attributes["Radix"], printOut);
                    }

                    // Add the "Value" attribute and its value for the current AOI parameter.
                    // Note: For AOIs, the only BOOL parameter with a value of 1 is "EnableIn".
                    // Note: For REAL datatype parameters, their intial zero value has the notation "0.0". All else is "0".
                    if (attributes["Name"] == "EnableIn")
                    {
                        AddAttributeToComplexElement(xmlFilePath, "DataValueMember", "Value", "1", printOut);
                    }
                    else if (attributes["DataType"] == "REAL")
                    {
                        AddAttributeToComplexElement(xmlFilePath, "DataValueMember", "Value", "0.0", printOut);
                    }
                    else
                    {
                        AddAttributeToComplexElement(xmlFilePath, "DataValueMember", "Value", "0", printOut);
                    }
                }
                if (printOut)
                {
                    ConsoleMessage("Complex elements added.", "STATUS");
                }
            }
            catch (Exception e)
            {
                ConsoleMessage(e.Message, "ERROR");
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

        private static void Print_AOIParameters(AOIParameter[] dataPointsArray, string aoiName, bool printPosition)
        {
            int arraySize = dataPointsArray.Length;
            ConsoleMessage($"Print {aoiName} parameters.", "STATUS");

            for (int i = 0; i < arraySize; i++)
            {
                if (i < arraySize - 1)
                {
                    Console.Write("           ├── ");
                }
                else
                {
                    Console.Write("           └── ");
                }

                if (printPosition == true)
                {
                    Console.WriteLine($"Name: {dataPointsArray[i].Name,-20} | Data Type: {dataPointsArray[i].DataType,-9} | " +
                        $"Scope: {dataPointsArray[i].Usage,-7} | Required: {dataPointsArray[i].Required,-5} | " +
                        $"Visible: {dataPointsArray[i].Visible,-5} |  Value: {dataPointsArray[i].Value,-20} | " +
                        $"Byte Position: {dataPointsArray[i].BytePosition,-3} | Bool Position: {dataPointsArray[i].BoolPosition}");
                }
                else if (printPosition == false)
                {
                    Console.WriteLine($"Name: {dataPointsArray[i].Name,-20} | Data Type: {dataPointsArray[i].DataType,-9} | " +
                        $"Scope: {dataPointsArray[i].Usage,-7} | Required: {dataPointsArray[i].Required,-5} | " +
                        $"Visible: {dataPointsArray[i].Visible,-5} |  Value: {dataPointsArray[i].Value,-20}");
                }
            }
        }


        #endregion

        #region METHODS: formatting text file
        /// <summary>
        /// Standardized method to print messages of varying categories to the console.
        /// </summary>
        /// <param name="messageContents">The contents of the message to be written to the console.</param>
        /// <param name="messageCategory">The name of the message category.</param>
        /// <param name="newLineForSection">
        /// A boolean input that determines whether to write the characters '---' to the console.<br/>
        /// (Note: only applicable if messageCateogry = "NEWSECTION")
        /// </param>
        public static void ConsoleMessage(string messageContents, string messageCategory = "", bool newLineForSection = true)
        {
            messageCategory = messageCategory.ToUpper().Trim();

            if ((messageCategory == "ERROR") || (messageCategory == "FAILURE") || (messageCategory == "FAIL"))
            {
                messageCategory = messageCategory.PadLeft(9, ' ') + ": ";
                Console.ForegroundColor = ConsoleColor.Red;
                Console.Write(messageCategory);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
            else if ((messageCategory == "SUCCESS") || (messageCategory == "PASS"))
            {
                messageCategory = messageCategory.PadLeft(9, ' ') + ": ";
                Console.ForegroundColor = ConsoleColor.Green;
                Console.Write(messageCategory);
                Console.ForegroundColor = ConsoleColor.Gray;
            }
            else if (messageCategory == "STATUS")
            {
                messageCategory = messageCategory.PadLeft(9, ' ') + ": ";
                Console.Write(messageCategory);
            }
            else if (messageCategory == "NEWSECTION")
            {
                if (newLineForSection)
                {
                    Console.Write($"---\n[{DateTime.Now.ToString("HH:mm:ss")}] ");
                }
                else
                {
                    Console.Write($"[{DateTime.Now.ToString("HH:mm:ss")}] ");
                }
            }
            else
            {
                messageCategory = messageCategory.PadLeft(9, ' ') + "  ";
                Console.Write(messageCategory);
            }

            messageContents = WrapText(messageContents, 11, 120);
            Console.WriteLine(messageContents);
        }

        /// <summary>
        /// Modify the input string to wrap the text to the next line after a certain length.<br/>
        /// The input string is seperated per word and then each line is incrementally added to per word.<br/>
        /// Start a new line when the character count of a line exceeds 125.
        /// </summary>
        /// <param name="inputString">The input string to be wrapped.</param>
        /// <param name="indentLength">An integer that defines the length of the characters in the indent starting each new line.</param>
        /// <param name="lineLimit">An integer that defines the maximum number of characters per line before a new line is created.</param>
        /// <returns>A modified string that wraps after a specified length of characters.</returns>
        private static string WrapText(string inputString, int indentLength, int lineLimit)
        {
            // Variables containing formatting information:
            StringBuilder newSentence = new StringBuilder(); // The properly formatted string to be returned.
            string[] words = inputString.Split(' ');         // An array where each element contains each word in an input string. 
            string indent = new string(' ', indentLength);   // An empty string to be used for indenting.
            string line = "";                                // The variable that will be modified and appended to the returned StringBuilder for each line.

            // Variables informing formatting logic:
            bool newLongWord = true;
            int numberOfNewLines = 0;
            int numberOfLongWords = 0;
            int indentedLineLimit = lineLimit - indentLength;

            // Cycle through each word in the input string.
            foreach (string word in words)
            {
                // The word (short or long) has any excess spaces removed. 
                string trimmedWord = word.Trim();

                // Required for "Long Word Splitting" Logic: This variable is used to wrap long words at the indentLength specified with indenting.
                int partLength = lineLimit - (indentLength + line.Length);

                // Required for "Long Word Splitting" Logic: The # of long words determine how a long word component is added to the console.
                // Long words for this method are defined as words that are above the character number of line limit minus indent length.
                if (trimmedWord.Length >= partLength)
                    numberOfLongWords++;

                // "Long Word Splitting" Logic
                // If the word is longer than the line limit # of characters, split it & wrap to the next line keeping indents.
                while (((line + trimmedWord).Length >= indentedLineLimit) && (trimmedWord.Length >= indentedLineLimit))
                {

                    string part = trimmedWord.Substring(0, partLength); // A peice of the long word to add to the existing line. 
                    trimmedWord = trimmedWord.Substring(partLength);    // The long word part is removed from trimmedWord.

                    // Long Word Scenario 1: This should only ever run once the first time a long word goes through the while loop.
                    if (((numberOfLongWords == 1) || (numberOfNewLines == 0)) && (newLongWord))
                    {
                        newSentence.AppendLine(line + part);            // Add line & part to return string. No indent b/c either the long word starts the message
                                                                        // or because the long word part gets added to the current line that already has words.
                        line = "";                                      // Reset the line string.
                        numberOfNewLines++;                             // Count up for number of new lines.
                        newLongWord = false;                            // Lock this if statement (Scenario 1) from being run again.
                        partLength = indentedLineLimit;
                    }
                    // Long Word Scenario 2: All other subsequent lines with long words (or long word components) need to be indented.
                    else
                    {
                        newSentence.AppendLine(indent + line + part);   // Add indented current line with part. (Note: line could be 0 chars if part is long enough.)
                        line = "";                                      // Reset the line string.
                        numberOfNewLines++;                             // Count up for number of new lines.
                        partLength = indentedLineLimit;
                    }
                }

                // Required for "Long Word Splitting" Logic: Determines how a long word component is added to the console.
                newLongWord = true;

                // "Adding Line" Logic
                // Check if the current line plus the next word (or the remaining part of a long word) exceeds the line limit (accounting for indenting).
                if ((line + trimmedWord).Length >= indentedLineLimit)
                {
                    // Line Scenario 1: If not the first line, add indented current line to return string. 
                    if (numberOfNewLines > 0)
                    {
                        newSentence.AppendLine(indent + line.TrimEnd());
                    }
                    // Line Scenario 2: If the first line, add the current line without indents to return string.
                    else
                    {
                        newSentence.AppendLine(line.TrimEnd());
                    }
                    line = "";           // Reset the line string.
                    numberOfNewLines++;  // Count up for number of new lines.
                }

                // Add the word (or the remaining part of a long word) to the current line.
                line += trimmedWord + " ";
            }

            // Same as "Adding Line" Logic where the line contents are the remaining input string contents under the line limit. 
            if (line.Length > 0)
            {
                if (numberOfNewLines > 0)
                    newSentence.Append(indent + line.TrimEnd());
                else
                    newSentence.Append(line.TrimEnd());
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
            ConsoleMessage($"'{folderPath}' set to retain {keepCount} test files.", "STATUS");
            string[] all_files = Directory.GetFiles(folderPath);
            var orderedFiles = all_files.Select(f => new FileInfo(f)).OrderBy(f => f.CreationTime).ToList();
            if (orderedFiles.Count > keepCount)
            {
                for (int i = 0; i < (orderedFiles.Count - keepCount); i++)
                {
                    FileInfo deleteThisFile = orderedFiles[i];
                    deleteThisFile.Delete();
                    ConsoleMessage($"Deleted '{deleteThisFile.FullName}'.", "STATUS");
                }
            }
            else
            {
                ConsoleMessage($"No files needed to be deleted (currently {orderedFiles.Count} test files).", "STATUS");
            }
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

            ConsoleMessage($"File created at '{excelFilePath}'.", "STATUS");
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
        private static async Task<string[]> GetTagValue_Async(string tagName, DataType type, string tagPath, LogixProject project, bool printout)
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
                {
                    ConsoleMessage($"Data type '{type}' not supported.", "ERROR");
                }
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage($"Could not get tag '{tagName}'.", "ERROR");
                Console.WriteLine(e.Message);
            }

            if (printout)
            {
                string online_message = $"online value: {return_array[1]}";
                string offline_message = $"offline value: {return_array[2]}";
                ConsoleMessage($"{tagName.PadRight(40, ' ')}{online_message.PadRight(35, ' ')}{offline_message.PadRight(35, ' ')}", "SUCCESS");
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
        private static string[] GetTagValue_Sync(string tagName, DataType type, string tagPath, LogixProject project, bool printout = false)
        {
            var task = GetTagValue_Async(tagName, type, tagPath, project, printout);
            task.Wait();
            return task.Result;
        }

        /// <summary>
        /// Asynchronously set either the online or offline value of a basic data type tag.<br/>
        /// (basic data types handled: boolean, single integer, integer, double integer, long integer, real, string)
        /// </summary>
        /// <param name="tagName">The name of the tag whose value will be set.</param>
        /// <param name="newTagValue">The value of the tag that will be set.</param>
        /// <param name="mode">This specifies whether the 'Online' or 'Offline' value of the tag is the one to set.</param>
        /// <param name="type">The data type of the tag whose value will be set.</param>
        /// <param name="tagPath">
        /// The tag path specifying the tag's scope and location in the Studio 5000 Logix Designer project.<br/>
        /// The tag path is based on the XML filetype (L5X) encapsulation of elements.
        /// </param>
        /// <param name="project">An instance of the LogixProject class.</param>
        /// <param name="printout">A boolean that, if True, prints the online and offline values to the console.</param>
        /// <returns>A Task that will set the online or offline value of a basic data type tag.</returns>
        private static async Task SetTagValue_Async(string tagName, string newTagValue, OperationMode mode, DataType type, string tagPath, LogixProject project, bool printout)
        {
            tagPath = tagPath + $"[@Name='{tagName}']";
            string[] old_tag_values = await GetTagValue_Async(tagName, type, tagPath, project, false);
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
                        ConsoleMessage($"Data type '{type}' not supported.", "ERROR");
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
                        ConsoleMessage($"Data type '{type}' not supported.", "ERROR");
                    old_tag_value = old_tag_values[2];
                }
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage("Unable to set tag value.", "ERROR");
                Console.WriteLine(e.Message);
            }

            try
            {
                await project.SaveAsync();
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage("Unable to save project", "ERROR");
                Console.WriteLine(e.Message);
            }

            if (printout)
            {
                string new_tag_value_string = Convert.ToString(newTagValue);
                if ((new_tag_value_string == "1") && (type == DataType.BOOL)) { new_tag_value_string = "True"; }
                if ((new_tag_value_string == "0") && (type == DataType.BOOL)) { new_tag_value_string = "False"; }

                string outputMessage = $"{old_tag_values[0],-40} {old_tag_value,20} -> {new_tag_value_string,-20}";
                ConsoleMessage(outputMessage, "STATUS");
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
        private static void SetTagValue_Sync(string tagName, string newTagValue, OperationMode mode, DataType type, string tagPath, LogixProject project, bool printout = false)
        {
            var task = SetTagValue_Async(tagName, newTagValue, mode, type, tagPath, project, printout);
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
        private static async Task<ByteString> GetAOIbytestring_Async(string fullTagPath, LogixProject project, OperationMode online_or_offline)
        {
            ByteString returnByteStringArray = ByteString.Empty;

            if (online_or_offline == OperationMode.Online)
                returnByteStringArray = await project.GetTagValueAsync(fullTagPath, OperationMode.Online, DataType.BYTE_ARRAY);
            else if (online_or_offline == OperationMode.Offline)
                returnByteStringArray = await project.GetTagValueAsync(fullTagPath, OperationMode.Offline, DataType.BYTE_ARRAY);
            else
                ConsoleMessage($"The input '{online_or_offline}' is not a valid selection. Input either 'OperationMode.Online' or 'OperationMode.Offline'.", "ERROR");

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
        private static ByteString GetAOIbytestring_Sync(string tagPath, LogixProject project, OperationMode online_or_offline)
        {
            var task = GetAOIbytestring_Async(tagPath, project, online_or_offline);
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

        private static async Task SetSingleValue_UDTorAOI(string newParameterValue, string aoiTagPath, string parameterName, OperationMode mode, AOIParameter[] input_TagDataArray, LogixProject project, bool printout = false)
        {
            ByteString input_ByteString = GetAOIbytestring_Sync(aoiTagPath, project, mode);
            byte[] new_byteArray = input_ByteString.ToByteArray();
            int inputArraySize = input_TagDataArray.Length;
            string oldParameterValue = "";

            for (int j = 0; j < inputArraySize; j++)
            {
                // Search the TagData[] array to get the associated newTagValue data needed.
                if (input_TagDataArray[j].Name == parameterName)
                {
                    DataType dataType = GetDataType(input_TagDataArray[j].DataType);
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
                        ConsoleMessage($"Data type '{dataType}' not supported.", "ERROR");
                    }
                }
            }

            await project.SetTagValueAsync(aoiTagPath, mode, new_byteArray, DataType.BYTE_ARRAY);
            string setParamIntro = $"Change {parameterName} value:".PadRight(40, ' ');

            if (printout)
                ConsoleMessage($"{setParamIntro} {oldParameterValue,20} -> {newParameterValue,-20}", "STATUS");
        }

        private static DataType GetDataType(string dataType)
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
                    ConsoleMessage($"Data type '{dataType}' not supported.", "ERROR");
                    throw new ArgumentException();
            }
            return type;
        }

        private static AOIParameter[] GetAOIParameterValues(AOIParameter[] input_TagDataArray, ByteString input_AOIorUDT_ByteString, bool printout)
        {
            // initialize values needed for this method
            AOIParameter[] output_TagDataArray = input_TagDataArray;
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
            mode = mode.ToUpper().Trim();

            var requestedControllerMode = default(LogixProject.RequestedControllerMode);
            if (mode == "PROGRAM")
            {
                requestedControllerMode = LogixProject.RequestedControllerMode.Program;
            }
            else if (mode == "RUN")
            {
                requestedControllerMode = LogixProject.RequestedControllerMode.Run;
            }
            else if (mode == "TEST")
            {
                requestedControllerMode = LogixProject.RequestedControllerMode.Test;
            }
            else
            {
                ConsoleMessage($"Mode '{mode}' is not supported.", "ERROR");
            }

            try
            {
                await project.SetCommunicationsPathAsync(commPath);
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage($"Unable to set communication path to '{commPath}'.", "ERROR");
                Console.WriteLine(e.Message);
            }

            try
            {
                await project.ChangeControllerModeAsync(requestedControllerMode);
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage($"Unable to set mode. Requested mode was '{mode}'.", "ERROR");
                Console.WriteLine(e.Message);
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
            catch (LogixSdkException e)
            {
                ConsoleMessage($"Unable to set communication path to '{commPath}'.", "ERROR");
                Console.WriteLine(e.Message);
            }

            try
            {
                LogixProject.ControllerMode controllerMode = await project.ReadControllerModeAsync();
                if (controllerMode != LogixProject.ControllerMode.Program)
                {
                    ConsoleMessage($"Controller mode is {controllerMode}. Downloading is possible only if the controller is in 'Program' mode.", "ERROR");
                }
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage("Unable to read ControllerMode.", "ERROR");
                Console.WriteLine(e.Message);
            }

            try
            {
                await project.DownloadAsync();
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage("Unable to download.", "ERROR");
                Console.WriteLine(e.Message);
            }

            // Download modifies the project.
            // Without saving, if used file will be opened again, commands which need correlation
            // between program in the controller and opened project like LoadImageFromSDCard or StoreImageOnSDCard
            // may not be able to succeed because project in the controller won't match opened project.
            try
            {
                await project.SaveAsync();
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage("Unable to save project.", "ERROR");
                Console.WriteLine(e.Message);
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
            catch (LogixSdkException e)
            {
                ConsoleMessage($"Unable to set commpath to '{commPath}'.", "ERROR");
                Console.WriteLine(e.Message);
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
                        throw new ArgumentOutOfRangeException("Controller mode is unrecognized.");
                }
            }
            catch (LogixSdkException e)
            {
                ConsoleMessage("Unable to read controller mode.", "ERROR");
                Console.WriteLine(e.Message);
            }

            return "";
        }
        #endregion

        #region METHODS: TEST & helper methods
        /// <summary>
        /// A test to compare the expected and actual values of a tag.
        /// </summary>
        /// <param name="tagName">The name of the tag to be tested.</param>
        /// <param name="expectedValue">The expected value of the tag under test.</param>
        /// <param name="actualValue">The actual value of the tag under test.</param>
        /// <returns>Return an integer value 1 for test failure and an integer value 0 for test success.</returns>
        /// <remarks>
        /// The integer output is added to an integer that tracks the total number of failures.<br/>
        /// At the end of all testing, the overall SUCCESS/FAILURE of this CI/CD test stage is determined whether its value is greater than 0.
        /// </remarks>
        private static int TEST_CompareForExpectedValue(string tagName, string expectedValue, string actualValue, bool printOut)
        {
            if (expectedValue != actualValue)
            {
                if (printOut)
                    ConsoleMessage($"{tagName} expected value '{expectedValue}' & actual value '{actualValue}' NOT equal.", "FAIL");

                return 1;
            }
            else
            {
                if (printOut)
                    ConsoleMessage($"{tagName} expected value '{expectedValue}' & actual value '{actualValue}' EQUAL.", "PASS");

                return 0;
            }
        }

        public static void EnsureFolderExists(string folderPath)
        {
            // Check if the folder exists
            if (!Directory.Exists(folderPath))
            {
                // Create the folder if it doesn't exist
                Directory.CreateDirectory(folderPath);
                ConsoleMessage($"Folder created at '{folderPath}'.", "STATUS");
            }
            else
            {
                ConsoleMessage($"Folder already exists at '{folderPath}'.", "STATUS");
            }
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

        public static uint ConvertStringToUint(string inputString)
        {
            double doubleValue;
            uint result = 0;

            if (double.TryParse(inputString, out doubleValue))
            {
                result = (uint)doubleValue;
            }
            else
            {
                ConsoleMessage("Conversion failed. The input string is not a valid double.", "ERROR");
            }

            return result;
        }
        #endregion
    }
}
