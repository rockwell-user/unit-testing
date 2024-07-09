// ---------------------------------------------------------------------------------------------------------------------------------------------------------------
//
// FileName: Echo_Program.cs
// FileType: Visual C# Source File
// Author : Rockwell Automation
// Created : 2024
// Description : This script provides supporting methods to set up an emulated controller using the Factory Talk Logix Echo SDK.
//
// ---------------------------------------------------------------------------------------------------------------------------------------------------------------

using RockwellAutomation.FactoryTalkLogixEcho.Api.Client;
using RockwellAutomation.FactoryTalkLogixEcho.Api.Interfaces;
using System.Globalization;

namespace LogixEcho_ClassLibrary
{
    /// <summary>
    /// Class containing Factory Talk Logix Echo SDK methods needed for CI/CD test stage execution.
    /// </summary>
    public class LogixEchoMethods
    {
        /// <summary>
        /// Script that sets up an emulated controller for CI/CD software in the loop (SIL) testing.<br/>
        /// If no emulated controller based on the ACD file path yet exists, create one, and then return the communication path.<br/>
        /// If an emulated controller based on the ACD file path exists, only return the communication path.
        /// </summary>
        /// <param name="acdFilePath">The file path pointing to the ACD project used for testing.</param>
        /// <returns>A string containing the communication path of the emulated controller that the ACD project file will go online with during testing.</returns>
        public static async Task<string> Main(string acdFilePath, string chassis_name, string controller_name)
        {
            var serviceClient = ClientFactory.GetServiceApiClientV2("CLIENT_TestStage_CICDExample");
            serviceClient.Culture = new CultureInfo("en-US");

            // Check if an emulated controller exists within an emulated chassis. If not, run through the if statement contents to create one.

            //if (CheckCurrentChassis_Sync("CICDtest_chassis", "CICD_test", serviceClient) == false)
            if (CheckCurrentChassis_Sync(chassis_name, controller_name, serviceClient) == false)
            {
                // Set up emulated chassis information.
                var chassisUpdate = new ChassisUpdate
                {
                    Name = chassis_name,
                    Description = "Test chassis for CI/CD demonstration."
                };
                ChassisData chassisCICD = await serviceClient.CreateChassis(chassisUpdate);

                // Set up emulated controller information.
                using (var fileHandle = await serviceClient.SendFile(acdFilePath))
                {
                    var controllerUpdate = await serviceClient.GetControllerInfoFromAcd(fileHandle);
                    controllerUpdate.ChassisGuid = chassisCICD.ChassisGuid;
                    var controllerData = await serviceClient.CreateController(controllerUpdate);
                }
            }
            // Get emulated controller information.
            string[] testControllerInfo = await Get_ControllerInfo_Async(chassis_name, controller_name, serviceClient);
            string commPath = @"EmulateEthernet\" + testControllerInfo[1];
            Console.WriteLine($"SUCCESS: project communication path specified is \"{commPath}\"");
            return commPath;
        }
        #region METHODS: setting up Logix Echo emulated controller
        /// <summary>
        /// Asynchronously check to see if a specific controller exists in a specific chassis.
        /// </summary>
        /// <param name="chassisName">The name of the emulated chassis to check the emulated controler in.</param>
        /// <param name="controllerName">The name of the emulated controller to check.</param>
        /// <param name="serviceClient">The Factory Talk Logix Echo interface.</param>
        /// <returns>A Task that returns a boolean value 'True' if the emulated controller already exists and a 'False' if it does not.</returns>
        public static async Task<bool> CheckCurrentChassis_Async(string chassisName, string controllerName, IServiceApiClientV2 serviceClient)
        {
            var chassisList = (await serviceClient.ListChassis()).ToList();
            for (int i = 0; i < chassisList.Count; i++)
            {
                if (chassisList[i].Name == chassisName)
                {
                    var chassisGuid = chassisList[i].ChassisGuid;
                    var controllerList = (await serviceClient.ListControllers(chassisGuid)).ToList();
                    for (int j = 0; j < controllerList.Count; j++)
                    {
                        if (controllerList[j].ControllerName == controllerName)
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Run the CheckCurrentChassisAsync method synchronously.<br/>
        /// Check to see if a specific controller exists in a specific chassis.
        /// </summary>
        /// <param name="chassisName">The name of the emulated chassis to check the emulated controler in.</param>
        /// <param name="controllerName">The name of the emulated controller to check.</param>
        /// <param name="serviceClient">The Factory Talk Logix Echo interface.</param>
        /// <returns>A boolean value 'True' if the emulated controller already exists and a 'False' if it does not.</returns>
        public static bool CheckCurrentChassis_Sync(string chassisName, string controllerName, IServiceApiClientV2 serviceClient)
        {
            var task = CheckCurrentChassis_Async(chassisName, controllerName, serviceClient);
            task.Wait();
            return task.Result;
        }

        /// <summary>
        /// Asynchronously get the emulated controller name, IP address, and project file path.
        /// </summary>
        /// <param name="chassisName">The emulated chassis to the emulatedcontroller information from.</param>
        /// <param name="controllerName">The emulated controller name.</param>
        /// <param name="serviceClient">The Factory Talk Logix Echo interface.</param>
        /// <returns>
        /// A Task that returns a string array containing controller information:<br/>
        /// return_array[0] = controller name<br/>
        /// return_array[1] = controller IP address<br/>
        /// return_array[2] = controller project file path
        /// </returns>
        public static async Task<string[]> Get_ControllerInfo_Async(string chassisName, string controllerName, IServiceApiClientV2 serviceClient)
        {
            string[] return_array = new string[3];
            var chassisList = (await serviceClient.ListChassis()).ToList();
            for (int i = 0; i < chassisList.Count; i++)
            {
                if (chassisList[i].Name == chassisName)
                {
                    var chassisGuid = chassisList[i].ChassisGuid;
                    var controllerList = (await serviceClient.ListControllers(chassisGuid)).ToList();
                    for (int j = 0; j < controllerList.Count; j++)
                    {
                        if (controllerList[j].ControllerName == controllerName)
                        {
                            return_array[0] = controllerList[j].ControllerName;
                            return_array[1] = controllerList[j].IPConfigurationData.Address.ToString() ?? "";
                            return_array[2] = controllerList[j].ProjectPath;
                        }
                    }
                }
            }
            return return_array;
        }
        #endregion
    }
}
