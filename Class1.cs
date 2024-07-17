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
        public static async Task<string> Main(string acdFilePath, string chassisName, string controllerName)
        {
            // Note: Implementing the v2 service client, as it allows for us to create chassis.
            // Only one Service Client is required per machine, anyways.
            IServiceApiClientV2 serviceClient = ClientFactory.GetServiceApiClientV2("ServiceClientV2");
            serviceClient.Culture = new CultureInfo("en-US");
            Console.WriteLine(CheckCurrentChassis_Sync(chassisName, serviceClient) == false);
            Console.WriteLine(CheckCurrentControllers_Sync(chassisName, controllerName, serviceClient) == false);
            // Check if an emulated controller exists within an emulated chassis. If not, run through the if statement contents to create one.
            if ((CheckCurrentChassis_Sync(chassisName, serviceClient) == false) && (CheckCurrentControllers_Sync(chassisName, controllerName, serviceClient) == false))
            {
                // Set up emulated chassis information.
                var chassisUpdate = new ChassisUpdate
                {
                    Name = chassisName,
                    Description = $"Test chassis for CI/CD created by Echo SDK: {System.DateTime.Now}"
                };
                ChassisData chassisCICD = await serviceClient.CreateChassis(chassisUpdate);
                Console.WriteLine("chassis!");
                if (CheckCurrentControllers_Sync(chassisName, controllerName, serviceClient) == false)
                {
                    // Set up emulated controller information.
                    using (var fileHandle = await serviceClient.SendFile(acdFilePath))
                    {
                        var controllerUpdate = await serviceClient.GetControllerInfoFromAcd(fileHandle);
                        controllerUpdate.ChassisGuid = chassisCICD.ChassisGuid;
                        var controllerData = await serviceClient.CreateController(controllerUpdate);
                    }
                    Console.WriteLine("controller!");
                }
            }

            // Get emulated controller information.
            string[] testControllerInfo = Get_ControllerInfo_Sync(chassisName, controllerName, serviceClient);
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
        public static async Task<bool> CheckCurrentChassis_Async(string chassisName, IServiceApiClientV2 serviceClient)
        {
            var chassisList = (await serviceClient.ListChassis()).ToList();
            for (int i = 0; i < chassisList.Count; i++)
            {
                if (chassisList[i].Name == chassisName)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Asynchronously check to see if a specific controller exists in a specific chassis.
        /// </summary>
        /// <param name="chassisName">The name of the emulated chassis to check the emulated controler in.</param>
        /// <param name="controllerName">The name of the emulated controller to check.</param>
        /// <param name="serviceClient">The Factory Talk Logix Echo interface.</param>
        /// <returns>A Task that returns a boolean value 'True' if the emulated controller already exists and a 'False' if it does not.</returns>
        public static async Task<bool> CheckCurrentControllers_Async(string chassisName, string controllerName, IServiceApiClientV2 serviceClient)
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
        public static bool CheckCurrentChassis_Sync(string chassisName, IServiceApiClientV2 serviceClient)
        {
            var task = CheckCurrentChassis_Async(chassisName, serviceClient);
            task.Wait();
            return task.Result;
        }

        /// <summary>
        /// Run the CheckCurrentChassisAsync method synchronously.<br/>
        /// Check to see if a specific controller exists in a specific chassis.
        /// </summary>
        /// <param name="chassisName">The name of the emulated chassis to check the emulated controler in.</param>
        /// <param name="controllerName">The name of the emulated controller to check.</param>
        /// <param name="serviceClient">The Factory Talk Logix Echo interface.</param>
        /// <returns>A boolean value 'True' if the emulated controller already exists and a 'False' if it does not.</returns>
        public static bool CheckCurrentControllers_Sync(string chassisName, string controllerName, IServiceApiClientV2 serviceClient)
        {
            var task = CheckCurrentControllers_Async(chassisName, controllerName, serviceClient);
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

        public static string[] Get_ControllerInfo_Sync(string chassisName, string controllerName, IServiceApiClientV2 serviceClient)
        {
            var task = Get_ControllerInfo_Async(chassisName, controllerName, serviceClient);
            task.Wait();
            return task.Result;
        }
        #endregion
    }
}
