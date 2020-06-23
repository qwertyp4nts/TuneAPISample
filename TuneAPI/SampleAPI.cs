using M1Tune;
using System;
using System.Diagnostics;

namespace TuneAPI
{
    class SampleAPI : OleMessageFilter
    {
        private Tune            m_tune;
        IMtcM1TuneApplication1  m_tuneApp;
        bool                    m_ECUConnectionState;
        uint                    m_ConnectedECUSerialNumber;

        [STAThread]
        public static void Main(string[] args)
        {
            Register();

            SampleAPI api = new SampleAPI();
            if (api.Initialise())
            {
                while (true)
                {
                    api.CheckForAPILicence();
                    api.PrintIntroHelp();
                    try
                    {
                        var option = int.Parse(Console.ReadLine());
                        api.ProcessOption(option);
                    }
                    catch (FormatException)
                    {
                        Console.WriteLine("Please enter a valid number");
                        break;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.ToString());
                    }
                }
            }

            Revoke();
        }

        bool Initialise()
        {
            try
            {
                m_tune = new Tune();
                m_tune.Visible = true; // Makes the Tune application visible to the user
                m_tuneApp = m_tune as IMtcM1TuneApplication1;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return m_tune != null;
        }

        void PrintIntroHelp()
        {
            Console.WriteLine("Choose from the following options to demo functionality:");
            Console.WriteLine("1: Print details of all installed packages to the console");
            Console.WriteLine("2: Connect to ECU and retrieve package");
            Console.WriteLine("3: Download Logged Data");
            Console.WriteLine("4: Load Most Recent Package");
            Console.WriteLine("5: Load Package by Name");
            Console.WriteLine("6: Send a GP package to the ECU");
            Console.WriteLine("7: Retrieve package from ECU and save it as a new file");
            Console.WriteLine("8: Get all live channel values and print to the console");
            Console.WriteLine("9: Get the live value of the 'ECU Uptime' channel");
            Console.WriteLine("10: Get the live value of the 'ECU Uptime' channel, at 1.5 seconds into the telemetry display");
            Console.WriteLine("11: Change the value of the 'Inlet Air Temperature Sensor Default' Parameter");
            Console.WriteLine("12: Print the Engine Efficiency table to the console");
            Console.WriteLine("13: Assign a resource to 'Airbox Temperature Sensor Resource' and set its translation");
            Console.WriteLine("14: Set the 'ADR CAN Bus' parameter to 'CAN Bus 1'");
            Console.WriteLine("15: Setup axis of Engine Efficiency table and tune sites");
            Console.WriteLine("16: Exit the program");
            Console.WriteLine("");
        }

        void ProcessOption(int option)
        {
            try
            {
                switch (option)
                {
                    case 1:
                        PrintAllPackages();
                        break;
                    case 2:
                        ConnectToECU();
                        break;
                    case 3:
                        DownloadLoggedData();
                        break;
                    case 4:
                        LoadRecentPackage();
                        break;
                    case 5:
                        LoadPackageByName();
                        break;
                    case 6:
                        SendPackage();
                        break;
                    case 7:
                        RetrievePackageFromECUandSave();
                        break;
                    case 8:
                        GetAllChannels();
                        break;
                    case 9:
                        GetChannelValue("ECU Uptime");
                        break;
                    case 10:
                        GetChannelValueAtSpecificTime("ECU Uptime", 1.5);
                        break;
                    case 11:
                        TuneIATParameter();
                        break;
                    case 12:
                        PrintTable("Engine Efficiency");
                        PrintTable("Engine Efficiency Main");
                        break;
                    case 13:
                        SetupAndTuneAirboxTemp();
                        break;
                    case 14:
                        TuneParameterByEnum("ADR CAN Bus", "CAN Bus 1");
                        break;
                    case 15:
                        SetupEngineEfficiencyTableAxis();
                        break;
                    case 16:
                        Exit();
                        break;

                    default:
                        break;
                }
            }
            catch (FormatException)
            {
                Debug.Assert(false);
            }
        }

        void PrintAllPackages()
        {
            CheckForInstalledPackage();

            var installedPkgs = m_tuneApp.InstalledPackages; 
            Console.WriteLine($"Total Installed Packages : {installedPkgs.Count}"); // Prints the number of installed packages to the console

            foreach (IMtcPackageInfo p in installedPkgs) // Prints details of all installed packages to the console
            {
                PrintInstalledPackage(p);
            }
        }

        void OpenWorkspace()
        {
            //Most functions open the most recent workspace. The following code is provided if the user wants to open a specific workspace
            if (m_tuneApp.RecentWorkspaces.Count > 0)
            {
                var m_recentWorkspace = m_tuneApp.RecentWorkspaces[0];

                if (m_recentWorkspace != null)
                {
                    var f = m_recentWorkspace.Path;
                    m_tuneApp.WorkspaceLoad(f); //Loads the most recently used workspace
                    //m_tuneApp.WorkspaceLoad("C:\\Users\\mila\\Documents\\MoTeC\\M1\\Tune\\Workspaces\\Tune 1"); //Loads workspace by file path
                }
            }
        }
        void OpenPasswordProtectedPackage()
        {
            //Pre-requisite: The password must already be set in the package. This function does not set password, it only allows login
            m_tuneApp.AutoLoginManager.Enabled = true; //Setting this to true hides the security dialog, so we can login silently with the API
            m_tuneApp.AutoLoginManager.SetLoginPassword("username", "password"); //Store the username and password, for silent login
            //m_tuneApp.AutoLoginManager.SetLoginKey("username", "C:\\Users\\mila\\Documents\\MoTeC\\M1\\Tune\\myKey.key"); //Login using key
            m_tuneApp.AutoLoginManager.SetAutoLoginName("username"); //Choose the user with which to login
            //m_tuneApp.AutoLoginManager.SetAutoLoginName(""); //Empty string is Guest user
            ConnectToECU();

            try
            {
                var pkg = GetMainPackage();
            }
            catch (Exception)
            {
                Console.WriteLine("Check username and password, or user permissions");
            }
        }

        void ConnectToECU()
        {
            CheckECUConnectionStatus();

            if (m_ECUConnectionState == false)
            {
                IMtcDevice currentDevice = GetDevice();

                uint serialNum = currentDevice.Serial;
                m_tuneApp.Devices.Connect(serialNum); //Connects to first device found
                //m_tuneApp.Devices.Connect(2851); // -> Connects to device with target serial number
                m_ConnectedECUSerialNumber = serialNum;
            }

            CheckECUConnectionStatus(true); //Ensure ECU connection is established
        }

        void DownloadLoggedData()
        {
            ConnectToECU();
            m_tuneApp.Devices.RetrieveLogData(m_ConnectedECUSerialNumber);
            //This takes us to screen where we select which sectors to extract from
            //It requires user interaction from here.
        }

        void LoadRecentPackage()
        {
            CheckForRecentPackage();
 
            IMtcRecentFile recentPkg = m_tuneApp.RecentPackages[0];
            m_tuneApp.Packages.Load(recentPkg.Path, false);
        }

        void LoadPackageByName()
        {
            const string pkgFileName = "Generic 4 cylinder, MAP based Efficiency Migration Base v5"; //The name of the package we want to open
            const string ecuModel = "M150"; //The hardware device type        
            bool foundPackage = false;

            CheckForInstalledPackage();

            var installedPkgs = m_tuneApp.InstalledPackages;

            foreach (IMtcPackageInfo p in installedPkgs)
            {
                PrintInstalledPackage(p);
                if (p.Comment.Equals(pkgFileName) && p.Hardware.Equals(ecuModel))
                {
                    foundPackage = true;
                    m_tune.Packages.Load(p.FileName, false);
                    break;
                }
            }

            if (!foundPackage)
                Console.WriteLine($"Did not find file with name: {pkgFileName} and ECU model {ecuModel}");
        }

        void SendPackage()
        {
            LoadPackageByName();
            var pkg = GetMainPackage();

            if (m_ECUConnectionState == false)
            {
                IMtcDevice currentDevice = GetDevice();
                m_ConnectedECUSerialNumber = currentDevice.Serial;
                Console.WriteLine($"Connected to ECU #{m_ConnectedECUSerialNumber}");
            }

            if (pkg.Send(m_ConnectedECUSerialNumber))
                Console.WriteLine("Package sent successfully");
            else
                Console.WriteLine("Package send failed");
        }

        void RetrievePackageFromECUandSave()
        {
            ConnectToECU();
            var pkg = GetMainPackage();
            pkg.SaveAs("Testing package save from API", "Generated by MoTeC Tune API");
        }

        void GetAllChannels()
        {
            ConnectToECU();
            CheckForMainPackage();

            var allChannels = m_tuneApp.Packages[0].DAQ.GetRealTimeValueAll();
            if (allChannels != null)
            {
                //Print all ECU channels and their values to the console
                foreach (IMtcDAQValue channel in allChannels)
                {
                    Console.WriteLine($"{channel.DisplayName}: {channel.DisplayValue} {channel.DisplayUnit}");
                }
            }
            else
            {
                Console.WriteLine("Failed to fetch channels. Check package");
            }
        }

        void GetChannelValue(string channel)
        {
            ConnectToECU();
            CheckForMainPackage();

            string formattedChannel = channel.Replace(' ', '.');
            var c = new string[] { formattedChannel };
            var allChannels = m_tuneApp.Packages[0].DAQ.GetRealTimeValueAll();

            if (allChannels != null && allChannels.Length > 0)
            {
                var liveChannelValue = m_tuneApp.Packages[0].DAQ.GetRealTimeValue(c);

                if (liveChannelValue == null)
                    Console.WriteLine($"Could not find {channel}");
                else
                    Console.WriteLine($"{channel} channel found. Current value is {liveChannelValue[0].DisplayValue} {liveChannelValue[0].DisplayUnit}"); 
            }
            else
            {
                Console.WriteLine("Failed to fetch channels. Check package");
            }
        }

        void GetChannelValueAtSpecificTime(string channel, double time)
        {
            ConnectToECU();
            CheckForMainPackage();
            System.Threading.Thread.Sleep(1500); //This is just for the sample. Delete when using function

            var pkg = m_tuneApp.Packages[0];
            string formattedChannel = channel.Replace(' ', '.');
            var c = new string[] { formattedChannel };
            var us = SecsToMicroSecs(time);

            if (pkg.DAQ.Time > us)
            {
                var v = pkg.DAQ.GetValue(c, us); //Gets the channel value <time> seconds into the telemetry display

                if (v != null && v[0].DisplayName.Equals(channel))
                    Console.WriteLine($"{channel} channel found. Value at {time.ToString()} seconds is {v[0].DisplayValue} {v[0].DisplayUnit}");
                else
                    Console.WriteLine($"Could not find {channel}");
            }
            else
                Console.WriteLine($"Requested telemetry time of {time} seconds has not elapsed yet. Maximum DAQ time is currently {(pkg.DAQ.Time / 1e6).ToString ("0.##")} seconds");
        }

        double SecsToMicroSecs(double s)
        {
            return s * 1e6;
        }

        void TuneIATParameter()
        {
            //Function pre-requisite: IAT channel has no resource assigned, Default value is 35.0C
            //This function changes the IAT default value and checks the IAT channel value after modifying its properties

            const string c = "Inlet Air Temperature";
            const string p = "Inlet Air Temperature Sensor Default";
 
            GetChannelValue(c); //Check value of IAT
            {
                var pkg = GetMainPackage();
                var parameterToChange = pkg.Parameters[p];
                if (parameterToChange != null)
                {
                    Console.WriteLine("Found :" + p);
                    parameterToChange.Site.Device.Value = 75.0; //Change IAT Sensor Default to 75.0

                    if (parameterToChange.Site.Device.Value != 75.0)
                    {
                        throw new Exception("Parameter set fail");
                    }
                }
            }
            GetChannelValue(c); //Ensure channel now reports updated value
        }

        void TuneParameter(string channelName, string channelValue)
        {
            //Generic version of tuneIATParameter() to use in other methods
            //Not only does this function tune parameters, it also sets dropdown values. 
            //For example: to set ADR CAN Bus from 'Not in Use' to 'CAN Bus 2' (which is the enumeration of index 2), write: tuneParameter("ADR CAN Bus", "2");
            //This method does this by index. See TuneParameterByEnum() to tune via enumeration name

            var pkg = GetMainPackage();

            var parameterToChange = pkg.Parameters[channelName];
            if (parameterToChange != null)
            {
                Console.WriteLine("Found :" + channelName);
                try
                {
                    double v = double.Parse(channelValue);
                    parameterToChange.Site.Device.Value = v;
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.ToString()); //make sure the input string channelValue can be converted to a double. If yes, might fail due to COM failure, or if resource is already assigned
                }
            }
            else
            {
                Console.WriteLine("Could not find parameter " + channelName);
            }
        }

        void TuneParameterByEnum(string channelName, string channelValue)
        {
            ConnectToECU();

            var pkg = GetMainPackage();

            var parameterToChange = pkg.Parameters[channelName];
            if (parameterToChange != null)
            {
                Console.WriteLine("Found: " + channelName);
                try
                {
                    var v = parameterToChange.Enumeration.EnumeratorByDisplayName[channelValue].Value;
                    parameterToChange.Site.Device.Value = v;
                }
                catch (Exception)
                {
                    Console.WriteLine($"ERROR: Ensure '{channelValue}' is a valid enumeration of: {channelName}");
                }
                // The following is commented out to make the demo self-contained. Re-enable to save the package after the change has been made
                // SavePackage(pkg); //Setting a resource requires package save
                // CheckECUConnectionStatus(true); //Setting resource requires ECU reset following a save. Ensure we re-connect successfully.
            }
            else
            {
                Console.WriteLine("Could not find parameter " + channelName);
            }
        }

        void PrintTable(string tableName)
        {
            ConnectToECU();

            var pkg = GetMainPackage();
            var tables = pkg.Tables;
            var t = tables[tableName];

            if (t != null)
                PrintTable(tables[tableName]);
            
            else
                Console.WriteLine($"{tableName} table could not be found");
        }

        void SetupAndTuneAirboxTemp()
        {
            ConnectToECU();
            
            var pkg = GetMainPackage();
            TuneParameterByEnum("Airbox Temperature Sensor Resource", "Analogue Voltage Input 11");

            SavePackage(pkg); //Setting a resource requires package save
            CheckECUConnectionStatus(true); //Setting resource requires ECU reset following a save. Ensure we re-connect successfully.

            var tables = pkg.Tables;
            IMtcTable t = tables["Airbox Temperature Sensor Translation"];
            if (t != null)
            {
                double[] x = { 1.000, 1.500, 2.000, 2.500, 3.000, 3.500, 4.000 }; //The voltage values we want on the x axis
                IMtcTableAxisShape xShape = t.XAxis.Shape;
                xShape.Values = x;

                t.ReShape(xShape, null, null, true);

                t.Site[0, 0, 0].Device.Value = -20;
                t.Site[1, 0, 0].Device.Value = 0;
                t.Site[2, 0, 0].Device.Value = 20;
                t.Site[3, 0, 0].Device.Value = 40;
                t.Site[4, 0, 0].Device.Value = 60;
                t.Site[5, 0, 0].Device.Value = 80;
                t.Site[6, 0, 0].Device.Value = 100;
            }
            else
            {
                Console.WriteLine($"{t.DisplayName} was null. Check table");
            }
            //SavePackage(pkg); //Save again after making changes to the table. This is commented out to make the demo self-contained.
        }

        void SetupEngineEfficiencyTableAxis()
        {
            var pkg = GetMainPackage();
            var tables = pkg.Tables;

            var t = tables["Engine Efficiency"];
            if (t == null)
            {
                t = tables["Engine Efficiency Main"];
            }

            double[] x = { 0, 500, 5000 };
            double[] y = { 30, 70, 90, 110};
            double[] z = { 100, 102 };
            SetupTableAxis(t.DisplayName, true, x, "rpm", true, y, "kPa", true, z, "kPa a");

            TuneEngineEfficiencyTable(t.DisplayName);
        }

        void TuneEngineEfficiencyTable(string tableName)
        {
            //call after SetupEngineEfficiencyTableAxis()
            var pkg = GetMainPackage();
            IMtcTable t = pkg.Tables[tableName];

            t.Site[0, 0, 0].Display.Value = 10;
            t.Site[1, 0, 0].Display.Value = 53;
            t.Site[2, 0, 0].Display.Value = 99;
            t.Site[0, 1, 0].Display.Value = 40;
            t.Site[1, 1, 0].Display.Value = 73;
            t.Site[2, 1, 0].Display.Value = 141.5;
            t.Site[0, 2, 0].Display.Value = 80;
            t.Site[1, 2, 0].Display.Value = 89.1;
            t.Site[2, 2, 0].Display.Value = 170.7;
            t.Site[0, 3, 0].Display.Value = 160;
            t.Site[1, 3, 0].Display.Value = 171;
            t.Site[2, 3, 0].Display.Value = 200.0;
        }

        void SetupTableAxis(string tableName, bool xEnabled, double[] x, string xUnit, bool yEnabled = false, double[] y = null, string yUnit = null, bool zEnabled = false, double[] z = null, string zUnit = null)
        { 
            //Generic function to setup table axis. Called by SetupEngineEfficiencyTableAxis()

            ConnectToECU();

            var pkg = GetMainPackage();

            var tables = pkg.Tables;
            IMtcTable t = tables[tableName];

            if (t != null)
            {
                IMtcTableAxisShape xShape = BuildShapeObject(t.XAxis.Shape, xEnabled, x, xUnit);
                IMtcTableAxisShape yShape = BuildShapeObject(t.YAxis.Shape, yEnabled, y, yUnit);
                IMtcTableAxisShape zShape = BuildShapeObject(t.ZAxis.Shape, zEnabled, z, zUnit);

                t.ReShape(xShape, yShape, zShape, true);
            }
            else
            {
                Console.WriteLine($"{t.DisplayName} was null. Check table");
            }
            //SavePackage(pkg); //Save again after making changes to the table. This is commented out to make the demo self-contained.
        }

        IMtcTableAxisShape BuildShapeObject(IMtcTableAxisShape shape, bool enabled = false, double[] values = null, string unit = null)
        {
            if (enabled == true)
            {
                shape.Enabled = true;
                if (values != null)
                {
                    shape.Values = values;
                    if (unit != null)
                    {
                        shape.Unit = unit;
                    }
                }
            }
            else
            {
                shape.Enabled = false;
            }
            return shape;
        }

        void Exit()
        {
            m_tuneApp.Exit();
            Environment.Exit(0);
        }

        void CheckForInstalledPackage()
        {
            if (m_tuneApp.InstalledPackages == null || m_tuneApp.InstalledPackages.Count == 0)
                throw new Exception("ERROR: No installed packages were found");
        }

        void CheckECUConnectionStatus(bool breakIfNotConnected = false)
        {
            if (m_tuneApp.Packages.Count > 0)
            {
                bool p = m_tuneApp.Packages[0].Connected;
                m_ECUConnectionState = p;
            }
            else
            {
                m_ECUConnectionState = false;
            }

            if (m_ECUConnectionState == false && breakIfNotConnected == true)
                throw new Exception("ERROR: Not connected to ECU");
        }

        IMtcDevice GetDevice()
        {
            if (m_tuneApp.Devices.Count == 0)
            {
                throw new Exception("ERROR: No ECU connections are found");
            }

            IMtcDevice d = m_tuneApp.Devices[0];

            if (d == null || d.Serial == 0)
            {
                throw new Exception("ERROR: No ECU connections are found");
            }
            else
            {
                return d;
            }
        }

        void CheckForRecentPackage()
        {
            if (m_tuneApp.RecentPackages.Count == 0 || m_tuneApp.RecentPackages[0] == null)
            {
                throw new Exception("ERROR: No recent package was found");
            }
        }

        void CheckForMainPackage()
        {
            if (m_tuneApp.Packages.Count == 0)
            {
                throw new Exception("ERROR: Main package not loaded");
            }
        }

        static void PrintInstalledPackage(IMtcPackageInfo pkg)
        {
            Console.WriteLine($"File Name : {pkg.FileName}");
            Console.WriteLine($"\tFile VehicleId : {pkg.VehicleId}");
            Console.WriteLine($"\tFile SerialNumber : {pkg.SerialNumber}");
            Console.WriteLine($"\tFile Comment : {pkg.Comment}");
            Console.WriteLine($"\tFile FirmwareVersionName : {pkg.FirmwareVersionName}");
            Console.WriteLine($"\tFile FirmwareVersion : {pkg.FirmwareVersion}");
            Console.WriteLine($"\tFile Hardware : {pkg.Hardware}");
            Console.WriteLine($"\tFile ModifiedDateTime : {pkg.ModifiedDateTime}");
            Console.WriteLine($"\tFile ImportedDateTime : {pkg.ImportedDateTime}");
        }

        IMtcPackage3 GetMainPackage()
        {
            if (m_tuneApp.Packages != null && m_tuneApp.Packages.Count > 0)
            {
                return m_tuneApp.Packages[0] as IMtcPackage3;
            }
            else
            {
                throw new Exception("FAILED: No main package loaded");
            }
        }

        void SavePackage(IMtcPackage3 pkg)
        {
            bool saved = pkg.Save();
            if (saved)
                Console.WriteLine("Package saved successfully");
            else
                throw new Exception("Failed to save package");
        }

        static void PrintTable(IMtcTable t)
        {
            if (t != null)
            {
                Console.WriteLine($"Table '{t.DisplayName}':");

                PrintAdjustItem(t);

                PrintTableAxis(t.XAxis, "X");
                PrintTableAxis(t.YAxis, "Y");
                PrintTableAxis(t.ZAxis, "Z");
            }
        }

        static void PrintAdjustItem(IMtcAdjustItem item)
        {
            Console.WriteLine($"{item.DisplayName} ({(item.ReadOnly ? "Readonly" : "Editable")} and {(item.Visible ? "Visible" : "Invisible")}) : {item.DataType}");
            PrintEnumeration(item.Enumeration);
        }

        static void PrintTableAxis(IMtcTableAxis axis, string name)
        {
            if (axis != null)
            {
                Console.WriteLine($"Axis {name}: ");
                Console.WriteLine($"{axis.DisplayName} ({(axis.Enabled ? "Enabled" : "Disabled")}, {(axis.ReadOnly ? "Readonly" : "Editable")} and {(axis.Visible ? "Visible" : "Invisible")}) : {axis.DataType}");
                PrintEnumeration(axis.Enumeration);

                Console.WriteLine($"\tMax Sites : {axis.MaxSites}");
                Console.WriteLine($"\tUsed Sites : {axis.UsedSites}");
                Console.WriteLine($"\tData Type : {axis.DataType}");

                Console.Write("\tValues: ");
                for (uint i = 0; i < axis.UsedSites; i++)
                {
                    Console.Write(axis.Site[i].Display.DisplayValue + " ");
                }
                Console.WriteLine();
            }
            else
                Console.WriteLine($"Axis {name}:  (none)");
        }

        static void PrintEnumeration(IMtcEnumeration e)
        {
            if (e != null)
            {
                Console.WriteLine(e.Name);
                foreach (IMtcEnumerator v in e)
                    Console.WriteLine($"{v.Value} : {v.DisplayName}");

                for (int i = 0; i < e.Count; i++)
                    Console.WriteLine($"{e[i].Value} : {e[i].DisplayName}");
            }
        }

        void CheckForAPILicence()
        {
            var activated = m_tuneApp.IsActivated();
            if (activated == false)
            {
                throw new Exception("Failed to find valid Tune API licence.");
            }
        }

    }
}
