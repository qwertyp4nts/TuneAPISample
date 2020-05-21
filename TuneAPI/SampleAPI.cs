using M1Tune;
using System;
using System.Diagnostics;
using System.Threading;

namespace TuneAPI
{
    class SampleAPI : OleMessageFilter
    {
        private Tune            m_tune;
        IMtcM1TuneApplication1  m_tuneApp;
        bool                    m_ECUConnectionState;
        uint                    m_ConnectedECUSerialNumber;
        IMtcRecentFile          m_recentWorkspace;

        [STAThread]
        public static void Main(string[] args)
        {
            Register();

            SampleAPI api = new SampleAPI();
            if (api.Initialise())
            {
                while (true)
                {
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
            Console.WriteLine("10: Change the value of the 'Inlet Air Temperature Sensor Default' Parameter");
            Console.WriteLine("11: Print the Engine Efficiency table to the console");
            Console.WriteLine("12: Assign a resource to 'Airbox Temperature Sensor Resource' and set its translation");
            Console.WriteLine("13: Set the 'ADR CAN Bus' parameter to 'CAN Bus 1'");
            Console.WriteLine("14: TESTING LOCKED ITEMS");
            Console.WriteLine("15: Exit the program");
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
                        TuneIATParameter();
                        break;
                    case 11:
                        PrintTable();
                        break;
                    case 12:
                        TuneTable();
                        break;
                    case 13:
                        TuneParameterByEnum("ADR CAN Bus", "CAN Bus 1");
                        break;
                    case 14:
                        LockTest();
                        break;
                    case 15:
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

        void TestResourceAssignment()
        {
            string channelName = "Airbox Temperature Sensor Resource";
            string channelValue = "11";
            var pkg = GetMainPackage();

            var parameterToChange = pkg.Parameters[channelName];

            foreach (IMtcParameter p in pkg.Parameters)
            {
                if(p.Site.Device.DisplayValue == "Analogue Voltage Input 12")
                {
                    throw new Exception("Resource already assigned to " + p.DisplayName);
                }
            }
            double v = double.Parse(channelValue);
            parameterToChange.Site.Device.Value = v;
        }

        void LockTest()
        {
            //  ConnectToECU();

            //  TestIMtcDAQInterface();
            //      m_tuneApp.Devices.Connect(2851);

            //    object a = DAQRealTImeValue("Engine Speed");

            // TestIMtcDAQValueInterface("Inlet Manifold Pressure");

            //   TestIMtcEnumeratorInterface("ADR CAN Bus", "2");

            /*Test IMtcMINMaxValidator
            var pkg = GetMainPackage();
            var tables = pkg.Tables;
            IMtcTable t = tables["Engine Efficiency"];
            TestIMtcMinMaxValidatorInterface(t.MinMaxValidator);
            */  //Test IMtcMINMaxValidator

            //   TestIMtcSiteValueInterface();
            TestResourceAssignment();
        }

        void TestIMtcM1TuneApplication1()
        {
            var instpkgs = m_tuneApp.InstalledPackages;
            var autologin = m_tuneApp.AutoLoginManager;
            var queryAPI = m_tuneApp.QueryAPI["sz"];
        }

        void TestIMtcM1TuneApplication()
        {
            /*    m_tuneApp.WorkspaceNew();
            m_tuneApp.WorkspaceOpen();
            m_tuneApp.WorkspaceLoad("C:\\Users\\mila\\Documents\\MoTeC\\M1\\Tune\\Workspaces\\Tune 4"); //Loads workspace by file path
            */
            var recents = m_tuneApp.RecentWorkspaces;
            var demofile = m_tuneApp.DemoFile;
            m_tuneApp.RunDemo();
            m_tuneApp.ShowOptions();
        }

        void TestIMtcM1TuneApplication0()
        {
            var recents = m_tuneApp.RecentPackages;
            var devices = m_tuneApp.Devices;
            var pkgs = m_tuneApp.Packages;
        }

        void TestIMtcQueryAPIInterface()
        {
            var a = m_tuneApp.QueryAPI["sz"];
            
        }

        void TestIMtcPackages1Interface()
        {
            var installedPkgs = m_tuneApp.InstalledPackages;
            var pkg = installedPkgs[0];
            var pkgs = (IMtcPackages1) m_tuneApp.Packages;
            pkgs.Load(pkg.FileName, false);
            pkgs.CloseAll();
            
        }

        void TestIMtcPackage3Interface()
        {
            var pkg = GetMainPackage();
            
            foreach (IMtcEnumeration e in pkg.Enumerations)
            {
                for (int i = 0; i < e.Count; i++)
                {
                    //Console.WriteLine(e.DisplayName + " " + e.Value);
                    Console.WriteLine(e.Name);
                    Console.WriteLine(e.EnumeratorByValue[i].DisplayName);
                }
                Console.WriteLine("");
            }
            //params tested in TestIMtcParametersInterface
            //tables tested in TestIMtcTablesInterface
            //save tested in many other methods of sample
        }

        void TestIMtcPackageInterface()
        {
            var pkg = GetMainPackage();
            Console.WriteLine("Name: " + pkg.Name);
            Console.WriteLine("Connected state: " + pkg.Connected);
            Console.WriteLine("DAQ: " + pkg.DAQ);
            Console.WriteLine("Firmware ID: " + pkg.FirmwareId);
            Console.WriteLine("Local file Name: " + pkg.LocalFileName);
         //   var a = pkg.HashValue[];
            Console.WriteLine("Vehicle ID: " + pkg.VehicleId);

            pkg.Disconnect();
        }

        void TestIMtcInstalledPackageInterface()
        {
            var installedPkgs = m_tuneApp.InstalledPackages;
            var pkg = installedPkgs[0];
            m_tune.Packages.Load(pkg.FileName, false);

            Console.WriteLine($"File Name : {pkg.FileName}");
            Console.WriteLine($"\tFile VehicleId : {pkg.VehicleId}");
            Console.WriteLine($"\tFile SerialNumber : {pkg.SerialNumber}");
            Console.WriteLine($"\tFile Comment : {pkg.Comment}");
            Console.WriteLine($"\tFile Firmware : {pkg.Firmware}");
            Console.WriteLine($"\tFile FirmwareVersionName : {pkg.FirmwareVersionName}");
            Console.WriteLine($"\tFile FirmwareVersion : {pkg.FirmwareVersion}");
            Console.WriteLine($"\tFile Hardware : {pkg.Hardware}");
            Console.WriteLine($"\tFile ModifiedDateTime : {pkg.ModifiedDateTime}");
            Console.WriteLine($"\tFile ImportedDateTime : {pkg.ImportedDateTime}");
        }

        void TestIMtcParametersInterface()
        {
            const string p = "Inlet Air Temperature Sensor Default";

            var pkg = GetMainPackage();
            var parameterToChange = pkg.Parameters[p];

            Console.WriteLine(p + ": " + parameterToChange.Site.Display.DisplayValue + parameterToChange.Site.Display.DisplayUnit);

            foreach (IMtcParameter pr in pkg.Parameters)
            {
                Console.WriteLine(pr.DisplayName + ": " + pr.Site.Display.DisplayValue + pr.Site.Display.DisplayUnit);
            }
           // var paramIndex0 = pkg.Parameters.Index[0];

         //   Console.WriteLine(paramIndex0.DisplayName + ": " + paramIndex0.Site.Display.DisplayValue + paramIndex0.Site.Display.DisplayUnit);
        }

        void TestIMtcTablesInterface()
        {
            var pkg = GetMainPackage();
            var tables = pkg.Tables;
         /*   const string tableName = "Engine Efficiency";
            var t = tables[tableName];
            PrintTable(t);*/

            var t2 = tables.Index[0];
            Console.WriteLine($"Table name: " + t2.DisplayName);
            PrintTableAxis(t2.XAxis, "X");
            PrintTableAxis(t2.YAxis, "Y");
            PrintTableAxis(t2.ZAxis, "Z");
            //PrintTable(t2);

            var t3 = tables.Index[1];
            Console.WriteLine($"Table name: " + t3.DisplayName);
            PrintTableAxis(t3.XAxis, "X");
            PrintTableAxis(t3.YAxis, "Y");
            PrintTableAxis(t3.ZAxis, "Z");
            // PrintTable(t3);

            var t4 = tables.Index[2];
            Console.WriteLine($"Table name: " + t4.DisplayName);
            PrintTableAxis(t4.XAxis, "X");
            PrintTableAxis(t4.YAxis, "Y");
            PrintTableAxis(t4.ZAxis, "Z");

            var t5 = tables.Index[3];
            Console.WriteLine($"Table name: " + t5.DisplayName);
            PrintTableAxis(t5.XAxis, "X");
            PrintTableAxis(t5.YAxis, "Y");
            PrintTableAxis(t5.ZAxis, "Z");

            var t6 = tables.Index[4];
            Console.WriteLine($"Table name: " + t6.DisplayName);
            PrintTableAxis(t6.XAxis, "X");
            PrintTableAxis(t6.YAxis, "Y");
            PrintTableAxis(t6.ZAxis, "Z");

            var t7 = tables.Index[5];
            Console.WriteLine($"Table name: " + t7.DisplayName);
            PrintTableAxis(t7.XAxis, "X");
            PrintTableAxis(t7.YAxis, "Y");
            PrintTableAxis(t7.ZAxis, "Z");

            var t8 = tables.Index[6];
            Console.WriteLine($"Table name: " + t8.DisplayName);
            PrintTableAxis(t8.XAxis, "X");
            PrintTableAxis(t8.YAxis, "Y");
            PrintTableAxis(t8.ZAxis, "Z");

            var t9 = tables.Index[7];
            Console.WriteLine($"Table name: " + t9.DisplayName);
            PrintTableAxis(t9.XAxis, "X");
            PrintTableAxis(t9.YAxis, "Y");
            PrintTableAxis(t9.ZAxis, "Z");

            var t10 = tables.Index[8];
            Console.WriteLine($"Table name: " + t10.DisplayName);
            PrintTableAxis(t10.XAxis, "X");
            PrintTableAxis(t10.YAxis, "Y");
            PrintTableAxis(t10.ZAxis, "Z");

            var t11 = tables.Index[9];
            Console.WriteLine($"Table name: " + t11.DisplayName);
            PrintTableAxis(t11.XAxis, "X");
            PrintTableAxis(t11.YAxis, "Y");
            PrintTableAxis(t11.ZAxis, "Z");
        }

        void TestInterpolate()
        {
            var pkg = GetMainPackage();
            var tables = pkg.Tables;
            const string tableName = "Boost Aim Main";
            var t = tables[tableName];

            double [] x = { 0, 18000 }; //eng sp * 6
            double [] y = { 0, 20, 40, 60, 80, 100}; //tpos
            t.ReShape(true, x, true, y, false, null, false);

            PrintTable(t);
            double [] x2 = { 0, 9000, 18000 }; //insert value
            t.ReShape(true, x2, true, y, false, null, true);

            PrintTable(t);
            //ensure values in 2500  rpm column are 50%

            //repeat test with interpolate off on 2nd reshape. ensure values are 100%
        }

        void TestIMtcTable()
        {
            var pkg = GetMainPackage();
            var tables = pkg.Tables;
            const string tableName = "Engine Efficiency";
            var engEff = tables[tableName];

            PrintTable(engEff);

            Console.WriteLine("X Axis: " + engEff.XAxis);
            Console.WriteLine("Y Axis: " + engEff.YAxis);
            Console.WriteLine("Z Axis: " + engEff.ZAxis);
            Console.WriteLine(engEff.Site[0, 0, 0].Display.Value);

            //  double[] x = { 0, 500, 1000, 5000 };
            //  double[] y = { 50, 70, 90, 110 };

            double[] x = { 0, 3000, 6000, 30000 }; //base units are deg/sec, so multiply all values by 6
            double[] y = { 50000, 70000, 90000, 110000 }; //base units are Pa, so multiply all values by 1000
            double[] z = { 90000, 100000, 110000 };
            engEff.ReShape(true, x, true, y, true, z, false);

            PrintTable(engEff);

            // t.Site[0, 0, 0].Device.Value = -20;
        }

        void TestIMtcTableAxisDisabledGroup()
        {
            var pkg = GetMainPackage();
            var tables = pkg.Tables;
            const string tableName = "Boost Aim Main";
            var engEff = tables[tableName];
            PrintTable(engEff);
        }

        void TestIMtcTableAxis()
        {
            var pkg = GetMainPackage();
            var tables = pkg.Tables;
            const string tableName = "Engine Efficiency";
            var engEff = tables[tableName];
            PrintTable(engEff);

            IMtcTableAxis itemX = engEff.XAxis;
            Console.WriteLine("Max sites: " + itemX.MaxSites);
            Console.WriteLine("Used sites: " + itemX.UsedSites);
            Console.WriteLine("Site value by index: " + itemX.Site[0].Display.DisplayValue);

            Console.WriteLine("ReadOnly: " + itemX.ReadOnly);
            Console.WriteLine("Visible: " + itemX.Visible);

            IMtcTableAxis itemY = engEff.YAxis;
            Console.WriteLine("Max sites: " + itemY.MaxSites);
            Console.WriteLine("Used sites: " + itemY.UsedSites);
            Console.WriteLine("Site value by index: " + itemY.Site[0].Display.DisplayValue);

            Console.WriteLine("ReadOnly: " + itemY.ReadOnly);
            Console.WriteLine("Visible: " + itemY.Visible);

            IMtcTableAxis itemZ = engEff.ZAxis;
            Console.WriteLine("Max sites: " + itemZ.MaxSites);
            Console.WriteLine("Used sites: " + itemZ.UsedSites);
            Console.WriteLine("Site value by index: " + itemZ.Site[0].Display.DisplayValue);

            Console.WriteLine("ReadOnly: " + itemZ.ReadOnly);
            Console.WriteLine("Visible: " + itemZ.Visible);
        }

        void TestIMtcAdjustItemInterface()
        {
            var pkg = GetMainPackage();
            var tables = pkg.Tables;
            const string tableName = "Engine Efficiency Adaption";
            var engEff = tables[tableName];
            PrintTable(engEff);

            IMtcAdjustItem item = engEff;
            Console.WriteLine("DisplayName: " + item.DisplayName);
            try
            {
                Console.WriteLine("Enumeration: " + item.Enumeration);
            }
            catch (Exception e) {
                Console.WriteLine(e.ToString());
            }
            try
            {
                Console.WriteLine("MinMaxValidator Min: " + item.MinMaxValidator.Min.DisplayValue + item.MinMaxValidator.Min.DisplayUnit + Environment.NewLine + "MinMaxValidator Max: " + item.MinMaxValidator.Max.DisplayValue + item.MinMaxValidator.Max.DisplayUnit);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            Console.WriteLine("ReadOnly: " + item.ReadOnly);
            Console.WriteLine("Visible: " + item.Visible);
            Console.WriteLine("DataType: " + item.DataType);

        }

        void Test21270()
        {
            //set password in ecu
            m_tuneApp.Devices.Connect(2851);

            IMtcRecentFile recentPkg = m_tuneApp.RecentPackages[0];
        }

            void Test21265()
        {
            m_tuneApp.Devices.Connect(2851);

            IMtcRecentFile recentPkg = m_tuneApp.RecentPackages[0];

            m_tuneApp.Packages.Load(recentPkg.Path, true);
        }

        void TestIMtcMinMaxValidatorInterface(IMtcMinMaxValidator v)
        {
            var min = v.Min;
            if (min != null)
            {
                Console.WriteLine("Min :" + min.DisplayValue + " " + min.DisplayUnit);
            }

            var max = v.Max;
            if (max != null)
            {
                Console.WriteLine("Max :" + max.DisplayValue + " " + max.DisplayUnit);
            }

            double value = 6;
            value = 6; Console.WriteLine("{0} is {1}", value, v.Validate(value, false));
            value = 6; Console.WriteLine("{0} is {1}", value, v.Validate(value));
            value = 15.0; Console.WriteLine("{0} is {1}", value, v.Validate(value, false));
            value = 15; Console.WriteLine("{0} is {1}", value, v.Validate(value));
            value = 40; Console.WriteLine("{0} is {1}", value, v.Validate(value));
            value = 200; Console.WriteLine("{0} is {1}", value, v.Validate(value));
            value = 201; Console.WriteLine("{0} is {1}", value, v.Validate(value));
        }

        void TestIMtcSiteValueInterface()
        {
            //testing of IMtcSite Display. Device is tested in SetParameter method
            var p = "Coolant Temperature Warning Maximum";
            var pkg = GetMainPackage();
            var parameterToChange = pkg.Parameters[p];
                Console.WriteLine("Found :" + p);
            Console.WriteLine("Value: " + parameterToChange.Site.Display.Value);
            Console.WriteLine("Unit: " + parameterToChange.Site.Display.Unit);
            Console.WriteLine("DisplayValue: " + parameterToChange.Site.Display.DisplayValue);
            Console.WriteLine("DisplayUnit: " + parameterToChange.Site.Display.DisplayUnit);
            Console.WriteLine("ValueType: " + parameterToChange.Site.Display.ValueType);
            parameterToChange.Site.Device.Value = 33.0;
            
        }

        void TestIMtcEnumeratorInterface(string channelName, string channelValue)
        {
            var pkg = GetMainPackage();

            var parameterToChange = pkg.Parameters[channelName];
            if (parameterToChange != null)
            {
                Console.WriteLine("Found: " + channelName);
                var v = parameterToChange.Enumeration.EnumeratorByDisplayName["CAN Bus 1"].Value;
                Console.WriteLine("Enum of CAN Bus 1: " + v);
                v = parameterToChange.Enumeration.EnumeratorByValue[v].Value;
                Console.WriteLine("Enum from value: " + v);
                var vv = parameterToChange.Enumeration.EnumeratorByValue[v].DisplayName;
                Console.WriteLine("Enum name from index: " + vv);
                v = parameterToChange.Enumeration.Name[v];
                Console.WriteLine("Whatever this is: " + v);

            }
        }

        void TestIMtcDAQValueInterface(string channelToSearchFor)
        {
            var allChannels = m_tuneApp.Packages[0].DAQ.GetRealTimeValueAll();

            string liveChannelValue = "";

            foreach (IMtcDAQValue channel in allChannels)
            {
                if (channel.DisplayName.Equals(channelToSearchFor))
                {
                    liveChannelValue = channel.DisplayValue + " " + channel.DisplayUnit;
                    Console.WriteLine(liveChannelValue);
                    Console.WriteLine("DisplayName: " + channel.DisplayName);
                    Console.WriteLine("DisplayValue: " + channel.DisplayValue);
                    Console.WriteLine("DisplayUnit: " + channel.DisplayUnit);
                    Console.WriteLine("Value: " + channel.Value);
                    Console.WriteLine("Time: " + channel.Time);
                    Console.WriteLine("Index: " + channel.Index);
                    break;
                }
            }
        }

        public object[] DAQRealTImeValue(string val)
        {
            return m_tuneApp.Packages[0].DAQ.GetRealTimeValue(new string[] { val });
        }

        void TestDAQGetRealTimeValueByChannelName()
        {
            ConnectToECU();
            //   string[] ids = { "Engine Speed", "ECU Uptime" };
            //   var ids = new string[] { "Engine Speed" };

            //       GetChannelValue("Engine Speed");
            var values = m_tuneApp.Packages[0].DAQ.GetRealTimeValue(new string[] { "Engine Speed" });
            /*
            foreach (IMtcDAQValue v in values)
            {
                double asd = v.Value;
                Console.WriteLine($"{v.DisplayName} = {asd} {v.DisplayUnit}");
                    }*/
            /*
            var e = m_tuneApp.Packages[0].DAQ.GetRealTimeValue("Engine Speed");

            Console.WriteLine(m_tuneApp.Packages[0].DAQ.Time);
        
            SecsToMicroSecs(m_tuneApp.Packages[0].DAQ.Time);
            var a = m_tuneApp.Packages[0].DAQ.GetValueAll(SecsToMicroSecs(5.12));
            */
        }
        void TestIMtcDAQInterface()
        {
            m_tuneApp.Devices.Connect(2851); //locked. Connect to device manually when testing without licence
            var a = m_tuneApp.Packages[0].DAQ.Active;
            var b = m_tuneApp.Packages[0].DAQ.Time;
            var c = m_tuneApp.Packages[0].DAQ.GetValue(1, 00.10);
            var d = m_tuneApp.Packages[0].DAQ.GetValueAll(00.10);
            var e = m_tuneApp.Packages[0].DAQ.GetRealTimeValue(1);
            var f = m_tuneApp.Packages[0].DAQ.GetRealTimeValueAll();
        }

        double SecsToMicroSecs(double s)
        {
            return s * 1e6;
        }

        void TestInterfaceIMtcDevices()
        {
            //breakpoint on all to allow for manual test setup
            var d = m_tuneApp.Devices[0]; //this is "Item" and it should be unlocked (return something) //unlocked
            m_tuneApp.Devices.RetrieveLogData(2851); //locked BUG: should be unlocked
            m_tuneApp.Devices.Connect(2851); //locked
            m_tuneApp.Devices.Disconnect(2851); //locked
        }

        void TestInterfaceIMtcDevice()
        {
            CheckForWorkspace(); //unlocked

            foreach (IMtcDevice d in m_tuneApp.Devices)
            {
                Console.WriteLine("Name: " + d.Name); // 15:2851:Start GP:5
                Console.WriteLine("DisplayName: " + d.DisplayName); //M150 #2851 Start GP
                Console.WriteLine("DisplaySerial: " + d.DisplaySerial); //2851
                Console.WriteLine("Serial: " + d.Serial); //2851
                Console.WriteLine("Type: " + d.Type); //M150
                Console.WriteLine("Description: " + d.Description); //
                Console.WriteLine("Address: " + d.Address); //[fe80::72b3:d5ff:fe71:e9c3%13]:5555
                Console.WriteLine("NetworkInterface: " + d.NetworkInterface); //13
            }
            //result in comments above. Same result without licence. Methods correctly unlocked
        }

        void PrintAllPackages()
        {
            CheckForInstalledPackage();

            var installedPkgs = m_tuneApp.InstalledPackages;
            Console.WriteLine($"Total Installed Packages : {installedPkgs.Count}"); // Prints the number of installed packages to the console

            foreach (IMtcInstalledPackage p in installedPkgs) // Prints details of all installed packages to the console
            {
                PrintInstalledPackage(p);
            }
        }

        void CheckForWorkspace()
        {
            if (m_recentWorkspace == null && m_tuneApp.RecentWorkspaces.Count > 0)
            {
                m_recentWorkspace = m_tuneApp.RecentWorkspaces[0];
                if (m_recentWorkspace != null)
                {
                    var f = m_recentWorkspace.Path;
                    m_tuneApp.WorkspaceLoad(f); //Loads the most recently used workspace}
                    //m_tuneApp.WorkspaceLoad("C:\\Users\\mila\\Documents\\MoTeC\\M1\\Tune\\Workspaces\\Tune 1"); //Loads workspace by file path
                }
            }

            if (m_recentWorkspace == null)
            {
                throw new Exception("ERROR: No recent workspace was found");
            }
        }

        void ConnectToECU()
        {
            CheckForWorkspace();
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
            //This takes us to screen where we select which sectors to extract from.
            //it requires user interaction from here.
        }

        void LoadRecentPackage()
        {
            CheckForWorkspace();
            CheckForRecentPackage();
 
            IMtcRecentFile recentPkg = m_tuneApp.RecentPackages[0];
            m_tuneApp.Packages.Load(recentPkg.Path, true);
        }

        void LoadPackageByName()
        {
            const string pkgFileName = "Generic 4 cylinder, MAP based Efficiency Migration Base v5"; //The name of the package we want to open
            const string ecuModel = "M150"; //The hardware device type          
            bool foundPackage = false;

            CheckForWorkspace();
            CheckForInstalledPackage();

            var installedPkgs = m_tuneApp.InstalledPackages;

            foreach (IMtcInstalledPackage p in installedPkgs)
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

        void GetChannelValue(string channelToSearchFor)
        {
            ConnectToECU();
            CheckForMainPackage();

            var allChannels = m_tuneApp.Packages[0].DAQ.GetRealTimeValueAll();
            if (allChannels != null && allChannels.Length > 0)
            {
                string liveChannelValue = "";

                foreach (IMtcDAQValue channel in allChannels)
                {
                    if (channel.DisplayName.Equals(channelToSearchFor))
                    {
                        liveChannelValue = channel.DisplayValue + " " + channel.DisplayUnit;
                        break;
                    }
                }
                if (string.IsNullOrEmpty(liveChannelValue))
                {
                    Console.WriteLine($"Searched {allChannels.Length} channels. Did not find {channelToSearchFor}");
                }
                else
                    Console.WriteLine($"{channelToSearchFor} channel found. Current value is {liveChannelValue}");
            }
            else
            {
                Console.WriteLine("Failed to fetch channels. Check package");
            }
        }

        void TuneIATParameter()
        {
            //IAT channel has no resource assigned
            //Default is 35.0C
            //Change default value and check IAT channel

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
                    Console.WriteLine(e.ToString()); //make sure the input string channelValue can be converted to a double. If yes, might fail due to COM failure
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
                catch (Exception e)
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

        void PrintTable()
        {
            ConnectToECU();
            
            var pkg = GetMainPackage();
            var tables = pkg.Tables;
            PrintTable(tables["Engine Efficiency"]);
        }

        void TuneTable()
        {
            ConnectToECU();
            
            var pkg = GetMainPackage();
            TuneParameter("Airbox Temperature Sensor Resource", "11");

            SavePackage(pkg); //Setting a resource requires package save
            CheckECUConnectionStatus(true); //Setting resource requires ECU reset following a save. Ensure we re-connect successfully.

            var tables = pkg.Tables;
            IMtcTable t = tables["Airbox Temperature Sensor Translation"];
            if (t != null)
            {
                double[] x = { 1.000, 1.500, 2.000, 2.500, 3.000, 3.500, 4.000 }; //The voltage values we want on the x axis

                t.ReShape(true, x, false, null, false, null, true);

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

        static void PrintInstalledPackage(IMtcInstalledPackage pkg)
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

                PrintAjustItem(t);

                PrintTableAxis(t.XAxis, "X");
                PrintTableAxis(t.YAxis, "Y");
                PrintTableAxis(t.ZAxis, "Z");
            }
        }

        static void PrintAjustItem(IMtcAdjustItem item)
        {
            Console.WriteLine($"{item.DisplayName} ({(item.ReadOnly ? "Readonly" : "Editable")} and {(item.Visible ? "Visible" : "Invisible")}) : {item.DataType}");
            PrintEnumeration(item.Enumeration);
        }

        static void PrintTableAxis(IMtcTableAxis axis, string name)
        {
            if (axis != null)
            {
                Console.WriteLine($"Axis {name}: ");
                PrintAjustItem(axis);
                Console.WriteLine($"\tMax Sites : {axis.MaxSites}");
                Console.WriteLine($"\tUsed Sites : {axis.UsedSites}");
                Console.WriteLine($"\tData Type : {axis.DataType}");

                Console.Write("\tValues:");
                for (uint i = 0; i < axis.UsedSites; i++)
                {
                    Console.Write(axis.Site[i].Display.DisplayValue);
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
    }
}
