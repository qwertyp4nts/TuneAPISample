using M1Tune;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace TuneAPI
{
    class SampleAPI
    {
        private Tune m_tune;
        IMtcM1TuneApplication1 m_tuneApp;
        uint m_ConnectedECUSerialNumber;
        IMtcRecentFile m_recentWorkspace;

        public static void Main(string[] args)
        {
            SampleAPI api = new SampleAPI();

            if (api.Initialise())
            {
                while (true)
                {
                    PrintIntroHelp();
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
        }

        bool Initialise()
        {
            try
            {
                m_tune = new Tune(); //#TODO should fail if activation doesnt exist. if not, add another line
                m_tune.Visible = true; // Makes the Tune application visible to the user
                m_tuneApp = m_tune as IMtcM1TuneApplication1;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return m_tune != null;
        }

        static void PrintIntroHelp()
        {
            Console.WriteLine("Choose from the following options to demo functionality:");
            Console.WriteLine("1: Print details of all installed packages to the console");
            Console.WriteLine("2: Check for Updates - @joe this will be removed");
            Console.WriteLine("3: Connect to ECU and retrieve package");
            Console.WriteLine("4: Download Logged Data");
            Console.WriteLine("5: Load Recent Package");
            Console.WriteLine("6: Load Package by Name");
            Console.WriteLine("7: Send a package to the ECU");
            Console.WriteLine("8: Retrieve package from ECU and save it as a new file");
            Console.WriteLine("9: Get all live channel values and print to console");
            Console.WriteLine("10: Get live value of channel");
            Console.WriteLine("11: Tune Parameter");
            Console.WriteLine("12: Print the engine efficiency table to the console");
            Console.WriteLine("13: Assign a resource and tune a table");
            Console.WriteLine("14: TEST FUNCTION - @joe this will be removed");
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
                        CheckForUpdates();
                        break;
                    case 3:
                        ConnectToECU();
                        break;
                    case 4:
                        DownloadLoggedData();
                        break;
                    case 5:
                        LoadRecentPackage();
                        break;
                    case 6:
                        LoadPackageByName();
                        break;
                    case 7:
                        SendPackage();
                        break;
                    case 8:
                        RetrievePackageFromECUandSave();
                        break;
                    case 9:
                        GetAllChannels();
                        break;
                    case 10:
                        GetChannelValue("ECU Uptime");
                        break;
                    case 11:
                        TuneIATParameter();
                        break;
                    case 12:
                        PrintTable();
                        break;
                    case 13:
                        TuneTable();
                        break;
                    case 14:
                        TestSave();
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

        void PrintAllPackages()
        {
            var installedPkgs = m_tuneApp.InstalledPackages;
            Console.WriteLine("Total Installed Packages : {0}", installedPkgs.Count); // Prints the number of installed packages to the console

            // Prints details of all installed packages to the console
            if (installedPkgs != null && installedPkgs.Count > 0)
            {
                foreach (IMtcInstalledPackage p in installedPkgs)
                {
                    PrintInstalledPackage(p);
                }
            }
            else
            {
                Console.WriteLine("No installed packages were found");
            }
        }

        void CheckForUpdates()
        {
            OpenWorkspace();
            m_tuneApp.CheckForUpdates();
        }

        void OpenWorkspace()
        {
            if (m_recentWorkspace == null && m_tuneApp.RecentWorkspaces.Count > 0)
            {
                m_recentWorkspace = m_tuneApp.RecentWorkspaces[0];
                if (m_recentWorkspace != null)
                {
                    var f = m_recentWorkspace.Path;
                    m_tuneApp.WorkspaceLoad(f); //Loads the most recently used workspace
                    //m_tuneApp.WorkspaceLoad("C:\\Users\\mila\\Documents\\MoTeC\\M1\\Tune\\Workspaces\\Tune 1"); //Loads workspace by file path
                }
                else
                {
                    Console.WriteLine("Recent workspace not available");
                }
            }
            else
            {
                Console.WriteLine("Workspace already open, or no recent workspaces available");
            }
        }

        void ConnectToECU()
        {
            OpenWorkspace();
            if (m_tuneApp.Devices.Count > 0)
            {
                IMtcDevice current = m_tuneApp.Devices[0];
                if (current != null)
                {
                    uint serialNum = current.Serial;
                    m_tuneApp.Devices.Connect(serialNum); //Connects to first device found
                    //m_tuneApp.Devices.Connect(2851); // -> Connects to device with target serial number

                    m_ConnectedECUSerialNumber = serialNum;
                }
            }
            else
            {
                Console.WriteLine("No ECU connections are found");
            }
        }

        void DownloadLoggedData()
        {
            ConnectToECU();

            if (m_ConnectedECUSerialNumber > 0)
            {
                m_tuneApp.Devices.RetrieveLogData(m_ConnectedECUSerialNumber);
                //This takes us to screen where we select which sectors to extract from. 
                //it requires user interaction from here.
            }
            else
            {
                Console.WriteLine("No ECU connections are found");
            }
        }

        void LoadRecentPackage()
        {
            OpenWorkspace();

            if (m_tuneApp.RecentPackages.Count > 0)
            {
                IMtcRecentFile recentPkg = m_tuneApp.RecentPackages[0];
                if (recentPkg != null)
                    m_tuneApp.Packages.Load(recentPkg.Path, true);
            }
            else
                Console.WriteLine("No recent packages available");
        }

        void LoadPackageByName()
        {
            const string pkgFileName = "Generic 4 cylinder, MAP based Efficiency Migration Base v5"; //The name of the package we want to open
            const string ecuModel = "M150"; //The hardware device type          

            OpenWorkspace();

            var installedPkgs = m_tuneApp.InstalledPackages;

            if (installedPkgs != null && installedPkgs.Count > 0)
            {
                bool foundPackage = false;

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
                {
                    Console.WriteLine($"Did not find file with name: {pkgFileName} and ECU model {ecuModel}");
                }
            }
            else
            {
                Console.WriteLine("No installed packages available");
            }
        }

        void SendPackage()
        {
            LoadPackageByName();
            var pkg = GetMainPackage();
            if (pkg != null)
            {
                if (m_tuneApp.Devices.Count > 0)
                {
                    IMtcDevice current = m_tuneApp.Devices[0];
                    if (current != null)
                    {
                        m_ConnectedECUSerialNumber = current.Serial;
                        Console.WriteLine($"Connected to ECU #{m_ConnectedECUSerialNumber}");
                        if (pkg.Send(m_ConnectedECUSerialNumber))
                            Console.WriteLine("Package sent successfully");
                        else
                            Console.WriteLine("Package send failed");
                    }
                }
                else
                {
                    Console.WriteLine("No valid ECUs found");
                }
            }
            else
            {
                Console.WriteLine("FAILED: No main package loaded");
            }
        }

        void RetrievePackageFromECUandSave()
        {
            ConnectToECU();
            var pkg = GetMainPackage();
            if (pkg != null)
            {
                pkg.SaveAs("Testing package save from API", "Generated by MoTeC Tune API");
            }
            else
            {
                Console.WriteLine("FAILED: No main package loaded");
            }
        }

        void GetAllChannels()
        {
            ConnectToECU();
            if (m_tuneApp.Packages.Count > 0)
            {
                var allChannels = m_tuneApp.Packages[0].DAQ.GetRealTimeValueAll();
                //Print all ECU channels and their values to the console
                if (allChannels != null)
                {
                    foreach (IMtcDAQValue channel in allChannels)
                    {
                        Console.WriteLine("{0}: {1} {2}", channel.DisplayName, channel.DisplayValue, channel.DisplayUnit);
                    }
                }
                else
                {
                    Console.WriteLine("Failed to fetch channels. Ensure package is loaded");
                }
            }
        }

        void GetChannelValue(string channelToSearchFor)
        {
            ConnectToECU();
            if (m_tuneApp.Packages.Count > 0)
            {
                var allChannels = m_tuneApp.Packages[0].DAQ.GetRealTimeValueAll();

                if (allChannels != null)
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
                        Console.WriteLine("Searched {0} channels. Did not find {1}", allChannels.Length, channelToSearchFor);
                    }
                    else
                        Console.WriteLine("{0} channel found. Current value is {1} {2}", channelToSearchFor, liveChannelValue);
                }
            }
            else
            {
                Console.WriteLine("Failed to fetch channels. Ensure package is loaded");
            }
        }

        void TuneIATParameter()
        {
            //IAT channel has no resource assigned
            //Default is 35.0C
            //Change default value and check IAT channel

            string c = "Inlet Air Temperature";
            string p = "Inlet Air Temperature Sensor Default";
            GetChannelValue(c); //Check value of IAT
            var pkg = GetMainPackage();
            if (pkg != null)
            {
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
                GetChannelValue(c); //Ensure channel now reports updated value
            }
            else
            {
                Console.WriteLine("FAILED: No main package loaded");
            }
        }

        void TuneParameter(string channelName, string channelValue, bool connected)
        {
            //Generic version of tuneIATParameter() to use in other methods
            //Not only does this function tune parameters, it also sets dropdown values. 
            //For example: to set ADR CAN Bus from 'Not in use' to 2, simply write: tuneParameter("ADR CAN Bus", "2");

            if (connected)
            {
                ConnectToECU(); //if connected is true, the package will be fetched from ECU and modified. If false, changes will be made to open offline package
            }

            var pkg = GetMainPackage();
            if (pkg != null)
            {
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
            else
            {
                Console.WriteLine("FAILED: No main package loaded");
            }
        }

        void PrintTable()
        {
            ConnectToECU();
            var pkg = GetMainPackage();
            if (pkg != null)
            {
                var tables = pkg.Tables;
                PrintTable(tables["Engine Efficiency"]);
            }
            else
            {
                Console.WriteLine("FAILED: No main package loaded");
            }

            throw new Exception("I'm AN EXCEPTIOn!");
        }

        void TuneTable()
        {
            ConnectToECU();
            var pkg = GetMainPackage();
            if (pkg != null)
            {
                TuneParameter("Airbox Temperature Sensor Resource", "11", false);

                SavePackage(pkg);

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
            }
            else
            {
                Console.WriteLine("FAILED: No main package loaded");
            }
        }

        void TestSave()
        {
            LoadPackageByName();
            var pkg = GetMainPackage();
            if (pkg != null)
            {
                TuneParameter("Airbox Temperature Sensor Resource", "11", false);
                SavePackage(pkg);
            }
            else
            {
                Console.WriteLine("FAILED: No main package loaded");
            }
        }

        void Exit()
        {
            m_tuneApp.Exit();
            Environment.Exit(0);
        }

        static void PrintInstalledPackage(IMtcInstalledPackage pkg)
        {
            Console.WriteLine("File Name : {0}", pkg.FileName);
            Console.WriteLine("\tFile VehicleId : {0}", pkg.VehicleId);
            Console.WriteLine("\tFile SerialNumber : {0}", pkg.SerialNumber);
            Console.WriteLine("\tFile Comment : {0}", pkg.Comment);
            Console.WriteLine("\tFile FirmwareVersionName : {0}", pkg.FirmwareVersionName);
            Console.WriteLine("\tFile FirmwareVersion : {0}", pkg.FirmwareVersion);
            Console.WriteLine("\tFile Hardware : {0}", pkg.Hardware);
            Console.WriteLine("\tFile ModifiedDateTime : {0}", pkg.ModifiedDateTime);
            Console.WriteLine("\tFile ImportedDateTime : {0}", pkg.ImportedDateTime);
        }

        IMtcPackage3 GetMainPackage()
        {
            if (m_tuneApp.Packages != null && m_tuneApp.Packages.Count > 0)
                return m_tuneApp.Packages[0] as IMtcPackage3;
            return null;
        }

        void SavePackage(IMtcPackage3 pkg)
        {
            bool saved = pkg.Save();
            if (saved)
                Console.WriteLine("Package saved successfully");
        }

        static void PrintTable(IMtcTable t)
        {
            if (t != null)
            {
                Console.WriteLine("Table '{0}':", t.DisplayName);

                PrintAjustItem(t);

                PrintTableAxis(t.XAxis, "X");
                PrintTableAxis(t.YAxis, "Y");
                PrintTableAxis(t.ZAxis, "Z");
            }
        }

        static void PrintAjustItem(IMtcAdjustItem item, bool print_enum = true)
        {
            Console.WriteLine("{0} ({1} {2}) : {3}", item.DisplayName, (item.ReadOnly ? "RO" : "RW"), item.Visible ? "VIS" : "INV", item.DataType);
            if (print_enum)
                PrintEnumeration(item.Enumeration);
        }

        static void PrintTableAxis(IMtcTableAxis axis, string name)
        {
            if (axis != null)
            {
                Console.WriteLine("Axis {0}: ", name);
                PrintAjustItem(axis);
                Console.WriteLine("\tMax Sites : {0}", axis.MaxSites);
                Console.WriteLine("\tUsed Sites : {0}", axis.UsedSites);
                Console.WriteLine("\tData Type : {0}", axis.DataType);

                Console.Write("\tValues:");

                for (uint i = 0; i < axis.UsedSites; i++)
                {
                    Console.Write("({0}) ", (axis.Site[i].Display.DisplayValue));
                }
                Console.WriteLine();
            }
            else
                Console.WriteLine("Axis {0}:  (none)", name);
        }

        static void PrintEnumeration(IMtcEnumeration e)
        {
            if (e != null)
            {
                Console.WriteLine(e.Name);
                foreach (IMtcEnumerator v in e)
                    Console.WriteLine("{0} : {1}", v.Value, v.DisplayName);

                for (int i = 0; i < e.Count; i++)
                    Console.WriteLine("{0} : {1}", e[i].Value, e[i].DisplayName);
            }
        }
    }
}
