using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ePubReader
{
    class WebBrowserUtility
    {
        public enum InternetExplorerEmulationMode
        {
            IE11NoMatterWhat = 11001,
            IE10NoMatterWhat = 10001,
            IE10IfPageHasCorrectDoctype = 10000,
            IE9NoMatterWhat = 9999,
            IE9IfPageHasCorrectDoctype = 9000,
            IE8NoMatterWhat = 8888,
            IE8IfPageHasCorrectDoctype = 8000,
            IE7IfPageHasCorrectDoctypeDEFAULT = 7000,
            BestIEForOS = 0
        }

        private const int MAJOR_WINDOWS_VERSION_WHEN_THIS_CODE_WAS_WRITTEN = 6;
        private const int MINOR_WINDOWS_VERSION_WHEN_THIS_CODE_WAS_WRITTEN = 2;

        /// <summary>
        /// Writes the HKEY_CURRENT_USER registry key necessary for any WebBrowser components used in this
        /// project to emulate a particular version of Internet Explorer. If you don't do this then any
        /// WebBrowser objects will run as IE7. 
        /// http://msdn.microsoft.com/en-us/library/ee330730(v=VS.85).aspx
        /// http://msdn.microsoft.com/en-us/library/ee330730%28VS.85%29.aspx#browser_emulation
        /// http://stackoverflow.com/questions/4612255/regarding-ie9-webbrowser-control
        /// </summary>
        static public void SetWebBrowserEmulation(InternetExplorerEmulationMode ieMode, bool applyOnFutureVersions)
        {
            int majorVersion = System.Environment.OSVersion.Version.Major;
            int minorVersion = System.Environment.OSVersion.Version.Minor;
            // Windows 2000 is version 5, Windows XP is version 5.1.
            // Vista is 6, Windows 7 is 6.1, Windows 8 is 6.2.
            // IE9 and IE10 are not available on Windows XP or earlier.
            // IE10 is not available on Windows Vista or earlier
            // If the version number is greater than X, where X = the versions released when this code was last
            // updated, then the setting will only be set if applyOnFutureVersions is true.
            if (ieMode == InternetExplorerEmulationMode.IE10IfPageHasCorrectDoctype || ieMode == InternetExplorerEmulationMode.IE10NoMatterWhat)
            {
                if (majorVersion < 6 && minorVersion < 1)
                    throw new ArgumentException("Cannot request IE10 on Windows versions before Windows 7");
            }
            else if (ieMode == InternetExplorerEmulationMode.IE9IfPageHasCorrectDoctype || ieMode == InternetExplorerEmulationMode.IE9NoMatterWhat)
            {
                if (majorVersion < 6)
                    throw new ArgumentException("Cannot request IE9 on Windows versions before Windows Vista");
            }
            if (majorVersion > MAJOR_WINDOWS_VERSION_WHEN_THIS_CODE_WAS_WRITTEN && !applyOnFutureVersions)
            {
                // Do not set. Later version of Windows. 
            }
            else if (majorVersion == MAJOR_WINDOWS_VERSION_WHEN_THIS_CODE_WAS_WRITTEN && minorVersion > MINOR_WINDOWS_VERSION_WHEN_THIS_CODE_WAS_WRITTEN && !applyOnFutureVersions)
            {
                // Do not set. Later version of Windows.
            }
            else
            {
                // OK, checked versions, we're good!
                Microsoft.Win32.RegistryKey regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION", true);
                if (regKey == null ) 
                {
                    Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Microsoft\\Internet Explorer\\Main\\FeatureControl");
                    regKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION");
                }
                string appExeName = AppDomain.CurrentDomain.FriendlyName;
                //appExeName = appExeName.Replace(".vshost", "");
                
                // What do we write? Depends on ieMode.
                if (ieMode == InternetExplorerEmulationMode.BestIEForOS)
                {
                    if (majorVersion < 6)
                    {
                        // Windows XP or earlier: Maximum Internet Explorer 8
                        regKey.SetValue(appExeName, InternetExplorerEmulationMode.IE8NoMatterWhat, Microsoft.Win32.RegistryValueKind.DWord);
                    }
                    else if (majorVersion == 6 && minorVersion < 2)
                    {
                        // Windows Vista: Maximum Internet Explorer 9
                        regKey.SetValue(appExeName, InternetExplorerEmulationMode.IE9NoMatterWhat, Microsoft.Win32.RegistryValueKind.DWord);
                    }
                    else
                    {
                        // Windows 7 or later: at time of writing, maximum Internet Explorer 10
                        regKey.SetValue(appExeName, InternetExplorerEmulationMode.IE10NoMatterWhat, Microsoft.Win32.RegistryValueKind.DWord);
                    }
                }
                else
                {
                    // Need a breakpoint here or this code does nothing. 
                    regKey.SetValue(appExeName, ieMode, Microsoft.Win32.RegistryValueKind.DWord);
                    
                }
                regKey.Close();
            }
        }

    }
}
