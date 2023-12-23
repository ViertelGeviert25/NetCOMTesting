using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyCOMLib
{


    [ComVisible(true)]
    [Guid("E7A98F6A-F8D7-47A1-8F0E-ECFCB8A07C98")]
    public interface IMyInterface
    {
        void MyMethod(string message);
        string MyProperty { get; set; }
    }

    [ComVisible(true)]
    [Guid("7FE8C9B1-0781-4AFB-951E-8B8C4787D1E1")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IMyInterface))]
    public class MyClass : IMyInterface
    {
        public void MyMethod(string message)
        {
            MessageBox.Show(message);
        }

        public string MyProperty { get; set; }
    }


    public static class ComRegistration
    {
        [ComRegisterFunction]
        public static void Register(Type t)
        {
            string keyName = string.Format(@"CLSID\{{{0}}}\Programmable", t.GUID);
            using (RegistryKey key = Registry.ClassesRoot.CreateSubKey(keyName))
            {
                key.SetValue(null, "");
                key.Close();
            }
        }

        [ComUnregisterFunction]
        public static void Unregister(Type t)
        {
            string keyName = string.Format(@"CLSID\{{{0}}}\Programmable", t.GUID);
            Registry.ClassesRoot.DeleteSubKey(keyName, false);
        }
    }
}
