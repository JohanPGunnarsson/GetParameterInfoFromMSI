using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using WindowsInstaller;



namespace GetProductInfoFromMSI
{
    class Program
    {
        public static string GetMsiProperty( string msiFile, string property )
        {
           string retVal = string.Empty;

           // Create an Installer instance
           Type classType = Type.GetTypeFromProgID( "WindowsInstaller.Installer" );
           Object installerObj = Activator.CreateInstance( classType );
           Installer installer = installerObj as Installer;

           // Open the msi file for reading
           // 0 - Read, 1 - Read/Write
           Database database = installer.OpenDatabase( msiFile, 0 );

           // Fetch the requested property
           string sql = String.Format(
              "SELECT Value FROM Property WHERE Property='{0}'", property );
           View view = database.OpenView( sql );
           view.Execute( null );

           // Read in the fetched record
           Record record = view.Fetch();
           if ( record != null )
              retVal = record.get_StringData( 1 );

           return retVal;
        }

        /// <summary>
        /// Obtain info from MSI archive
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            if (2 == args.Length)
                Console.WriteLine(GetMsiProperty(args[0], args[1]));
            else
                Console.WriteLine("no msi specified");
        }
    }
}
