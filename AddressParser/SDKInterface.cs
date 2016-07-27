using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LOSAutomation;

namespace AddressParser
{
    class SDKInterface
    {
        public static void Login(string username, string password, string connection)
        {
            SDKSession session = new SDKSession();
            session.Login(username, password, connection);
            session.Authorize("Atlantic Bay Mortgage", "36040", "773329768226018262");
        }

        public static SDKApplication GetApplication(SDKSession session)
        {
            return session.GetApplication();
        }

        public static SDKFile GetFile(SDKApplication app, string filename)
        {
            return app.OpenFile(filename, false);
        }

        public static void AlterField(SDKFile file, string fieldname, string value)
        {
            file.SetFieldValue(fieldname, value);
        }

        public static void CloseAndSave(SDKApplication app, SDKFile file)
        {
            file.Save();
            app.CloseFile(file);
        }
    }
}
