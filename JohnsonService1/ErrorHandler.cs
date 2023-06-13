using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace JohnsonService
{
    public enum ErrorType
    {
        Critical,
        Minor,
        Information
    }

    class ErrorHandler
    {
        private const String origen = "Servicio Johnson";
        private const String log = "Application"; //Aplicación para Windows en español

        internal static void RegistrarError(String msg, ErrorType type, Exception ex)
        {
            try
            {
                if (!EventLog.SourceExists(origen,"."))
                    EventLog.CreateEventSource(origen, log);

                if (type == ErrorType.Information)
                    EventLog.WriteEntry(origen, ex.ToString(), EventLogEntryType.Information);
                else if (type == ErrorType.Minor)
                    EventLog.WriteEntry(origen, ex.ToString(), EventLogEntryType.Warning);
                else
                    EventLog.WriteEntry(origen, ex.ToString(), EventLogEntryType.Error);
            }
            catch { }
        }

        internal static void ManejarError(String msg, ErrorType type, Exception ex)
        {
            //registrar error
            RegistrarError(msg, type, ex);

            //mostrar mensaje si es conveniente
            if (type == ErrorType.Critical)
            {
                EventLog.WriteEntry(origen, msg, EventLogEntryType.Error);
            }
        }

    }
}
