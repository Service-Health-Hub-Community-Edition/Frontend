using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Runtime.InteropServices;

namespace Microsoft.ServiceHealthHub.Core
{
    [Guid("B343065D-6F77-4B7C-8348-49F6400DB9A2")]
    public interface IMSDatabaseConnector
    {
        List<MSPropertyBag> GetDataSet(string command, MSPropertyBag parameters = null);
        void ExecuteQuerySequence(string command);
        int ExecuteQuery(string command, MSPropertyBag parameters = null);
        int ExecuteStoredProcedure(string storedProcedureName, MSPropertyBag parameters = null);
    }
}