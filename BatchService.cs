using ADBatchImportForInterProd;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;

namespace ADBatchImportforInterProd
{
    public class BatchService : IBatchService
    {
        private static string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;

        public string getElapsedTime(Stopwatch stopWatch) => String.Format("{0:00}:{1:00}:{2:00}.{3:00}", stopWatch.Elapsed.Hours, stopWatch.Elapsed.Minutes, stopWatch.Elapsed.Seconds, stopWatch.Elapsed.Milliseconds / 10);

        public void ExecuteNonQuery(string v)
        {

            SqlCommand cmd = new SqlCommand(v, new SqlConnection(connectionString));
            cmd.Connection.Open();
            cmd.ExecuteNonQuery();
            cmd.Connection.Close();
        }

        public string[] GetColumnNames() =>
         new string[]
         {
                "cn",
                "extensionAttribute1",
                "extensionAttribute13",
                "extensionAttribute3",
                "extensionAttribute4",
                "extensionAttribute7",
                "givenName",
                "initials",
                "l",
                "postalCode",
                "sn",
                "streetAddress",
                "telephoneNumber",
                "whenChanged",
                "whenCreated"
         };

        public string GetDBValue(string source) => (string.IsNullOrEmpty(source) || string.IsNullOrWhiteSpace(source)) ? null : source;

        public void WriteToEventLog(string message)
        {
            SqlConnection con = new SqlConnection(connectionString);
            con.Open();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = con;
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "Insert into Portal.EventLog (EventMessage, [Type], [Source], CreateDate) values (@message, 'Batch','ADExtract', GETDATE())";
            cmd.Parameters.AddWithValue("message", message);
            cmd.ExecuteNonQuery();
            con.Close();
        }

        public string getCWOPAInsertStatement() =>
         @"INSERT INTO cwopa_agency_file
            select REPLICATE('0', 8-LEN([extensionAttribute1])) + [extensionAttribute1] AS [extensionAttribute1_padded]
            -- old method:
            --        case len(extensionAttribute1)        --employee_num contains employeeID from active directory
             --         when 6 then '00' + substring(extensionAttribute1 , 1 ,6) 
             --         when 8 then '00' + substring(extensionAttribute1 , 3 ,6) 
             --         end
            ,substring(givenName , 1, 50)         --first_name  
                  ,substring(sn , 1, 50)               --last_name
                  ,substring(initials , 1, 30)         --name_middle
                  ,substring(extensionAttribute13 , 1, 50) --email_address
                  ,substring(cn , 1, 20)                   --domain_name
                  ,substring(telephoneNumber , 1, 30)      --work_phone
                  ,substring(streetAddress , 1, 100)       --work_address
                  ,substring(l  , 1, 100)                  --work_city
                  ,substring(postalCode , 1, 15)           --work_zip
                  ,substring(extensionAttribute3  , 1, 255)  --deputate
                  ,substring(extensionAttribute4 , 1, 255)   --bureau
                  ,substring(extensionAttribute7 , 1, 255)   --division
    from  ad_dep_users 
    where extensionAttribute1 is  not null
      and extensionAttribute1 like REPLICATE('0', 8-LEN([extensionAttribute1])) + [extensionAttribute1]   
    --     and extensionAttribute1 not like 'X%'
--    and extensionAttribute1 in (select  PERNR from    
    -- EIS_SAP_VIEW)
     ;
";       

    }
}
