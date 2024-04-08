using LSEXT;
using LSSERVICEPROVIDERLib;
using Oracle.ManagedDataAccess.Client;
using Patholab_DAL_V1;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PrintVegaCassete
{
    [ComVisible(true)]
    [ProgId("PrintVegaCassete.PrintVegaCassete")]
    public class PrintVegaCassete : IWorkflowExtension
    {
        INautilusServiceProvider sp;
        INautilusDBConnection ntlsCon;
        OracleConnection connection = null;
        OracleCommand cmd = null;
        DataLayer dal = null;

        public void Execute(ref LSExtensionParameters Parameters)
        {
            try
            {
                sp = Parameters["SERVICE_PROVIDER"];

                long wnid = Parameters["WORKFLOW_NODE_ID"]; // wnid = Workflow Node ID

                ntlsCon = null;
                if (sp != null)
                {
                    ntlsCon = sp.QueryServiceProvider("DBConnection") as NautilusDBConnection;
                }
                else
                {
                    return;
                }

                if (ntlsCon != null)
                {
                    connection = GetConnection(ntlsCon);
                    cmd = connection.CreateCommand();

                    dal = new DataLayer();
                    dal.Connect(ntlsCon);
                }
                else 
                {
                    throw new Exception("can't get nautilus connection.");
                }

                var sql = string.Format("select parent_id from lims_sys.workflow_node where workflow_node_id={0}", wnid);

                cmd = new OracleCommand(sql, connection);
                var parentNodeId = cmd.ExecuteScalar();

                string printerEventName = string.Empty;
                if (parentNodeId != null)
                {
                    sql = string.Format("select LONG_NAME from lims_sys.workflow_node where workflow_node_id={0}", parentNodeId);

                    cmd = new OracleCommand(sql, connection);
                    printerEventName = Convert.ToString(cmd.ExecuteScalar());

                }

                PHRASE_ENTRY entry = dal.FindBy<PHRASE_ENTRY>(pe => pe.PHRASE_HEADER.NAME.Equals("Vega Printer") && pe.PHRASE_INFO.Equals(printerEventName)).FirstOrDefault();
                MessageBox.Show(entry.PHRASE_DESCRIPTION.ToString());
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        public OracleConnection GetConnection(INautilusDBConnection ntlsCon)
        {
            OracleConnection connection = null;
            if (ntlsCon != null)
            {
                //initialize variables
                string rolecommand;
                //try catch block
                try
                {

                    string connectionString;
                    string server = ntlsCon.GetServerDetails();
                    string user = ntlsCon.GetUsername();
                    string password = ntlsCon.GetPassword();

                    connectionString =
                        string.Format("Data Source={0};User ID={1};Password={2};", server, user, password);

                    if (string.IsNullOrEmpty(user))
                    {
                        connectionString = "User Id=/;Data Source=" + server + ";Connection Timeout=60;";
                    }

                    //create connection
                    connection = new OracleConnection(connectionString);

                    //open the connection
                    connection.Open();

                    //get lims user password
                    string limsUserPassword = ntlsCon.GetLimsUserPwd();

                    //set role lims user
                    if (limsUserPassword == "")
                    {
                        //lims_user is not password protected 
                        rolecommand = "set role lims_user";
                    }
                    else
                    {
                        //lims_user is password protected
                        rolecommand = "set role lims_user identified by " + limsUserPassword;
                    }

                    //set the oracle user for this connection
                    OracleCommand command = new OracleCommand(rolecommand, connection);

                    //try/catch block
                    try
                    {
                        //execute the command
                        command.ExecuteNonQuery();
                    }
                    catch (Exception f)
                    {
                        //throw the exeption
                    }

                    //get session id
                    double sessionId = ntlsCon.GetSessionId();

                    //connect to the same session 
                    string sSql = string.Format("call lims.lims_env.connect_same_session({0})", sessionId);

                    //Build the command 
                    command = new OracleCommand(sSql, connection);

                    //execute the command
                    command.ExecuteNonQuery();
                }
                catch (Exception e)
                {
                    //throw the exeption
                }
            }
            return connection;
        }
    }
}
