using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Security;
using System.Configuration;
using System.Xml;
using System.IO;
using System.Globalization;
using CxCompress;
using System.IO.Compression;




/*! \mainpage Overview
*
* \section first_sec Class Information
*
* This class contains the methods, which are used for the connection and query of the Camaix Tool Database. <br>The Database exists on a SQL-Express Server and 
* can be connected by using the methods from this class. 
*
* \subsection first_subsec Versionhistory
*
* Version: 1.00<br>
* Author: S. Schwark<br>
* Date: 23.09.2014<br>
* Copyright: Camaix Gmbh, Eschweiler 2014<br>
* 
* \subsection second_subsec Additional Documents
*
*
*/

namespace CSqlServerDb
{
    //!  A class for the communication with the Camaix Tool Database.
    /*!
      This class contains the methods for the commnication with a MS-SQL Server, where the Database is stored. The Camaix Tool Database is a DIN 4000
      compliant Database. All the following methods are written for the communication with the Database Server. 
    */
    public class CSqlQueries
    {


        public class filterParams
        {
            public bool pCheckDia;
            public double pDiaMin;
            public double pDiaMax;
            public bool pCheckLen;
            public double pLenMin;
            public double pLenMax;
            public bool pCheckShaftDia;
            public double pShaftDiameterMin;
            public double pShaftDiameterMax;
            public bool pCheckTipDia;
            public double pTipDiameterMin;
            public double pTipDiameterMax;
            public bool pCheckTaperHeight;
            public double pTaperHeightMin;
            public double pTaperHeightMax;
            public bool pCheckCoreDia;
            public double pCoreDiaMin;
            public double pCoreDiaMax;
            public bool pCheckCoreHeight;
            public double pCoreHeightMin;
            public double pCoreHeightMax;
            public bool pCheckTipAngle;
            public double pTipAngleMin;
            public double pTipAngleMax;
            public bool pCheckCuttingLength;
            public double pCuttingLengthMin;
            public double pCuttingLengthMax;
            public bool pCheckShoulderLength;
            public double pShoulderLengthMin;
            public double pShoulderLengthMax;
        }

        private string debugMessage = null;


        //!  A class for the storage of a Mill Process Data.
        /*!
          This class contains the milling processes which can be selected for the technological datasets 
        */
        public class MillProcessData
        {

            private Dictionary<string, TechData> internalDictionary = new Dictionary<string, TechData>();

            public void Add(string millProcess, TechData technology)
            {
                if (!this.internalDictionary.ContainsKey(millProcess))
                {
                    this.internalDictionary.Add(millProcess, technology);
                }
            }


            public int Count()
            {
                return this.internalDictionary.Count();
            }

            public string Get_MillProcess(int id)
            {

                return internalDictionary.Keys.ElementAt(id);
            }

            public TechData Get_MillProcessData(int technology_id)
            {
                TechData technology = new TechData();
                for (int i = 1; i < (millProcessNames.Count + 1); i++)
                {
                    if (internalDictionary.ElementAt(technology_id).Key.Equals(millProcessNames[i.ToString()]))
                    {
                        technology = internalDictionary.ElementAt(technology_id).Value;
                    }

                }

                return technology;
            }

        }

        //!  A class for the storage of materials .
        /*!
          This class contains the techonological process data which is bound to a material  
        */

        public class TechData
        {

            private Dictionary<string, ProcessData> internalDictionary = new Dictionary<string, ProcessData>();

            public void Add(string material, ProcessData process)
            {
                if (!this.internalDictionary.ContainsKey(material))
                {
                    this.internalDictionary.Add(material, process);
                }
            }


            public int Count()
            {
                return this.internalDictionary.Count();
            }

            public string Get_Material(int id)
            {
                return internalDictionary.Keys.ElementAt(id);
            }

            public ProcessData Get_Process(int material_id)
            {
                ProcessData process = new ProcessData();
                for (int i = 0; i < materialNames.Count; i++)
                {
                    if (internalDictionary.ElementAt(material_id).Key.Equals(materialNames[(i + 1).ToString()]))
                    {
                        process = internalDictionary.ElementAt(material_id).Value;
                    }

                }

                return process;
            }

        }

        //!  A class for the storage of processing data.
        /*!
          This class contains the techonological process data. The process data is the class where the technological values can be found  
        */
        public class ProcessData
        {

            private Dictionary<string, Dictionary<string, string>> internalDictionary = new Dictionary<string, Dictionary<string, string>>();

            public void Add(string process, Dictionary<string, string> techvalues)
            {
                if (!this.internalDictionary.ContainsKey(process))
                {
                    this.internalDictionary.Add(process, techvalues);
                }
            }


            public int Count()
            {
                return this.internalDictionary.Count();
            }

            public Dictionary<string, string> Get_TechData(int num)
            {
                return internalDictionary.ElementAt(num).Value;
            }

            public string Get_Process(int num)
            {
                 return internalDictionary.ElementAt(num).Key;
            }
        }



        //!  Class for the storage of an assemly tool
        /*!
          This class contains the data which is needed to store a assembly into the database  
        */

        public class CAssembly
        {
            public long id;             //! The id of the assembly
            public long toolId;         //! Id of the tool
            public long holderId;       //! Id of the holder
            public string description;  //! Assembly description
            public string name;         //! Assembly name
            public string assemblyNr;   //! Number of the tool    
            public double reach;        //! tool reach
            public double diameter;     //! tool diameter
            public string assemblyId;   //! RFID of the tool    
        }

        //!  Class which contains the coordinates of a point in a 2D space.
        /*!
          This class is used for the creation of contour graphics.   
        */

        public class CPoint2d
        {
            public double X;
            public double Y;

            public CPoint2d(double ValueX, double ValueY)
            {
                X = ValueX;
                Y = ValueY;
            }
            public double m_GetX()
            {
                return X;
            }
            public double m_GetY()
            {
                return Y;
            }
        }   
     
        //private string m_pLastError;
        //public static SqlConnection m_pConnection = null;
        public static string userId = "sa";
        public static string userPassword = "cdiT21E#";
        public static string sqlServer = "localhost\\TOOLDB";
        public static string TrustedConnection = "no";

        public CSqlServerDb.CDefinition.CTool m_pTool = new CSqlServerDb.CDefinition.CTool();
        public static Dictionary<string, string> din4000keys = new Dictionary<string, string>();
        public static SortedDictionary<string, string> din4000names = new SortedDictionary<string, string>();
        public static Dictionary<string, string> materialNames = new Dictionary<string, string>();
        public static Dictionary<string, string> millProcessNames = new Dictionary<string, string>();
        public static Dictionary<string, string> processNames = new Dictionary<string, string>();
        public static Dictionary<string, string> techKeys = new Dictionary<string, string>();

        //----------------------------------------------------------------------------------

        /*! getSQLiteConnection initiates a connection to the SQL-Server and returns the 
         * open SQLConnection. If an error has occured, the m_pLastError Variable contains the 
         * last errormessage and the method returns null.
         * @return SQLConnection to the database 
        */

        public SqlConnection getSqlConnection()
        {
            try
            {
                SqlConnection sqlConnection = new SqlConnection("user id=" + userId + ";" +
                                                       "password=" + userPassword + ";server=" + sqlServer + ";" +
                                                       "Trusted_Connection=" + TrustedConnection + ";" +
                                                       "database=TOOLDB; " +
                                                       "connection timeout=30;" +
                                                       "MultipleActiveResultSets=True");
                sqlConnection.Open();
                return sqlConnection;
            }
            catch (Exception ex)
            {
                return null;
            }
        }


        //----------------------------------------------------------------------------------

        /*! getToolList() returns a list of the tools, which are stored in the database. 
          * @param SQLConnection to the database
          * @return Dictionary<string, Tuple<string, string>> containing id, name, type 
         */

        public Dictionary<string, Tuple<string, string>> getToolList(SqlConnection sqlConnection)
        {
            // Instantiate a dictionary for the tool list
            Dictionary<string, Tuple<string, string>> toolList = new Dictionary<string,Tuple<string,string>>();

            // Using block for the database queries
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the list of the tools.
                SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[tool].[id], [TOOLDB].[dbo].[tool].[name], [TOOLDB].[dbo].[tool].[type] FROM [TOOLDB].[dbo].[tool];", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = command.ExecuteReader();

                // Read the data until there are no more rows
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        // Add the description to the toolDescription dictionary
                        toolList.Add(Convert.ToString(sqlDataReader.GetInt64(0)), new Tuple<string, string>(sqlDataReader.GetString(1), sqlDataReader.GetString(2)));
                    }
                }
            }
            sqlConnection.Close();
            sqlConnection.Dispose();
            return toolList;
        }

        /*! getToolList() returns a list of the tools, which are stored in the database. */

        public Dictionary<string, Tuple<string, string>> getToolListFilter(SqlConnection sqlConnection)
        {
            // Instantiate a dictionary for the tool list
            Dictionary<string, Tuple<string, string>> toolList = new Dictionary<string, Tuple<string, string>>();

            // Using block for the database queries
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the list of the tools.
                SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[tool].[id], [TOOLDB].[dbo].[tool].[name], [TOOLDB].[dbo].[din4000_key].[name], [TOOLDB].[dbo].[tool_characteristic].[data_value] FROM [TOOLDB].[dbo].[tool] LEFT JOIN [TOOLDB].[dbo].[tool_characteristic] ON tool.id = tool_characteristic.tool_id LEFT JOIN [TOOLDB].[dbo].[din4000_key] ON [TOOLDB].[dbo].[din4000_key].[id] = [TOOLDB].[dbo].[tool_characteristic].[key_id] WHERE ([TOOLDB].[dbo].[tool_characteristic].[key_id] = 3 AND CAST ([TOOLDB].[dbo].[tool_characteristic].[data_value] AS DECIMAL(9,2)) = 10)", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = command.ExecuteReader();

                // Read the data until there are no more rows
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        // Add the description to the toolDescription dictionary
                        toolList.Add(Convert.ToString(sqlDataReader.GetInt64(0)), new Tuple<string, string>(sqlDataReader.GetString(1), sqlDataReader.GetString(2)));
                    }
                }
            }
            sqlConnection.Close();
            sqlConnection.Dispose();
            return toolList;
        }


        //----------------------------------------------------------------------------------

        /*! getToolTable() returns a list of the tools, which are stored in the database. */

        public DataTable getToolTable(SqlConnection sqlConnection, int NodeIndex)
        {
            // Instantiate a dictionary for the tool list
            DataTable toolList = new DataTable();

            string ToolType = "";

            switch (NodeIndex)
            {
                case 0:
                    ToolType = "ALL";
                    break;
                case 1:
                    ToolType = "BNN";
                    break;
                case 2:
                    ToolType = "FSN";
                    break;
                case 3:
                    ToolType = "BGN";
                    break;
                case 4:
                    ToolType = "RNN";
                    break;
            }

            // Using block for the database queries
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the list of the tools.
                string Filter = "";

                if (ToolType != "ALL")
                    Filter = "tool.type = '" + ToolType + "' AND ";

                SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[tool].[id], [TOOLDB].[dbo].[tool].[name] AS Name, [TOOLDB].[dbo].[tool].[type], [TOOLDB].[dbo].[din4000_key].[name], [TOOLDB].[dbo].[tool_characteristic].[data_value] AS Beschreibung FROM [TOOLDB].[dbo].[tool] LEFT JOIN [TOOLDB].[dbo].[tool_characteristic] ON tool.id = tool_characteristic.tool_id LEFT JOIN [TOOLDB].[dbo].[din4000_key] ON [TOOLDB].[dbo].[din4000_key].[id] = [TOOLDB].[dbo].[tool_characteristic].[key_id] WHERE " + Filter + "[TOOLDB].[dbo].[din4000_key].[name] = 'J22' AND [TOOLDB].[dbo].[din4000_key].[name] = 'J22'", sqlConnection);
                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = command.ExecuteReader();

                toolList.Load(sqlDataReader);

            }
            sqlConnection.Close();
            sqlConnection.Dispose();
            return toolList;
        }
        //----------------------------------------------------------------------------------

        //----------------------------------------------------------------------------------

        /*! getToolTable() returns a list of the tools, which are stored in the database. */

        public DataTable getToolTableComplete(SqlConnection sqlConnection)
        {
            // Instantiate a dictionary for the tool list
            DataTable toolList = new DataTable();

            // Using block for the database queries
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the list of the tools.
                //SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[tool].[id], [TOOLDB].[dbo].[tool].[name], [TOOLDB].[dbo].[tool].[type] FROM [TOOLDB].[dbo].[tool] WHERE [TOOLDB].[dbo].[tool].[type] = '" + ToolType + "' AND [TOOLDB].[dbo].[din4000_key].[name] = 'J22';", sqlConnection);
                //SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[tool].[id], [TOOLDB].[dbo].[tool].[name], [TOOLDB].[dbo].[tool].[type], [TOOLDB].[dbo].[din4000_key].[name], [TOOLDB].[dbo].[tool_characteristic].[data_value] FROM [TOOLDB].[dbo].[tool] LEFT JOIN [TOOLDB].[dbo].[tool_characteristic] ON tool.id = tool_characteristic.tool_id LEFT JOIN [TOOLDB].[dbo].[din4000_key] ON [TOOLDB].[dbo].[din4000_key].[id] = [TOOLDB].[dbo].[tool_characteristic].[key_id] WHERE [TOOLDB].[dbo].[tool].[type] = '" + ToolType + "' AND [TOOLDB].[dbo].[din4000_key].[name] = 'J22' AND [TOOLDB].[dbo].[din4000_key].[name] = 'J22'", sqlConnection);
                SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[tool].[id], [TOOLDB].[dbo].[tool].[name], [TOOLDB].[dbo].[tool].[type], [TOOLDB].[dbo].[din4000_key].[name], [TOOLDB].[dbo].[tool_characteristic].[data_value] FROM [TOOLDB].[dbo].[tool] LEFT JOIN [TOOLDB].[dbo].[tool_characteristic] ON tool.id = tool_characteristic.tool_id LEFT JOIN [TOOLDB].[dbo].[din4000_key] ON [TOOLDB].[dbo].[din4000_key].[id] = [TOOLDB].[dbo].[tool_characteristic].[key_id] WHERE [TOOLDB].[dbo].[din4000_key].[name] = 'J22' AND [TOOLDB].[dbo].[din4000_key].[name] = 'J22'", sqlConnection);
                //SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[tool].[id], [TOOLDB].[dbo].[tool].[name], [TOOLDB].[dbo].[din4000_key].[name], [TOOLDB].[dbo].[tool_characteristic].[data_value] FROM [TOOLDB].[dbo].[tool] LEFT JOIN [TOOLDB].[dbo].[tool_characteristic] ON tool.id = tool_characteristic.tool_id LEFT JOIN [TOOLDB].[dbo].[din4000_key] ON [TOOLDB].[dbo].[din4000_key].[id] = [TOOLDB].[dbo].[tool_characteristic].[key_id] WHERE ([TOOLDB].[dbo].[tool_characteristic].[key_id] = 3 AND CAST ([TOOLDB].[dbo].[tool_characteristic].[data_value] AS DECIMAL(9,2)) = 10)", sqlConnection);
                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = command.ExecuteReader();

                toolList.Load(sqlDataReader);

            }
            sqlConnection.Close();
            sqlConnection.Dispose();
            return toolList;
        }
        //----------------------------------------------------------------------------------

        /*! getToolTable() returns a list of the tools, which are stored in the database. */

        public DataTable getToolTableFilter(SqlConnection sqlConnection, string ToolType)
        {
            // Instantiate a dictionary for the tool list
            DataTable toolList = new DataTable();

            // Using block for the database queries
            using (sqlConnection)
            {
                string Filter = "";

                if (ToolType != "ALL")
                    Filter = "WHERE [TOOLDB].[dbo].[tool].[type] = '" + ToolType + "'";

                SqlCommand command = new SqlCommand("SELECT[TOOLDB].[dbo].[tool_characteristic].[tool_id] AS ID, [TOOLDB].[dbo].[tool].[name] AS Name, MAX(CASE WHEN[TOOLDB].[dbo].[tool_characteristic].[key_id] = '64' THEN[TOOLDB].[dbo].[tool_characteristic].[data_value] END) Beschreibung FROM[TOOLDB].[dbo].[tool_characteristic] LEFT JOIN[TOOLDB].[dbo].[tool] ON[TOOLDB].[dbo].[tool].[id] = [TOOLDB].[dbo].[tool_characteristic].[tool_id] " + Filter + " GROUP BY[TOOLDB].[dbo].[tool_characteristic].[tool_id], [TOOLDB].[dbo].[tool].[name] ORDER BY[TOOLDB].[dbo].[tool_characteristic].[tool_id]", sqlConnection);
                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = command.ExecuteReader();

                toolList.Load(sqlDataReader);

            }
            sqlConnection.Close();
            sqlConnection.Dispose();
            return toolList;
        }

        //----------------------------------------------------------------------------------

        /*! getToolTable() returns a list of the tools, which are stored in the database. */

        public DataTable getTurningToolTableFilter(SqlConnection sqlConnection, string ToolType)
        {
            // Instantiate a dictionary for the tool list
            DataTable toolList = new DataTable();

            // Using block for the database queries
            using (sqlConnection)
            {
                string Filter = "";

                if (ToolType != "ALL")
                    Filter = "WHERE [TOOLDB].[dbo].[tool].[type] = '" + ToolType + "'";

                SqlCommand command = new SqlCommand("SELECT[TOOLDB].[dbo].[turning_tool_characteristic].[turning_tool_id] AS ID, [TOOLDB].[dbo].[turning_tool].[name] AS Name, MAX(CASE WHEN[TOOLDB].[dbo].[turning_tool_characteristic].[key_id] = '64' THEN[TOOLDB].[dbo].[turning_tool_characteristic].[data_value] END) Beschreibung FROM[TOOLDB].[dbo].[turning_tool_characteristic] LEFT JOIN[TOOLDB].[dbo].[turning_tool] ON[TOOLDB].[dbo].[turning_tool].[id] = [TOOLDB].[dbo].[turning_tool_characteristic].[turning_tool_id] " + Filter + " GROUP BY[TOOLDB].[dbo].[turning_tool_characteristic].[turning_tool_id], [TOOLDB].[dbo].[turning_tool].[name] ORDER BY[TOOLDB].[dbo].[turning_tool_characteristic].[turning_tool_id]", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = command.ExecuteReader();

                toolList.Load(sqlDataReader);

            }
            sqlConnection.Close();
            sqlConnection.Dispose();
            return toolList;
        }
        //----------------------------------------------------------------------------------
        //----------------------------------------------------------------------------------

        /*! getToolTable() returns a list of the tools, which are stored in the database. */

        public DataTable getAssemblyTableFilter(SqlConnection sqlConnection, string ToolType)
        {
            // Instantiate a dictionary for the tool list
            DataTable toolList = new DataTable();

            // Using block for the database queries
            using (sqlConnection)
            {
                string Filter = "";

                if (ToolType != "ALL")
                    Filter = "WHERE [TOOLDB].[dbo].[tool].[type] = '" + ToolType + "'";

                SqlCommand command = new SqlCommand("SELECT[TOOLDB].[dbo].[assembly].[id], [TOOLDB].[dbo].[assembly].[assembly_nr], [TOOLDB].[dbo].[assembly].[name], [TOOLDB].[dbo].[assembly].[description], [TOOLDB].[dbo].[assembly].[assembly_id] FROM [TOOLDB].[dbo].[assembly] LEFT JOIN [TOOLDB].[dbo].[tool] ON [TOOLDB].[dbo].[assembly].[tool_id] = [TOOLDB].[dbo].[tool].[id] " + Filter, sqlConnection);
                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = command.ExecuteReader();

                toolList.Load(sqlDataReader);

            }
            sqlConnection.Close();
            sqlConnection.Dispose();
            return toolList;
        }

        //----------------------------------------------------------------------------------

        /*! getToolList() returns a list of the tools, which are stored in the database. */

        public Dictionary<string, Tuple<string, string>> getToolListByParam(SqlConnection sqlConnection, Dictionary<string, string> toolData)
        {
            // Instantiate a dictionary for the tool list
            Dictionary<string, Tuple<string, string>> toolList = new Dictionary<string, Tuple<string, string>>();

            // Using block for the database queries
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the list of the tools.
                //SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[tool].[id], [TOOLDB].[dbo].[tool].[name], [TOOLDB].[dbo].[tool].[type] FROM [TOOLDB].[dbo].[tool];", sqlConnection);
                SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[tool].[id], [TOOLDB].[dbo].[tool].[name], [TOOLDB].[dbo].[din4000_key].[name], [TOOLDB].[dbo].[tool_characteristic].[data_value] FROM [TOOLDB].[dbo].[tool] LEFT JOIN [TOOLDB].[dbo].[tool_characteristic] ON tool.id = tool_characteristic.tool_id LEFT JOIN [TOOLDB].[dbo].[din4000_key] ON [TOOLDB].[dbo].[din4000_key].[id] = [TOOLDB].[dbo].[tool_characteristic].[key_id] WHERE ([TOOLDB].[dbo].[tool_characteristic].[key_id] = 3 AND CAST ([TOOLDB].[dbo].[tool_characteristic].[data_value] AS DECIMAL(9,2)) = " + toolData["3"] + ")", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = command.ExecuteReader();

                // Read the data until there are no more rows
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        // Add the description to the toolDescription dictionary
                        toolList.Add(Convert.ToString(sqlDataReader.GetInt64(0)), new Tuple<string, string>(sqlDataReader.GetString(1), sqlDataReader.GetString(2)));
                    }
                }
            }
            sqlConnection.Close();
            sqlConnection.Dispose();
            return toolList;
        }

        /*! getHolderList() returns a list of the holders, which are stored in the database. */

        public Dictionary<string, string> getHolderList(SqlConnection sqlConnection)
        {
            // Instantiate a dictionary for the tool list
            Dictionary<string, string> holderList = new Dictionary<string, string>();

            // Using block for the database queries
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the list of the tools.
                SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[holder].[id], [TOOLDB].[dbo].[holder].[name] FROM [TOOLDB].[dbo].[holder];", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = command.ExecuteReader();

                // Read the data until there are no more rows
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        // Add the description to the toolDescription dictionary
                        holderList.Add(Convert.ToString(sqlDataReader.GetInt64(0)), sqlDataReader.GetString(1));
                    }
                }
            }
            sqlConnection.Close();
            sqlConnection.Dispose();
            return holderList;
        }
        //----------------------------------------------------------------------------------

        /*! getHolderTable() returns a list of the holders, which are stored in the database. */

        public DataTable getHolderTable(SqlConnection sqlConnection)
        {
            // Instantiate a dictionary for the tool list
            DataTable holderList = new DataTable();

            // Using block for the database queries
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the list of the tools.
                SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[holder].[id], [TOOLDB].[dbo].[holder].[name] AS Name, [TOOLDB].[dbo].[holder_characteristic].[data_value] AS Beschreibung FROM [TOOLDB].[dbo].[holder]LEFT JOIN [TOOLDB].[dbo].[holder_characteristic] ON holder_id = holder.id WHERE key_id = 64;", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = command.ExecuteReader();

                holderList.Load(sqlDataReader);
                
            }
            sqlConnection.Close();
            sqlConnection.Dispose();
            return holderList;
        }
        //----------------------------------------------------------------------------------
        
        /*! getAssemblyList() returns a list of the assemblies, which are stored in the database. */

        public Dictionary<string, string> getAssemblyList(SqlConnection sqlConnection)
        {
            // Instantiate a dictionary for the assembly list
            Dictionary<string, string> assemblyList = new Dictionary<string, string>();

            // Using block for the database queries
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the list of the tools.
                SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[assembly].[id], [TOOLDB].[dbo].[assembly].[description] FROM [TOOLDB].[dbo].[assembly];", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = command.ExecuteReader();

                // Read the data until there are no more rows
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        // Add the description to the toolDescription dictionary
                        assemblyList.Add(Convert.ToString(sqlDataReader.GetInt64(0)), sqlDataReader.GetString(1));
                    }
                }
            }
            sqlConnection.Close();
            sqlConnection.Dispose();
            return assemblyList;
        }

        /*! getSheetList() returns a list of the assemblies, which are stored in the database. */

        public DataTable getSheetList(SqlConnection sqlConnection)
        {
            // Instantiate a dictionary for the assembly list
            DataTable SheetList = new DataTable();

            // Using block for the database queries
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the list of the tools.
                SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[sheet].[sheet_id], [TOOLDB].[dbo].[sheet].[description] FROM [TOOLDB].[dbo].[sheet];", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = command.ExecuteReader();

                // Read the data until there are no more rows
                SheetList.Load(sqlDataReader); 
            }
            sqlConnection.Close();
            sqlConnection.Dispose();
            return SheetList;
        }

        //----------------------------------------------------------------------------------

        /*! getToolById uses the connection to the SQL-Server and gets the data of the tool with the given
         * toolID. The method returns a dictionary which contains all the characteristic data including the tool
           id and the description. */


        public Dictionary<string, string> getToolById(SqlConnection sqlConnection, long toolId)
        {
            // Instantiate a dictionary for the tool description
            Dictionary<string, string> toolDescription = new Dictionary<string, string>();

            // Add the Id to the dictionary
            toolDescription.Add("id", toolId.ToString());

            // Using block for the database queries
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the description of the tool.
                SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[tool].[id], [TOOLDB].[dbo].[tool].[name], [TOOLDB].[dbo].[tool].[type] FROM [TOOLDB].[dbo].[tool] WHERE [TOOLDB].[dbo].[tool].[id] = " + toolId + ";", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = command.ExecuteReader();

                // Read the data until there are no more rows
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        // Add the description to the toolDescription dictionary
                        toolDescription.Add("description", sqlDataReader.GetString(1));
                        toolDescription.Add("type", sqlDataReader.GetString(2));
                    }
                }

                // Create a second SqlCommand and query the characteristics of the tool.
                command = new SqlCommand("SELECT [TOOLDB].[dbo].[tool].[id], [TOOLDB].[dbo].[tool].[name], [TOOLDB].[dbo].[din4000_key].[name], [TOOLDB].[dbo].[tool_characteristic].[data_value] FROM [TOOLDB].[dbo].[tool] LEFT JOIN [TOOLDB].[dbo].[tool_characteristic] ON tool.id = tool_characteristic.tool_id LEFT JOIN [TOOLDB].[dbo].[din4000_key] ON [TOOLDB].[dbo].[din4000_key].[id] = [TOOLDB].[dbo].[tool_characteristic].[key_id] WHERE [TOOLDB].[dbo].[tool].[id] = " + toolId + ";", sqlConnection);

                // Execute the query
                sqlDataReader = command.ExecuteReader();

                // Read the characteristic data until there are no more rows
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        try
                        {
                            // Store the data into the dictionary
                            toolDescription.Add(sqlDataReader.GetString(2), sqlDataReader.GetString(3));
                        }
                        catch (Exception ex)
                        {
                            debugMessage = ex.Message;

                        }
                    }
                }
            }
            sqlConnection.Close();
            sqlConnection.Dispose();
            return toolDescription;
        }


       //----------------------------------------------------------------------------------

        /*! getToolTechnologyById uses the connection to the SQL-Server and gets the technological data of the tool with the given
         * toolID. The method returns a dictionary which contains all the technological data including the tool
           id and the description. */


        public TechData getToolTechnologyById(SqlConnection sqlConnection, int toolId)
        {
            // Instantiate a dictionary for the material
            TechData techMaterial = new TechData();


            // Add the Id to the dictionary
            techMaterial.Add(toolId.ToString(), null);

            // Using block for the database queries
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the technology of the tool.
                SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[tool].[id], [TOOLDB].[dbo].[tool].[name] FROM [TOOLDB].[dbo].[tool] WHERE [TOOLDB].[dbo].[tool].[id] = " + toolId + ";", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = command.ExecuteReader();

                // Read the data until there are no more rows
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        // Add the description to the toolTechnology dictionary
                        techMaterial.Add(sqlDataReader.GetString(1), null);
                    }
                }

                string queryString = "SELECT [TOOLDB].[dbo].[tool].[id], [TOOLDB].[dbo].[tool].[name], [TOOLDB].[dbo].[tool_technology].[material_id], [TOOLDB].[dbo].[tool_technology].[mill_process_id], [TOOLDB].[dbo].[tool_technology].processing_id, [TOOLDB].[dbo].[tool_technology].[value], [TOOLDB].[dbo].[tool_technology].[description] FROM [TOOLDB].[dbo].[tool] LEFT JOIN [TOOLDB].[dbo].[tool_technology] ON tool.id = [tool_technology].tool_id LEFT JOIN [TOOLDB].[dbo].[din4000_key] ON [TOOLDB].[dbo].[din4000_key].[id] = [TOOLDB].[dbo].[tool_technology].[processing_id] WHERE [TOOLDB].[dbo].[tool].[id] = " + toolId + ";";

                SqlDataAdapter adapter = new SqlDataAdapter(queryString, sqlConnection);

                DataSet TechTable = new DataSet();

                adapter.Fill(TechTable, "TechTable");

                //// Create a second SqlCommand and query the characteristics of the tool.
                //command = new SqlCommand("SELECT [TOOLDB].[dbo].[tool].[id], [TOOLDB].[dbo].[tool].[name], [TOOLDB].[dbo].[tool_technology].[material_id], [TOOLDB].[dbo].[tool_technology].processing_id, [TOOLDB].[dbo].[tool_technology].[value], [TOOLDB].[dbo].[tool_technology].[description] FROM [TOOLDB].[dbo].[tool] LEFT JOIN [TOOLDB].[dbo].[tool_technology] ON tool.id = [tool_technology].tool_id LEFT JOIN [TOOLDB].[dbo].[din4000_key] ON [TOOLDB].[dbo].[din4000_key].[id] = [TOOLDB].[dbo].[tool_technology].[processing_id] WHERE [TOOLDB].[dbo].[tool].[id] = " + toolId + ";", sqlConnection);

                //// Execute the query
                //sqlDataReader = command.ExecuteReader();

                //DataSet dataSet = new DataSet(); 

                ////DataTable internalData = new DataTable();
                //DataTable internalData = sqlDataReader.GetSchemaTable();

                //dataSet.Tables.Add(internalData);

                //dataSet.EnforceConstraints = false;

                //internalData.Load(sqlDataReader);

                //DataRow[] result = internalData.Select("material_id = 1");

                // Instantiate a new List for the millprocess
                List<int> millprocesses = new List<int>();

                if (TechTable.Tables[0].Rows.Count > 1)
                {
                    // Build a list for all the millprocess in the table
                    foreach (DataRow row in TechTable.Tables[0].Rows)
                    {
                        // Get the millprocess id
                        int millprocess_id = int.Parse(row[2].ToString());

                        // If the id isn't in the list, add it
                        if (!millprocesses.Contains(millprocess_id))
                            millprocesses.Add(millprocess_id);
                    }

                    // Run through the millprocess list
                    foreach (int millprocess in millprocesses)
                    {
                        // Instantiate a dictionary for the materials
                        //MillProcessData millProcess = new MillProcessData();

                        // List for the processes
                        List<int> materials = new List<int>();

                        // Fetch every row in the TechTable
                        foreach (DataRow row in TechTable.Tables[0].Rows)
                        {
                            // The variables for the material_id and the process_id
                            int material_id = int.Parse(row[3].ToString());
                            int millprocess_id = int.Parse(row[2].ToString());

                            // If its the actual material then add the process to the process List
                            if (millprocess_id == millprocess)
                                if (!materials.Contains(material_id))
                                    materials.Add(material_id);
                        }



                        // Instantiate a new List for the materials
                        //List<int> materials = new List<int>();

                        if (TechTable.Tables[0].Rows.Count > 1)
                        {
                            // Build a list for all the materials in the table
                            foreach (DataRow row in TechTable.Tables[0].Rows)
                            {
                                // Get the material id
                                int material_id = int.Parse(row[3].ToString());

                                // If the id isn't in the list, add it
                                if (!materials.Contains(material_id))
                                    materials.Add(material_id);
                            }

                            // Run through the material list
                            foreach (int material in materials)
                            {
                                // Instantiate a new List for the materials
                                //List<int> materials = new List<int>();

                                // Instantiate a dictionary for the process
                                ProcessData techProcess = new ProcessData();

                                // List for the processes
                                List<int> processes = new List<int>();

                                // Fetch every row in the TechTable
                                foreach (DataRow row in TechTable.Tables[0].Rows)
                                {
                                    // The variables for the material_id and the process_id
                                    int millprocess_id = int.Parse(row[2].ToString());
                                    int material_id = int.Parse(row[3].ToString());
                                    int process_id = int.Parse(row[4].ToString());

                                    // If its the actual material then add the process to the process List
                                    if (material_id == material && millprocess_id == millprocess)
                                        if (!processes.Contains(process_id))
                                            processes.Add(process_id);
                                }

                                // Run through the process list
                                foreach (int process in processes)
                                {
                                    // Instantiate a Dictionary for Technological Values 
                                    Dictionary<string, string> dicTechValues = new Dictionary<string, string>();

                                    // Run through the table
                                    foreach (DataRow row in TechTable.Tables[0].Rows)
                                    {
                                        int millprocess_id = int.Parse(row[2].ToString());
                                        int material_id = int.Parse(row[3].ToString());
                                        int process_id = int.Parse(row[4].ToString());

                                        if ((material_id == material) && (process_id == process) && millprocess_id == millprocess)
                                            dicTechValues.Add(row[6].ToString(), row[5].ToString());
                                    }

                                    techProcess.Add(process.ToString(), dicTechValues);
                                }

                                techMaterial.Add(materialNames[material.ToString()].ToString(), techProcess);
                            }


                            Console.WriteLine("Stop");

                        }
                        else
                            techMaterial = null;


                    }
                }
            }
            return techMaterial;
        }
                
            



        private DataSet getTechTableTool(SqlConnection sqlConnection, long toolId)
        {
            // Using block for the database queries
            using (sqlConnection)
            {
                //// Create a new SqlCommand and query the technology of the tool.
                //SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[tool].[id], [TOOLDB].[dbo].[tool].[name] FROM [TOOLDB].[dbo].[tool] WHERE [TOOLDB].[dbo].[tool].[id] = " + toolId + ";", sqlConnection);

                //// Create a new SqlDataReader and execute the query
                //SqlDataReader sqlDataReader = command.ExecuteReader();

                //// Read the data until there are no more rows
                //if (sqlDataReader.HasRows)
                //{
                //    while (sqlDataReader.Read())
                //    {
                //        // Add the description to the toolTechnology dictionary
                //        millingProcess.Add(sqlDataReader.GetString(1), null);
                //    }
                //}

                string queryString = "SELECT [TOOLDB].[dbo].[tool].[id], [TOOLDB].[dbo].[tool].[name], [TOOLDB].[dbo].[tool_technology].[mill_process_id], [TOOLDB].[dbo].[tool_technology].[material_id], [TOOLDB].[dbo].[tool_technology].processing_id, [TOOLDB].[dbo].[tool_technology].[value], [TOOLDB].[dbo].[tool_technology].[description] FROM [TOOLDB].[dbo].[tool] LEFT JOIN [TOOLDB].[dbo].[tool_technology] ON tool.id = [tool_technology].tool_id LEFT JOIN [TOOLDB].[dbo].[din4000_key] ON [TOOLDB].[dbo].[din4000_key].[id] = [TOOLDB].[dbo].[tool_technology].[processing_id] WHERE [TOOLDB].[dbo].[tool].[id] = " + toolId + ";";

                SqlDataAdapter adapter = new SqlDataAdapter(queryString, sqlConnection);

                DataSet TechTable = new DataSet();

                adapter.Fill(TechTable, "TechTable");

                return TechTable;
            }


        }



        private DataSet getTechTableAssembly(SqlConnection sqlConnection, long assemblyId)
        {
            // Using block for the database queries
            using (sqlConnection)
            {
                //// Create a new SqlCommand and query the technology of the tool.
                //SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[tool].[id], [TOOLDB].[dbo].[tool].[name] FROM [TOOLDB].[dbo].[tool] WHERE [TOOLDB].[dbo].[tool].[id] = " + toolId + ";", sqlConnection);

                //// Create a new SqlDataReader and execute the query
                //SqlDataReader sqlDataReader = command.ExecuteReader();

                //// Read the data until there are no more rows
                //if (sqlDataReader.HasRows)
                //{
                //    while (sqlDataReader.Read())
                //    {
                //        // Add the description to the toolTechnology dictionary
                //        millingProcess.Add(sqlDataReader.GetString(1), null);
                //    }
                //}

                string queryString = "SELECT [TOOLDB].[dbo].[assembly].[id], [TOOLDB].[dbo].[assembly].[name], [TOOLDB].[dbo].[assembly_technology].[mill_process_id], [TOOLDB].[dbo].[assembly_technology].[material_id], [TOOLDB].[dbo].[assembly_technology].processing_id, [TOOLDB].[dbo].[assembly_technology].[value], [TOOLDB].[dbo].[assembly_technology].[description] FROM [TOOLDB].[dbo].[assembly] LEFT JOIN [TOOLDB].[dbo].[assembly_technology] ON assembly.id = [assembly_technology].assembly_id LEFT JOIN [TOOLDB].[dbo].[din4000_key] ON [TOOLDB].[dbo].[din4000_key].[id] = [TOOLDB].[dbo].[assembly_technology].[processing_id] WHERE [TOOLDB].[dbo].[assembly].[id] = " + assemblyId + ";";

                SqlDataAdapter adapter = new SqlDataAdapter(queryString, sqlConnection);

                DataSet TechTable = new DataSet();

                adapter.Fill(TechTable, "TechTable");

                return TechTable;
            }


        }



        //----------------------------------------------------------------------------------

        /*! getToolTechnologyById uses the connection to the SQL-Server and gets the technological data of the tool with the given
         * toolID. The method returns a dictionary which contains all the technological data including the tool
           id and the description. */


        public MillProcessData LoadTechDataTool(SqlConnection sqlConnection, long toolId)
        {
            // Read the complete technology table from the database
            DataSet TechTable = getTechTableTool(sqlConnection, toolId);

            // Instantiate a dictionary for the millingProcess
            MillProcessData millingProcess = new MillProcessData();

            millingProcess.Add(toolId.ToString(), null);
            millingProcess.Add(toolId.ToString(), null);

            // If there at least one entry, we read the milling process ids
            if (TechTable.Tables[0].Rows.Count > 1)
            {

                 // Instantiate a new List for the milling processes
                List<int> millProcesses = new List<int>();

                // Build a list for all the millingProcesses in the table
                foreach (DataRow row in TechTable.Tables[0].Rows)
                {
                    // Get the millingProcess id
                    int millProcess_id = int.Parse(row[2].ToString());

                    // If the id isn't in the list, add it
                    if (!millProcesses.Contains(millProcess_id))
                        millProcesses.Add(millProcess_id);

                }

                // At this point we have a list of all milling processes

                // Now we create a list of all the materials
                List<int> materials = new List<int>();

                // Run through the millingProcess list
                foreach (int millProcess in millProcesses)
                {
                    materials.Clear();

                    // Build a list for all the materials in the actual millingProcess
                    foreach (DataRow row in TechTable.Tables[0].Rows)
                    {
                        if(millProcess == int.Parse(row[2].ToString()))
                        {
                            // Get the material id
                            int material_id = int.Parse(row[3].ToString());

                            // If the id isn't in the list, add it
                            if (!materials.Contains(material_id))
                                materials.Add(material_id);
                        }

                    }

                    // At this point we have a list of all materials for the milling process

                     // Now we create a list of all the materials
                    List<int> processes = new List<int>();

                    TechData MaterialData = new TechData();



                    // Run through the millingProcess list
                    foreach (int material in materials)
                    {
                        processes.Clear();
                        
                        // Build a list for all the processes for the actual material
                        foreach (DataRow row in TechTable.Tables[0].Rows)
                        {
                            if (millProcess == int.Parse(row[2].ToString()) && material == int.Parse(row[3].ToString()))
                            {
                                // Get the process id
                                int process_id = int.Parse(row[4].ToString());

                                // If the id isn't in the list, add it
                                if (!processes.Contains(process_id))
                                    processes.Add(process_id);
                            }

                        }

                        // At this point we have a list of all processes for the material

                        ProcessData processData = new ProcessData();

                        // Run through the millingProcess list
                        foreach (int process in processes)
                        {
                            Dictionary<string, string> techData = new Dictionary<string, string>();

                            foreach (DataRow row in TechTable.Tables[0].Rows)
                            {
                                if (millProcess == int.Parse(row[2].ToString()) && material == int.Parse(row[3].ToString()) && process == int.Parse(row[4].ToString()))
                                {
                                    techData.Add(din4000names[row[6].ToString()], row[5].ToString());
                                }

                            }

                            processData.Add(processNames[process.ToString()], techData);
                        }


                        MaterialData.Add(materialNames[material.ToString()], processData);

                    }

                    millingProcess.Add(millProcessNames[millProcess.ToString()], MaterialData);

                }

            }

            return millingProcess;

        }

        public MillProcessData getToolMillProcessById(SqlConnection sqlConnection, int toolId)
        {
            // Instantiate a dictionary for the millingProcess
            MillProcessData millingProcess = new MillProcessData();

            // Read the complete technology table from the database
            DataSet TechTable = getTechTableTool(sqlConnection, toolId);

            // Add the Id to the dictionary
            millingProcess.Add(toolId.ToString(), null);


            // Instantiate a new List for the milling processes
            List<int> millProcesses = new List<int>();

            
            // If there at least one entry, we read the milling process ids
            if (TechTable.Tables[0].Rows.Count > 1)
            {
                // Build a list for all the millingProcesses in the table
                foreach (DataRow row in TechTable.Tables[0].Rows)
                {
                    // Get the millingProcess id
                    int millProcess_id = int.Parse(row[2].ToString());

                    // If the id isn't in the list, add it
                    if (!millProcesses.Contains(millProcess_id))
                        millProcesses.Add(millProcess_id);
                }

                List<int> materials = new List<int>();

                // Run through the millingProcess list
                foreach (int millProcess in millProcesses)
                {
                    // Instantiate a dictionary for the process
                    TechData ProcessMaterial = millingProcess.Get_MillProcessData(millProcess);

                    millingProcess.Add(millProcessNames[millProcess.ToString()].ToString(), ProcessMaterial);


                }


            }
            
                //// List for the processes
                //List<int> ProcessMaterials = new List<int>();


                //if (TechTable.Tables[0].Rows.Count > 1)
                //{
                //    // Build a list for all the millingProcesses in the table
                //    foreach (DataRow row in TechTable.Tables[0].Rows)
                //    {
                //        // Get the millingProcess id
                //        int millProcess_id = int.Parse(row[2].ToString());

                //        // If the id isn't in the list, add it
                //        if (!millProcesses.Contains(millProcess_id))
                //            millProcesses.Add(millProcess_id);
                //    }

                    
            //        // Run through the millingProcess list
            //        foreach (int millProcess in millProcesses)
            //        {
            //            // Instantiate a dictionary for the process
            //            TechData ProcessMaterial = millingProcess.Get_MillProcessData(millProcess);

            //            // Fetch every row in the TechTable
            //            foreach (DataRow row in TechTable.Tables[0].Rows)
            //            {
            //                // The variables for the material_id and the process_id
            //                int millingProcess_id = int.Parse(row[2].ToString());
            //                int material_id = int.Parse(row[3].ToString());

            //                // If its the actual material then add the process to the process List
            //                if (millingProcess_id == millProcess)
            //                    if (!ProcessMaterials.Contains(material_id))
            //                    {
            //                        ProcessMaterials.Add(material_id);
            //                        TechData tmpTechData = millingProcess.Get_MillProcessData(material_id);
            //                        millingProcess.Add(millProcessNames[millingProcess_id.ToString()].ToString(), tmpTechData);
            //                    }

            //            }
            //        }

                    

            //        // Build a list for all the materials in the table
            //        foreach (DataRow row in TechTable.Tables[0].Rows)
            //        {
            //            // Get the material id
            //            int material_id = int.Parse(row[3].ToString());

            //            // If the id isn't in the list, add it
            //            if (!materials.Contains(material_id))
            //                materials.Add(material_id);
            //        }

            //        // Run through the material list
            //        foreach (int material in materials)
            //        {
            //            // Instantiate a dictionary for the process
            //            ProcessData techProcess = new ProcessData();

            //            // List for the processes
            //            List<int> processes = new List<int>();

            //            // Fetch every row in the TechTable
            //            foreach (DataRow row in TechTable.Tables[0].Rows)
            //            {
            //                // The variables for the material_id and the process_id
            //                int material_id = int.Parse(row[3].ToString());
            //                int process_id = int.Parse(row[4].ToString());

            //                // If its the actual material then add the process to the process List
            //                if (material_id == material)
            //                    if (!processes.Contains(process_id))
            //                        processes.Add(process_id);
            //            }

            //            // Run through the process list
            //            foreach (int process in processes)
            //            {
            //                // Instantiate a Dictionary for Technological Values 
            //                Dictionary<string, string> dicTechValues = new Dictionary<string, string>();

            //                // Run through the table
            //                foreach (DataRow row in TechTable.Tables[0].Rows)
            //                {
            //                    int material_id = int.Parse(row[3].ToString());
            //                    int process_id = int.Parse(row[4].ToString());

            //                    if ((material_id == material) && (process_id == process))
            //                        dicTechValues.Add(row[6].ToString(), row[5].ToString());
            //                }

            //                techProcess.Add(process.ToString(), dicTechValues);
            //            }

            //            //techMaterial.Add(materialNames[material.ToString()].ToString(), techProcess);
            //            //millProcess.Add(millProcessNames[material_id.ToString()].ToString(), millProcess);
            //        }


            //        Console.WriteLine("Stop");

            //    }
            //    else
            //        millingProcess = null;
            //}


            return millingProcess;
        }





        //----------------------------------------------------------------------------------

        /*! getToolTechnologyById uses the connection to the SQL-Server and gets the technological data of the tool with the given
         * toolID. The method returns a dictionary which contains all the technological data including the tool
           id and the description. */


        public MillProcessData LoadTechDataAssembly(SqlConnection sqlConnection, long toolId)
        {
            // Read the complete technology table from the database
            DataSet TechTable = getTechTableAssembly(sqlConnection, toolId);

            // Instantiate a dictionary for the millingProcess
            MillProcessData millingProcess = new MillProcessData();

            millingProcess.Add(toolId.ToString(), null);
            millingProcess.Add(toolId.ToString(), null);

            // If there at least one entry, we read the milling process ids
            if (TechTable.Tables[0].Rows.Count > 1)
            {

                // Instantiate a new List for the milling processes
                List<int> millProcesses = new List<int>();

                // Build a list for all the millingProcesses in the table
                foreach (DataRow row in TechTable.Tables[0].Rows)
                {
                    // Get the millingProcess id
                    int millProcess_id = int.Parse(row[2].ToString());

                    // If the id isn't in the list, add it
                    if (!millProcesses.Contains(millProcess_id))
                        millProcesses.Add(millProcess_id);

                }

                // At this point we have a list of all milling processes

                // Now we create a list of all the materials
                List<int> materials = new List<int>();

                // Run through the millingProcess list
                foreach (int millProcess in millProcesses)
                {
                    materials.Clear();

                    // Build a list for all the materials in the actual millingProcess
                    foreach (DataRow row in TechTable.Tables[0].Rows)
                    {
                        if (millProcess == int.Parse(row[2].ToString()))
                        {
                            // Get the material id
                            int material_id = int.Parse(row[3].ToString());

                            // If the id isn't in the list, add it
                            if (!materials.Contains(material_id))
                                materials.Add(material_id);
                        }

                    }

                    // At this point we have a list of all materials for the milling process

                    // Now we create a list of all the materials
                    List<int> processes = new List<int>();

                    TechData MaterialData = new TechData();



                    // Run through the millingProcess list
                    foreach (int material in materials)
                    {
                        processes.Clear();

                        // Build a list for all the processes for the actual material
                        foreach (DataRow row in TechTable.Tables[0].Rows)
                        {
                            if (millProcess == int.Parse(row[2].ToString()) && material == int.Parse(row[3].ToString()))
                            {
                                // Get the process id
                                int process_id = int.Parse(row[4].ToString());

                                // If the id isn't in the list, add it
                                if (!processes.Contains(process_id))
                                    processes.Add(process_id);
                            }

                        }

                        // At this point we have a list of all processes for the material

                        ProcessData processData = new ProcessData();

                        // Run through the millingProcess list
                        foreach (int process in processes)
                        {
                            Dictionary<string, string> techData = new Dictionary<string, string>();

                            foreach (DataRow row in TechTable.Tables[0].Rows)
                            {
                                if (millProcess == int.Parse(row[2].ToString()) && material == int.Parse(row[3].ToString()) && process == int.Parse(row[4].ToString()))
                                {
                                    string TechDataKey = techKeys[row[6].ToString()];
                                    if (!techData.ContainsKey(TechDataKey))
                                        techData.Add(TechDataKey, row[5].ToString());
                                }

                            }

                            processData.Add(processNames[process.ToString()], techData);
                        }


                        MaterialData.Add(materialNames[material.ToString()], processData);

                    }

                    millingProcess.Add(millProcessNames[millProcess.ToString()], MaterialData);

                }

            }

            return millingProcess;

        }


        public void DeleteAssemblyTechnology(SqlConnection sqlConnection, long AssemblyId, string MillProcessKey, string ProcessNameKey, string MaterialKey)
        {
            // Use the existing SQLiteConnection
            using (sqlConnection)
            {
                // Create a new SqLiteCommand and query the description of the tool.
                using (SqlCommand SQLCommand = new SqlCommand("DELETE FROM [TOOLDB].[dbo].[assembly_technology] WHERE assembly_id = " + AssemblyId + " AND mill_process_id=" + MillProcessKey + " AND processing_id=" + ProcessNameKey + " AND material_id=" + MaterialKey + ";", sqlConnection))
                {

                    // Create a new SQLiteDataReader and execute the query
                    SqlDataReader SQLDataReader = null;

                    // Execute the command
                    SQLDataReader = SQLCommand.ExecuteReader();
                }
            }

            //return true;




        }



        //----------------------------------------------------------------------------------

        /*! DeleteAssemblyTechnology() uses the connection to the DB-Server and deletes the technological data of the Tool with the given
         * tool_id.
         * @param SQLiteConnection to the database
         * @param The Tool id as long value
         * @param MillProcessKey 
         * @param ProcessNameKey 
         * @param MaterialKey 
         */

        public void DeleteToolTechnology(SqlConnection SQLConnection, long ToolId, string MillProcessKey, string ProcessNameKey, string MaterialKey)
        {
            // Use the existing SQLiteConnection
            using (SQLConnection)
            {
                // Create a new SqLiteCommand and query the description of the Tool.
                using (SqlCommand SQLCommand = new SqlCommand("DELETE FROM tool_technology WHERE tool_id = " + ToolId + " AND mill_process_id=" + MillProcessKey + " AND processing_id=" + ProcessNameKey + " AND material_id=" + MaterialKey + ";", SQLConnection))
                {

                    // Create a new SQLiteDataReader and execute the query
                    SqlDataReader SQLDataReader = null;

                    // Execute the command
                    SQLDataReader = SQLCommand.ExecuteReader();
                }
            }
        }






        public bool CheckForExistingTechnology(SqlConnection sqlConnection, long AssemblyId, string MillProcessKey, string ProcessNameKey, string MaterialKey)
        {
            bool TechnologyExists = false;

            // Use the existing SQLiteConnection
            using (sqlConnection)
            {
                // Create a new SqLiteCommand and query the description of the tool.
                using (SqlCommand SQLCommand = new SqlCommand("SELECT * FROM assembly_technology WHERE assembly_id = " + AssemblyId + " AND mill_process_id=" + MillProcessKey + " AND processing_id=" + ProcessNameKey + " AND material_id=" + MaterialKey + ";", sqlConnection))
                {

                    // Create a new SQLiteDataReader and execute the query
                    SqlDataReader SQLDataReader = null;

                    // Execute the command
                    SQLDataReader = SQLCommand.ExecuteReader();

                    TechnologyExists = SQLDataReader.HasRows;

                    SQLDataReader.Close();
                    SQLDataReader.Dispose();
                }
            }

            return TechnologyExists;
        }




        public MillProcessData getAssemblyMillProcessById(SqlConnection sqlConnection, int toolId)
        {
            // Instantiate a dictionary for the millingProcess
            MillProcessData millingProcess = new MillProcessData();

            // Read the complete technology table from the database
            DataSet TechTable = getTechTableTool(sqlConnection, toolId);

            // Add the Id to the dictionary
            millingProcess.Add(toolId.ToString(), null);


            // Instantiate a new List for the milling processes
            List<int> millProcesses = new List<int>();


            // If there at least one entry, we read the milling process ids
            if (TechTable.Tables[0].Rows.Count > 1)
            {
                // Build a list for all the millingProcesses in the table
                foreach (DataRow row in TechTable.Tables[0].Rows)
                {
                    // Get the millingProcess id
                    int millProcess_id = int.Parse(row[2].ToString());

                    // If the id isn't in the list, add it
                    if (!millProcesses.Contains(millProcess_id))
                        millProcesses.Add(millProcess_id);
                }

                List<int> materials = new List<int>();

                // Run through the millingProcess list
                foreach (int millProcess in millProcesses)
                {
                    // Instantiate a dictionary for the process
                    TechData ProcessMaterial = millingProcess.Get_MillProcessData(millProcess);

                    millingProcess.Add(millProcessNames[millProcess.ToString()].ToString(), ProcessMaterial);


                }


            }


            return millingProcess;
        }

        
        
        //----------------------------------------------------------------------------------

        /*! getHolderById uses the connection to the SQL-Server and gets the data of the holder with the given
         * holderID. The method returns a dictionary which contains all the characteristic data including the holder
           id and the description. */


        public Dictionary<string, string> getHolderById(SqlConnection sqlConnection, long holderId)
        {
            // Instantiate a dictionary for the tool description
            Dictionary<string, string> holderDescription = new Dictionary<string, string>();

            // Add the Id to the dictionary
            holderDescription.Add("id", holderId.ToString());

            // Using block for the database queries
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the description of the holder.
                SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[holder].[id], [TOOLDB].[dbo].[holder].[name] FROM [TOOLDB].[dbo].[holder] WHERE [TOOLDB].[dbo].[holder].[id] = " + holderId + ";", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = command.ExecuteReader();

                // Read the data until there are no more rows
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        // Add the description to the toolDescription dictionary
                        holderDescription.Add("description", sqlDataReader.GetString(1));
                    }
                }

                // Create a second SqlCommand and query the characteristics of the tool.
                command = new SqlCommand("SELECT [TOOLDB].[dbo].[holder].[id], [TOOLDB].[dbo].[holder].[name], [TOOLDB].[dbo].[din4000_key].[name], [TOOLDB].[dbo].[holder_characteristic].[data_value] FROM [TOOLDB].[dbo].[holder] LEFT JOIN [TOOLDB].[dbo].[holder_characteristic] ON holder.id = holder_characteristic.holder_id LEFT JOIN [TOOLDB].[dbo].[din4000_key] ON [TOOLDB].[dbo].[din4000_key].[id] = [TOOLDB].[dbo].[holder_characteristic].[key_id] WHERE [TOOLDB].[dbo].[holder].[id] = " + holderId + ";", sqlConnection);

                // Execute the query
                sqlDataReader = command.ExecuteReader();

                // Read the characteristic data until there are no more rows
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        // Store the data into the dictionary
                        if (!sqlDataReader.IsDBNull(2))
                        {
                            string NewKey = sqlDataReader.GetString(2);
                            if(!holderDescription.ContainsKey(NewKey))
                                holderDescription.Add(sqlDataReader.GetString(2), sqlDataReader.GetString(3));
                        }
                    }
                }
            }
            sqlConnection.Close();
            sqlConnection.Dispose();
            return holderDescription;
        }

        //----------------------------------------------------------------------------------
 
        /*! insertTool() stores a new tool into the database. The method uses the given SqlConnection and inserts 
         the the data in the given toolDescription dictionary via a SQL-query into the data tables.*/

        public long insertTool(SqlConnection sqlConnection, Dictionary<string, string> toolDescription, long ToolId = 0)
        {
            // Initialize the variable for the last id
            long lastToolId = 0;
            long lastCharacteristicId = 0;

            //// Load the DIN 4000 Keys into the dictionary
            //loadDinNames();

            // Create a new SqlCommand and query the description of the tool.
            SqlCommand sqlCommand = new SqlCommand("SELECT MAX(Id) FROM tool;", sqlConnection);

            // Create a new SqlDataReader and execute the query
            SqlDataReader sqlDataReader = null;


            // Use the existing SqlConnection
            using (sqlConnection)
            {
                if (ToolId == 0)
                {

                    // Execute the command
                    sqlDataReader = sqlCommand.ExecuteReader();

                    // Read the given id
                    if (sqlDataReader.HasRows)
                    {
                        while (sqlDataReader.Read())
                        {
                            try
                            {
                                lastToolId = sqlDataReader.GetInt64(0);
                            }
                            catch (Exception ex)
                            {
                                debugMessage = ex.Message;
                                lastToolId = 0;
                            }
                        }
                    }
                }
                else
                    lastToolId = ToolId - 1;

                // Create a new SqlCommand and query the description of the tool.
                sqlCommand = new SqlCommand("INSERT INTO tool (id, name, type) VALUES (@id, @name, @type)", sqlConnection);
                
                // Add the parameter to the insert command.
                sqlCommand.Parameters.AddWithValue("@id", lastToolId + 1);
                sqlCommand.Parameters.AddWithValue("@name", toolDescription["name"]);
                sqlCommand.Parameters.AddWithValue("@type", toolDescription["type"]);
                
                // Create a new SqlDataReader and execute the query
                sqlDataReader = sqlCommand.ExecuteReader();

                // Create a new SqlCommand and query the description of the tool.
                sqlCommand = new SqlCommand("SELECT MAX(Id) FROM tool_characteristic;", sqlConnection);

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

                // Read the given id
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        try
                        {
                            lastCharacteristicId = sqlDataReader.GetInt64(0);
                        }
                        catch (Exception ex)
                        {
                            debugMessage = ex.Message;
                            lastCharacteristicId = 0;
                        }
                    }
                }

                for(int i = 1; i < toolDescription.Count; i++)
                {
                    sqlCommand = new SqlCommand("INSERT INTO tool_characteristic (id, tool_id, key_id, data_value) VALUES (@id, @tool_id, @key_id, @data_value)", sqlConnection);
                    
                    sqlCommand.Parameters.AddWithValue("@id", lastCharacteristicId + i);
                    sqlCommand.Parameters.AddWithValue("@tool_id", lastToolId + 1);
                    string keyId = null;
                    foreach (string key in din4000names.Keys)
                    {
                        //Console.WriteLine(importCharValues.ElementAt(i).Key + "-" + key + Environment.NewLine);
                        if (key.Equals(toolDescription.ElementAt(i).Key))
                        {
                            //Console.WriteLine(toolDescription.ElementAt(i).Key + "-" + key + Environment.NewLine);
                            keyId = key;
                        }
                    }

                    if(i > 0 && !string.IsNullOrEmpty(keyId))
                    {
                        sqlCommand.Parameters.AddWithValue("@key_id", keyId);
                        sqlCommand.Parameters.AddWithValue("@data_value", toolDescription.ElementAt(i).Value);

                        // don't forget to take care of connection - I omit this part for clearness
                        // Execute the command
                        sqlDataReader = sqlCommand.ExecuteReader();
                    }
                }                
            }

            return lastToolId + 1;
        }


        /*! insertTool() stores a new tool into the database. The method uses the given SqlConnection and inserts 
         the the data in the given toolDescription dictionary via a SQL-query into the data tables.*/

        public void UpdateToolCharacteristic(SqlConnection sqlConnection, string Key, string Value, long ToolId)
        {

            // Create a new SqlCommand and query the description of the tool.
            SqlCommand sqlCommand = new SqlCommand("SELECT * FROM [TOOLDB].[dbo].[tool_characteristic] WHERE tool_id = " + ToolId.ToString() + " AND key_id = '" + Key + "';", sqlConnection);

            // Create a new SqlDataReader and execute the query
            SqlDataReader sqlDataReader = null;

            long IsPresent = 0;


            // Use the existing SqlConnection
            using (sqlConnection)
            {

                sqlDataReader = sqlCommand.ExecuteReader();

                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        IsPresent = sqlDataReader.GetInt64(0);
                    }
                }

                if (IsPresent != 0)
                {
                    sqlCommand = new SqlCommand("UPDATE tool_characteristic SET data_value = '" + Value + "' WHERE tool_id = " + ToolId.ToString() + " AND key_id = '" + Key + "';", sqlConnection);
                }
                else
                {
                    long lastCharacteristicId = 0;
                    
                    // Create a new SqlCommand and query the description of the tool.
                    sqlCommand = new SqlCommand("SELECT MAX(Id) FROM tool_characteristic;", sqlConnection);

                    // Execute the command
                    sqlDataReader = sqlCommand.ExecuteReader();

                    // Read the given id
                    if (sqlDataReader.HasRows)
                    {
                        while (sqlDataReader.Read())
                        {
                            try
                            {
                                lastCharacteristicId = sqlDataReader.GetInt64(0);
                            }
                            catch (Exception ex)
                            {
                                debugMessage = ex.Message;
                                lastCharacteristicId = 0;
                            }
                        }
                    }
                    sqlCommand = new SqlCommand("INSERT INTO tool_characteristic (id, tool_id, key_id, data_value) VALUES (@id, @tool_id, @key_id, @data_value)", sqlConnection);
                    sqlCommand.Parameters.AddWithValue("@id", lastCharacteristicId + 1);
                    sqlCommand.Parameters.AddWithValue("@tool_id", ToolId);
                    sqlCommand.Parameters.AddWithValue("@key_id", Key);
                    sqlCommand.Parameters.AddWithValue("@data_value", Value);

                }

                    // Execute the command
                    sqlDataReader = sqlCommand.ExecuteReader();

            }

        }





        public void UpdateHolderCharacteristic(SqlConnection sqlConnection, string Key, string Value, long HolderId)
        {

            // Create a new SqlCommand and query the description of the tool.
            SqlCommand sqlCommand = new SqlCommand("SELECT * FROM [TOOLDB].[dbo].[holder_characteristic] WHERE holder_id = " + HolderId.ToString() + " AND key_id = '" + Key + "';", sqlConnection);

            // Create a new SqlDataReader and execute the query
            SqlDataReader sqlDataReader = null;

            long IsPresent = 0;


            // Use the existing SqlConnection
            using (sqlConnection)
            {

                sqlDataReader = sqlCommand.ExecuteReader();

                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        IsPresent = sqlDataReader.GetInt64(0);
                    }
                }

                if (IsPresent != 0)
                {
                    sqlCommand = new SqlCommand("UPDATE holder_characteristic SET data_value = '" + Value + "' WHERE holder_id = " + HolderId.ToString() + " AND key_id = '" + Key + "';", sqlConnection);
                }
                else
                {
                    long lastCharacteristicId = 0;

                    // Create a new SqlCommand and query the description of the tool.
                    sqlCommand = new SqlCommand("SELECT MAX(Id) FROM holder_characteristic;", sqlConnection);

                    // Execute the command
                    sqlDataReader = sqlCommand.ExecuteReader();

                    // Read the given id
                    if (sqlDataReader.HasRows)
                    {
                        while (sqlDataReader.Read())
                        {
                            try
                            {
                                lastCharacteristicId = sqlDataReader.GetInt64(0);
                            }
                            catch (Exception ex)
                            {
                                debugMessage = ex.Message;
                                lastCharacteristicId = 0;
                            }
                        }
                    }
                    sqlCommand = new SqlCommand("INSERT INTO holder_characteristic (id, holder_id, key_id, data_value) VALUES (@id, @holder_id, @key_id, @data_value)", sqlConnection);
                    sqlCommand.Parameters.AddWithValue("@id", lastCharacteristicId + 1);
                    sqlCommand.Parameters.AddWithValue("@holder_id", HolderId);
                    sqlCommand.Parameters.AddWithValue("@key_id", Key);
                    sqlCommand.Parameters.AddWithValue("@data_value", Value);

                }

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

            }

        }



        public void UpdateAssemblyReach(SqlConnection SQLConnection, double ToolReach, string AssemblyNr)
        {
            // Create a new SqlDataReader and execute the query
            SqlDataReader SQLDataReader = null;

            // Use the existing SqlConnection
            using (SQLConnection)
            {
                string Reach = Convert.ToString(ToolReach, CultureInfo.InvariantCulture);

                SqlCommand SQLCommand = new SqlCommand("UPDATE assembly SET reach = '" + Reach + "' WHERE assembly_nr = '" + AssemblyNr + "';", SQLConnection);


                SQLDataReader = SQLCommand.ExecuteReader();

            }

        }


        public void UpdateAssemblyDiameter(SqlConnection SQLConnection, double ToolDiameter, string AssemblyNr)
        {
            // Create a new SqlDataReader and execute the query
            SqlDataReader SQLDataReader = null;

            // Use the existing SqlConnection
            using (SQLConnection)
            {
                string Diameter = Convert.ToString(ToolDiameter, CultureInfo.InvariantCulture);

                SqlCommand SQLCommand = new SqlCommand("UPDATE assembly SET diameter = '" + Diameter.Replace(",", ".") + "' WHERE assembly_nr = '" + AssemblyNr + "';", SQLConnection);


                SQLDataReader = SQLCommand.ExecuteReader();

            }

        }





        //----------------------------------------------------------------------------------

        /*! insertHolder() stores a new holder into the database. The method uses the given SqlConnection and inserts 
         the the data in the given holderDescription dictionary via a SQL-query into the data tables.*/

        public long insertHolder(SqlConnection sqlConnection, Dictionary<string, string> holderDescription)
        {
            // Initialize the variable for the last id
            long lastHolderId = 0;
            long lastCharacteristicId = 0;

            // Load the DIN 4000 Keys into the dictionary
            //loadDinKeys();

            // Use the existing SqlConnection
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the description of the holder.
                SqlCommand sqlCommand = new SqlCommand("SELECT MAX(Id) FROM holder;", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = null;

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

                // Read the given id
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        try
                        {
                            lastHolderId = sqlDataReader.GetInt64(0);
                        }
                        catch (Exception ex)
                        {
                            debugMessage = ex.Message;
                            lastHolderId = 0;
                        }
                    }
                }

                // Create a new SqlCommand and query the description of the holder.
                sqlCommand = new SqlCommand("INSERT INTO holder (id, name) VALUES (@id, @name)", sqlConnection);

                // Add the parameter to the insert command.
                sqlCommand.Parameters.AddWithValue("@id", lastHolderId + 1);
                if(holderDescription.ContainsKey("J21"))
                    sqlCommand.Parameters.AddWithValue("@name", holderDescription["J21"]);
                //if (holderDescription.ContainsKey("description"))
                //    sqlCommand.Parameters.AddWithValue("@name", holderDescription["description"]);

                // Create a new SqlDataReader and execute the query
                sqlDataReader = sqlCommand.ExecuteReader();

                // Create a new SqlCommand and query the description of the holder.
                sqlCommand = new SqlCommand("SELECT MAX(Id) FROM holder_characteristic;", sqlConnection);

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

                // Read the given id
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        try
                        {
                            lastCharacteristicId = sqlDataReader.GetInt64(0);
                        }
                        catch (Exception ex)
                        {
                            debugMessage = ex.Message; 
                            lastCharacteristicId = 0;
                        }
                    }
                }

                for (int i = 1; i < holderDescription.Count; i++)
                {
                    sqlCommand = new SqlCommand("INSERT INTO holder_characteristic (id, holder_id, key_id, data_value) VALUES (@id, @holder_id, @key_id, @data_value)", sqlConnection);
                    sqlCommand.Parameters.AddWithValue("@id", lastCharacteristicId + i);
                    sqlCommand.Parameters.AddWithValue("@holder_id", lastHolderId + 1);
                    string keyId = null;
                    foreach (string key in din4000keys.Keys)
                    {
                        //Console.WriteLine(importCharValues.ElementAt(i).Key + "-" + key + Environment.NewLine);
                        if (key.Equals(holderDescription.ElementAt(i).Key))
                        {
                            //Console.WriteLine(holderDescription.ElementAt(i).Key + "-" + key + Environment.NewLine);
                            keyId = din4000keys[key];
                        }
                    }
                    if (!String.IsNullOrEmpty(keyId))
                    {
                        sqlCommand.Parameters.AddWithValue("@key_id", keyId);
                        sqlCommand.Parameters.AddWithValue("@data_value", holderDescription.ElementAt(i).Value);

                        // don't forget to take care of connection - I omit this part for clearness
                        // Execute the command
                        sqlDataReader = sqlCommand.ExecuteReader();
                    }
                }
            }

            return lastHolderId + 1;
        }

        //----------------------------------------------------------------------------------

        /*! deleteTool() removes a tool from the database. Parameters are the connection to the 
         * database server and the id-number of the tool. */

        public bool deleteTool(SqlConnection sqlConnection, long toolId)
        {
            SqlConnection m_sqlConnection = getSqlConnection();

            long dataCount = GetLastId(m_sqlConnection, "assembly");

            for (int i = 0; i < dataCount; i++)
            {
                SqlConnection assemblySqlConnection = getSqlConnection();
                if (getAssemblyById(assemblySqlConnection, i + 1).toolId == toolId)
                {
                    MessageBox.Show("Das ausgewählte Werkzeug ist Teil einer Baugruppe. Bitte löschen sie zuerst die Baugruppe und dann das Werkzeug.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }

            // Use the existing SqlConnection
            using (sqlConnection)
            {   
                // Create a new SqlCommand and query the description of the tool.
                SqlCommand sqlCommand = new SqlCommand("DELETE [dbo].[tool_characteristic] FROM tool_characteristic WHERE tool_id = " + toolId + ";", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = null;

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

                // Create a new SqlCommand and query the description of the tool.
                sqlCommand = new SqlCommand("DELETE [dbo].[tool_technology] FROM tool_technology WHERE tool_id = " + toolId + ";", sqlConnection);

                // Create a new SqlDataReader and execute the query
                sqlDataReader = null;

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

                // Create a new SqlCommand and query the description of the tool.
                sqlCommand = new SqlCommand("DELETE [dbo].[tool_profile] FROM tool_profile WHERE tool_id = " + toolId + ";", sqlConnection);

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

                // Create a new SqlCommand and query the description of the tool.
                sqlCommand = new SqlCommand("DELETE [dbo].[tool_model] FROM tool_model WHERE tool_id = " + toolId + ";", sqlConnection);

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

                // Create a new SqlCommand and query the description of the tool.
                sqlCommand = new SqlCommand("DELETE [dbo].[tool] FROM tool WHERE id = " + toolId + ";", sqlConnection);

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

            }

            return true;
        }


        //----------------------------------------------------------------------------------

        /*! deleteHolder() removes a holder from the database. Parameters are the connection to the 
         * database server and the id-number of the holder. */

        public bool deleteHolder(SqlConnection sqlConnection, long holderId)
        {

            // Use the existing SqlConnection
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the description of the holder.
                SqlCommand sqlCommand = new SqlCommand("DELETE [dbo].[holder_characteristic] FROM holder_characteristic WHERE holder_id = " + holderId + ";", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = null;

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

                // Create a new SqlCommand and query the profile of the holder.
                sqlCommand = new SqlCommand("DELETE [dbo].[holder_profile] FROM holder_profile WHERE holder_id = " + holderId + ";", sqlConnection);

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

                // Create a new SqlCommand and query the description of the holder.
                sqlCommand = new SqlCommand("DELETE [dbo].[holder] FROM holder WHERE id = " + holderId + ";", sqlConnection);

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

            }

            return true;
        }

        //----------------------------------------------------------------------------------

        /*! deleteToolList() removes a toollist from the database. Parameters are the connection to the 
         * database server and the id-number of the toollist.
         * @param SQLConnection to the database
         * @param id-number of the holder  as a long variable.  
         * @return bool flag true if succesful / false if not
         */

        public bool deleteToolList(SqlConnection sqlConnection, string sheetId)
        {

            // Use the existing SQLiteConnection
            using (sqlConnection)
            {
                // Create a new SqLiteCommand
                SqlCommand Command = new SqlCommand("DELETE FROM sheet_entries WHERE sheet_id = '" + sheetId + "';", sqlConnection);

                // Create a new SQLiteDataReader and execute the query
                SqlDataReader SqlDataReader = Command.ExecuteReader();

                Command = new SqlCommand("DELETE FROM sheet WHERE sheet_id = '" + sheetId + "';", sqlConnection);

                // Execute the command
                SqlDataReader = Command.ExecuteReader();

            }
            return true;
        }

        //----------------------------------------------------------------------------------

        /*! getAssemblyById() uses the connection to the SQL-Server and gets the data of the assembly with the given
         * ID. The method returns a class which contains all the data of the assembly, like the ids of the tool and 
         * the holder and a description of the assembly. */


        public CAssembly getAssemblyById(SqlConnection sqlConnection, long assemblyId)
        {
            // Instantiate a CAssembly for the Assembly description
            CAssembly assemblyDescription = new CAssembly();

            // Add the Id to the assemblyDescription
            assemblyDescription.id = assemblyId;

            // Using block for the database queries
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the description of the assembly.
                SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[assembly].[id], [TOOLDB].[dbo].[assembly].[description], [TOOLDB].[dbo].[assembly].[assembly_nr], [TOOLDB].[dbo].[assembly].[tool_id], [TOOLDB].[dbo].[assembly].[holder_id], [TOOLDB].[dbo].[assembly].[name], [TOOLDB].[dbo].[assembly].[reach], [TOOLDB].[dbo].[assembly].[diameter], [TOOLDB].[dbo].[assembly].[assembly_id] FROM [TOOLDB].[dbo].[assembly] WHERE [TOOLDB].[dbo].[assembly].[id] = " + assemblyId + ";", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = command.ExecuteReader();

                // Read the data until there are no more rows
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        assemblyDescription.description = sqlDataReader.GetString(1);
                        assemblyDescription.assemblyNr = sqlDataReader.GetString(2);
                        assemblyDescription.toolId = sqlDataReader.GetInt64(3);
                        assemblyDescription.holderId = sqlDataReader.GetInt64(4);
                        try
                        {
                            assemblyDescription.name = sqlDataReader.GetString(5);
                            assemblyDescription.reach = sqlDataReader.GetDouble(6);
                            assemblyDescription.diameter = sqlDataReader.GetDouble(7);
                            assemblyDescription.assemblyId = sqlDataReader.GetString(8);
                        }
                        catch (Exception ex)
                        {

                            debugMessage = ex.Message;
                            //assemblyDescription.holderId = 0;
                            assemblyDescription.reach = 0;
                        }
                    }
                }
            }

            sqlConnection.Close();
            sqlConnection.Dispose();
            return assemblyDescription;
        }

        //----------------------------------------------------------------------------------

        /*! getAssemblyById() uses the connection to the SQL-Server and gets the data of the assembly with the given
         * ID. The method returns a class which contains all the data of the assembly, like the ids of the tool and 
         * the holder and a description of the assembly. */


        public CAssembly getAssemblyByName(SqlConnection sqlConnection, string assemblyName)
        {
            // Instantiate a CAssembly for the Assembly description
            CAssembly assemblyDescription = new CAssembly();

            // Add the Id to the assemblyDescription
            assemblyDescription.description = assemblyName;

            // Using block for the database queries
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the description of the assembly.
                SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[assembly].[id], [TOOLDB].[dbo].[assembly].[name], [TOOLDB].[dbo].[assembly].[description], [TOOLDB].[dbo].[assembly].[tool_id], [TOOLDB].[dbo].[assembly].[holder_id], [TOOLDB].[dbo].[assembly].[assembly_nr], [TOOLDB].[dbo].[assembly].[reach], [TOOLDB].[dbo].[assembly].[diameter], [TOOLDB].[dbo].[assembly].[assembly_id] FROM [TOOLDB].[dbo].[assembly] WHERE [TOOLDB].[dbo].[assembly].[name] = '" + assemblyName + "';", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = command.ExecuteReader();

                // Read the data until there are no more rows
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        assemblyDescription.id = sqlDataReader.GetInt64(0);

                        if (!sqlDataReader.IsDBNull(sqlDataReader.GetOrdinal("name")))
                            assemblyDescription.name = sqlDataReader.GetString(1);

                        if (!sqlDataReader.IsDBNull(sqlDataReader.GetOrdinal("description")))
                            assemblyDescription.description = sqlDataReader.GetString(2);

                        if (!sqlDataReader.IsDBNull(sqlDataReader.GetOrdinal("tool_id")))
                            assemblyDescription.toolId = sqlDataReader.GetInt64(3);

                        if (!sqlDataReader.IsDBNull(sqlDataReader.GetOrdinal("holder_id")))
                            assemblyDescription.holderId = sqlDataReader.GetInt64(4);

                        if (!sqlDataReader.IsDBNull(sqlDataReader.GetOrdinal("assembly_nr")))
                            assemblyDescription.assemblyNr = sqlDataReader.GetString(5); 

                        if (!sqlDataReader.IsDBNull(sqlDataReader.GetOrdinal("reach")))
                            assemblyDescription.reach = sqlDataReader.GetDouble(6);

                        if (!sqlDataReader.IsDBNull(sqlDataReader.GetOrdinal("diameter")))
                            assemblyDescription.diameter = sqlDataReader.GetDouble(7);

                        if (!sqlDataReader.IsDBNull(sqlDataReader.GetOrdinal("assembly_id")))
                            assemblyDescription.assemblyId = sqlDataReader.GetString(8);
                        //assemblyDescription.id = sqlDataReader.GetInt64(0);
                        //assemblyDescription.name = sqlDataReader.GetString(1);
                        //assemblyDescription.description = sqlDataReader.GetString(2);
                        //assemblyDescription.toolId = sqlDataReader.GetInt64(3);
                        //assemblyDescription.holderId = sqlDataReader.GetInt64(4);
                        //assemblyDescription.assemblyNr = sqlDataReader.GetString(5);
                        //assemblyDescription.reach = sqlDataReader.GetDouble(6);
                        //assemblyDescription.diameter = sqlDataReader.GetDouble(7);
                        //assemblyDescription.assemblyId = sqlDataReader.GetString(8);
                    }
                }
            }

            sqlConnection.Close();
            sqlConnection.Dispose();
            return assemblyDescription;
        }

        //----------------------------------------------------------------------------------

        /*! getAssemblyById() uses the connection to the SQL-Server and gets the data of the assembly with the given
         * ID. The method returns a class which contains all the data of the assembly, like the ids of the tool and 
         * the holder and a description of the assembly. */


        public CAssembly getAssemblyByNr(SqlConnection sqlConnection, string assemblyNr)
        {
            // Instantiate a CAssembly for the Assembly description
            CAssembly assemblyDescription = new CAssembly();

            // Add the Id to the assemblyDescription
            assemblyDescription.description = assemblyNr;

            // Using block for the database queries
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the description of the assembly.
                SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[assembly].[id], [TOOLDB].[dbo].[assembly].[name], [TOOLDB].[dbo].[assembly].[description], [TOOLDB].[dbo].[assembly].[tool_id], [TOOLDB].[dbo].[assembly].[holder_id], [TOOLDB].[dbo].[assembly].[assembly_nr], [TOOLDB].[dbo].[assembly].[reach], [TOOLDB].[dbo].[assembly].[diameter], [TOOLDB].[dbo].[assembly].[assembly_id] FROM [TOOLDB].[dbo].[assembly] WHERE [TOOLDB].[dbo].[assembly].[assembly_nr] = '" + assemblyNr + "';", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = command.ExecuteReader();

                // Read the data until there are no more rows
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        assemblyDescription.id = sqlDataReader.GetInt64(0);

                        assemblyDescription.id = sqlDataReader.GetInt64(0);

                        if (!sqlDataReader.IsDBNull(sqlDataReader.GetOrdinal("name")))
                            assemblyDescription.name = sqlDataReader.GetString(1);

                        if (!sqlDataReader.IsDBNull(sqlDataReader.GetOrdinal("description")))
                            assemblyDescription.description = sqlDataReader.GetString(2);

                        if (!sqlDataReader.IsDBNull(sqlDataReader.GetOrdinal("tool_id")))
                            assemblyDescription.toolId = sqlDataReader.GetInt64(3);

                        if (!sqlDataReader.IsDBNull(sqlDataReader.GetOrdinal("holder_id")))
                            assemblyDescription.holderId = sqlDataReader.GetInt64(4);

                        if (!sqlDataReader.IsDBNull(sqlDataReader.GetOrdinal("assembly_nr")))
                            assemblyDescription.assemblyNr = assemblyNr;

                        if (!sqlDataReader.IsDBNull(sqlDataReader.GetOrdinal("reach")))
                            assemblyDescription.reach = sqlDataReader.GetDouble(6);

                        if (!sqlDataReader.IsDBNull(sqlDataReader.GetOrdinal("diameter")))
                            assemblyDescription.diameter = sqlDataReader.GetDouble(7);

                        if (!sqlDataReader.IsDBNull(sqlDataReader.GetOrdinal("assembly_id")))
                            assemblyDescription.assemblyId = sqlDataReader.GetString(8);
                    }
                }
            }

            sqlConnection.Close();
            sqlConnection.Dispose();
            return assemblyDescription;
        }

        //----------------------------------------------------------------------------------

        /*! insertAssembly() stores a new assembly into the database. The method uses the given SqlConnection and inserts 
         the given toolId, holderId and the assembly description via a SQL-query into the data tables.*/

        public long insertAssembly(SqlConnection sqlConnection, string assemblyNr, long toolId, long holderId, string assemblyDescription, string assemblyName, double toolReach, double toolDiameter = 0, long oldId = 0, string RfidCode = "")
        {
            // Initialize the variable for the last id
            long lastAssemblyId = 0;

            // Use the existing SqlConnection
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the description of the tool.
                SqlCommand sqlCommand = new SqlCommand("SELECT MAX(Id) FROM assembly;", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = null;

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

                // Read the given id
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        try
                        {
                            lastAssemblyId = sqlDataReader.GetInt64(0);
                        }
                        catch (Exception ex)
                        {
                            debugMessage = ex.Message; 
                            lastAssemblyId = 0;
                        }
                    }
                }

                if (oldId != 0)
                    lastAssemblyId = oldId - 1;

                // Create a new SqlCommand and query the description of the tool.
                sqlCommand = new SqlCommand("INSERT INTO assembly (id, assembly_nr, tool_id, holder_id, description, name, reach, diameter, assembly_id) VALUES (@id, @assembly_nr, @tool_id, @holder_id, @description, @name, @reach, @diameter, @assembly_id)", sqlConnection);

                // Add the parameter to the insert command.
                sqlCommand.Parameters.AddWithValue("@id", lastAssemblyId + 1);
                sqlCommand.Parameters.AddWithValue("@assembly_nr", assemblyNr);
                sqlCommand.Parameters.AddWithValue("@description", assemblyDescription);
                sqlCommand.Parameters.AddWithValue("@tool_id", toolId);
                sqlCommand.Parameters.AddWithValue("@holder_id", holderId);
                sqlCommand.Parameters.AddWithValue("@name", assemblyName);
                sqlCommand.Parameters.AddWithValue("@reach", toolReach);
                sqlCommand.Parameters.AddWithValue("@diameter", toolDiameter);
                sqlCommand.Parameters.AddWithValue("@assembly_id", RfidCode);

                // Create a new SqlDataReader and execute the query
                sqlDataReader = sqlCommand.ExecuteReader();

            }

            return lastAssemblyId + 1;
        }


        /*! insertArchive() stores a .cxa archive into the database. The method uses the given SqlConnection and inserts 
         the given date, creator, description, the company name and the filepath  via a SQL-query into the data tables.*/

        public long insertArchive(SqlConnection sqlConnection, CManifest Manifest)
        {
            // Initialize the variable for the last id
            long lastManifestId = 0;

            // Use the existing SqlConnection
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the description of the tool.
                SqlCommand sqlCommand = new SqlCommand("SELECT MAX(Id) FROM archive;", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = null;

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

                // Read the given id
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        try
                        {
                            lastManifestId = sqlDataReader.GetInt64(0);
                        }
                        catch (Exception ex)
                        {
                            debugMessage = ex.Message; 
                            lastManifestId = 0;
                        }
                    }
                }

                // Create a new SqlCommand and query the description of the tool.
                sqlCommand = new SqlCommand("INSERT INTO archive (id, date, creator, description, path) VALUES (@id, @date, @creator, @description, @path)", sqlConnection);

                // Add the parameter to the insert command.
                sqlCommand.Parameters.AddWithValue("@id", lastManifestId + 1);
                sqlCommand.Parameters.AddWithValue("@date", Manifest.Date);
                sqlCommand.Parameters.AddWithValue("@creator", Manifest.Creator);
                sqlCommand.Parameters.AddWithValue("@description", Manifest.Description);
                sqlCommand.Parameters.AddWithValue("@company", Manifest.Company);
                sqlCommand.Parameters.AddWithValue("@path", (lastManifestId + 1).ToString() + ".cxa");

                // Create a new SqlDataReader and execute the query
                sqlDataReader = sqlCommand.ExecuteReader();

            }

            return lastManifestId + 1;
        }

        public CManifest getArchiveById(SqlConnection sqlConnection, long archiveId)
        {
            // Instantiate a CManifest for the Manifest description
            CManifest ArchiveDescription = new CManifest();


            // Using block for the database queries
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the description of the assembly.
                SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[archive].[id], [TOOLDB].[dbo].[archive].[creator], [TOOLDB].[dbo].[archive].[description], [TOOLDB].[dbo].[archive].[company], [TOOLDB].[dbo].[archive].[path] FROM [TOOLDB].[dbo].[archive] WHERE [TOOLDB].[dbo].[archive].[id] = " + archiveId + ";", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = command.ExecuteReader();

                // Read the data until there are no more rows
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        ArchiveDescription.Date = sqlDataReader.GetString(1);
                        ArchiveDescription.Creator = sqlDataReader.GetString(2);
                        ArchiveDescription.Description = sqlDataReader.GetString(3);
                        ArchiveDescription.Company = sqlDataReader.GetString(4);
                        ArchiveDescription.ProjectFileName = sqlDataReader.GetString(0) + ".cxa";
                    }
                }
            }

            return ArchiveDescription;
        }

        //----------------------------------------------------------------------------------

        /*! deleteAssembly() removes an assembly from the database. Parameters are the connection to the 
         * database server and the id-number of the assembly. */

        public bool deleteAssembly(SqlConnection sqlConnection, long assemblyId)
        {

            // Use the existing SqlConnection
            using (sqlConnection)
            {
                // Create a new SqlCommand and delete the assembly.
                SqlCommand sqlCommand = new SqlCommand("DELETE FROM assembly WHERE id = " + assemblyId + ";", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = null;

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

                // Create a new SqlCommand and delete the assembly technology.
                sqlCommand = new SqlCommand("DELETE FROM assembly_technology WHERE assembly_id = " + assemblyId + ";", sqlConnection);

                // Create a new SqlDataReader and execute the query
                sqlDataReader = null;

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

            }

            return true;
        }

        //----------------------------------------------------------------------------------
        /* Part with the private class methods. These methods are not intended for the use 
         * outside of this class. */

        public void loadDinKeys()
        {

            SqlDataReader myDataReader = null;

            SqlConnection mySqlConnection = getSqlConnection();

            SqlCommand mySqlCommand = new SqlCommand("SELECT [TOOLDB].[dbo].[din4000_key].[name],[TOOLDB].[dbo].[din4000_key].[id] FROM [TOOLDB].[dbo].[din4000_key];", mySqlConnection);
            myDataReader = mySqlCommand.ExecuteReader(CommandBehavior.CloseConnection);

            if (myDataReader.HasRows)
            {
                while (myDataReader.Read())
                {
                    din4000keys.Add(myDataReader.GetString(0), myDataReader.GetInt64(1).ToString());

                }
            }

            // Always call Close when done reading.
            myDataReader.Close();

            // Close the connection when done with it.
            mySqlConnection.Close();

        }

        public Dictionary<string, string> loadKeyMapping(int camsystem, int tooltype)
        {
            Dictionary<string, string> KeyMapping = new Dictionary<string, string>();

            SqlDataReader myDataReader = null;

            SqlConnection mySqlConnection = getSqlConnection();

            SqlCommand mySqlCommand = new SqlCommand("SELECT [TOOLDB].[dbo].[import_mapping].[camsystem_id],[TOOLDB].[dbo].[import_mapping].[tooltype_id],[TOOLDB].[dbo].[import_mapping].[parameter],[TOOLDB].[dbo].[import_mapping].[key_id] FROM [TOOLDB].[dbo].[import_mapping] WHERE [TOOLDB].[dbo].[import_mapping].camsystem_id = " + camsystem.ToString() + " AND [TOOLDB].[dbo].[import_mapping].tooltype_id = " + tooltype.ToString() + ";", mySqlConnection);
            myDataReader = mySqlCommand.ExecuteReader(CommandBehavior.CloseConnection);

            if (myDataReader.HasRows)
            {
                while (myDataReader.Read())
                {
                    if(!KeyMapping.ContainsKey(myDataReader.GetString(2)))
                        KeyMapping.Add(myDataReader.GetString(2), (myDataReader.GetInt32(3)).ToString());
                }
            }

            // Always call Close when done reading.
            myDataReader.Close();

            // Close the connection when done with it.
            mySqlConnection.Close();

            return KeyMapping;

        }


        //public Dictionary<string, string> loadKeyMappingKeys(int camsystem, int tooltype)
        //{
        //    Dictionary<string, string> KeyMapping = new Dictionary<string, string>();

        //    SqlDataReader myDataReader = null;

        //    SqlConnection mySqlConnection = getSqlConnection();

        //    SqlCommand mySqlCommand = new SqlCommand("SELECT [TOOLDB].[dbo].[import_mapping].[camsystem_id], [TOOLDB].[dbo].[import_mapping].[tooltype_id], [TOOLDB].[dbo].[import_mapping].[parameter], [TOOLDB].[dbo].[import_mapping].[key_id] FROM [TOOLDB].[dbo].[import_mapping] WHERE [TOOLDB].[dbo].[import_mapping].[camsystem_id] = " + camsystem.ToString() + " AND [TOOLDB].[dbo].[import_mapping].[tooltype_id] = " + tooltype.ToString() + ";", mySqlConnection);
        //    myDataReader = mySqlCommand.ExecuteReader(CommandBehavior.CloseConnection);

        //    if (myDataReader.HasRows)
        //    {
        //        while (myDataReader.Read())
        //        {
        //            string ActualMappingKey = GetDINKey(mySqlConnection, myDataReader.GetInt32(3));
        //            if(!KeyMapping.ContainsKey(ActualMappingKey))
        //                KeyMapping.Add(ActualMappingKey, myDataReader.GetString(2));
        //        }
        //    }

        //    // Always call Close when done reading.
        //    myDataReader.Close();

        //    // Close the connection when done with it.
        //    mySqlConnection.Close();

        //    return KeyMapping;

        //}


        /*! insertArchive() stores a .cxa archive into the database. The method uses the given SqlConnection and inserts 
         the given date, creator, description, the company name and the filepath  via a SQL-query into the data tables.*/

        public long insertToolList(SqlConnection sqlConnection, string sheet_id, string creator, string description, string nc_program, string date, List<string> AssemblyIdList)
        {
            // Initialize the variable for the last id
            long lastSheetId = 0;
            long lastSheetEntrieId = 0;
            bool isUS = false;
            bool isDE = false;

            string[] dateParts = date.Split('.');

            isUS = TestDateUS(getSqlConnection());//JAN
            isDE = TestDateDE(getSqlConnection());//JAN  

            if (isUS == true)
            {
                date = dateParts[0] + "." + dateParts[1] + "." + dateParts[2];
            }
            if (isDE == true)
            {
                date = dateParts[1] + "." + dateParts[0] + "." + dateParts[2];
            }
            else
            {
                MessageBox.Show("Der Server auf dem die Datenbank abgelegt ist, verfügt nicht über die korrekte Regionseinstellung! Bitte wenden Sie sich an Ihren Administrator.");
            }

            // Use the existing SqlConnection
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the description of the tool.
                SqlCommand sqlCommand = new SqlCommand("SELECT MAX(Id) FROM sheet;", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = null;

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

                // Read the given id
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        try
                        {
                            lastSheetId = sqlDataReader.GetInt64(0);
                        }
                        catch (Exception ex)
                        {
                            debugMessage = ex.Message;
                            lastSheetId = 0;
                        }
                    }
                }

                // Create a new SqlCommand and query the description of the tool.
                sqlCommand = new SqlCommand("INSERT INTO sheet (id, sheet_id, creator, description, nc_program, date) VALUES (@id, @sheet_id, @creator, @description, @nc_program, @date)", sqlConnection);

                // Add the parameter to the insert command.
                sqlCommand.Parameters.AddWithValue("@id", lastSheetId + 1);
                sqlCommand.Parameters.AddWithValue("@sheet_id", sheet_id);
                sqlCommand.Parameters.AddWithValue("@creator", creator);
                sqlCommand.Parameters.AddWithValue("@description",description);
                sqlCommand.Parameters.AddWithValue("@nc_program", nc_program);
                sqlCommand.Parameters.AddWithValue("@date", date);

                // Create a new SqlDataReader and execute the query
                sqlDataReader = sqlCommand.ExecuteReader();


                // Create a new SqlCommand and query the description of the tool.
                sqlCommand = new SqlCommand("SELECT MAX(Id) FROM sheet_entries;", sqlConnection);

                // Create a new SqlDataReader and execute the query
                sqlDataReader = null;

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

                // Read the given id
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        try
                        {
                            lastSheetEntrieId = sqlDataReader.GetInt64(0);
                        }
                        catch (Exception ex)
                        {
                            debugMessage = ex.Message;
                            lastSheetEntrieId = 0;
                        }
                    }
                }

                foreach (string AssemblyId in AssemblyIdList)
                {
                    // Create a new SqlCommand and query the description of the tool.
                    sqlCommand = new SqlCommand("INSERT INTO [sheet_entries] (id, sheet_id, tool_id) VALUES (@id, @sheet_id, @tool_id)", sqlConnection);

                    // Add the parameter to the insert command.
                    sqlCommand.Parameters.AddWithValue("@id", lastSheetEntrieId + 1);
                    sqlCommand.Parameters.AddWithValue("@sheet_id", sheet_id);
                    sqlCommand.Parameters.AddWithValue("@tool_id", AssemblyId);

                    // Create a new SqlDataReader and execute the query
                    sqlDataReader = sqlCommand.ExecuteReader();
                    lastSheetEntrieId++;
                }




            }

            return lastSheetId + 1;
        }

        public Dictionary<string, string> getSheetListById(SqlConnection sqlConnection, string sheetId)
        {
            // Instantiate a CManifest for the Manifest description
            Dictionary<string, string> ToolList = new Dictionary<string, string>();


            // Using block for the database queries
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the description of the assembly.
                SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[sheet].[sheet_id], [TOOLDB].[dbo].[sheet].[description], [TOOLDB].[dbo].[sheet].[creator], [TOOLDB].[dbo].[sheet].[nc_program], [TOOLDB].[dbo].[sheet].[date] FROM [TOOLDB].[dbo].[sheet] WHERE [TOOLDB].[dbo].[sheet].[sheet_id] = '" + sheetId + "';", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = command.ExecuteReader();

                // Read the data until there are no more rows
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        ToolList.Add("sheet_id", (sqlDataReader.GetValue(0)).ToString());
                        ToolList.Add("description", sqlDataReader.GetString(1));
                        ToolList.Add("creator", sqlDataReader.GetString(2));
                        ToolList.Add("nc_program", sqlDataReader.GetString(3));
                        ToolList.Add("date", sqlDataReader.GetDateTime(4).ToShortDateString());
                    }
                }
                // Create a new SqlCommand and query the description of the assembly.
                command = new SqlCommand("SELECT [TOOLDB].[dbo].[sheet].[sheet_id], [TOOLDB].[dbo].[sheet_entries].[tool_id] FROM [TOOLDB].[dbo].[sheet] LEFT JOIN [TOOLDB].[dbo].[sheet_entries] ON  [TOOLDB].[dbo].[sheet].[sheet_id] = [TOOLDB].[dbo].[sheet_entries].[sheet_id] WHERE [TOOLDB].[dbo].[sheet].[sheet_id] = '" + sheetId + "';", sqlConnection);

                // Create a new SqlDataReader and execute the query
                sqlDataReader = command.ExecuteReader();

                // Read the data until there are no more rows
                string ToolId = "";

                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        try
                        {
                            ToolId = ToolId + (sqlDataReader.GetValue(1)).ToString() + "|";
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message + " - Keine Ergebnisse. Daher Wandlungsfehler wg. null");
                        }
                    }
                }
                ToolList.Add("tools", ToolId);

            }




            return ToolList;
        }

        //----------------------------------------------------------------------------------



        public void loadDinNames()
        {

            SqlDataReader myDataReader = null;

            SqlConnection mySqlConnection = getSqlConnection();

            SqlCommand mySqlCommand = new SqlCommand("SELECT [TOOLDB].[dbo].[din4000_key].[key], [TOOLDB].[dbo].[din4000_key].[id] FROM [TOOLDB].[dbo].[din4000_key];", mySqlConnection);
            myDataReader = mySqlCommand.ExecuteReader(CommandBehavior.CloseConnection);

            if (myDataReader.HasRows)
            {
                while (myDataReader.Read())
                {
                    if (!din4000names.ContainsKey(myDataReader.GetString(0)))
                    {
                        din4000names.Add(myDataReader.GetInt64(1).ToString(), myDataReader.GetString(0));
                    }
                }
            }

            //din4000names.Sort();

            // Always call Close when done reading.
            myDataReader.Close();

            // Close the connection when done with it.
            mySqlConnection.Close();

        }

        public void loadProcessNames()
        {

            SqlDataReader myDataReader = null;

            SqlConnection mySqlConnection = getSqlConnection();

            SqlCommand mySqlCommand = new SqlCommand("SELECT [TOOLDB].[dbo].[process].[id], [TOOLDB].[dbo].[process].[process] FROM [TOOLDB].[dbo].[process];", mySqlConnection);
            myDataReader = mySqlCommand.ExecuteReader(CommandBehavior.CloseConnection);

            if (myDataReader.HasRows)
            {
                while (myDataReader.Read())
                {
                    if (!processNames.ContainsKey(myDataReader.GetInt64(0).ToString()))
                    {
                        processNames.Add(myDataReader.GetInt64(0).ToString(), myDataReader.GetString(1));
                    }
                }
            }

            // Always call Close when done reading.
            myDataReader.Close();

            // Close the connection when done with it.
            mySqlConnection.Close();

        }

        public void loadMaterials()
        {

            SqlDataReader myDataReader = null;

            SqlConnection mySqlConnection = getSqlConnection();

            SqlCommand mySqlCommand = new SqlCommand("SELECT [TOOLDB].[dbo].[material].[material], [TOOLDB].[dbo].[material].[id] FROM [TOOLDB].[dbo].[material];", mySqlConnection);
            myDataReader = mySqlCommand.ExecuteReader(CommandBehavior.CloseConnection);

            if (myDataReader.HasRows)
            {
                while (myDataReader.Read())
                {
                    if (!materialNames.ContainsKey(myDataReader.GetString(0)))
                    {
                        materialNames.Add(myDataReader.GetInt64(1).ToString(), myDataReader.GetString(0));
                    }
                }
            }

            //din4000names.Sort();

            // Always call Close when done reading.
            myDataReader.Close();

            // Close the connection when done with it.
            mySqlConnection.Close();

        }


        public void loadTechKeys()
        {

            SqlDataReader myDataReader = null;

            SqlConnection mySQLConnection = getSqlConnection();
            if (mySQLConnection == null)
                return;


            SqlCommand mySQLiteCommand = new SqlCommand("SELECT [TOOLDB].[dbo].[tech_keys].[id], [TOOLDB].[dbo].[tech_keys].[key] FROM [TOOLDB].[dbo].[tech_keys];", mySQLConnection);
            myDataReader = mySQLiteCommand.ExecuteReader(CommandBehavior.CloseConnection);

            if (myDataReader.HasRows)
            {
                while (myDataReader.Read())
                {
                    if (!techKeys.ContainsKey((myDataReader.GetInt64(0)).ToString()))
                    {
                        techKeys.Add(myDataReader.GetInt64(0).ToString(), myDataReader.GetString(1));
                    }
                }
            }

            //din4000names.Sort();

            // Always call Close when done reading.
            myDataReader.Close();

            // Close the connection when done with it.
            mySQLConnection.Close();

        }

        
        
        
        public void loadMillingProcessNames()
        {

            SqlDataReader myDataReader = null;

            SqlConnection mySqlConnection = getSqlConnection();

            SqlCommand mySqlCommand = new SqlCommand("SELECT [TOOLDB].[dbo].[mill_process].[name], [TOOLDB].[dbo].[mill_process].[id] FROM [TOOLDB].[dbo].[mill_process];", mySqlConnection);
            myDataReader = mySqlCommand.ExecuteReader(CommandBehavior.CloseConnection);

            if (myDataReader.HasRows)
            {
                while (myDataReader.Read())
                {
                    if (!millProcessNames.ContainsKey(myDataReader.GetString(0)))
                    {
                        millProcessNames.Add(myDataReader.GetInt64(1).ToString(), myDataReader.GetString(0));
                    }
                }
            }

            //din4000names.Sort();

            // Always call Close when done reading.
            myDataReader.Close();

            // Close the connection when done with it.
            mySqlConnection.Close();

        }


        public string m_GetDINKey(SqlConnection sqlConnection, long m_pKey)
        {
            SqlDataReader myDataReader = null;
            string m_pKeyWord = null;

            using (sqlConnection)
            {
                SqlCommand mySqlCommand = new SqlCommand("SELECT name, id, [key] FROM din4000_key WHERE id = " + m_pKey + ";", sqlConnection);
                myDataReader = mySqlCommand.ExecuteReader();

                if (myDataReader.HasRows)
                {
                    while (myDataReader.Read())
                    {
                        m_pKeyWord = myDataReader.GetString(2);

                    }
                }

            }

            // Always call Close when done reading.
           //myDataReader.Close();

            // Close the connection when done with it.
            //mySqlConnection.Close();
            sqlConnection.Close();
            sqlConnection.Dispose();
            return m_pKeyWord;
        }


        public long GetLastId(SqlConnection sqlConnection, string TableName)
        {
            long lastId = 0;

            // Use the existing SqlConnection
            using (sqlConnection)
            {
                SqlCommand sqlCommand = new SqlCommand("SELECT MAX(Id) FROM " + TableName + ";", sqlConnection);
                
                SqlDataReader myDataReader = sqlCommand.ExecuteReader();

                if (myDataReader.HasRows)
                {
                    while (myDataReader.Read())
                    {
                        try
                        {
                            lastId = myDataReader.GetInt64(0);
                        }
                        catch (Exception ex)
                        {
                            debugMessage = ex.Message;
                            lastId = 0;
                        }

                    }
                }
            }

            // Always call Close when done reading.
            //myDataReader.Close();

            // Close the connection when done with it.
            //mySqlConnection.Close();
            sqlConnection.Close();
            sqlConnection.Dispose();
            return lastId;
        }

        public string GetDINName(SqlConnection SQLConnection, int Key)
        {
            SqlDataReader SQLDataReader = null;
            string KeyWord = null;

            //SQLiteConnection SQLConnection = getSQLiteConnection();

            SqlCommand SQLCommand = new SqlCommand("SELECT [TOOLDB].[dbo].[din4000_key].[name], [TOOLDB].[dbo].[din4000_key].[id], [TOOLDB].[dbo].[din4000_key].[key] FROM [TOOLDB].[dbo].[din4000_key] WHERE [TOOLDB].[dbo].[din4000_key].[id] = " + Key + ";", SQLConnection);
            SQLDataReader = SQLCommand.ExecuteReader();

            if (SQLDataReader.HasRows)
            {
                while (SQLDataReader.Read())
                {
                    KeyWord = SQLDataReader.GetString(2);

                }
            }

            // Always call Close when done reading.
            SQLDataReader.Close();

            // Close the connection when done with it.
            SQLConnection.Close();
            return KeyWord;
        }

        public string GetDINKey(SqlConnection SQLoldConnection, int Key)
        {
            SqlDataReader SQLDataReader = null;
            string KeyWord = null;

            SqlConnection SQLConnection = getSqlConnection();

            SqlCommand SQLCommand = new SqlCommand("SELECT [TOOLDB].[dbo].[din4000_key].[name], [TOOLDB].[dbo].[din4000_key].[id], [TOOLDB].[dbo].[din4000_key].[key] FROM [TOOLDB].[dbo].[din4000_key] WHERE id = " + Key + ";", SQLConnection);
            SQLDataReader = SQLCommand.ExecuteReader();

            if (SQLDataReader.HasRows)
            {
                while (SQLDataReader.Read())
                {
                    KeyWord = SQLDataReader.GetString(0);

                }
            }

            // Always call Close when done reading.
            SQLDataReader.Close();

            // Close the connection when done with it.
            SQLConnection.Close();
            return KeyWord;
        }



        public Dictionary<long, bool> loadVisibleToolEntries(int tooltype)
        {
            Dictionary<long, bool> ToolEntries = new Dictionary<long, bool>();

            SqlDataReader myDataReader = null;

            SqlConnection mySqlConnection = getSqlConnection();

            SqlCommand mySqlCommand = new SqlCommand("SELECT [TOOLDB].[dbo].[tool_entries].[wz_typ], [TOOLDB].[dbo].[tool_entries].[entry], [TOOLDB].[dbo].[tool_entries].[mandatory], [TOOLDB].[dbo].[tool_entries].[visible] FROM [TOOLDB].[dbo].[tool_entries] WHERE [TOOLDB].[dbo].[tool_entries].[wz_typ] = " + tooltype.ToString() + ";", mySqlConnection);
            myDataReader = mySqlCommand.ExecuteReader(CommandBehavior.CloseConnection);

            if (myDataReader.HasRows)
            {
                while (myDataReader.Read())
                {
                    if (myDataReader.GetBoolean(3))
                        ToolEntries.Add(myDataReader.GetInt32(1), (myDataReader.GetBoolean(2)));
                }
            }

            // Always call Close when done reading.
            myDataReader.Close();

            // Close the connection when done with it.
            mySqlConnection.Close();

            return ToolEntries;

        }



        public Dictionary<long, bool> loadToolEntries(int tooltype)
        {
            Dictionary<long, bool> ToolEntries = new Dictionary<long, bool>();

            SqlDataReader myDataReader = null;

            SqlConnection mySqlConnection = getSqlConnection();

            SqlCommand mySqlCommand = new SqlCommand("SELECT [TOOLDB].[dbo].[tool_entries].[wz_typ], [TOOLDB].[dbo].[tool_entries].[entry], [TOOLDB].[dbo].[tool_entries].[mandatory], [TOOLDB].[dbo].[tool_entries].[visible] FROM [TOOLDB].[dbo].[tool_entries] WHERE [TOOLDB].[dbo].[tool_entries].[wz_typ] = " + tooltype.ToString() + ";", mySqlConnection);
            myDataReader = mySqlCommand.ExecuteReader(CommandBehavior.CloseConnection);

            if (myDataReader.HasRows)
            {
                while (myDataReader.Read())
                {
                        ToolEntries.Add(myDataReader.GetInt32(1), (myDataReader.GetBoolean(2)));
                }
            }

            // Always call Close when done reading.
            myDataReader.Close();

            // Close the connection when done with it.
            mySqlConnection.Close();

            return ToolEntries;

        }

        /*! UpdateToolEntries() updates the state of the tool entries in the database
         * @param SQLConnection to the database
         * @param The ToolEntries of the current tooltype
         * @param The tooltype as integer
         */

        public void UpdateToolEntries(SqlConnection SqlConnection, Dictionary<long, bool> ToolEntries, int ToolType)
        {
            // Create a new SqlDataReader and execute the query
            SqlDataReader SQLDataReader = null;

            // Use the existing SqlConnection
            using (SqlConnection)
            {

                for (int i = 0; i < ToolEntries.Count; i++)
                {
                    string MandatoryEntry = "0";
                    if (ToolEntries.ElementAt(i).Value)
                        MandatoryEntry = "1";

                    SqlCommand SQLCommand = new SqlCommand("UPDATE tool_entries SET visible = '" + MandatoryEntry + "' WHERE entry = '" + ToolEntries.ElementAt(i).Key.ToString() + "' AND wz_typ = '" + ToolType + "';", SqlConnection);
                    SQLDataReader = SQLCommand.ExecuteReader();


                }

            }
        }

        public long InsertTechDataTool(SqlConnection SqlConnection, Dictionary<string, string> dicTechData)
        {

            long lastId = GetLastId(getSqlConnection(), "tool_technology") + 1;

            // Use the existing SqlConnection
            using (SqlConnection)
            {
                SqlDataReader DataReader = null;

                // Create a new SqlCommand and query the description of the tool.
                SqlCommand sqlCommand = new SqlCommand("INSERT INTO [TOOLDB].[dbo].[tool_technology] (id, tool_id, mill_process_id, material_id, processing_id, description, value) VALUES (@id, @tool_id, @mill_process_id, @material_id, @processing_id, @description, @value);", SqlConnection);

                // Add the parameter to the insert command.
                sqlCommand.Parameters.AddWithValue("@id", lastId);
                sqlCommand.Parameters.AddWithValue("@tool_id", dicTechData["tool_id"]);
                sqlCommand.Parameters.AddWithValue("@mill_process_id", dicTechData["mill_process_id"]);
                sqlCommand.Parameters.AddWithValue("@material_id", dicTechData["material_id"]);
                sqlCommand.Parameters.AddWithValue("@processing_id", dicTechData["processing_id"]);
                sqlCommand.Parameters.AddWithValue("@description", dicTechData["description"]);
                sqlCommand.Parameters.AddWithValue("@value", dicTechData["value"]);

                DataReader = sqlCommand.ExecuteReader();


            }
            return lastId;
        }

        public long InsertTechDataAssembly(SqlConnection SqlConnection, Dictionary<string, string> dicTechData)
        {

            long lastId = GetLastId(getSqlConnection(), "assembly_technology") + 1;

            // Use the existing SqlConnection
            using (SqlConnection)
            {
                SqlDataReader DataReader = null;

                // Create a new SqlCommand and query the description of the tool.
                SqlCommand sqlCommand = new SqlCommand("INSERT INTO [TOOLDB].[dbo].[assembly_technology] (id, assembly_id, mill_process_id, material_id, processing_id, description, value) VALUES (@id, @assembly_id, @mill_process_id, @material_id, @processing_id, @description, @value);", SqlConnection);

                // Add the parameter to the insert command.
                sqlCommand.Parameters.AddWithValue("@id", lastId);
                sqlCommand.Parameters.AddWithValue("@assembly_id", dicTechData["assembly_id"]);
                sqlCommand.Parameters.AddWithValue("@mill_process_id", dicTechData["mill_process_id"]);
                sqlCommand.Parameters.AddWithValue("@material_id", dicTechData["material_id"]);
                sqlCommand.Parameters.AddWithValue("@processing_id", dicTechData["processing_id"]);
                sqlCommand.Parameters.AddWithValue("@description", dicTechData["description"]);
                sqlCommand.Parameters.AddWithValue("@value", dicTechData["value"]);

                DataReader = sqlCommand.ExecuteReader();


            }
            return lastId;
        }


        public bool InsertToolProfile(SqlConnection SqlConnection, string xmlProfile, long toolId)
        {
 
            long lastId = GetLastId(getSqlConnection(), "tool_profile") + 1;

            // Use the existing SqlConnection
            using (SqlConnection)
            {
                SqlDataReader DataReader = null;

                // Create a new SqlCommand and query the description of the tool.
                SqlCommand sqlCommand = new SqlCommand("INSERT INTO [TOOLDB].[dbo].[tool_profile] VALUES (" + lastId + "," + toolId + ",'" + xmlProfile + "');", SqlConnection);

                DataReader = sqlCommand.ExecuteReader();


            }
            return true;
        }

        public bool InsertHolderProfile(SqlConnection m_pSqlConnection, string xmlProfile, long toolId)
        {

            long lastId = GetLastId(getSqlConnection(), "holder_profile") + 1;

            // Use the existing SqlConnection
            using (m_pSqlConnection)
            {
                SqlDataReader DataReader = null;

                // Create a new SqlCommand and query the description of the tool.
                SqlCommand sqlCommand = new SqlCommand("INSERT INTO [TOOLDB].[dbo].[holder_profile] VALUES (" + lastId + "," + toolId + ",'" + xmlProfile + "');", m_pSqlConnection);

                DataReader = sqlCommand.ExecuteReader();


            }
            return true;
        }

        public string GetToolProfileById(SqlConnection m_pSqlConnection, long profileId)
        {
            // Instantiate a dictionary for the tool description
            string xmlProfile = null;


            // Using block for the database queries
            using (m_pSqlConnection)
            {
                // Create a new SqlCommand and query the description of the tool.
                SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[tool_profile].[id], [TOOLDB].[dbo].[tool_profile].[profile] FROM [TOOLDB].[dbo].[tool_profile] WHERE [TOOLDB].[dbo].[tool_profile].[tool_id] = " + profileId + ";", m_pSqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = command.ExecuteReader();

                // Read the data until there are no more rows
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        // Add the description to the toolDescription dictionary
                        xmlProfile = sqlDataReader.GetString(1);
                    }
                }
            }

            if (String.IsNullOrEmpty(xmlProfile))
                return "";
            else
                return xmlProfile.Replace("><", ">" + Environment.NewLine + "<");

        }

        public string GetHolderProfileById(SqlConnection m_pSqlConnection, long profileId)
        {
            // Instantiate a dictionary for the tool description
            string xmlProfile = null;


            // Using block for the database queries
            using (m_pSqlConnection)
            {
                // Create a new SqlCommand and query the description of the tool.
                SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[holder_profile].[id], [TOOLDB].[dbo].[holder_profile].[profile] FROM [TOOLDB].[dbo].[holder_profile] WHERE [TOOLDB].[dbo].[holder_profile].[holder_id] = " + profileId + ";", m_pSqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = command.ExecuteReader();

                // Read the data until there are no more rows
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        // Add the description to the toolDescription dictionary
                        xmlProfile = sqlDataReader.GetString(1);
                    }
                }
            }

            if (String.IsNullOrEmpty(xmlProfile))
                return "";
            else
                return xmlProfile.Replace("><", ">" + Environment.NewLine + "<");

        }

        //----------------------------------------------------------------------------------

        /*! GetHolderModelById() Reads the stl graphic of a tool from the database.
         * @param SQLiteConnection to the database
         * @param long integer containing the tool id
         * @return string variable which contains the tl-contour
         */


        public byte[] GetHolderModelById(SqlConnection SQLConnection, long modelId)
        {
            // Instantiate a dictionary for the Tool description
            byte[] HolderModel = null;

            using (SQLConnection)
            {
                    // Create a new SqLiteCommand and query the description of the Tool.
                    SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[holder_model].[model] FROM [TOOLDB].[dbo].[holder_model] WHERE [TOOLDB].[dbo].[holder_model].[holder_id] = " + modelId + ";");

                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (reader["model"] != null && !Convert.IsDBNull(reader["model"]))
                            {
                                HolderModel = (byte[])reader["model"];
                            }
                        }
                    }

                if (HolderModel != null)
                {
                    MemoryStream input = new MemoryStream(HolderModel);
                    MemoryStream output = new MemoryStream();
                    using (DeflateStream dstream = new DeflateStream(input, CompressionMode.Decompress))
                    {
                        dstream.CopyTo(output);
                    }
                    return output.ToArray();
                }
                else
                    return null;
            }
        }


        //----------------------------------------------------------------------------------

        /*! GetToolModelById() Reads the stl graphic of a tool from the database.
         * @param SQLiteConnection to the database
         * @param long integer containing the tool id
         * @return string variable which contains the tl-contour
         */


        public byte[] GetToolModelById(SqlConnection SQLConnection, long modelId)
        {
            // Instantiate a dictionary for the Tool description
            byte[] ToolModel = null;

            using (SQLConnection)
            {
                // Create a new SqLiteCommand and query the description of the Tool.
                SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[tool_model].[model] FROM [TOOLDB].[dbo].[tool_model] WHERE [TOOLDB].[dbo].[tool_model].[tool_id] = " + modelId + ";");

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        if (reader["model"] != null && !Convert.IsDBNull(reader["model"]))
                        {
                            ToolModel = (byte[])reader["model"];
                        }
                    }
                }

                if (ToolModel != null)
                {
                    MemoryStream input = new MemoryStream(ToolModel);
                    MemoryStream output = new MemoryStream();
                    using (DeflateStream dstream = new DeflateStream(input, CompressionMode.Decompress))
                    {
                        dstream.CopyTo(output);
                    }
                    return output.ToArray();
                }
                else
                    return null;
            }
        }

        //----------------------------------------------------------------------------------

        /*! GetAssemblyModelById() Reads the stl graphic of a assembly from the database.
         * @param SQLiteConnection to the database
         * @param long integer containing the tool id
         * @return string variable which contains the tl-contour
         */
        public byte[] GetAssemblyModelById(SqlConnection SQLConnection, long modelId)
        {
            // Instantiate a dictionary for the Tool description
            byte[] ToolModel = null;

            using (SQLConnection)
            {
                // Create a new SqLiteCommand and query the description of the Tool.
                SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[assembly_model].[model] FROM [TOOLDB].[dbo].[assembly_model] WHERE [TOOLDB].[dbo].[assembly_model].[assembly_id] = " + modelId + ";");

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        if (reader["model"] != null && !Convert.IsDBNull(reader["model"]))
                        {
                            ToolModel = (byte[])reader["model"];
                        }
                    }
                }

                if (ToolModel != null)
                {
                    MemoryStream input = new MemoryStream(ToolModel);
                    MemoryStream output = new MemoryStream();
                    using (DeflateStream dstream = new DeflateStream(input, CompressionMode.Decompress))
                    {
                        dstream.CopyTo(output);
                    }
                    return output.ToArray();
                }
                else
                    return null;
            }
        }


        //----------------------------------------------------------------------------------

        /*! GetTurningToolModelById() Reads the stl graphic of a turning tool from the database.
         * @param SQLiteConnection to the database
         * @param long integer containing the tool id
         * @return string variable which contains the tl-contour
         */
        public byte[] GetTurningToolModelById(SqlConnection SQLConnection, long modelId)
        {
            // Instantiate a dictionary for the Tool description
            byte[] ToolModel = null;

            using (SQLConnection)
            {
                // Create a new SqLiteCommand and query the description of the Tool.
                SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[turning_tool_model].[model] FROM [TOOLDB].[dbo].[turning_tool_model] WHERE [TOOLDB].[dbo].[turning_tool_model].[tool_id] = " + modelId + ";");

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        if (reader["model"] != null && !Convert.IsDBNull(reader["model"]))
                        {
                            ToolModel = (byte[])reader["model"];
                        }
                    }
                }

                if (ToolModel != null)
                {
                    MemoryStream input = new MemoryStream(ToolModel);
                    MemoryStream output = new MemoryStream();
                    using (DeflateStream dstream = new DeflateStream(input, CompressionMode.Decompress))
                    {
                        dstream.CopyTo(output);
                    }
                    return output.ToArray();
                }
                else
                    return null;
            }
        }

        public string GetDictionaryKey(Dictionary<string, string> Dict, string value)
        {
            string key = null;
            foreach (KeyValuePair<string, string> data in Dict)
            {
                if (data.Value.Equals(value))
                    key = data.Key.ToString();
            }
            return key;
        }

        public string GetSortedDictionaryKey(SortedDictionary<string, string> Dict, string value)
        {
            string key = null;
            foreach (KeyValuePair<string, string> data in Dict)
            {
                if (data.Value.Equals(value))
                    key = data.Key.ToString();
            }
            return key;
        }


        //----------------------------------------------------------------------------------

        /*! InsertToolModel() Inserts the svg profile string of a tool into the database.
         * @param SQLiteConnection to the database
         * @param string variable which contains the svg-contour
         * @param long integer containing the tool id
         * @return bool flag, true if succesful #/ false if not
         */

        public bool InsertToolModel(SqlConnection SQLConnection, byte[] ToolModel, long toolId)
        {

            MemoryStream output = new MemoryStream();
            using (DeflateStream dstream = new DeflateStream(output, CompressionMode.Compress))
            {
                dstream.Write(ToolModel, 0, ToolModel.Length);
            }

            long lastId = GetLastId(getSqlConnection(), "tool_model") + 1;

            // Use the existing SQLiteConnection
            using (SQLConnection)
            {
                try
                {
                    // Create a new SqLiteCommand and query the description of the Tool.
                    SqlCommand SQLCommand = new SqlCommand();

                    SQLCommand.CommandText = "INSERT INTO tool_model (id, tool_id, model) VALUES (@id, @tool_id, @model)";
                    SQLCommand.Parameters.AddWithValue("@id", lastId);
                    SQLCommand.Parameters.AddWithValue("@tool_id", toolId);
                    SQLCommand.Parameters.Add("@model", DbType.Binary).Value = output.ToArray();
                    SQLCommand.ExecuteNonQuery();
                }
                catch(Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

            }
            return true;
        }

        //----------------------------------------------------------------------------------

        /*! InsertAssemblyModel() Inserts the svg profile string of a assembly into the database.
         * @param SQLiteConnection to the database
         * @param string variable which contains the svg-contour
         * @param long integer containing the tool id
         * @return bool flag, true if succesful #/ false if not
         */

        public bool InsertAssemblyModel(SqlConnection SQLConnection, byte[] AssemblyModel, long AssemblyId)
        {

            MemoryStream output = new MemoryStream();
            using (DeflateStream dstream = new DeflateStream(output, CompressionMode.Compress))
            {
                dstream.Write(AssemblyModel, 0, AssemblyModel.Length);
            }

            long lastId = GetLastId(getSqlConnection(), "assembly_model") + 1;

            // Use the existing SQLiteConnection
            using (SQLConnection)
            {
                try
                {
                    // Create a new SqLiteCommand and query the description of the Tool.
                    SqlCommand SQLCommand = new SqlCommand();

                    SQLCommand.CommandText = "INSERT INTO assembly_model (id, assembly_id, model) VALUES (@id, @assembly_id, @model)";
                    SQLCommand.Parameters.AddWithValue("@id", lastId);
                    SQLCommand.Parameters.AddWithValue("@assembly_id", AssemblyId);
                    SQLCommand.Parameters.Add("@model", DbType.Binary).Value = output.ToArray();
                    SQLCommand.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

            }
            return true;
        }

        //----------------------------------------------------------------------------------

        /*! InsertHolderModel() Inserts the svg profile string of a tool into the database.
         * @param SQLiteConnection to the database
         * @param string variable which contains the svg-contour
         * @param long integer containing the tool id
         * @return bool flag, true if succesful #/ false if not
         */

        public bool InsertHolderModel(SqlConnection SQLConnection, byte[] HolderModel, long holderId)
        {

            MemoryStream output = new MemoryStream();
            using (DeflateStream dstream = new DeflateStream(output, CompressionMode.Compress))
            {
                dstream.Write(HolderModel, 0, HolderModel.Length);
            }

            long lastId = GetLastId(getSqlConnection(), "holder_model") + 1;

            // Use the existing SQLiteConnection
            using (SQLConnection)
            {
                try
                {
                    // Create a new SqLiteCommand and query the description of the Tool.
                    SqlCommand SQLCommand = new SqlCommand();

                    SQLCommand.CommandText = "INSERT INTO holder_model (id, holder_id, model) VALUES (@id, @holder_id, @model)";
                    SQLCommand.Parameters.AddWithValue("@id", lastId);
                    SQLCommand.Parameters.AddWithValue("@holder_id", holderId);
                    SQLCommand.Parameters.Add("@model", DbType.Binary).Value = output.ToArray();
                    SQLCommand.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

            }
            return true;
        }

        //----------------------------------------------------------------------------------

        /*! InsertTurningToolModel() Inserts the svg profile string of a tool into the database.
         * @param SQLiteConnection to the database
         * @param byte array which contains the stl model
         * @param long integer containing the tool id
         * @return bool flag, true if succesful #/ false if not
         */

        public bool InsertTurningToolModel(SqlConnection SQLConnection, byte[] TurningToolModel, long TuningToolId)
        {

            MemoryStream output = new MemoryStream();
            using (DeflateStream dstream = new DeflateStream(output, CompressionMode.Compress))
            {
                dstream.Write(TurningToolModel, 0, TurningToolModel.Length);
            }

            long lastId = GetLastId(getSqlConnection(), "turning_tool_model") + 1;

            // Use the existing SQLiteConnection
            // Use the existing SQLiteConnection
            using (SQLConnection)
            {
                try
                {
                    // Create a new SqLiteCommand and query the description of the Tool.
                    SqlCommand SQLCommand = new SqlCommand();

                    SQLCommand.CommandText = "INSERT INTO turning_tool_model (id, tool_id, model) VALUES (@id, @tool_id, @model)";
                    SQLCommand.Parameters.AddWithValue("@id", lastId);
                    SQLCommand.Parameters.AddWithValue("@tool_id", TuningToolId);
                    SQLCommand.Parameters.Add("@model", DbType.Binary).Value = output.ToArray();
                    SQLCommand.ExecuteNonQuery();

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    return false;
                }
            }
            return true;
        }



        //-------------------------------------------------------------------------
        // ExportSvgData reads the given XML File and parses the content.
        //-------------------------------------------------------------------------
        public List<CPoint2d> getHolderProfileSimById(string svgText)
        {
            // This is the offset for the conversion of the file
            float rOffsetX = 0;

            // Instantiate the XML elements
            List<CPoint2d> PointList = new List<CPoint2d>();

            // If there is no text to parse, then stop here            
            if (String.IsNullOrEmpty(svgText) || (svgText.Length < 23))
                return null;

            //// Cut out the header of the SVG-Text
            //if (!string.IsNullOrEmpty(svgText))
            //    svgText = svgText.Substring(22);

            // Instantiate a new XML Reader
            XmlReader pXmlReader = XmlReader.Create(new StringReader(svgText));

            // Create Lists for the different elements and two lists for the cut and nocut part of the svg file. 
            List<CPoint2d> lCut = new List<CPoint2d>();

            // Begin reading the svg data
            while (pXmlReader.Read())
            {
                // Instantiate a new class for the elements


                switch (pXmlReader.Name)
                {
                    // The g-attribute shows if it is a cut or nocut area
                    case "g":
                        break;

                    // If the attribut name is line, we create a line node.
                    case "line":
                        PointList.AddRange(convertSvgLineValue(pXmlReader));
                        break;
                    // If the attribut name is a path, we create a cwarc or ccwarc node.
                    case "path":
                        PointList.AddRange(convertSvgPathValue(pXmlReader));

                        break;
                    // The first parameter of the view box attribute is used for the calculation of the tool transformation
                    case "svg":
                        if (pXmlReader.IsStartElement())
                            try
                            {
                                rOffsetX = Convert.ToSingle(pXmlReader.GetAttribute("viewBox").Split()[0].Replace(",", "."), CultureInfo.InvariantCulture);
                            }
                            catch (Exception ex)
                            {
                                debugMessage = ex.Message;
                            }
                        break;

                }

            }

            return PointList;
        }


        //-------------------------------------------------------------------------
        // Reads the svg line string and converts it to line
        //-------------------------------------------------------------------------
        private static List<CPoint2d> convertSvgLineValue(XmlReader reader)
        {
            List<CPoint2d> PointList = new List<CPoint2d>();
            
            CPoint2d elem2D = new CPoint2d(Convert.ToDouble(reader.GetAttribute("x1").Replace(",", "."), CultureInfo.InvariantCulture),
                Convert.ToDouble(reader.GetAttribute("y1").Replace(",", "."), CultureInfo.InvariantCulture));

            PointList.Add(elem2D);

            elem2D.X = Convert.ToDouble(reader.GetAttribute("x2").Replace(",", "."), CultureInfo.InvariantCulture);
            elem2D.Y = Convert.ToDouble(reader.GetAttribute("y2").Replace(",", "."), CultureInfo.InvariantCulture);

            PointList.Add(elem2D);

            return PointList;
        }

        //-------------------------------------------------------------------------
        // Reads the Path string and converts it to cwarc or ccwarc
        //-------------------------------------------------------------------------
        private static List<CPoint2d> convertSvgPathValue(XmlReader reader)
        {
            List<CPoint2d> PointList = new List<CPoint2d>();

            string[] pathString = reader.GetAttribute("d").Split();

            //string startAngleString = reader.GetAttribute("start_angle");
            //string endAngleString = reader.GetAttribute("end_angle");

            //elem2D.rCx = Convert.ToDouble((pathString[0].Split(',')[0]).Substring(1).Replace(",", "."), CultureInfo.InvariantCulture);
            //elem2D.rCy = Convert.ToDouble((pathString[0].Split(',')[1]).Replace(",", "."), CultureInfo.InvariantCulture);

            //elem2D.rEx = Convert.ToDouble((pathString[4].Split(',')[0]).Replace(",", "."), CultureInfo.InvariantCulture);
            //elem2D.rEy = Convert.ToDouble((pathString[4].Split(',')[1]).Replace(",", "."), CultureInfo.InvariantCulture);

            //elem2D.rSx = Convert.ToDouble((pathString[1].Split(',')[0]).Substring(1).Replace(",", "."), CultureInfo.InvariantCulture);
            //elem2D.rSy = Convert.ToDouble((pathString[1].Split(',')[1]).Replace(",", "."), CultureInfo.InvariantCulture);

            //if (pathString[3].Split(',')[2] == "1")
            //    elem2D.type = "cwarc";
            //else
            //    elem2D.type = "ccwarc";

            //if (double.Parse(pathString[4].Split(',')[0]) < (double.Parse(pathString[4].Split(',')[1])))
            ////{
            ////    if (elem2D.type == "cwarc")
            //        elem2D.type = "cwarc";
            //    else
            //        elem2D.type = "ccwarc";

            //}

            return PointList;
        }

        public List<CPoint2d> CreateDummyHolder(bool IsMetric)
        {
            List<CPoint2d> PointList = new List<CPoint2d>();

            double MeasurementValue = 0;

            if (IsMetric)
            {
                MeasurementValue = 25;
            }
            else
            {
                MeasurementValue = 0.984;
            }

            CPoint2d elem2D1 = new CPoint2d(0, 0);
            CPoint2d elem2D2 = new CPoint2d(MeasurementValue, 0);
            CPoint2d elem2D3 = new CPoint2d(MeasurementValue, MeasurementValue);
            CPoint2d elem2D4 = new CPoint2d(0, MeasurementValue);
            PointList.Add(elem2D1);
            PointList.Add(elem2D2);
            PointList.Add(elem2D3);
            PointList.Add(elem2D4);


            return PointList;
        }
        /*! DeleteToolContour() removes a tool from the database. Parameters are the connection to the 
         * database server and the id-number of the tool. */

        public bool DeleteToolContour(SqlConnection sqlConnection, long toolId)
        {

            // Use the existing SqlConnection
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the description of the tool.
                SqlCommand sqlCommand = new SqlCommand("DELETE [dbo].[tool_profile] FROM tool_profile WHERE tool_id = " + toolId + ";", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = null;

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

            }

            return true;
        }

        /*! DeleteToolModel() removes a tool from the database. Parameters are the connection to the 
         * database server and the id-number of the tool. */

        public bool DeleteToolModel(SqlConnection sqlConnection, long toolId)
        {

            // Use the existing SqlConnection
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the description of the tool.
                SqlCommand sqlCommand = new SqlCommand("DELETE [dbo].[model] FROM tool_model WHERE tool_id = " + toolId + ";", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = null;

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

            }

            return true;
        }

        /*! DeleteAssemblyModel() removes a tool from the database. Parameters are the connection to the 
         * database server and the id-number of the tool. */

        public bool DeleteAssemblyModel(SqlConnection sqlConnection, long AssemblyId)
        {

            // Use the existing SqlConnection
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the description of the tool.
                SqlCommand sqlCommand = new SqlCommand("DELETE [dbo].[model] FROM assembly_model WHERE assembly_id = " + AssemblyId + ";", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = null;

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

            }

            return true;
        }


        //----------------------------------------------------------------------------------

        /*! DeleteHolderModel() removes a HolderModel from the database.  
         * @param  SQLConnection to the database
         * @param The Holder id as long value
         * @return bool flag, true if succesful #/ false if not
         */

        public bool DeleteHolderModel(SqlConnection SQLConnection, long holderId)
        {

            // Use the existing SqlConnection
            using (SQLConnection)
            {
                // Create a new SqlCommand and query the description of the Tool.
                SqlCommand SQLCommand = new SqlCommand("DELETE FROM [dbo].[holder_model] WHERE [dbo].[holder_model].[holder_id] = " + holderId + ";", SQLConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader SQLDataReader = null;

                // Execute the command
                SQLDataReader = SQLCommand.ExecuteReader();

            }

            return true;
        }




        /*! DeleteToolContour() removes a tool from the database. Parameters are the connection to the 
         * database server and the id-number of the tool. */

        public bool DeleteHolderContour(SqlConnection sqlConnection, long HolderId)
        {

            // Use the existing SqlConnection
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the description of the tool.
                SqlCommand sqlCommand = new SqlCommand("DELETE [dbo].[holder_profile] FROM holder_profile WHERE holder_id = " + HolderId + ";", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = null;

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

            }

            return true;
        }

        //----------------------------------------------------------------------------------

        /*! IsAssemblyInToolList() checks if one or more of the setting sheets contains the given assembly.
         * @param SQLiteConnection to the database
         * @param string containing the assembly id
         * @return List<string> which contains ids of the sheets where the assembly is been used
         */


        public List<string> IsAssemblyInToolList(SqlConnection SQLConnection, string AssemblyId)
        {
            List<string> SheetList = new List<string>();

            // Using block for the database queries
            using (SQLConnection)
            {
                // Create a new SqlCommand and query the description of the assembly.
                SqlCommand command = new SqlCommand("SELECT [dbo].[sheet_entries].[sheet_id], [dbo].[sheet_entries].[tool_id] FROM [dbo].[sheet_entries] WHERE [dbo].[sheet_entries].[tool_id] = " + AssemblyId + ";", SQLConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader SQLDataReader = command.ExecuteReader();

                // Read the data until there are no more rows
                if (SQLDataReader.HasRows)
                {
                    while (SQLDataReader.Read())
                    {
                        SheetList.Add(SQLDataReader.GetString(0));
                    }
                }
            }

            return SheetList;
        }

        //----------------------------------------------------------------------------------

        /*! getTurningToolById uses the connection to the SQL-Server and gets the data of the turningtool with the given
         * toolID. The method returns a dictionary which contains all the characteristic data including the tool
           id and the description. */


        public Dictionary<string, string> getTurningToolById(SqlConnection sqlConnection, long toolId)
        {
            // Instantiate a dictionary for the tool description
            Dictionary<string, string> toolDescription = new Dictionary<string, string>();

            // Add the Id to the dictionary
            toolDescription.Add("id", toolId.ToString());

            // Using block for the database queries
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the description of the tool.
                SqlCommand command = new SqlCommand("SELECT [TOOLDB].[dbo].[turning_tool].[id], [TOOLDB].[dbo].[turning_tool].[name], [TOOLDB].[dbo].[turning_tool].[type] FROM [TOOLDB].[dbo].[turning_tool] WHERE [TOOLDB].[dbo].[turning_tool].[id] = " + toolId + ";", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = command.ExecuteReader();

                // Read the data until there are no more rows
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        // Add the description to the toolDescription dictionary
                        toolDescription.Add("description", sqlDataReader.GetString(1));
                        toolDescription.Add("type", sqlDataReader.GetString(2));
                    }
                }

                // Create a second SqlCommand and query the characteristics of the tool.
                command = new SqlCommand("SELECT [TOOLDB].[dbo].[turning_tool].[id], [TOOLDB].[dbo].[turning_tool].[name], [TOOLDB].[dbo].[din4000_key].[name], [TOOLDB].[dbo].[turning_tool_characteristic].[data_value] FROM [TOOLDB].[dbo].[turning_tool] LEFT JOIN [TOOLDB].[dbo].[turning_tool_characteristic] ON turning_tool.id = turning_tool_characteristic.tool_id LEFT JOIN [TOOLDB].[dbo].[din4000_key] ON [TOOLDB].[dbo].[din4000_key].[id] = [TOOLDB].[dbo].[turning_tool_characteristic].[key_id] WHERE [TOOLDB].[dbo].[turning_tool].[id] = " + toolId + ";", sqlConnection);

                // Execute the query
                sqlDataReader = command.ExecuteReader();

                // Read the characteristic data until there are no more rows
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        try
                        {
                            // Store the data into the dictionary
                            toolDescription.Add(sqlDataReader.GetString(2), sqlDataReader.GetString(3));
                        }
                        catch (Exception ex)
                        {
                            debugMessage = ex.Message;

                        }
                    }
                }
            }
            sqlConnection.Close();
            sqlConnection.Dispose();
            return toolDescription;
        }

        //----------------------------------------------------------------------------------

        /*! insertTurningTool() stores a new turningtool into the database. The method uses the given SqlConnection and inserts 
         the the data in the given toolDescription dictionary via a SQL-query into the data tables.*/

        public long insertTurningTool(SqlConnection sqlConnection, Dictionary<string, string> toolDescription, long ToolId = 0)
        {
            // Initialize the variable for the last id
            long lastToolId = 0;
            long lastCharacteristicId = 0;

            //// Load the DIN 4000 Keys into the dictionary
            //loadDinNames();

            // Create a new SqlCommand and query the description of the tool.
            SqlCommand sqlCommand = new SqlCommand("SELECT MAX(Id) FROM turning_tool;", sqlConnection);

            // Create a new SqlDataReader and execute the query
            SqlDataReader sqlDataReader = null;


            // Use the existing SqlConnection
            using (sqlConnection)
            {
                if (ToolId == 0)
                {

                    // Execute the command
                    sqlDataReader = sqlCommand.ExecuteReader();

                    // Read the given id
                    if (sqlDataReader.HasRows)
                    {
                        while (sqlDataReader.Read())
                        {
                            try
                            {
                                lastToolId = sqlDataReader.GetInt64(0);
                            }
                            catch (Exception ex)
                            {
                                debugMessage = ex.Message;
                                lastToolId = 0;
                            }
                        }
                    }
                }
                else
                    lastToolId = ToolId - 1;

                // Create a new SqlCommand and query the description of the tool.
                sqlCommand = new SqlCommand("INSERT INTO turning_tool (id, name, type) VALUES (@id, @name, @type)", sqlConnection);

                // Add the parameter to the insert command.
                sqlCommand.Parameters.AddWithValue("@id", lastToolId + 1);
                sqlCommand.Parameters.AddWithValue("@name", toolDescription["name"]);
                sqlCommand.Parameters.AddWithValue("@type", toolDescription["type"]);

                // Create a new SqlDataReader and execute the query
                sqlDataReader = sqlCommand.ExecuteReader();

                // Create a new SqlCommand and query the description of the tool.
                sqlCommand = new SqlCommand("SELECT MAX(Id) FROM turning_tool_characteristic;", sqlConnection);

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

                // Read the given id
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        try
                        {
                            lastCharacteristicId = sqlDataReader.GetInt64(0);
                        }
                        catch (Exception ex)
                        {
                            debugMessage = ex.Message;
                            lastCharacteristicId = 0;
                        }
                    }
                }

                for (int i = 1; i < toolDescription.Count; i++)
                {
                    sqlCommand = new SqlCommand("INSERT INTO turning_tool_characteristic (id, tool_id, key_id, data_value) VALUES (@id, @tool_id, @key_id, @data_value)", sqlConnection);

                    sqlCommand.Parameters.AddWithValue("@id", lastCharacteristicId + i);
                    sqlCommand.Parameters.AddWithValue("@tool_id", lastToolId + 1);
                    string keyId = null;
                    foreach (string key in din4000names.Keys)
                    {
                        //Console.WriteLine(importCharValues.ElementAt(i).Key + "-" + key + Environment.NewLine);
                        if (key.Equals(toolDescription.ElementAt(i).Key))
                        {
                            //Console.WriteLine(toolDescription.ElementAt(i).Key + "-" + key + Environment.NewLine);
                            keyId = key;
                        }
                    }

                    if (i > 0 && !string.IsNullOrEmpty(keyId))
                    {
                        sqlCommand.Parameters.AddWithValue("@key_id", keyId);
                        sqlCommand.Parameters.AddWithValue("@data_value", toolDescription.ElementAt(i).Value);

                        // don't forget to take care of connection - I omit this part for clearness
                        // Execute the command
                        sqlDataReader = sqlCommand.ExecuteReader();
                    }
                }
            }

            return lastToolId + 1;
        }

        //----------------------------------------------------------------------------------

        /*! deleteTurningTool() removes a turningtool from the database. Parameters are the connection to the 
         * database server and the id-number of the tool. */

        public bool deleteTurningTool(SqlConnection sqlConnection, long toolId)
        {
            SqlConnection m_sqlConnection = getSqlConnection();

            // Use the existing SqlConnection
            using (sqlConnection)
            {
                // Create a new SqlCommand and query the description of the tool.
                SqlCommand sqlCommand = new SqlCommand("DELETE [dbo].[turning_tool_characteristic] FROM turning_tool_characteristic WHERE tool_id = " + toolId + ";", sqlConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader sqlDataReader = null;

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

                // Create a new SqlCommand and query the description of the tool.
                sqlCommand = new SqlCommand("DELETE [dbo].[turning_tool_technology] FROM turning_tool_technology WHERE tool_id = " + toolId + ";", sqlConnection);

                // Create a new SqlDataReader and execute the query
                sqlDataReader = null;

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

                // Create a new SqlCommand and query the description of the tool.
                sqlCommand = new SqlCommand("DELETE [dbo].[turning_tool_profile] FROM turning_tool_profile WHERE tool_id = " + toolId + ";", sqlConnection);

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

                // Create a new SqlCommand and query the description of the tool.
                sqlCommand = new SqlCommand("DELETE [dbo].[turning_tool] FROM turning_tool WHERE id = " + toolId + ";", sqlConnection);

                // Execute the command
                sqlDataReader = sqlCommand.ExecuteReader();

            }

            return true;
        }

        //----------------------------------------------------------------------------------

        /*! DeleteTurningToolModel() removes a TurningToolModel from the database.  
         * @param SQLiteConnection to the database
         * @param The TurningTool id as long value
         * @return bool flag, true if succesful #/ false if not
         */

        public bool DeleteTurningToolModel(SqlConnection SQLConnection, long toolId)
        {

            // Use the existing SqlConnection
            using (SQLConnection)
            {
                // Create a new SqlCommand and query the description of the Tool.
                SqlCommand SQLCommand = new SqlCommand("DELETE FROM  [dbo].[turning_tool_model] WHERE tool_id = " + toolId + ";", SQLConnection);

                // Create a new SqlDataReader and execute the query
                SqlDataReader SQLDataReader = null;

                // Execute the command
                SQLDataReader = SQLCommand.ExecuteReader();

            }

            return true;
        }

        //----------------------------------------------------------------------------------

        /*! RunSQLCommands() runs a set of SQL commands given in an text array.  
         * @param SQLiteConnection to the database
         * @param String array containing SQL Statemants
         * @return bool flag, true if succesful #/ false if not
         */


        public bool RunSQLCommands(SqlConnection sqlConnection, string[] CommandList)
        {

            // Use the existing SqlConnection
            try
            {
                foreach (string Command in CommandList)
                {

                    // Create a new SqlCommand and query the description of the Tool.
                    SqlCommand SQLCommand = new SqlCommand(Command, sqlConnection);

                    // Execute the command
                    SQLCommand.ExecuteNonQuery();
                }

                sqlConnection.Close();
                sqlConnection.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return true;
        }

        //----------------------------------------------------------------------------------

        /*! StoreMaterial() Stores a new set of materials in the material table.  
         * @param MaterialList a list of the materials to store in the database   
         * @return bool flag, true if succesful #/ false if not
         */

        public bool StoreMaterial(List<string> MaterialList)
        {

            SqlConnection mySqlConnection = getSqlConnection();

            SqlCommand DeleteSqlCommand = new SqlCommand("DELETE FROM [dbo].[material]", mySqlConnection);
            DeleteSqlCommand.ExecuteReader();

            for (int i = 0; i < MaterialList.Count; i++)
            {
                SqlCommand InsertSqlCommand = new SqlCommand("INSERT INTO [dbo].[material] (id, material) VALUES (" + (i + 1).ToString() + ", '" + MaterialList.ElementAt(i) + "');", mySqlConnection);
                InsertSqlCommand.ExecuteNonQuery();
            }

            // Close the connection when done with it.
            mySqlConnection.Close();

            return true;
        }


        //----------------------------------------------------------------------------------

        /*! StoreProcess() Stores a new set of processes in the process table.  
         * @param ProcessList a list of the processes to store in the database   
         * @return bool flag, true if succesful #/ false if not
         */

        public bool StoreProcess(List<string> ProcessList)
        {
            SqlConnection mySqlConnection = getSqlConnection();

            SqlCommand DeleteSQLiteCommand = new SqlCommand("DELETE FROM [dbo].[process]", mySqlConnection);
            DeleteSQLiteCommand.ExecuteReader();

            for (int i = 0; i < ProcessList.Count; i++)
            {
                SqlCommand InsertSqlCommand = new SqlCommand("INSERT INTO [dbo].[process] (id, process) VALUES (" + (i + 1).ToString() + ", " + ProcessList.ElementAt(i) + ");", mySqlConnection);
                InsertSqlCommand.ExecuteNonQuery();

            }


            // Close the connection when done with it.
            mySqlConnection.Close();

            return true;
        }


        //----------------------------------------------------------------------------------

        /*! StoreMillProcess() Stores a new set of millprocesses in the millprocess table.  
         * @param MillProcessList a list of the millprocesses to store in the database   
         * @return bool flag, true if succesful #/ false if not
         */

        public bool StoreMillProcess(List<string> MillProcessList)
        {
            SqlConnection mySqlConnection = getSqlConnection();

            SqlCommand DeleteSqlCommand = new SqlCommand("DELETE FROM [dbo].[mill_process]", mySqlConnection);
            DeleteSqlCommand.ExecuteReader();

            for (int i = 0; i < MillProcessList.Count; i++)
            {
                SqlCommand InsertSqlCommand = new SqlCommand("INSERT INTO [dbo].[mill_process] (id, name) VALUES (" + (i + 1).ToString() + ", " + MillProcessList.ElementAt(i) + ");", mySqlConnection);
                InsertSqlCommand.ExecuteNonQuery();
            }


            // Close the connection when done with it.
            mySqlConnection.Close();

            return true;
        }



        //----------------------------------------------------------------------------------

        /*! GetDbVersion() Reads the version string from the database.
         * @param SQLiteConnection to the database
         * @return int variable containig the version code
         */

        public int GetDbVersion(SqlConnection sqlConnection)
        {
            Int32 VersionCode = 0;

            try
            {
                // Create a new SqLiteCommand and query the version.
                SqlCommand command = new SqlCommand("SELECT  [TOOLDB].[dbo].[version].[db_version] FROM [TOOLDB].[dbo].[version];", sqlConnection);

                // Create a new SQLiteDataReader and execute the query
                SqlDataReader SQLDataReader = command.ExecuteReader();

                // Read the data until there are no more rows
                if (SQLDataReader.HasRows)
                {
                    while (SQLDataReader.Read())
                    {
                        // Add the description to the toolDescription dictionary
                        VersionCode = SQLDataReader.GetInt32(0);
                    }
                }
                else
                {
                    return 0;
                }
            }
            catch (Exception ex)
            {
                return 99;
            }

            return VersionCode;

        }

        public bool TestDateUS(SqlConnection sqlConnection)
        {
            bool isValid = false;
            string KeyWord = "";
            // Use the existing SqlConnection
            using (sqlConnection)
            {
                SqlCommand sqlCommand = new SqlCommand("IF ISDATE('2019-29-06 10:19:41.177') = 1 " +
                    "SELECT 'VALID' " +
                    "ELSE " +
                    "SELECT 'INVALID';", sqlConnection);

                //'2009-05-12 10:19:41.177'

                SqlDataReader myDataReader = sqlCommand.ExecuteReader();


                if (myDataReader.HasRows)
                {
                    while (myDataReader.Read())
                    {
                        KeyWord = myDataReader.GetString(0);
                        Console.WriteLine(KeyWord);
                        if (KeyWord.Equals("VALID"))
                        {
                            isValid = true;
                        }

                    }
                }
            }

            sqlConnection.Close();
            sqlConnection.Dispose();
            return isValid;
        }

        public bool TestDateDE(SqlConnection sqlConnection)
        {
            bool isValid = false;
            string KeyWord = "";
            // Use the existing SqlConnection
            using (sqlConnection)
            {
                SqlCommand sqlCommand = new SqlCommand("IF ISDATE('2019-06-29 10:19:41.177') = 1 " +
                    "SELECT 'VALID' " +
                    "ELSE " +
                    "SELECT 'INVALID';", sqlConnection);

                //'2009-05-12 10:19:41.177'

                SqlDataReader myDataReader = sqlCommand.ExecuteReader();


                if (myDataReader.HasRows)
                {
                    while (myDataReader.Read())
                    {
                        KeyWord = myDataReader.GetString(0);
                        Console.WriteLine(KeyWord);
                        if (KeyWord.Equals("VALID"))
                        {
                            isValid = true;
                        }

                    }
                }
            }

            sqlConnection.Close();
            sqlConnection.Dispose();
            return isValid;
        }

        /*! UpdateAssemblyToolID()        
          * @param SQLConnection to the database
          * @param id-number of the Tool  as a long variable.  
          * @return the id of the Assembly as string
          */
        public void UpdateAssemblyToolID(SqlConnection SQLConnection, long toolID, string AssemblyNr)
        {
            // Create a new SqlDataReader and execute the query
            SqlDataReader SQLDataReader = null;

            //SQLConnection = getSqlConnection();

            // Use the existing SqlConnection
            using (SQLConnection)
            {
                //string Diameter = Convert.ToString(toolID, CultureInfo.InvariantCulture);

                SqlCommand SQLCommand = new SqlCommand("UPDATE assembly SET tool_id = '" + toolID + "' WHERE id = '" + AssemblyNr + "';", SQLConnection);


                SQLDataReader = SQLCommand.ExecuteReader();
            }
        }

        //----------------------------------------------------------------------------------


        /*! UpdateAssemblyHolderID()        
         * @param SQLConnection to the database
         * @param id-number of the Tool  as a long variable.  
         * @return the id of the Assembly as string
         */
        public void UpdateAssemblyHolderID(SqlConnection SQLConnection, long holderID, string AssemblyNr)
        {
            // Create a new SqlDataReader and execute the query
            SqlDataReader SQLDataReader = null;

            //SQLConnection = getSqlConnection();

            // Use the existing SqlConnection
            using (SQLConnection)
            {
                //string Diameter = Convert.ToString(toolID, CultureInfo.InvariantCulture);

                SqlCommand SQLCommand = new SqlCommand("UPDATE assembly SET holder_id = '" + holderID + "' WHERE id = '" + AssemblyNr + "';", SQLConnection);


                SQLDataReader = SQLCommand.ExecuteReader();
            }
        }

        //----------------------------------------------------------------------------------



        /*! checkIfPartOfAssembly() checks if selected tool is part of an assembly and returns a dict with the assembly id and the tool_id        
        * @param SQLiteConnection to the database
        * @param id-number of the Tool as a long variable.  
        * @return AssemblyList as a Dictionary<string, string>
        */

        public Dictionary<string, string> checkIfToolIsPartOfAssembly(SqlConnection SqlConnection, long toolId)
        {
            Dictionary<string, string> AssemblyList = new Dictionary<string, string>();
            AssemblyList.Clear();
            long counter = 0;

            SqlConnection m_SqlConnection = getSqlConnection();

            long dataCount = GetLastId(m_SqlConnection, "assembly");

            for (int i = 0; i < dataCount; i++)
            {
                if (getAssemblyById(m_SqlConnection, i + 1).toolId == toolId)
                {
                    //MessageBox.Show("Das ausgewählte Werkzeug ist Teil einer Baugruppe. Bitte löschen sie zuerst die Baugruppe und dann das Werkzeug.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    counter++;
                    AssemblyList.Add(getAssemblyById(SqlConnection, i + 1).id.ToString(), toolId.ToString());
                }
            }
            return AssemblyList;
        }


        //----------------------------------------------------------------------------------


        /*! checkIfHolderIsPartOfAssembly() checks if selected tool is part of an assembly and returns a dict with the assembly id and the tool_id        
        * @param SQLiteConnection to the database
        * @param id-number of the Tool as a long variable.  
        * @return AssemblyList as a Dictionary<string, string>
        */

        public Dictionary<string, string> checkIfHolderIsPartOfAssembly(SqlConnection SqlConnection, long holderID)
        {
            Dictionary<string, string> AssemblyList = new Dictionary<string, string>();
            AssemblyList.Clear();
            long counter = 0;

            SqlConnection m_SqlConnection = getSqlConnection();

            long dataCount = GetLastId(m_SqlConnection, "assembly");

            for (int i = 0; i < dataCount; i++)
            {
                if (getAssemblyById(m_SqlConnection, i + 1).holderId == holderID)
                {
                    //MessageBox.Show("Das ausgewählte Werkzeug ist Teil einer Baugruppe. Bitte löschen sie zuerst die Baugruppe und dann das Werkzeug.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    counter++;
                    AssemblyList.Add(getAssemblyById(SqlConnection, i + 1).id.ToString(), holderID.ToString());
                }
            }
            return AssemblyList;
        }


        //----------------------------------------------------------------------------------



        ///*! insertHolder() stores a new holder into the database. The method uses the given SqlConnection and inserts 
        // the the data in the given holderDescription dictionary via a SQL-query into the data tables.*/

        //public long insertHolder(SqlConnection sqlConnection, Dictionary<string, string> holderDescription)
        //{
        //    // Initialize the variable for the last id
        //    long lastHolderId = 0;
        //    long lastCharacteristicId = 0;

        //    // Load the DIN 4000 Keys into the dictionary
        //    //loadDinKeys();

        //    // Use the existing SqlConnection
        //    using (sqlConnection)
        //    {
        //        // Create a new SqlCommand and query the description of the holder.
        //        SqlCommand sqlCommand = new SqlCommand("SELECT MAX(Id) FROM holder;", sqlConnection);

        //        // Create a new SqlDataReader and execute the query
        //        SqlDataReader sqlDataReader = null;

        //        // Execute the command
        //        sqlDataReader = sqlCommand.ExecuteReader();

        //        // Read the given id
        //        if (sqlDataReader.HasRows)
        //        {
        //            while (sqlDataReader.Read())
        //            {
        //                try
        //                {
        //                    lastHolderId = sqlDataReader.GetInt64(0);
        //                }
        //                catch (Exception ex)
        //                {
        //                    debugMessage = ex.Message;
        //                    lastHolderId = 0;
        //                }
        //            }
        //        }

        //        // Create a new SqlCommand and query the description of the holder.
        //        sqlCommand = new SqlCommand("INSERT INTO holder (id, name) VALUES (@id, @name)", sqlConnection);

        //        // Add the parameter to the insert command.
        //        sqlCommand.Parameters.AddWithValue("@id", lastHolderId + 1);
        //        if(holderDescription.ContainsKey("J21"))
        //            sqlCommand.Parameters.AddWithValue("@name", holderDescription["J21"]);
        //        //if (holderDescription.ContainsKey("description"))
        //        //    sqlCommand.Parameters.AddWithValue("@name", holderDescription["description"]);

        //        // Create a new SqlDataReader and execute the query
        //        sqlDataReader = sqlCommand.ExecuteReader();

        //        // Create a new SqlCommand and query the description of the holder.
        //        sqlCommand = new SqlCommand("SELECT MAX(Id) FROM holder_characteristic;", sqlConnection);

        //        // Execute the command
        //        sqlDataReader = sqlCommand.ExecuteReader();

        //        // Read the given id
        //        if (sqlDataReader.HasRows)
        //        {
        //            while (sqlDataReader.Read())
        //            {
        //                try
        //                {
        //                    lastCharacteristicId = sqlDataReader.GetInt64(0);
        //                }
        //                catch (Exception ex)
        //                {
        //                    debugMessage = ex.Message; 
        //                    lastCharacteristicId = 0;
        //                }
        //            }
        //        }

        //        for (int i = 1; i < holderDescription.Count; i++)
        //        {
        //            sqlCommand = new SqlCommand("INSERT INTO holder_characteristic (id, holder_id, key_id, data_value) VALUES (@id, @holder_id, @key_id, @data_value)", sqlConnection);
        //            sqlCommand.Parameters.AddWithValue("@id", lastCharacteristicId + i);
        //            sqlCommand.Parameters.AddWithValue("@holder_id", lastHolderId + 1);
        //            string keyId = null;
        //            foreach (string key in din4000keys.Keys)
        //            {
        //                //Console.WriteLine(importCharValues.ElementAt(i).Key + "-" + key + Environment.NewLine);
        //                if (key.Equals(holderDescription.ElementAt(i).Key))
        //                {
        //                    //Console.WriteLine(holderDescription.ElementAt(i).Key + "-" + key + Environment.NewLine);
        //                    keyId = din4000keys[key];
        //                }
        //            }
        //            if (!String.IsNullOrEmpty(keyId))
        //            {
        //                sqlCommand.Parameters.AddWithValue("@key_id", keyId);
        //                sqlCommand.Parameters.AddWithValue("@data_value", holderDescription.ElementAt(i).Value);

        //                // don't forget to take care of connection - I omit this part for clearness
        //                // Execute the command
        //                sqlDataReader = sqlCommand.ExecuteReader();
        //            }
        //        }
        //    }

        //    return lastHolderId + 1;
        //}

        ////----------------------------------------------------------------------------------


        /*! insertHolder() stores a new holder into the database. The method uses the given SqlConnection and inserts 
         the the data in the given holderDescription dictionary via a SQL-query into the data tables.*/

        public long insertHolder(SqlConnection sqlConnection, Dictionary<string, string> holderDescription, long holderID = 0)
        {
            // Initialize the variable for the last id
            long lastHolderId = 0;
            long lastCharacteristicId = 0;

            // Load the DIN 4000 Keys into the dictionary
            //loadDinKeys();

            if (holderID == 0)
            {

                // Use the existing SqlConnection
                using (sqlConnection)
                {
                    // Create a new SqlCommand and query the description of the holder.
                    SqlCommand sqlCommand = new SqlCommand("SELECT " + holderID + " FROM holder;", sqlConnection);

                    // Create a new SqlDataReader and execute the query
                    SqlDataReader sqlDataReader = null;

                    // Execute the command
                    sqlDataReader = sqlCommand.ExecuteReader();

                    // Read the given id
                    if (sqlDataReader.HasRows)
                    {
                        while (sqlDataReader.Read())
                        {
                            try
                            {
                                lastHolderId = sqlDataReader.GetInt64(0);
                            }
                            catch (Exception ex)
                            {
                                debugMessage = ex.Message;
                                lastHolderId = 0;
                            }
                        }
                    }

                    // Create a new SqlCommand and query the description of the holder.
                    sqlCommand = new SqlCommand("INSERT INTO holder (id, name) VALUES (@id, @name)", sqlConnection);

                    // Add the parameter to the insert command.
                    sqlCommand.Parameters.AddWithValue("@id", holderID);
                    if (holderDescription.ContainsKey("J21"))
                        sqlCommand.Parameters.AddWithValue("@name", holderDescription["J21"]);
                    //if (holderDescription.ContainsKey("description"))
                    //    sqlCommand.Parameters.AddWithValue("@name", holderDescription["description"]);

                    // Create a new SqlDataReader and execute the query
                    sqlDataReader = sqlCommand.ExecuteReader();

                    // Create a new SqlCommand and query the description of the holder.
                    sqlCommand = new SqlCommand("SELECT MAX(Id) FROM holder_characteristic;", sqlConnection);

                    // Execute the command
                    sqlDataReader = sqlCommand.ExecuteReader();

                    // Read the given id
                    if (sqlDataReader.HasRows)
                    {
                        while (sqlDataReader.Read())
                        {
                            try
                            {
                                lastCharacteristicId = sqlDataReader.GetInt64(0);
                            }
                            catch (Exception ex)
                            {
                                debugMessage = ex.Message;
                                lastCharacteristicId = 0;
                            }
                        }
                    }

                    for (int i = 1; i < holderDescription.Count; i++)
                    {
                        sqlCommand = new SqlCommand("INSERT INTO holder_characteristic (id, holder_id, key_id, data_value) VALUES (@id, @holder_id, @key_id, @data_value)", sqlConnection);
                        sqlCommand.Parameters.AddWithValue("@id", lastCharacteristicId + i);
                        sqlCommand.Parameters.AddWithValue("@holder_id", holderID);
                        string keyId = null;
                        foreach (string key in din4000keys.Keys)
                        {
                            //Console.WriteLine(importCharValues.ElementAt(i).Key + "-" + key + Environment.NewLine);
                            if (key.Equals(holderDescription.ElementAt(i).Key))
                            {
                                //Console.WriteLine(holderDescription.ElementAt(i).Key + "-" + key + Environment.NewLine);
                                keyId = din4000keys[key];
                            }
                        }
                        if (!String.IsNullOrEmpty(keyId))
                        {
                            sqlCommand.Parameters.AddWithValue("@key_id", keyId);
                            sqlCommand.Parameters.AddWithValue("@data_value", holderDescription.ElementAt(i).Value);

                            // don't forget to take care of connection - I omit this part for clearness
                            // Execute the command
                            sqlDataReader = sqlCommand.ExecuteReader();
                        }
                    }
                }
            }
            else
            {
                // Use the existing SqlConnection
                using (sqlConnection)
                {
                    // Create a new SqlCommand and query the description of the holder.
                    SqlCommand sqlCommand = new SqlCommand("SELECT MAX(Id) FROM holder;", sqlConnection);

                    // Create a new SqlDataReader and execute the query
                    SqlDataReader sqlDataReader = null;

                    // Execute the command
                    sqlDataReader = sqlCommand.ExecuteReader();

                    // Read the given id
                    if (sqlDataReader.HasRows)
                    {
                        while (sqlDataReader.Read())
                        {
                            try
                            {
                                lastHolderId = sqlDataReader.GetInt64(0);
                            }
                            catch (Exception ex)
                            {
                                debugMessage = ex.Message;
                                lastHolderId = 0;
                            }
                        }
                    }

                    // Create a new SqlCommand and query the description of the holder.
                    sqlCommand = new SqlCommand("INSERT INTO holder (id, name) VALUES (@id, @name)", sqlConnection);

                    // Add the parameter to the insert command.
                    sqlCommand.Parameters.AddWithValue("@id", lastHolderId + 1);
                    if (holderDescription.ContainsKey("J21"))
                        sqlCommand.Parameters.AddWithValue("@name", holderDescription["J21"]);
                    //if (holderDescription.ContainsKey("description"))
                    //    sqlCommand.Parameters.AddWithValue("@name", holderDescription["description"]);

                    // Create a new SqlDataReader and execute the query
                    sqlDataReader = sqlCommand.ExecuteReader();

                    // Create a new SqlCommand and query the description of the holder.
                    sqlCommand = new SqlCommand("SELECT MAX(Id) FROM holder_characteristic;", sqlConnection);

                    // Execute the command
                    sqlDataReader = sqlCommand.ExecuteReader();

                    // Read the given id
                    if (sqlDataReader.HasRows)
                    {
                        while (sqlDataReader.Read())
                        {
                            try
                            {
                                lastCharacteristicId = sqlDataReader.GetInt64(0);
                            }
                            catch (Exception ex)
                            {
                                debugMessage = ex.Message;
                                lastCharacteristicId = 0;
                            }
                        }
                    }

                    for (int i = 1; i < holderDescription.Count; i++)
                    {
                        sqlCommand = new SqlCommand("INSERT INTO holder_characteristic (id, holder_id, key_id, data_value) VALUES (@id, @holder_id, @key_id, @data_value)", sqlConnection);
                        sqlCommand.Parameters.AddWithValue("@id", lastCharacteristicId + i);
                        sqlCommand.Parameters.AddWithValue("@holder_id", lastHolderId + 1);
                        string keyId = null;
                        foreach (string key in din4000keys.Keys)
                        {
                            //Console.WriteLine(importCharValues.ElementAt(i).Key + "-" + key + Environment.NewLine);
                            if (key.Equals(holderDescription.ElementAt(i).Key))
                            {
                                //Console.WriteLine(holderDescription.ElementAt(i).Key + "-" + key + Environment.NewLine);
                                keyId = din4000keys[key];
                            }
                        }
                        if (!String.IsNullOrEmpty(keyId))
                        {
                            sqlCommand.Parameters.AddWithValue("@key_id", keyId);
                            sqlCommand.Parameters.AddWithValue("@data_value", holderDescription.ElementAt(i).Value);

                            // don't forget to take care of connection - I omit this part for clearness
                            // Execute the command
                            sqlDataReader = sqlCommand.ExecuteReader();
                        }
                    }
                }
            }
            return lastHolderId + 1;
        }

        //----------------------------------------------------------------------------------


    }
}
