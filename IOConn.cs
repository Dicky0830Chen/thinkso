using System;
using System.Data;
using System.Collections;
using System.IO;
using System.Data.SqlClient;
using System.Data.Common;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Security.Cryptography;
using System.Data.Odbc;

namespace DocumentFile
{
    public class IOConn
    {
        public DataSet Ds;
        public ArrayList ArrStr;
        public int ActCnt;
        public int TotCnt;
        public int RtnCnt;
        public int FstCnt = 0;
        public int SelCnt = 0;
        public string cnnStr = System.Configuration.ConfigurationManager.AppSettings.Get("ConnectionStr");

        private byte[] m_IV = new byte[8];
        private byte[] m_Key = new byte[8];


        public DataSet SelectSet(string Sql)
        {
            OdbcConnection Conn = getConnection();
            OdbcCommand cmd;
            OdbcDataAdapter Dap;

            try
            {
                cmd = new OdbcCommand(Sql, Conn);
                cmd.CommandTimeout = 6000;
                Dap = new OdbcDataAdapter();
                Ds = new DataSet();
                Dap.SelectCommand = cmd;
                Dap.Fill(Ds, FstCnt, FstCnt + SelCnt, "Table0");

                cmd.Cancel();
            }
            catch (Exception)
            {
                Ds = new DataSet();
            }

            Conn.Close();
            Dap = null;
            cmd = null;
            Conn = null;

            return Ds;
        }

        public ArrayList SelectAry(string Sql)
        {
            OdbcConnection Conn = getConnection();
            OdbcCommand cmd;
            OdbcDataReader DaRd;
            int i = 0;
            int col = 0;
            string str;

            ArrStr = new ArrayList();
            TotCnt = 0;
            ActCnt = 0;
            RtnCnt = 0;
            try
            {
                cmd = new OdbcCommand(Sql, Conn);
                cmd.CommandTimeout = 6000;
                DaRd = cmd.ExecuteReader();
                col = DaRd.FieldCount;
                string cnm = "";
                for (i = 0; i < col; i++) cnm += DaRd.GetName(i) + ",";
                while (DaRd.Read())
                {
                    TotCnt++;
                    if (TotCnt >= FstCnt)
                    {
                        ActCnt++;
                        if ((SelCnt != 0) && (ActCnt > SelCnt)) break;
                        str = "";
                        try
                        {
                            for (i = 0; i < col; i++)
                            {
                                if (DaRd.GetValue(i) == null)
                                    str += "^";
                                else
                                    str += DaRd.GetValue(i).ToString().Replace("^", "") + "^";
                            }
                            ArrStr.Add(str);
                        }
                        catch (Exception) { }
                    }
                }
                DaRd.Close();
                cmd.Cancel();
            }
            catch (Exception)
            {
                ArrStr = new ArrayList();
            }

            Conn.Close();
            DaRd = null;
            cmd = null;
            Conn = null;

            return ArrStr;
        }

        /*************************************************************************/

        public List<Dictionary<string, string>> SelectDct(string Sql)
        {
            OdbcConnection Conn =  getConnection();
            OdbcCommand cmd;
            OdbcDataReader DaRd;
            int i = 0;
            int col = 0;
            string str;
            List<Dictionary<string, string>> ListDct = new List<Dictionary<string, string>>();

            try
            {
                cmd = new OdbcCommand(Sql, Conn);
                cmd.CommandTimeout = 6000;
                DaRd = cmd.ExecuteReader();
                col = DaRd.FieldCount;
                while (DaRd.Read())
                {
                    Dictionary<string, string> DctStr = new Dictionary<string, string>();
                    for (i = 0; i < col; i++)
                    {
                        str = DaRd.IsDBNull(i) ? "" : DaRd.GetValue(i).ToString();
                        DctStr.Add(DaRd.GetName(i), str);
                    }
                    ListDct.Add(DctStr);
                }
            }
            catch (Exception)
            {
                ListDct = new List<Dictionary<string, string>>();
            }
            Conn.Close();
            return ListDct;
        }

        public int ExecuteSql(string Sql)
        {
            OdbcConnection Conn =  getConnection();
            OdbcCommand cmd = null;

            try
            {
                cmd = new OdbcCommand(Sql, Conn);
                cmd.CommandTimeout = 6000;
                ActCnt = cmd.ExecuteNonQuery();
                cmd.Cancel();
            }
            catch (Exception)
            {
                ActCnt = -9;
            }
            Conn.Close();
            cmd = null;
            Conn = null;

            return ActCnt;
        }

        public OdbcConnection getConnection()
        {
            OdbcConnection Conn = null;

            try
            {
                Conn = new OdbcConnection(cnnStr);
                Conn.Open();
            }
            catch (Exception)
            {
                Conn = null;
            }

            return Conn;
        }



    }
}