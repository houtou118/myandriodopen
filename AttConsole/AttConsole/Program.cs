using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Collections;
using System.Data.SqlClient;
using System.Globalization;

namespace AttConsole
{
    class Program
    {
        static DbHelperSQL DBSQL = new DbHelperSQL();

        //Create Standalone SDK class dynamicly.
        static zkemkeeper.CZKEMClass axCZKEM1 = new zkemkeeper.CZKEMClass();

        static bool bIsConnected = false;//the boolean value identifies whether the device is connected
        static int iMachineNumber = 1;//the serial number of the device.After connecting the device ,this value will be changed.


        static void Main(string[] args)
        {
            try
            {
                Hashtable ht = new Hashtable();

                bIsConnected = axCZKEM1.Connect_Net("125.34.40.33", Convert.ToInt32("4370"));
                if (bIsConnected == true)
                {
                    Console.WriteLine("连接状态:{0}", "连接");
                }
                else
                {
                    Console.WriteLine("连接状态:{0}", "未连接，请联系管理员");
                    DBSQL.ExecuteSql("insert into OAAttConcent values (1,'远程指纹机未能连接','" + DateTime.Now + "')");
                    return;
                }
                DataTable dtRN = DBSQL.Query("select Id,RealName from OAUsers where DeptId = 21");

                foreach (DataRow dr in dtRN.Rows)
                {
                    if (ht.Contains(dr["id"]) == false)
                    {
                        ht.Add(dr["id"], dr["RealName"]);
                    }
                }

                int idwEnrollNumber = 0;
                string sName = "";
                string sPassword = "";
                int iPrivilege = 0;
                bool bEnabled = false;

                axCZKEM1.EnableDevice(iMachineNumber, false);
                axCZKEM1.ReadAllUserID(iMachineNumber);//read all the user information to the memory

                DataTable dt = new DataTable();
                DataColumn idColumn = new DataColumn("id", Type.GetType("System.Int32"));
                DataColumn nameColumn = new DataColumn("name");
                DataColumn OAidColumn = new DataColumn("oaId", Type.GetType("System.Int32"));
                dt.Columns.Add(idColumn);
                dt.Columns.Add(nameColumn);
                dt.Columns.Add(OAidColumn);

                while (axCZKEM1.GetAllUserInfo(iMachineNumber, ref idwEnrollNumber, ref sName, ref sPassword, ref iPrivilege, ref bEnabled))//get all the users' information from the memory
                {
                    foreach (DictionaryEntry de in ht)
                    {
                        if (sName.Contains(de.Value.ToString()) == true)
                        {
                            DataRow newRow = dt.NewRow();
                            newRow["id"] = idwEnrollNumber;
                            newRow["name"] = de.Value;
                            newRow["oaId"] = de.Key;
                            dt.Rows.Add(newRow);
                        }
                    }
                }

                DataTable attDt = new DataTable();
                DataColumn id1Column = new DataColumn("id", Type.GetType("System.Int32"));
                DataColumn id2Column = new DataColumn("Userid", Type.GetType("System.Int32"));
                DataColumn dateColumn = new DataColumn("oadate");
                DataColumn oaweekColumn = new DataColumn("oaweek");
                DataColumn deptIdColumn = new DataColumn("depId", Type.GetType("System.Int32"));
                DataColumn sbqdColumn = new DataColumn("sbqd");
                DataColumn kqqkColumn = new DataColumn("kqqk");
                DataColumn xbqdColumn = new DataColumn("xbqd");
                DataColumn xbqkColumn = new DataColumn("xbqk");
                DataColumn bzColumn = new DataColumn("bz");

                attDt.Columns.Add(id1Column);
                attDt.Columns.Add(dateColumn);
                attDt.Columns.Add(oaweekColumn);
                attDt.Columns.Add(id2Column);
                attDt.Columns.Add(deptIdColumn);
                attDt.Columns.Add(sbqdColumn);
                attDt.Columns.Add(kqqkColumn);
                attDt.Columns.Add(xbqdColumn);
                attDt.Columns.Add(xbqkColumn);
                attDt.Columns.Add(bzColumn);

                int idwVerifyMode = 0;
                int idwInOutMode = 0;
                string sTime = "";

                if (axCZKEM1.ReadGeneralLogData(iMachineNumber))//read all the attendance records to the memory
                {

                    while (axCZKEM1.GetGeneralLogDataStr(iMachineNumber, ref idwEnrollNumber, ref idwVerifyMode, ref idwInOutMode, ref sTime))//get the records from memory
                    {
                        if (Convert.ToDateTime(sTime).Date == DateTime.Today)
                        {
                            if (attDt.Rows.Count == 0)
                            {
                                DataRow newRow = attDt.NewRow();
                                newRow["Userid"] = idwEnrollNumber;
                                newRow["oadate"] = Convert.ToDateTime(sTime).Date;
                                newRow["oaweek"] = CultureInfo.CurrentCulture.DateTimeFormat.GetDayName(Convert.ToDateTime(sTime).DayOfWeek);
                                newRow["depId"] = 21;
                                newRow["sbqd"] = sTime;
                                newRow["kqqk"] = "正常";
                                newRow["xbqd"] = sTime;
                                newRow["xbqk"] = "正常";
                                attDt.Rows.Add(newRow);
                            }
                            else
                            {
                                DataRow[] drh = attDt.Select("Userid='" + idwEnrollNumber + "'");
                                if (drh.Length > 0)
                                {
                                    if (Convert.ToDateTime(sTime) < Convert.ToDateTime(drh[0][5].ToString()))
                                    {
                                        drh[0][5] = sTime;
                                    }
                                    if (Convert.ToDateTime(sTime) > Convert.ToDateTime(drh[0][7].ToString()))
                                    {
                                        drh[0][7] = sTime;
                                    }
                                }
                                else
                                {
                                    DataRow newRow = attDt.NewRow();
                                    newRow["Userid"] = idwEnrollNumber;
                                    newRow["oadate"] = Convert.ToDateTime(sTime).Date;
                                    newRow["oaweek"] = CultureInfo.CurrentCulture.DateTimeFormat.GetDayName(Convert.ToDateTime(sTime).DayOfWeek);
                                    newRow["depId"] = 21;
                                    newRow["sbqd"] = sTime;
                                    newRow["kqqk"] = "正常";
                                    newRow["xbqd"] = sTime;
                                    newRow["xbqk"] = "正常";
                                    attDt.Rows.Add(newRow);
                                }
                            }
                        }
                    }
                }

                for (int i = 0; i < attDt.Rows.Count; i++)
                {
                    DataRow[] dtr1 = dt.Select("id='" + attDt.Rows[i][3].ToString() + "'");
                    if (dtr1.Length > 0)
                    {
                        attDt.Rows[i][3] = dtr1[0][2].ToString();
                        if (Convert.ToDateTime("09:10") < Convert.ToDateTime(attDt.Rows[i]["sbqd"].ToString()))
                        {
                            attDt.Rows[i]["kqqk"] = "迟到";
                        }
                        if (Convert.ToDateTime(attDt.Rows[i]["xbqd"].ToString()) < Convert.ToDateTime("17:00"))
                        {
                            attDt.Rows[i]["xbqk"] = "早退";
                        }
                        //dt.Rows.RemoveAt(i);
                        //i--;
                    }
                    else
                    {
                        attDt.Rows.RemoveAt(i);
                        i--;
                    }
                }

                bool result = SqlBulkCopyInsert(attDt);
                if (result == false)
                {
                    DBSQL.ExecuteSql("insert into OAAttConcent values (1,'数据批量插入时出错','" + DateTime.Now + "')");
                    return;
                }
                else
                {
                    DBSQL.ExecuteSql("insert into OAAttConcent values (0,'远程考勤数据(" + DateTime.Today.ToString("yyyy-MM-dd") + ")已成功录入','" + DateTime.Now + "')");
                    return;
                }
            }
            catch (Exception e)
            {
                DBSQL.ExecuteSql("insert into OAAttConcent values (1,'提取远程数据出错," + e.ToString() + "','" + DateTime.Now + "')");
                return;
            }

        }

        /// <summary>
        /// 使用SqlBulkCopy方式插入数据
        /// </summary>
        /// <param name="dataTable"></param>
        /// <returns></returns>
        private static bool SqlBulkCopyInsert(DataTable dt)
        {
            try
            {
                DataTable dataTable = dt;

                string connectionString = "server=(local);database=hydroOA;uid=sa;pwd=jianshen";
                SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(connectionString);
                sqlBulkCopy.DestinationTableName = "OAAttendance";
                sqlBulkCopy.BatchSize = dataTable.Rows.Count;
                SqlConnection sqlConnection = new SqlConnection(connectionString);
                sqlConnection.Open();
                if (dataTable != null && dataTable.Rows.Count != 0)
                {
                    sqlBulkCopy.WriteToServer(dataTable);
                }
                sqlBulkCopy.Close();
                sqlConnection.Close();
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }
    }
}
