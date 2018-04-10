using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data.SqlClient;

namespace MyNameSpace
{
	/// <summary>
	/// 用于数据库的增删改查
	/// </summary>
	public class SqlHelper
	{
        //<add connectionString="Data Source=DESKTOP-GMUPG6E;Initial Catalog=nononodeleteImportant;Integrated Security=True" name="constr" />
        //获取配置文件中的sql连接字符串

        public static string[] GetSQLinfor()
        {
            string[] str = new string[3];
            var temp = System.IO.File.OpenText(@"L:\Users\sheng.nan\Tools\MyTest_Do not Deleted.txt");
            str = temp.ReadLine().Split(new char[] { '|' });
            temp.Close();
            return str;
        }
        public static string[] conPwd = GetSQLinfor();
        public static readonly string constr = $@"server={conPwd[0]};DataBase=db_MTOALT1;uid={conPwd[1]};pwd={conPwd[2]}";
        /// <summary>
        /// 增加，删除，更改数据库
        /// </summary>
        /// <param name="sqlcmd"> sql语句</param>
        /// <param name="param">sql语句使用到的参数，防止sql注入</param>
        /// <returns>受影响的行数</returns>
        public static int MyExecuteNonQuery(string sqlcmd, params SqlParameter[] param)
		{
			using (SqlConnection con = new SqlConnection(constr))
			{
				using (SqlCommand cmd = new SqlCommand(sqlcmd, con))
				{
					con.Open();
					if (param != null)
					{
						cmd.Parameters.AddRange(param);
					}
					return cmd.ExecuteNonQuery();
				}
			}
		}
		/// <summary>
		///返回受影响的首行的首列值
		/// </summary>
		/// <param name="sqlcmd"> sql语句</param>
		/// <param name="param">sql语句使用到的参数，防止sql注入</param>
		/// <returns>受影响的首行的列值</returns>
		public static object MyExecuteScalar(string sqlcmd, params SqlParameter[] param)
		{
			using (SqlConnection con = new SqlConnection(constr))
			{
				using (SqlCommand cmd = new SqlCommand(sqlcmd, con))
				{
					con.Open();
					if (param != null)
					{
						cmd.Parameters.AddRange(param);
					}
					return cmd.ExecuteScalar(); ;
				}
			}
		}
		/// <summary>
		/// 读取数据库的内容，返回sqldatareader对象
		/// </summary>
		/// <param name="sqlcmd">sql查询语句</param>
		/// <param name="param">sql查询语句用到的参数，防止sql注入</param>
		/// <returns></returns>
		public static SqlDataReader MyExecuteReader(string sqlcmd, params SqlParameter[] param)
		{
			SqlConnection con = new SqlConnection(constr);
			using (SqlCommand cmd = new SqlCommand(sqlcmd, con))
			{
				if (param != null)
				{
					cmd.Parameters.AddRange(param);
				}
				try
				{
					con.Open();
					return cmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection);
				}
				catch (Exception)
				{
					con.Close();
					con.Dispose();
					throw;
				}
			}
		}
	}
}
