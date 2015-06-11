using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using MySql.Data.MySqlClient;
using Office;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;  


/*
 * SignPress程序的服务器程序
 * 
 
 */
namespace server
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("SignPress服务器程序");
            
            /// 测试连接服务器以及查询测试
            /**
            DBTools dbtools = new DBTools();
            MySqlDataReader mysqlread = dbtools.getmysqlread("SELECT * FROM department;");
            while (mysqlread.Read( ))   // 一次读一条记录
            {
                Console.WriteLine(mysqlread["id"].ToString( ) + "  " + mysqlread["name"].ToString( ));
            }**/

            #region 第六行显示
            //Start Word and create a new document.  
            /*
             * 首先添加COM组件Office，
             * 然后添加.Net组件引用Microsoft.Office.Interop.Word;  
             * 接着添加如下的代码
             * using Office；
             * using Microsoft.Office.Interop.Word;
             * using Word = Microsoft.Office.Interop.Word;  
             */
            /// 对word的操作信息
            object oMissing = System.Reflection.Missing.Value;  
            object oEndOfDoc = "//endofdoc"; /* /endofdoc is a predefined bookmark */
            object filename = @".\拨款会签单--航道局.doc";
            //Start Word and create a new document.  
            Word.Application oWord = new Word.Application();
            oWord.Visible = true;  
            // Word.Document oDoc = oWord.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);              
            Word.Documents oDoc = oWord.Documents.Open(

            
            #endregion
        }
    }
}
