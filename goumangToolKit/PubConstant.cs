using System;
using System.Configuration;

namespace GoumangToolKit
{
    
    public class PubConstant
    {

         private static string _connectionString;
       
        /// <summary>
        /// ��ȡ�����ַ���
        /// </summary>
        public static string ConnectionString
        {           
            get 
            {
                if (_connectionString!=null)
                {
                    return _connectionString;
                }


                  try
    {
                _connectionString=Properties.Settings.Default.datastr;
    }
                catch
                  {
                      _connectionString = " Database='autorivet';Data Source='192.168.3.32';User Id='autorivet';Password='222222';CharSet=utf8;Allow User Variables=True;Allow Zero Datetime=True";
                }
        

      
;       
                return _connectionString; 
            }
            set
            {
                _connectionString=value;
                
            }




        }

        /// <summary>
        /// �õ�web.config������������ݿ������ַ�����
        /// </summary>
        /// <param name="configName"></param>
        /// <returns></returns>



    }
}
