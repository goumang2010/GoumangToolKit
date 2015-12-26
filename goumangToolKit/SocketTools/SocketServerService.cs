using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Sockets;
using System.Net;  // IP，IPAddress, IPEndPoint，端口等；
using System.Threading;
using System.Windows.Forms;
using System.IO;
using System.ComponentModel;

namespace GoumangToolKit
{
  public  class SocketServerService:INotifyPropertyChanged

    {
        Thread threadWatch = null; // 负责监听客户端连接请求的 线程；
        Socket socketWatch = null;
        Dictionary<string, Socket> dict = new Dictionary<string, Socket>();
        Dictionary<string, Thread> dictThread = new Dictionary<string, Thread>();
        string ServerIP = localMethod.GetConfigValue("ServerIP");
        string ServerPort = localMethod.GetConfigValue("ServerPort");
        public event PropertyChangedEventHandler PropertyChanged;
        public List<string> ClientList
        {
            get
            {
                return dict.Keys.ToList();
            }


        }
       

        protected virtual void OnPropertyChanged(string propertyName)
        {
            var handler = PropertyChanged;
            if (handler!=null)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }

        }
        //Start to listen
        public void btnBeginListen()
        {
            // 创建负责监听的套接字，注意其中的参数；
            socketWatch = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);

            IPAddress address = IPAddress.Parse(ServerIP);
            // 创建包含ip和端口号的网络节点对象；
            IPEndPoint endPoint = new IPEndPoint(address, int.Parse(ServerPort));
            try
            {
                // 将负责监听的套接字绑定到唯一的ip和端口上；
                socketWatch.Bind(endPoint);
            }
            catch (SocketException se)
            {
                throw se;
            }
            // 设置监听队列的长度；
            socketWatch.Listen(10);
            // 创建负责监听的线程；
            threadWatch = new Thread(WatchConnecting);
            threadWatch.IsBackground = true;
            threadWatch.Start();
          //  ShowMsg("服务器启动监听成功！");
            //}
        }
        void WatchConnecting()
        {
            while (true)  // 持续不断的监听客户端的连接请求；
            {
                // 开始监听客户端连接请求，Accept方法会阻断当前的线程；
                Socket sokConnection = socketWatch.Accept(); // 一旦监听到一个客户端的请求，就返回一个与该客户端通信的 套接字；
                // 想列表控件中添加客户端的IP信息；
               // lbOnline.Add(sokConnection.RemoteEndPoint.ToString());
                // 将与客户端连接的 套接字 对象添加到集合中；
                dict.Add(sokConnection.RemoteEndPoint.ToString(), sokConnection);
                OnPropertyChanged("ClientList");
               // ShowMsg("客户端连接成功！");
                Thread thr = new Thread(RecMsg);
                thr.IsBackground = true;
                thr.Start(sokConnection);
                dictThread.Add(sokConnection.RemoteEndPoint.ToString(), thr);  //  将新建的线程 添加 到线程的集合中去。
            }
        }


        void RecMsg(object sokConnectionparn)
        {
            Socket sokClient = sokConnectionparn as Socket;
            while (true)
            {
                // 定义一个2M的缓存区；
                byte[] arrMsgRec = new byte[1024 * 1024 * 2];
                // 将接受到的数据存入到输入  arrMsgRec中；
                int length = -1;
                try
                {
                    length = sokClient.Receive(arrMsgRec); // 接收数据，并返回数据的长度；
                }
                catch (SocketException se)
                {
                   // ShowMsg("异常：" + se.Message);
                    // 从 通信套接字 集合中删除被中断连接的通信套接字；
                    dict.Remove(sokClient.RemoteEndPoint.ToString());
                    OnPropertyChanged("ClientList");
                    // 从通信线程集合中删除被中断连接的通信线程对象；
                    dictThread.Remove(sokClient.RemoteEndPoint.ToString());
                    // 从列表中移除被中断的连接IP
                  //  lbOnline.Items.Remove(sokClient.RemoteEndPoint.ToString());
                    break;
                }

                if (arrMsgRec[0] == 0)  // 表示接收到的是数据；
                {
                    if (length == 0)
                    {
                        // 从 通信套接字 集合中删除被中断连接的通信套接字；
                        dict.Remove(sokClient.RemoteEndPoint.ToString());
                        OnPropertyChanged("ClientList");
                        // 从通信线程集合中删除被中断连接的通信线程对象；
                        dictThread.Remove(sokClient.RemoteEndPoint.ToString());
                        // 从列表中移除被中断的连接IP
                 //       lbOnline.Items.Remove(sokClient.RemoteEndPoint.ToString());
                        break;
                    }
                    else
                    {
                        string strMsg = System.Text.Encoding.UTF8.GetString(arrMsgRec, 1, length - 1);
                        if (strMsg.Contains("notify:"))
                        {

                            Notify(strMsg.Replace("notify:", ""));
                        }
                  // ShowMsg(strMsg);
                    }

                }
                if (arrMsgRec[0] == 1) // 表示接收到的是文件；
                {
                    SaveFileDialog sfd = new SaveFileDialog();

                    if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {// 在上边的 sfd.ShowDialog（） 的括号里边一定要加上 this 否则就不会弹出 另存为 的对话框，而弹出的是本类的其他窗口，，这个一定要注意！！！【解释：加了this的sfd.ShowDialog(this)，“另存为”窗口的指针才能被SaveFileDialog的对象调用，若不加thisSaveFileDialog 的对象调用的是本类的其他窗口了，当然不弹出“另存为”窗口。】

                        string fileSavePath = sfd.FileName;// 获得文件保存的路径；
                                                           // 创建文件流，然后根据路径创建文件；
                        using (FileStream fs = new FileStream(fileSavePath, FileMode.Create))
                        {
                            fs.Write(arrMsgRec, 1, length - 1);
                            ShowMsg("文件保存成功：" + fileSavePath);
                        }
                    }
                }
            }
        }

        private void ShowMsg(string v)
        {
            MessageBox.Show(v);
        }

        public void SendFile(string filepath, string strKey)
        {
            using (FileStream fs = new FileStream(filepath, FileMode.Open))
            {
                string fileName = System.IO.Path.GetFileName(filepath);
                string fileExtension = System.IO.Path.GetExtension(filepath);
                string strMsg = "我给你发送的文件为:" + fileName + "\r\n";
                byte[] arrMsg = System.Text.Encoding.UTF8.GetBytes(strMsg); // 将要发送的字符串转换成Utf-8字节数组；
                byte[] arrSendMsg = new byte[arrMsg.Length + 1];
                arrSendMsg[0] = 0; // 表示发送的是消息数据
                Buffer.BlockCopy(arrMsg, 0, arrSendMsg, 1, arrMsg.Length);
                //  bool fff = true;

                if (string.IsNullOrEmpty(strKey))   // 判断是不是选择了发送的对象；
                {
                    MessageBox.Show("请选择你要发送的好友！！！");
                }
                else
                {
                    dict[strKey].Send(arrSendMsg);
                    byte[] arrFile = new byte[1024 * 1024 * 2];
                    int length = fs.Read(arrFile, 0, arrFile.Length); 
                    byte[] arrFileSend = new byte[length + 1];
                    arrFileSend[0] = 1; 
                    Buffer.BlockCopy(arrFile, 0, arrFileSend, 1, length);
                    dict[strKey].Send(arrFileSend);

                }
            }
        }
        public void SendMsg(string strKey, string strMsg)
        {
            byte[] arrMsg = System.Text.Encoding.UTF8.GetBytes(strMsg); // 将要发送的字符串转换成Utf-8字节数组；
            byte[] arrSendMsg = new byte[arrMsg.Length + 1];
            arrSendMsg[0] = 0; // 表示发送的是消息数据
            Buffer.BlockCopy(arrMsg, 0, arrSendMsg, 1, arrMsg.Length);

            if (string.IsNullOrEmpty(strKey))   // 判断是不是选择了发送的对象；
            {
                MessageBox.Show("请选择你要发送的好友！！！");
            }
            else
            {
                dict[strKey].Send(arrSendMsg);

            }
        }
        public void SendMsgToAll(string strMsg)
        {

            byte[] arrMsg = System.Text.Encoding.UTF8.GetBytes(strMsg); 

            byte[] arrSendMsg = new byte[arrMsg.Length + 1]; 
            arrSendMsg[0] = 0; 
            Buffer.BlockCopy(arrMsg, 0, arrSendMsg, 1, arrMsg.Length);
            foreach (Socket s in dict.Values)
            {

                s.Send(arrSendMsg);
            }
           
        }

        //通知所有服务器IP的用户
        public void Notify(string msg)
        {
            string strMsg = "show:" + msg + "\r\n";
            SendMsgToAll(strMsg);
        }




 


    }
}
