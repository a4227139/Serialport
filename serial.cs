using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
//using System.Threading;
using System.Runtime.InteropServices;
using System.Collections;
using System.IO;
using System.Diagnostics;
using System.Data.SqlClient;


namespace Serialport
{
   
   
    public partial class serial : Form
    {

        public class Device
        {
            public string name;
            public DateTime intime;

            public Device(string n, DateTime dt)
            {
                name = n;
                intime = dt;
            }
        };
        public static int Count=0;
        public static int Count1 = 0;
        private StringBuilder builder = new StringBuilder();
        private bool Listening = false;//是否没有执行完invoke相关操作
        private bool Closing = false;//是否正在关闭串口，执行Application.DoEvents，并阻止再次invoke
        private List<byte> buffer = new List<byte>(4096);//默认分配1页内存，并始终限制不允许超过
        private long received_count = 0;
        DateTime flag_time = DateTime.Now;
        string[] date_array;
        Device[] device = new Device[30];
        struct LastXY //记录上一个数据包处理后的X，Y值
        {
            public string net;
            public double lastX;
            public double lastY;
        }
        LastXY[] lastxy = new LastXY[30];
        struct BtoM
        {
            public string beice;
            public string maodian;
            public double x;
            public double y;
            public double R;
        };
        BtoM[] btom = new BtoM[3];
        struct Point 
        {
            public double x;
            public double y;//x,y为坐标
            public int i;
            public int j;//i,j表示是那两个圆的交点
        };
        Point[] point = new Point[6];
        private static string sqlData = "Data Source=.;Initial Catalog=Contiki;Integrated Security=True";//连接数据库的语句
       // SqlCommand cmd = new SqlCommand("select * from Record_Online", conn);
        static SqlConnection conn = new SqlConnection(sqlData);//建立连接
        static SqlConnection conn1 = new SqlConnection(sqlData);
        SerialPort myserialPort = new System.IO.Ports.SerialPort();
            
        TimeSpan noSignal_time = new TimeSpan();

           void procTimer_Tick(object sender, EventArgs e)        
          {    

            for (int i = 0; i <= Count; i++)
            {
                if (device[i] == null)
                    continue;
                noSignal_time = DateTime.Now - device[i].intime;
                if (Convert.ToInt32(noSignal_time.TotalMinutes) > 5)
                {
                    //Device temp = new Device("0", flag_time);
                    richTextBox1.Text += device[i].name + " 异常！"+"\r\n";
                    for (int j = i; j < Count;j++ )
                     device[j] = device[j + 1]; 
                        Count--;
                }
            }
         }
        public serial()
        {
            InitializeComponent();
            comboBox1.Items.AddRange(System.IO.Ports.SerialPort.GetPortNames());
            button3.Enabled = false;
            
           /* if (!myserialPort.IsOpen)
            {
                myserialPort.Open();
                richTextBox1.Text = "";
            }
            else
                richTextBox1.Text = "port busy:";*/
        }
        //private string rsSTRING;

        private void button1_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
        }

        private void myserialPort_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            
            if (Closing) return;//如果正在关闭，忽略操作，直接返回，尽快的完成串口监听线程的一次循环
            int n; byte[] buf;
            try
            {
                Listening = true;//设置标记，说明我已经开始处理数据，一会儿要使用系统UI的。
                n = myserialPort.BytesToRead;//先记录下来，避免某种原因，人为的原因，操作几次之间时间长，缓存不一致
                buf = new byte[n];//声明一个临时数组存储当前来的串口数据
                received_count += n;//增加接收计数
                myserialPort.Read(buf, 0, n);//读取缓冲数据

                /////////////////////////////////////////////////////////////////////////////////////////////////////////////
                //1.缓存数据
                buffer.AddRange(buf);
                //有数据才反应
                flag_time = DateTime.Now;
                //2.完整性判断
                bool flag = true;
                bool flag2 = true;
                while (buffer.Count >= 30)//>= 29至少要包含头（4字节）+长度（2字节）+校验（1字节）直到将buffer的数据处理完！
                {
                    //请不要担心使用>=，因为>=已经和>,<,=一样，是独立操作符，并不是解析成>和=2个符号
                    //2.1 查找数据头
                    if (buffer[0] == 'F' && buffer[1] == 'E')
                    {
                        int date_len = 2;
                        string date_str = "";

                        while (buffer[date_len] != 'E' || buffer[date_len + 1] != 'E') //&& buffer[date_len + 1] != 'E'
                        {
                            if (buffer[date_len] == 'F' && buffer[date_len + 1] == 'E')
                            {
                                buffer.RemoveRange(0, date_len);
                                flag = false;
                                break;
                            }
                            int code = Convert.ToInt32(buffer[date_len]);
                            if (code != 44) //将10进制转换为字符
                            {
                                code = code - 48;
                                date_str += code.ToString();
                            }
                            else
                            { date_str += ","; }

                            // date_str += buffer[date_len];
                            date_len++;
                            if (buffer.Count < date_len + 2)
                            {
                                flag2 = false;
                                break;
                            }
                        }
                        if (flag == false)
                        {
                            flag = true;
                            continue;
                        }
                        if (flag2 == false)
                        {
                            flag2 = true;
                            break;
                        }
                        //MessageBox.Show(date_str);
                        date_array = date_str.Split(',');
                        if (date_array.Length < 7)  //接收的数据不完整
                        {
                            int jj;
                            for (jj = 2; jj < buffer.Count; jj++)
                            {
                                if (buffer[jj - 1] == 'F' && buffer[jj] == 'E')
                                    break;
                            }
                            buffer.RemoveRange(0, jj);//从缓存中删除错误数据
                            continue;//继续下一次循环
                        }
                        if (date_array.Length == 7)
                        {
                            this.Invoke(new EventHandler(SaveSql));
                            this.Invoke(new EventHandler(Device_Exception));

                            buffer.RemoveRange(0, date_len);
                        }

                    }
                    else
                    {
                        //这里是很重要的，如果数据开始不是头，则删除数据
                        buffer.RemoveAt(0);
                    }
                }
                /////////////////////////////////////////////////////////////////////////////////////////////////////////////

                builder.Clear();//清除字符串构造器的内容
            }
            finally
            {
                Listening = false;//我用完了，ui可以关闭串口了。
            }

        }
        private void Device_Exception(object o, EventArgs e) 
        {
            
            DateTime dt=DateTime.Now;
            int count = Count; int i;
            //arr[0]=new int[]{10,20,30};
            for (i = 0; i <= count; i++) 
            {
                if (device[i] == null)
                    continue;
                if (device[i].name == date_array[0] + date_array[1])
                {
                    device[i].intime = dt;
                    break;
                }
            }
            if (i == count+1 )
            {
                device[i-1] = new Device(date_array[0] + date_array[1], dt);
                Count++;
            }
           // richTextBox1.Text += "dt:" + dt.ToString()+"Count:"+Count.ToString()+"\r\n";
        }
        double GetR(string Rssi)//由rssi值求距离的函数
        {
            //double A = 45, N;   //N = 2.24;   //2 .... 2.24
            double R;
            int rssi = int.Parse(Rssi.Replace(" ",""));
            if (rssi <= 20)
            {
                R = 0.1;
            }
            else if (rssi < 45)
            {
                R = (rssi - 20) * 0.04;
            }
            else if (rssi >= 45 && rssi <= 47)
            {
                R = 1;
            }
            else if (rssi < 58)
            {
                R = (rssi - 47) * 0.36 + 1;
            }
            else if (rssi >= 58 && rssi <= 60)
            {
                R = 5;
            }
            else if (rssi < 62)
            {
                R = (rssi - 60) * 2.5 + 5;
            }
            else if (rssi >= 62 && rssi <= 64)
            {
                R = 10;
            }
            else if (rssi < 69)
            {
                R = (rssi - 64) * 2 + 10;
            }
            else if (rssi >= 69 && rssi <= 70)
            {
                R = 20;
            }
            else if (rssi < 77)
            {
                R = (rssi - 70) * 1.4 + 20;
            }
            else if (rssi >= 77 && rssi <= 82)
            {
                R = 30;
            }
            else if (rssi == 83)
            {
                R = 35;
            }
            else if (rssi >= 84 && rssi <= 87)
            {
                R = 40;
            }
            else
            {
                R = 50;
            }
            return R;
        }
        int SelectPoint()//选择最适合的交点,返回该点在point数组的下标
        {
            int i, j = 0, k; double[] d = new double[6]; double ymax=btom[0].y,ymin=btom[0].y,min;
            for (i = 0; i < btom.Length; i++)//找出隧道宽度上下限
            {
                if (ymax < btom[i].y)
                    ymax = btom[i].y;
                if (ymin > btom[i].y)
                    ymin = btom[i].y;
            }
            //richTextBox1.Text += "ymax:" + ymax.ToString()+"  ymin:"+ymin.ToString()+"\r\n";
                for (i = 0; i < point.Length; i++)
                {
                    //richTextBox1.Text +="before point下标为："+i.ToString()+ "  x:" + point[i].x.ToString() + "  y:" + point[i].y.ToString() + "\r\n";
                    if (point[i].y < ymin-1 || point[i].y > ymax+1)//如果交点在矩形（即隧道）外，则不算
                        continue;
                    
                    k = 0;
                    if (k == point[i].i || k == point[i].j) k++;
                    if (k == point[i].i || k == point[i].j) k++;//经过两次判断可得第三圆在btom数组的下标
                    if (Math.Pow(point[i].x - btom[k].x, 2) + Math.Pow(point[i].y - btom[k].y, 2) == Math.Pow(btom[k].R, 2))//如果三圆交于同一点，返回该点下标。
                    {
                        return i;
                    }
                    else if (Math.Pow(point[i].x - btom[k].x, 2) + Math.Pow(point[i].y - btom[k].y, 2) < Math.Pow(btom[k].R, 2)) //判断交点在第三圆的圆内还是圆外,如果在园内
                    {
                        //richTextBox1.Text += "d=" + Math.Pow(btom[k].R, 2).ToString() +"-"+ (Math.Pow(point[i].x - btom[k].x, 2) + Math.Pow(point[i].y - btom[k].y, 2)).ToString()+"\r\n";
                        d[i] = Math.Pow(btom[k].R, 2) - (Math.Pow(point[i].x - btom[k].x, 2) + Math.Pow(point[i].y - btom[k].y, 2));//半径的平方减去该交点到圆心距离的平方
                        //j++;
                    }
                    else //判断交点在第三圆的圆内还是圆外,如果在园外
                    {
                       // richTextBox1.Text += "d=" + (Math.Pow(point[i].x - btom[k].x, 2) + Math.Pow(point[i].y - btom[k].y, 2)).ToString()+"-" +Math.Pow(btom[k].R, 2).ToString()+"\r\n";
                        d[i] = (Math.Pow(point[i].x - btom[k].x, 2) + Math.Pow(point[i].y - btom[k].y, 2)) - Math.Pow(btom[k].R, 2);//该交点到圆心距离的平方减去半径的平方
                        //j++;
                    }
                    //richTextBox1.Text += "before  d:" + d[i].ToString()+"  i:"+i.ToString()+"\r\n";
                    //richTextBox1.Text += "x:" + point[i].x.ToString() + "  y:" + point[i].y.ToString() + "\r\n";
                   // richTextBox1.Text += "after point下标为：" + i.ToString()+"  j为"+j.ToString() + "  x:" + point[i].x.ToString() + "  y:" + point[i].y.ToString() + "  d:" + d[j - 1].ToString() + "\r\n";
                }
                min = d[0];
            for (i = 0; i < d.Length-1; i++)
            {

                if (min > d[i + 1] && d[i + 1]>0)
                {
                    min = d[i + 1];
                    j = i + 1;
                }
                //richTextBox1.Text += "after  d:" + d[i].ToString() + "  min:" + min.ToString() + "\r\n";
            }
           // richTextBox1.Text += "返回的point坐标为：" + j.ToString()+"\r\n";
            return j;
        }
        Point  GetPoint()//求三个圆的全部交点
        {
            int i, j, k=0; double d,p,q,a,b,c,rsqu,rsum;//a,b,c分别为圆和交点直线联立后所得的一元二次方程的二次项系数，一次项系数，常数项。p,q为中间量。rsum为两圆半径之和，rsqu为两圆半径之和的平方。
            for(i=0,j=1;i<=2;i++,j++)//循环3次分别求三个圆中的某圆与另两圆的交点
            {
                if (i == 2) j = 0;
                d = Math.Pow(btom[i].x - btom[j].x, 2)+Math.Pow(btom[i].y - btom[j].y, 2);//d为两圆心距离的平方
                rsum = btom[i].R + btom[j].R; rsqu = rsum * rsum;
                //richTextBox1.Text += "xi:" + btom[i].x.ToString() + "  yi:" + btom[i].y.ToString() + "  xj:" + btom[j].x.ToString() + "  yj:" + btom[j].y.ToString() + "\r\n";
                //richTextBox1.Text += "iR:" + btom[i].R.ToString() + j.ToString() + "  jR:" + btom[j].R.ToString() + "\r\n"+"两圆心距离的平方:" + d.ToString() + "  两圆半径相加的平方："+rsqu.ToString() + "\r\n";
                if (d <= rsqu)//如果两圆心距离的平方小于等于两半径相加的平方，即两圆心的距离小于等于两半径之和，也就是两圆有交点
                {
                    if (btom[i].y == btom[j].y)
                    {
                        point[k].x =(Math.Pow(btom[i].R,2)- Math.Pow(btom[j].R,2)+ Math.Pow(btom[j].x,2)- Math.Pow(btom[i].x,2))/(2*(btom[j].x-btom[i].x));
                        point[k].y = Math.Pow(Math.Pow(btom[i].R, 2) - Math.Pow(point[k].x - btom[i].x, 2), 1 / 2) + btom[i].y;
                        point[k].i = i;
                        point[k].j = j;
                        //richTextBox1.Text += "纵坐标相等：x:" + point[k].x.ToString() + " y:" + point[k].y.ToString() +"下标：" + k.ToString() + "\r\n"; ;
                        k++;
                        point[k].x = point[k - 1].x;
                        point[k].y = -Math.Pow(Math.Pow(btom[i].R, 2) - Math.Pow(point[k].x - btom[i].x, 2), 1 / 2) + btom[i].y;
                        point[k].i = i;
                        point[k].j = j;
                        k++;
                    }     
                    else if (btom[i].x == btom[j].x)
                    {
                        point[k].y = (Math.Pow(btom[i].R, 2) - Math.Pow(btom[j].R, 2) + Math.Pow(btom[j].y, 2) - Math.Pow(btom[i].y, 2)) / (2 * (btom[j].y - btom[i].y));
                        point[k].x = Math.Pow(Math.Pow(btom[i].R, 2) - Math.Pow(point[k].y - btom[i].y, 2), 1 / 2) + btom[i].x;
                        point[k].i = i;
                        point[k].j = j;
                        //richTextBox1.Text += "横坐标相等：x:" + point[k].x.ToString() + " y:" + point[k].y.ToString() + "下标：" + k.ToString() + "\r\n"; ;
                        k++;
                        point[k].y = point[k - 1].y;
                        point[k].x = -Math.Pow(Math.Pow(btom[i].R, 2) - Math.Pow(point[k].y - btom[i].y, 2), 1 / 2) + btom[i].x;
                        point[k].i = i;
                        point[k].j = j;
                        k++;
                        //richTextBox1.Text += "横坐标相等：x:" + point[k].x.ToString() + " y:" + point[k].y.ToString() + "下标：" + k.ToString() + "\r\n"; ;
                    }    
                    /*else if (d == Math.Pow((btom[i].R + btom[j].R), 2))//如果两圆相切
                    {
                        p = (btom[i].x - btom[j].x) / (btom[j].y - btom[i].y);
                        q = (Math.Pow(btom[i].R, 2) - Math.Pow(btom[j].R, 2) + Math.Pow(btom[j].x, 2) - Math.Pow(btom[i].x, 2) + Math.Pow(btom[j].y, 2) - Math.Pow(btom[i].y, 2)) / (2 * (btom[j].y - btom[i].y));
                        a = 1 + Math.Pow(p, 2);
                        b = 2 * (p * q - p * btom[i].y - btom[i].x);
                        c = Math.Pow(q - btom[i].y, 2) + Math.Pow(btom[i].x, 2) - Math.Pow(btom[i].R, 2);
                        point[k].x=-b/(2*a);
                        point[k].y = -Math.Pow(b, 2) /( 4 * a) + c;
                        point[k].i = i;
                        point[k].j = j;
                        k++;
                    }*/
                    else//两圆相交或相切
                    {
                        p = (btom[i].x - btom[j].x) / (btom[j].y - btom[i].y);
                        q = (Math.Pow(btom[i].R, 2) - Math.Pow(btom[j].R, 2) + Math.Pow(btom[j].x, 2) - Math.Pow(btom[i].x, 2) + Math.Pow(btom[j].y, 2) - Math.Pow(btom[i].y, 2)) / (2 * (btom[j].y - btom[i].y));
                        a = 1 + Math.Pow(p, 2);
                        b = 2 * (p * q - p * btom[i].y - btom[i].x);
                        c = Math.Pow(q - btom[i].y, 2) + Math.Pow(btom[i].x, 2) - Math.Pow(btom[i].R, 2);
                        //richTextBox1.Text += "p:" + p.ToString() + " q:" + q.ToString() + " a:" + a.ToString() + " b:" + b.ToString() + " c:" + c.ToString()+"\r\n";
                        point[k].x = (-b + Math.Pow((Math.Pow(b, 2) - 4 * a * c), 1 / 2)) / (2 * a);
                        point[k].y = p * point[k].x + q;//a * Math.Pow(point[k].x, 2) + b * point[k].x + c;
                        point[k].i = i;
                        point[k].j = j;
                        //richTextBox1.Text += "相交或相切x:" + point[k].x.ToString() + " y:" + point[k].y.ToString()+"下标："+k.ToString()+"\r\n";
                        k++;
                        point[k].x = (-b - Math.Pow((Math.Pow(b, 2) - 4 * a * c), 1 / 2)) / (2 * a);
                        point[k].y = p * point[k].x + q;//a * Math.Pow(point[k].x, 2) + b * point[k].x + c;
                        point[k].i = i;
                        point[k].j = j;
                        //richTextBox1.Text += "相交或相切x:" + point[k].x.ToString() + " y:" + point[k].y.ToString() + "下标：" + k.ToString() + "\r\n";
                        k++;
                    }
                }
            }
            /*for (i = 0; i <= k; i++)
                richTextBox1.Text += "point" + i.ToString() + " x:"+point[i].x.ToString("#.00") + " y:"+point[i].y.ToString("#.00")+"\r\n";*/
                i = SelectPoint();
            return point[i];
            
        }
        int Get_Location(ref double X, ref double Y)//定位函数，返回节点node的x,y坐标
        {
            int i = 0; Point p=new Point();
            SqlCommand cmd_MBR_Exist = new SqlCommand("select* from MBR where Beice='" + date_array[0] + date_array[1] + "' ", conn);//从锚点-被测-RSSI表中找出被测节点地址是新收到的数据包发送者的记录（应有3条）
            SqlDataReader reader1 = null;
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
            conn.Open();
            
            reader1 = cmd_MBR_Exist.ExecuteReader();//执行MBR表操作
            while (reader1.Read()) i += 1;//统计符合上述条件的记录的个数
            if (i < 3)//如果小于3，无法定位，退出
            {
                reader1.Close();
                conn.Close();
                return 0;
            }
            else//如果有3条符合，即一个被测对应有了3个锚点
            {
                SqlCommand cmd_Maodian_Exist = new SqlCommand("select Net,X,Y,Beice,Rssi from Maodian,MBR where Maodian.Net=MBR.Maodian and Beice="+date_array[0]+date_array[1]+"", conn1);//关联锚点表和MBR表，找出锚点地址，坐标，被测节点地址，RSSI值
                SqlDataReader reader2 = null;
                conn1.Open();
                i = 0;
                reader2 = cmd_Maodian_Exist.ExecuteReader();
              while (reader2.Read())//将上述值赋值到btom结构数组
                {
                    btom[i].maodian = reader2[0].ToString();
                    btom[i].x = double.Parse(reader2[1].ToString());
                    btom[i].y = double.Parse(reader2[2].ToString());
                    btom[i].beice = reader2[3].ToString();
                    btom[i].R = GetR(reader2[4].ToString());
                    i++;
                  //richTextBox1.Text += "i:" + i.ToString() + "  R:" + btom[i].R.ToString() + "\r\n"; 
                  //richTextBox1.Text += "i:"+i.ToString()+" reader2[0]:" + reader2[0].ToString() + "reader2[1]:" + reader2[1].ToString() + "reader2[2]:" + reader2[2].ToString() + "reader2[3]:" + reader2[3].ToString() + "reader2[4]:" + reader2[4].ToString() +  "\r\n";
                }
                p=GetPoint();
                X = p.x;
                Y = p.y;
                reader2.Close();
                conn1.Close();
                conn.Close();
            }
            return 0;
        }
        private void SaveSql(object o, EventArgs e)//收到数据包进入该函数
        {
            string Voltage_End, Voltage_Maodian, date, leavetime; double X = 0, Y = 0; int i;
            DateTime dt = DateTime.Now;  
            SqlCommand cmd_Check_Exist = new SqlCommand("select * from Check1", conn1);//浏览Check1表
            SqlDataReader reader = null;
            string Save_MBR = "if exists(select * from MBR where Maodian='" + date_array[2] + date_array[3] + "'and Beice='" + date_array[0] + date_array[1] + "')begin update MBR set Rssi='" + date_array[4] + "'where Maodian='" + date_array[2] + date_array[3] + "'and Beice='" + date_array[0] + date_array[1] + "' end else begin insert into MBR(Maodian,Beice,Rssi) values('" + date_array[2] + date_array[3] + "','" + date_array[0] + date_array[1] + "' ,'" + date_array[4] + "') end";//将锚地址，被测节点地址，它们之间的rssi值写入MBR表
            SqlCommand cmd_Save_MBR = new SqlCommand(Save_MBR, conn);
            string Insert_Check1 = "insert into Check1(Net,Intime,Outtime,Date) values('" + date_array[0] + date_array[1] + "','" + dt.ToString("HH:mm") + "' ,' "+dt.ToString("HH:mm")+"',' "+ dt.ToString("yyyy/MM/dd") +"')";//将新纪录（新的时间）插入考勤表
            SqlCommand cmd_Insert_Check1 = new SqlCommand(Insert_Check1, conn);
            string Initialise_Check1 = "if not exists(select * from Check1 where Net='" + date_array[0] + date_array[1] + "')begin insert into Check1(Net,Intime,Outtime,Date) values('" + date_array[0] + date_array[1] + "','" + dt.ToString("HH:mm") + "' ,' " + dt.ToString("HH:mm") + "',' " + dt.ToString("yyyy/MM/dd") + "') end";
            SqlCommand cmd_Initialise_Check1 = new SqlCommand(Initialise_Check1, conn);
            
            richTextBox1.Text += date_array[0] + "," + date_array[1] + "," + date_array[2] + "," + date_array[3] + "," + date_array[4] + "," + date_array[5] + "," + date_array[6] + "\r\n";
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
            conn.Open();
            cmd_Initialise_Check1.ExecuteNonQuery();
            cmd_Save_MBR.ExecuteNonQuery();
            Get_Location( ref  X, ref Y);//确定X,Y的值
            if (Y < 0) Y = 0;
            for (i = 0; i < Count1; i++)
            if (Math.Abs(lastxy[i].lastX - X) > 20 || Math.Abs(lastxy[i].lastY - Y) >20 )
            {
                X = lastxy[i].lastX; Y = lastxy[i].lastY;
            }
            conn.Open();
            string Save_Location = "insert into  Location(Net,X,Y,Date,Time) values('" + date_array[0] + date_array[1] + "' ,'" + X.ToString("#.00") + "','" + Y.ToString("#.00") + "','" + dt.ToString("yyyy/MM/dd") + "' ,' " + dt.ToString("HH:mm:ss") + "')";//将节点地址和XY坐标写入定位表
            SqlCommand cmd_Save_Location = new SqlCommand(Save_Location, conn);
            //richTextBox1.Text += "X:" + X.ToString() + "  Y:" + Y.ToString()+"\r\n";
            cmd_Save_Location.ExecuteNonQuery();//执行定位表操作
            Voltage_End = date_array[5];
            Voltage_Maodian = date_array[6];
            double voltage_end=double.Parse(Voltage_End)/100;//计算被测节点电压
            double voltage_maodian = double.Parse(Voltage_Maodian) / 100;//计算锚点电压
            //以下四句将节点地址和电压写入设备表
            string Save_Device_End = "if exists(select * from Device where Net='" + date_array[0] + date_array[1] + "')begin update Device set Voltage='" + voltage_end.ToString() + "'where Net='" + date_array[0] + date_array[1] + "' end else begin insert into Device(NET,Voltage) values('" + date_array[0] + date_array[1] + "','" + voltage_end.ToString() + "') end";
            string Save_Device_Maodian = "if exists(select * from Device where Net='" + date_array[2] + date_array[3] + "')begin update Device set Voltage='" + voltage_maodian.ToString() + "'where Net='" + date_array[2] + date_array[3] + "' end else begin insert into Device(NET,Voltage) values('" + date_array[2] + date_array[3] + "','" + voltage_maodian.ToString() + "') end";
            SqlCommand cmd_Save_Device_End = new SqlCommand(Save_Device_End, conn);
            SqlCommand cmd_Save_Device_Maodian = new SqlCommand(Save_Device_Maodian, conn);
            cmd_Save_Device_End.ExecuteNonQuery();
            cmd_Save_Device_Maodian.ExecuteNonQuery();//执行设备表操作
            conn1.Open();//为reader新开一个连接
            reader = cmd_Check_Exist.ExecuteReader();//定义浏览check表的对象
            while (reader.Read())//执行浏览check表操作
            {
                SqlCommand cmd_Check_Intime_Update = new SqlCommand("update Check1 set Intime='" + dt.ToString("HH:mm") + "' where Net='" + date_array[0] + date_array[1] + "'", conn);//更新进入时间
                SqlCommand cmd_Check_Date_Update = new SqlCommand("update Check1 set Date='" + dt.ToString("yyyy/MM/dd") + "' where Net='" + date_array[0] + date_array[1] + "'", conn);//更新日期
                SqlCommand cmd_Check_Outtime_Update = new SqlCommand("update Check1 set Outtime='" + dt.ToString("HH:mm") + "' where Net='" + date_array[0] + date_array[1] + "'", conn);//更新离开时间
                SqlCommand cmd_Check_Flag_Update = new SqlCommand("update Check1 set Flag='" + 1 + "' where Net='" + date_array[0] + date_array[1] + "'", conn);//将Flag位设1
                if (reader[4].ToString().Replace(" ","") == "1")
                    continue;
                date = reader[3].ToString();
                //richTextBox1.Text += "date;" + date.Replace(" ", "") +"intime:" + intime.ToString("yyyy/MM/dd") + "\r\n";
                string a = reader[0].ToString(), b = date_array[0] + date_array[1];
                if (double.Parse(a)==double.Parse(b))//reader从头开始浏览check表中各行，如果reader【0】即Net中有匹配该数据包发送者节点地址的才能进行下面操作，否则什么也不做
                {
                    if (reader[1].ToString() == "")//如果intime为空，更新
                    {
                        //richTextBox1.Text += "not exist" + "\r\n";
                        cmd_Check_Intime_Update.ExecuteNonQuery();
                        cmd_Check_Date_Update.ExecuteNonQuery();
                        cmd_Check_Outtime_Update.ExecuteNonQuery();
                    }
                    else if (date.Replace(" ", "") == dt.ToString("yyyy/MM/dd"))//如果是同一天
                    {
                        leavetime=reader[2].ToString();
                        DateTime leave = DateTime.ParseExact(leavetime.Replace(" ",""), "HH:mm", null);//将数据库中的字符串格式的outtime转换格式
                        TimeSpan ts_now = new TimeSpan(dt.Ticks);
                        TimeSpan ts_outtime = new TimeSpan(leave.Ticks);
                        string diff = ts_now.Subtract(ts_outtime).TotalMinutes.ToString();//求出新的进入时间和之前的离开时间相差的间隔
                        //richTextBox1.Text += diff + "\r\n";
                        if (double.Parse(diff) > 120)
                        {//如果间隔大于2个小时新建一行
                            //richTextBox1.Text += reader[0].ToString()+"置一，建行"+"\r\n"; 
                            cmd_Check_Flag_Update.ExecuteNonQuery();
                            cmd_Insert_Check1.ExecuteNonQuery();
                        }
                        else cmd_Check_Outtime_Update.ExecuteNonQuery();//否则，只更新outtime
                    }
                    else //如果不是同一天,与大与两小时时相同处理
                    {
                        cmd_Check_Flag_Update.ExecuteNonQuery();
                        cmd_Insert_Check1.ExecuteNonQuery();
                    }
                }
            }
            reader.Close();
            conn1.Close();
            conn.Close();
            
            for(i=0;i<Count1;i++)
                if (lastxy[i].net == date_array[0] + date_array[1])
                {
                    lastxy[i].lastX = X;
                    lastxy[i].lastY = Y;
                    break;
                }
            if (i == Count1)
            {
                lastxy[i].lastX = X;
                lastxy[i].lastY = Y;
                lastxy[i].net = date_array[0] + date_array[1];
                Count1++;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            Timer procTimer = new Timer();
            procTimer.Interval = 10000;
            procTimer .Start();
            procTimer .Tick += new EventHandler(procTimer_Tick);
            /*TimeSpan noSignal_time = new TimeSpan();
            for (int i = 0; i <= Count; i++)
            {
                noSignal_time = DateTime.Now - device[i].intime;
                if (Convert.ToInt32(noSignal_time.TotalMinutes) < 5)
                {
                    richTextBox1.Text += device[i].name + "异常！";
                    Count--;
                }
            }*/
            
            /*myserialPort.DataReceived += myserialPort_DataReceived;
            SqlCommand cmd_Clear_Location = new SqlCommand("delete from Location", conn);
            SqlCommand cmd_Clear_Device = new SqlCommand("delete from Device", conn);
            SqlCommand cmd_Clear_Check1 = new SqlCommand("update Check1 set Intime=null,Outtime=null", conn);
            conn.Open();
            cmd_Clear_Location.ExecuteNonQuery();
            cmd_Clear_Device.ExecuteNonQuery();
            cmd_Clear_Check1.ExecuteNonQuery();
            conn.Close();*/
        }

        private void button2_Click(object sender, EventArgs e)
        {
            myserialPort.BaudRate = 115200;
            myserialPort.NewLine = "\r\n";
            myserialPort.RtsEnable = true;//根据实际情况吧。
            myserialPort.PortName = comboBox1.SelectedItem.ToString();
            myserialPort.Open();
            myserialPort.DataReceived += myserialPort_DataReceived;
            button2.Enabled = false;
            button3.Enabled = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            myserialPort.Close();
            button2.Enabled = true;
            button3.Enabled = false;
        }
    }
}
