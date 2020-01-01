using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;//添加类库
using System.IO;//添加类库 输入和输出


namespace _20170606_信一哲
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public double dmstorad(string s)//封装角度转弧度的计算方法
        {
            string[] ss = s.Split(new char[3] { '°', '′', '″' }, StringSplitOptions.RemoveEmptyEntries);
            //用°、′、″将输入的角度值字符ss分为三个字符，并且删掉空格
            double[] d = new double[ss.Length];//定义一个双精度数组d，存放ss的长度
            for (int i = 0; i < d.Length; i++)
                d[i] = Convert.ToDouble(ss[i]);//将分割后的每个字符利用循环结构转为双精度double类型
            double sign = d[0] >= 0.0 ? 1.0 : -1.0;//判断首个数值是否大于零
            double rad = 0;
            if (d.Length == 1)
                rad = Math.Abs(d[0]) * Math.PI / 180;//若只有°前有数字，则转化后的弧度值等于该数字取绝对值后，乘以π，除以180
            else if (d.Length == 2)//当只有°、′前有数字的情况
                rad = (Math.Abs(d[0]) + d[1] / 60) * Math.PI / 180;
            else//当°、′、″处都有数字时
                rad = (Math.Abs(d[0]) + d[1] / 60 + d[2] / 60 / 60) * Math.PI / 180;
            rad = sign * rad;//将°前的正负号“还”回来
            return rad;//返回所得的弧度值
        }

        public string radtodms(double rad)//封装弧度转角度的计算方法
        {
            double sign = rad >= 0.0 ? 1.0 : -1.0;//判断弧度的正负
            rad = Math.Abs(rad) * 180 / Math.PI;//将弧度转化为角度值
            double[] d = new double[3];//定义一个双精度数组d
            d[0] = (int)rad;//取整得度
            d[1] = (int)((rad - d[0]) * 60);//将弧度值减去°后，求′，并保留整数
            d[2] = (rad - d[0] - d[1] / 60) * 60 * 60;//将弧度值减去°、′后，求″，不取整，保留精度
            d[2] = Math.Round(d[2], 2);//取″保留精度后的两位小数
            if (d[2] == 60) //特殊情况判断，若″前数字为60
            {
                d[1] += 1;//向′进一位
                d[2] -= 60;//″的值再减去60
                if (d[1] == 60)//若′前数字为60
                {
                    d[0] += 1;//向°进一位
                    d[1] -= 60;//′的值再减去60
                }
            }
            d[0] = sign * d[0];//将弧度值的正负号再“还给”角度值
            string s = Convert.ToString(d[0]) + '°' + Convert.ToString(d[1]) + '′' + Convert.ToString(d[2]) + '″';
            //定义一个字符串s，存放将所得的°、′、″前的数字加上符号后整合后，并双精度类型转化为字符串类型后的值
            return s;//返回s的值
        }
        public double fangweijiaotuisuan(double[] sdr, double[] cr)//封装推算坐标方位角的计算方法
        {
            double sum = 0;//定义一个双精度变量sum，初值为0
            for (int i = 1; i < sdr.Length; i++)//计算坐标方位角的循环结构
            {
                cr[i] = cr[i - 1] + sdr[i] - Math.PI;
                //坐标方位角计算公式，坐标方位角=上一个角度的坐标方位角+此时的观测角度（左角）-π
                if (cr[i] >= Math.PI * 2) cr[i] -= Math.PI * 2;//特殊情况判断，当坐标方位角>360°时，应减去360°
                else if (cr[i] < 0.0) cr[i] += Math.PI * 2;//特殊情况判断，当坐标方位角<0°时，应加上360°
                sum += sdr[i];//求总和
            }
            return sum;//返回sum的最终结果
        }
        private void button1_Click(object sender, EventArgs e)//计算主程序
        {
          //角度计算主程序
            string[] sd = new string[dataGridView1.RowCount - 5];//新建一个数组，存放观测角度的原始值
            double[] sdr = new double[sd.Length];//新建一个数组，存放观测角度的弧度值
            double[] cr = new double[sdr.Length];//新建一个数组，存放计算的坐标方位角
            double sum = 0;//定义一个双精度变量，，存放角度总值，初值为0
            cr[0] = dmstorad(Convert.ToString(dataGridView1.Rows[0].Cells[4].Value));
            //获取第一个坐标方位角，并将其转化为弧度，放入cr[]数组第一个元素中
            double acd = dmstorad(Convert.ToString(dataGridView1.Rows[dataGridView1.RowCount - 6].Cells[4].Value));
            //获取终边坐标方位角，并将其转化为弧度，放入acd中用于计算和检核
            for (int i=1; i<sd.Length; i++)//从第二行开始循环，将观测角度的原始值放入sd[]数组中，并转化成弧度值存放在sdr数组中
            {
                sd[i] = Convert.ToString(dataGridView1.Rows[i].Cells[1].Value);
                sdr[i] = dmstorad(sd[i]);
            }
            sum = fangweijiaotuisuan(sdr, cr);//计算改正前坐标方位角和观测角度的总和，分别存储在cr数组和sum中
            dataGridView1.Rows[dataGridView1.RowCount - 4].Cells[1].Value = radtodms(sum);//将观测角度总和放入表格中
            double fd, fdx;//定义两个双精度类型的变量fd,fdx
            fd = cr[cr.Length - 1] - acd;//计算角度闭合差，单位弧度
            fdx = 60 * Math.Sqrt(sd.Length - 1);//计算角度闭合差限差，单位是秒
            dataGridView1.Rows[dataGridView1.RowCount - 3].Cells[1].Value = 
                Convert.ToString(Math.Round(fd * 180 / Math.PI * 3600, 2)) + "″";//将角度闭合差存入表格中
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[1].Value = Convert.ToString(Math.Round(fdx,2))+"″";
            //将角度闭合限差存入表格中
            if(Math.Abs(fd * 180 / Math.PI * 3600) > fdx)//检查角度闭合差是否满足要求
                MessageBox.Show("角度闭合差超限！");
            else 
            {
                double vd = -fd / (sd.Length - 1);//分配角度闭合差（观测左角）
                double sumvd = 0;
                for (int i = 1; i<sdr.Length; i++)
                {
                    sdr[i] += vd;//计算改正后的观测角度，并存入sdr数组中
                    sumvd += vd;
                    dataGridView1.Rows[i].Cells[2].Value = Convert.ToString(Math.Round(vd * 180 / Math.PI *3600,2))+"″";
                    //将角度改正数存入表格中
                    dataGridView1.Rows[i].Cells[3].Value = radtodms(sdr[i]);
                }            
                if (Math.Round(sumvd,8)!= Math.Round(-fd, 8))//秒保留2位对应弧度是8位
                    MessageBox.Show("角度改正数分配有误！");
                else dataGridView1.Rows[dataGridView1.RowCount - 4 ].Cells[2].Value = 
                    Convert.ToString(Math.Round(sumvd * 180 / Math.PI * 3600,2)) +"″";//将角度改正数总和存入表格中
                sum = fangweijiaotuisuan(sdr,cr);//推算改正后的坐标方位角
                if (Math.Round(cr[cr.Length - 1], 8) != Math.Round(acd, 8))
                    MessageBox.Show("坐标方位角推算有误！");
               else
                {
                    dataGridView1.Rows[dataGridView1.RowCount - 4].Cells[3].Value = radtodms(sum);
                    //将改正后的观测角度总和存入表格
                    for (int i =1; i< cr.Length-1; i++)//将改正后的坐标方位角存入表格
                       dataGridView1.Rows[i].Cells[4].Value = radtodms(cr[i]);
                }
            }       
            double[] juli = new double[dataGridView1.RowCount - 6];//新建一个数组，存放距离的原始值
            double[] xzengliang = new double[juli.Length + 1];//新建一个数组，存放x坐标增量的计算值
            double[] yzengliang = new double[juli.Length + 1];//新建一个数组，存放y坐标增量的计算值
            double sumjuli = 0;//定义一个双精度变量，存放距离总值，初值为0
            double sumxzengliang = 0;//定义一个双精度变量，存放x坐标增量总值总值，初值为0
            double sumyzengliang = 0;//定义一个双精度变量，存放y坐标增量总值总值，初值为0
            for (int i = 1; i < juli.Length; i++)//从第二行开始循环
            {
                juli[i] = Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value);//将表中获得的观测距离的原始值转化为双精度类型，并存入juli[]数组中
                sumjuli += juli[i];//计算距离总和，存入sumjuli双精度变量中
                xzengliang[i] = Math.Cos(cr[i]) * juli[i];//利用△x=cosα*l求得各个点x的坐标增量，依次存入xzengliang的数组
                sumxzengliang += xzengliang[i];//求得△x的总和，并存入sumxzengliang的双精度变量
                yzengliang[i] = Math.Sin(cr[i]) * juli[i];//利用△y=sinα*l求得各个点y的坐标增量，依次存入yzengliang的数组
                sumyzengliang += yzengliang[i];//求得△y的总和，并存入sumyzengliang的双精度变量
                dataGridView1.Rows[i].Cells[6].Value = Convert.ToString(Math.Round(xzengliang[i], 3));//将每个求得的△x的值转化为字符型，保留三位小数后，存入表格中第五列
                dataGridView1.Rows[i].Cells[7].Value = Convert.ToString(Math.Round(yzengliang[i], 3));//将每个求得的△y的值转化为字符型，保留三位小数后，存入表格中第六列
            }//循环体结束
            dataGridView1.Rows[dataGridView1.RowCount - 4].Cells[5].Value = Convert.ToString(sumjuli);//将距离总和的值转化为字符类型，存入表格的倒数第四行、第五列
            dataGridView1.Rows[dataGridView1.RowCount - 4].Cells[6].Value = Convert.ToString(Math.Round(sumxzengliang,3));//将△x总和的值转化为字符类型，存入表格的倒数第四行、第六列，并保留三位小数
            dataGridView1.Rows[dataGridView1.RowCount - 4].Cells[7].Value = Convert.ToString(Math.Round(sumyzengliang,3));//将△y总和的值转化为字符类型，存入表格的倒数第四行、第七列，并保留三位小数
            double x1 = Convert.ToDouble(dataGridView1.Rows[1].Cells[12].Value);//定义一个双精度变量x1，存放表格第二行第十三列的字符转化为双精度后的值，即第一个已知点的x坐标
            double y1 = Convert.ToDouble(dataGridView1.Rows[1].Cells[13].Value);//定义一个双精度变量y1，存放表格第二行第十四列的字符转化为双精度后的值，即第一个已知点的y坐标
            double x2 = Convert.ToDouble(dataGridView1.Rows[juli.Length + 1].Cells[12].Value);//定义一个双精度变量x2，存放表格倒数第六行第十三列的字符转化为双精度后的值，即第二个已知点的x坐标
            double y2 = Convert.ToDouble(dataGridView1.Rows[juli.Length + 1].Cells[13].Value); //定义一个双精度变量y2，存放表格倒数第六行第十四列的字符转化为双精度后的值，即第二个已知点的y坐标
            double fx = sumxzengliang - (x2 - x1);//定义一个双精度变量fx，存放x坐标增量闭合差
            double fy = sumyzengliang - (y2 - y1);//定义一个双精度变量fy，存放y坐标增量闭合差
            double fxy = Math.Sqrt(fx * fx + fy * fy);//定义一个双精度变量fxy，存放导线全长闭合差
            double k = sumjuli / fxy;//定义一个双精度变量k，存放导线全长相对闭合差，现求其倒数，填表时求反倒数
            dataGridView1.Rows[dataGridView1.RowCount - 3].Cells[7].Value = Convert.ToString(Math.Round(fx, 3));//将坐标增量闭合差fx的值转化为字符类型，存入表格的倒数第三行、第八列，并保留三位小数
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[7].Value = Convert.ToString(Math.Round(fy, 3));//将坐标增量闭合差fy的值转化为字符类型，存入表格的倒数第二行、第八列，并保留三位小数
            dataGridView1.Rows[dataGridView1.RowCount - 3].Cells[9].Value = Convert.ToString(Math.Round(fxy, 3));//将导线全长闭合差fxy的值转化为字符类型，存入表格的倒数第三行、第十一列，并保留三位小数
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[9].Value = Convert.ToString((int)k);//将导线全长相对闭合差分母取整，导线全长相对闭合差=1/整数
            double[] xgz = new double[juli.Length + 1]; //定义一个双精度数组xgz，存放x坐标增量的改正数 
            double[] ygz = new double[juli.Length + 1];//定义一个 双精度数组ygz，存放y坐标增量的改正数 
            double sumxgz = 0;//定义一个双精度变量sumxgz，存放x坐标增量改正数的总和，初值为0
            double sumygz = 0;//定义一个双精度变量sumygz，存放y坐标增量改正数的总和，初值为0
            double[] xz = new double[juli.Length + 1]; //定义一个双精度数组xz，存放改正后的x坐标增量
            double[] yz = new double[juli.Length + 1];//定义一个双精度数组xy，存放改正后的y坐标增量
            double sumxz = 0;//定义一个双精度变量sumxz，存放改正后的x坐标增量，初值为0
            double sumyz = 0;//定义一个双精度变量sumyz，存放改正后的y坐标增量，初值为0
            double[] x = new double[juli.Length + 1]; //定义一个双精度数组x，存放计算所得的各点x坐标 
            double[] y = new double[juli.Length + 1];//定义一个双精度数组y，存放计算所得的各点y坐标 
            x[0] = x1;//其中，x数组的第一个元素x[0]的值为之前从表中获取的x1的值
            y[0] = y1;//其中，y数组的第一个元素y[0]的值为之前从表中获取的y1的值
            if (k < 2000) //判断导线全长相对闭合差是否超限 
                MessageBox.Show("导线全长相对闭合差超限！");//当k<2000时，即导线全长相对中误差<1/2000时，超限
            else
            {
                for (int j = 1; j < xgz.Length + 1; j++)
                {
                    xgz[j] = -fx * juli[j] / sumjuli; //计算坐标增量改正数 
                    ygz[j] = -fy * juli[j] / sumjuli;
                    sumxgz += xgz[j]; //计算坐标增量改正数总和 
                    sumygz += ygz[j];
                    dataGridView1.Rows[j].Cells[8].Value = Convert.ToString(Math.Round(xgz[j], 4));
                    //将坐标增量改正数放入表格 
                    dataGridView1.Rows[j].Cells[9].Value = Convert.ToString(Math.Round(ygz[j], 4));
                    xz[j] = xzengliang[j] + xgz[j]; //计算改正后坐标增量 
                    yz[j] = yzengliang[j] + ygz[j];
                    sumxz += xz[j]; //计算改正后坐标增量总和 
                    sumyz += yz[j];
                    dataGridView1.Rows[j].Cells[10].Value = Convert.ToString(Math.Round(xz[j], 3));
                    //将改正后坐标增量放入表格 
                    dataGridView1.Rows[j].Cells[11].Value = Convert.ToString(Math.Round(yz[j], 3));
                }
                //C#中 1/2000 = 0 两个整数相除，可以写 成1.0/2000 保留小数位数 
                if (Math.Round(sumxgz, 4) != Math.Round(-fx, 4) || Math.Round(sumygz, 4) !=
                Math.Round(-fy, 4))
                    MessageBox.Show("坐标增量分配有误！");
                if (Math.Round(sumxz, 4) != Math.Round(x2 - x1, 4) || Math.Round(sumyz, 4) != Math.Round(y2 - y1, 4))
                    MessageBox.Show("改正后坐标增量计算有误！");
                dataGridView1.Rows[dataGridView1.RowCount - 4].Cells[8].Value =
                Convert.ToString(Math.Round(sumxgz, 3)); //将坐标增量改正数总和放入表格中 
                dataGridView1.Rows[dataGridView1.RowCount - 4].Cells[9].Value =
                Convert.ToString(Math.Round(sumygz, 3));
                dataGridView1.Rows[dataGridView1.RowCount - 4].Cells[10].Value =
                Convert.ToString(Math.Round(sumxz, 3)); //将改正后坐标增量总和放入表格中 
                dataGridView1.Rows[dataGridView1.RowCount - 4].Cells[11].Value =
                Convert.ToString(Math.Round(sumyz, 3));
                for (int j = 1; j < x.Length - 1; j++)
                {
                    x[j + 1] = x[j] + xz[j]; //计算x,y坐标 
                    y[j + 1] = y[j] + yz[j];
                    dataGridView1.Rows[j + 1].Cells[12].Value = Convert.ToString(Math.Round(x[j + 1], 3));
                    //将x,y坐标放入表格 
                    dataGridView1.Rows[j + 1].Cells[13].Value = Convert.ToString(Math.Round(y[j + 1], 3));
                }
                if (Math.Round(x[x.Length - 1] + xz[xz.Length - 1], 3) != Math.Round(x2, 3) || Math.Round
                    (y[y.Length - 1] + yz[yz.Length - 1], 3) != Math.Round(y2, 3))
                    MessageBox.Show("坐标计算有误！");
            } 
        }
        private void button2_Click(object sender, EventArgs e)//“关闭”按钮程序
        {
            Application.Exit();//关闭窗口
        }

        private void excel文件ToolStripMenuItem_Click(object sender, EventArgs e)//Excel文件导入程序
        {
            dataGridView1.DataSource = null;//清除数据源，清空表格里的所有数据
            dataGridView1.Rows.Clear();//清空行
            dataGridView1.Columns.Clear();//清空列
            OpenFileDialog file = new OpenFileDialog();//声明，打开文件对话框file
            file.Filter = "Excel文件|*.xls|Excel文件|*.xlsx";//文件过滤器，只显示Excel文件，限制打开的文件后缀只能为xls或xlsx
            if (file.ShowDialog() == DialogResult.OK)//如果文件正常打开时
            {
                string fname = file.FileName;//定义一个字符串fname，获取打开的文件名称
                string strSource = "provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source="
                    + fname + ";Extended Properties='Excel 8.0;    HDR=Yes;IMEX=1'";//准备文件来源信息
                OleDbConnection conn = new OleDbConnection(strSource);//Excel文件源放入conn中
                string sqlstring = "SELECT * FROM [Sheet1$]";//准备选择表中的sheet1
                OleDbDataAdapter adapter = new OleDbDataAdapter(sqlstring, conn);//声明数据配置器adapter
                DataSet da = new DataSet();//声明数据集da
                adapter.Fill(da);//使用adapter填充方法
                dataGridView1.DataSource = da.Tables[0];//将da.Tables[0]作为dataGridView1的数据源
            }
            else
                return;
        }
        private void txt文件ToolStripMenuItem_Click(object sender, EventArgs e)//txt文件导入程序
        {
            //txt输入代码
            dataGridView1.DataSource = null;//清除数据源
            dataGridView1.Rows.Clear();//清空数据表格的行
            dataGridView1.Columns.Clear();//清空数据表格的列
            OpenFileDialog file = new OpenFileDialog();//声明 打开文件类file
            file.Filter = "文本文件|*.txt";//文件过滤器，只显示txt文件
            if (file.ShowDialog() == DialogResult.OK)//如果文件正常打开
            {
                StreamReader sr = new StreamReader(file.FileName, System.Text.Encoding.Default);
                //声明文本读取流，并以文本编码格式读取
                textBox1.Text = sr.ReadToEnd();//将sr中的内容全部放到textBox中
                sr.Close();
            }
            else
                return;

            //textBox1.Text 存入数组，然后存入dataGridView1
            string[] str = textBox1.Text.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            //将textBox1.Text中按行分割，并放在一维字符串数组中
            string[][] k = new string[str.Length][];//定义字符串交错数组，行数与str的长度相同
            for (int i = 0; i < str.Length; i++)
                k[i] = str[i].Split(',');//将str中对应的字符串以逗号（英文状态）分割，并放在k中
            dataGridView1.RowCount = k.Length;//定义表格控件的行数，与str长度相同
            dataGridView1.ColumnCount = k[0].Length;//定义表格列数，与k[0]长度相同
            for (int i = 0; i < k[0].Length; i++)
                dataGridView1.Columns[i].HeaderText = k[0][i];//将k中第0行元素放入表格的表头
            for (int i = 1; i < k.Length; i++)//将k中数据元素放入对应的表格中
            {
                for (int j = 0; j < k[i].Length; j++)
                    dataGridView1.Rows[i - 1].Cells[j].Value = k[i][j];
            }
        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
