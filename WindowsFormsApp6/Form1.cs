using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using ORiN2.interop.CAO;


namespace WindowsFormsApp6
{
    public partial class Form1 : Form
    {
        private CaoEngine caoEng;
        private CaoWorkspaces caoWss;
        private CaoWorkspace caoWs;

        // Robot objects
        private CaoController caoCtrl;  
        private CaoRobot caoRobot;      
       
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            caoEng = new CaoEngine();
            caoWss = caoEng.Workspaces;
            caoWs = caoWss.Item(0);

        }

        private void button2_Click(object sender, EventArgs e)
        {
           
            caoCtrl = caoWs.AddController("Robot11", "CaoProv.DENSO.RC8", "", "Server=192.168.0.1");
            caoRobot = caoCtrl.AddRobot("lph", "");
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            caoRobot.Execute("Motor", new object[] { 1, 0 });
        }

        private void button4_Click(object sender, EventArgs e)
        {
            caoRobot.Execute("TakeArm", new object[] { 0, 1 });           
            caoRobot.Move(1, "P0", "");
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
               
           
                // Delete CaoEngine
                if (caoWs != null)
                {
                    Marshal.ReleaseComObject(caoWs);
                    caoWs = null;
                }

                if (caoWss != null)
                {
                    Marshal.ReleaseComObject(caoWss);
                    caoWss = null;
                }

                if (caoEng != null)
                {
                    Marshal.ReleaseComObject(caoEng);
                    caoEng = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            caoRobot.Execute("Motor", new object[] { 0, 0 });
            caoRobot.Execute("GiveArm", "");
            if (caoRobot != null)
            {
                if (caoCtrl != null)
                {

                    CaoRobots robots = caoCtrl.Robots;
                    robots.Remove(caoRobot.Name);

                    Marshal.ReleaseComObject(caoRobot);
                    caoRobot = null;

                    Marshal.ReleaseComObject(robots);
                    robots = null;
                }
            }
            if (caoCtrl != null)
            {
                CaoControllers ctrls = caoWs.Controllers;
                ctrls.Remove(caoCtrl.Name);

                Marshal.ReleaseComObject(caoCtrl);
                caoCtrl = null;

                Marshal.ReleaseComObject(ctrls);
                ctrls = null;
            }

            if (caoWs != null)
            {
                Marshal.ReleaseComObject(caoWs);
                caoWs = null;
            }

            if (caoWss != null)
            {
                Marshal.ReleaseComObject(caoWss);
                caoWss = null;
            }

            if (caoEng != null)
            {
                Marshal.ReleaseComObject(caoEng);
                caoEng = null;

            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            caoRobot.Execute("TakeArm", new object[] { 0, 1 });
            ///caoRobot.Move(1, "J(44, -6.6, 114, 213", "");
            caoRobot.Move(1, "P1", "");
        }
    }
}
