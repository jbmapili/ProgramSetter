using DxpSimpleAPI;
using Microsoft.VisualBasic.FileIO;
using OpcRcw.Da;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
[assembly: InternalsVisibleTo("ProgPatternUnitTest")]


namespace 運転パターン設定
{
    public partial class ProgramPattern : Form
    {
        internal static readonly int MAX_STEP = 100;

        DxpSimpleClass opc = new DxpSimpleClass();
        string devPrefix = Properties.Settings.Default.DevicePrefix;
        string tmpCurReg = Properties.Settings.Default.TmpCurReg;
        string hmdCurReg = Properties.Settings.Default.HmdCurReg;
        string sunCurReg = Properties.Settings.Default.SunCurReg;
        string tmpSetReg = Properties.Settings.Default.TmpSetReg;
        string hmdSetReg = Properties.Settings.Default.HmdSetReg;
        string sunSetReg = Properties.Settings.Default.SunSetReg;
        string[] curRegs = null;
        string[] setRegs = null;

        internal List<Label> stps = new List<Label>(MAX_STEP);
        internal List<TextBox> tmps = new List<TextBox>(MAX_STEP);
        internal List<TextBox> hmds = new List<TextBox>(MAX_STEP);
        internal List<TextBox> suns = new List<TextBox>(MAX_STEP);
        internal List<TextBox> mins = new List<TextBox>(MAX_STEP);

        int curStp = 1;
        int spentMin = 0;


        public ProgramPattern()
        {
            InitializeComponent();

            curRegs = new string[] {
                devPrefix + tmpCurReg,
                devPrefix + hmdCurReg,
                devPrefix + sunCurReg,
            };
            setRegs = new string[] {
                devPrefix + tmpSetReg,
                devPrefix + hmdSetReg,
                devPrefix + sunSetReg,                
            };

            mainTable.RowCount = MAX_STEP;
            mainTable.Height = mainTable.Height * MAX_STEP;
            for (int i = 0; i < MAX_STEP; i++)
            {
                Label step = new Label();
                step.Text = (i + 1).ToString();
                step.TextAlign = ContentAlignment.MiddleRight;
                step.BackColor = Color.White;
                mainTable.Controls.Add(step, 0, i);
                stps.Add(step);

                TextBox tmp = new TextBox();
                mainTable.Controls.Add(tmp, 1, i);
                tmp.TextAlign = HorizontalAlignment.Right;
                tmp.Text = "0";
                tmp.Validating += tmpTxt_Validating;
                tmps.Add(tmp);

                TextBox hmd = new TextBox();
                mainTable.Controls.Add(hmd, 2, i);
                hmd.TextAlign = HorizontalAlignment.Right;
                hmd.Text = "30";
                hmd.Validating += hmdTxt_Validating;
                hmds.Add(hmd);

                TextBox sun = new TextBox();
                mainTable.Controls.Add(sun, 3, i);
                sun.TextAlign = HorizontalAlignment.Right;
                sun.Text = "200";
                sun.Validating += sunTxt_Validating;
                suns.Add(sun);

                TextBox min = new TextBox();
                mainTable.Controls.Add(min, 4, i);
                min.TextAlign = HorizontalAlignment.Right;
                min.Text = "0";
                min.Validating += minTxt_Validating;
                errorProvider.SetIconAlignment(min, ErrorIconAlignment.MiddleLeft);
                mins.Add(min);

            }
        }

        /// <summary>
        /// Temperature Validation
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tmpTxt_Validating(object sender, CancelEventArgs e)
        {
            string errorTxt = "-30～50℃の値を入力してください。";
            string errorTxt2 = "Cannot contain letters or symbols";
            TextBox tb = sender as TextBox;
            if (tb == null)
            {
                Debug.Assert(false, "テキストボックス認識失敗");
                return;
            }


            int ret = 0;
            if (int.TryParse(tb.Text, out ret))
            {
                if (-30 <= ret && ret <= 50)
                {
                    // Success
                    errorProvider.SetError(tb, null);
                    return;
                }
                // Error
                tb.Text = "0";
                errorProvider.SetError(tb, errorTxt);
            }
            else
            {
                tb.Text = "0";
                errorProvider.SetError(tb, errorTxt2);
            }
        }

        /// <summary>
        /// Humidity Validation
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void hmdTxt_Validating(object sender, CancelEventArgs e)
        {
            string errorTxt = "30～90%の値を入力してください。";
            string errorTxt2 = "Cannot contain letters or symbols";
            TextBox tb = sender as TextBox;
            if (tb == null)
            {
                Debug.Assert(false, "テキストボックス認識失敗");
                return;
            }


            int ret = 0;
            if (int.TryParse(tb.Text, out ret))
            {
                if (30 <= ret && ret <= 90)
                {
                    // Success
                    errorProvider.SetError(tb, null);
                    return;
                }
                // Error
                tb.Text = "30";
                errorProvider.SetError(tb, errorTxt);
            }
            else
            {
                tb.Text = "30";
                errorProvider.SetError(tb, errorTxt2);
            }
        }

        /// <summary>
        /// Solar Radiation Validation
        /// </summary>
        /// <param name="sender">メッセージ元</param>
        /// <param name="e">イベント</param>
        private void sunTxt_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            string errorTxt = "200～1200W/m2の値を入力してください。";
            string errorTxt2 = "Cannot contain letters or symbols";
            TextBox tb = sender as TextBox;
            if (tb == null)
            {
                Debug.Assert(false, "テキストボックス認識失敗");
                return;
            }


            int ret = 0;
            if (int.TryParse(tb.Text, out ret))
            {
                if (200 <= ret && ret <= 1200)
                {
                    // Success
                    errorProvider.SetError(tb, null);
                    return;
                }
                // Error
                tb.Text = "200";
                errorProvider.SetError(tb, errorTxt);
            }
            
            else
            {
                tb.Text = "200";
                errorProvider.SetError(tb, errorTxt2);
            }
        }

        /// <summary>
        /// Time(min) Validation
        /// </summary>
        /// <param name="sender">メッセージ元</param>
        /// <param name="e">イベント</param>
        private void minTxt_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            string errorTxt = "正しい時間(分)を入力してください。";
            string errorTxt2 = "Cannot contain letters or symbols";
            TextBox tb = sender as TextBox;
            if (tb == null)
            {
                Debug.Assert(false, "テキストボックス認識失敗");
                return;
            }


            int ret = 0;
            if (int.TryParse(tb.Text, out ret))
            {
                if (ret >= 0)
                {
                    // Success
                    errorProvider.SetError(tb, null);
                    return;
                }
                
            // Error
            tb.Text = "0";
            errorProvider.SetError(tb, errorTxt);
            }

            else
            {
                tb.Text = "0";
                errorProvider.SetError(tb, errorTxt2);
            }
        }



        private void SetBtn_Click(object sender, EventArgs e)
        {
            if (timer.Enabled)
            {
                programEnd();
            }
            else
            {
                int[] errs = null;
                object[] writeVals = new object[] {
                    Convert.ToInt32(tmps[0].Text),
                    Convert.ToInt32(hmds[0].Text),
                    Convert.ToInt32(suns[0].Text),                   
                };
                if (opc.Write(setRegs, writeVals, out errs) == false)
                {
                    Debug.Assert(false, "Opc.Write failed.");
                    return;
                }

                curStp = 1;
                //stps[curStp - 1].BackColor = Color.LimeGreen;
                StepProcess();
                SetBtn.Text = "設定停止";
                mainTable.Enabled = false;
                
            }

            timer.Enabled = !timer.Enabled;

        }

        private void programEnd()
        {
            timer.Enabled = false;
            curStp = 1;
            SetBtn.Text = "設定開始";
            foreach (Label l in stps)
            {
                l.BackColor = Color.White;
            }
            mainTable.Enabled = true;
        }

        private void timer_Tick(object sender, EventArgs e)
        {
            MessageBox.Show("Call method already");
            StepProcess();
            Read_Write();
        }

        private void StepProcess()
        {
            /*
             kung spentMin<0 execute nya ung if.
             Tapos, check nya kung ung current step ay mababa sa 100(MAX).
             Tapos, kinukuha nya ung value ng Time
             Pag nakuha ung time, lalabas sa forloop pag hindi mas mataas sa 0 ung min.
             * Tapos, Compare nya kung 100 na ung step tapos tatawagin ung program end
             * Tapos, spentMin mag aadd ng 0.
             */
            if (spentMin <= 0)
            {
                //MessageBox.Show("If SpentMin: " + spentMin);
                for (; curStp < MAX_STEP; curStp++)
                {
                    int min = 0;
                    if (int.TryParse(mins[curStp - 1].Text, out min))
                    {
                        if (min > 0)
                        {
                            break;  // Enabled step found
                        }
                    }
                    else
                    {
                        Debug.Assert(false, "時間(分)の数値変換失敗");
                    }
                }
                if (curStp >= MAX_STEP) // Out of Step
                {
                    programEnd();
                    return;     // End of Cycle
                }
                spentMin++;
                //MessageBox.Show("If SpentMin: " + spentMin);
            }
                /*
                 *Kung ung spentMin hindi 0
                 *Tapos min=0
                 *Tapos, check nya ung time
                 *Kung ung spentMin ay 0 or mataas, spentMin=0 Tapos curStp++
                 * */
            else
            {
                int min = 0;
                if (int.TryParse(mins[curStp - 1].Text, out min))
                {
                    if (spentMin >= min)
                    {
                        //MessageBox.Show("Else SpentMin: " + spentMin);
                        spentMin=0;
                        curStp++;
                    }
                }
            }
            foreach (Label l in stps)
            {
                l.BackColor = Color.White;
            }
            stps[curStp - 1].BackColor = Color.LightGreen;

            //MessageBox.Show(" stps[curStp - 1]: " + stps[curStp - 1]);

        }

        private int GetTarget2StepValue(int currentValue, int currentStepValue, int currentStepTime, int nextStepValue)
        {
           // return currentValue + ((nextStepValue - currentValue) - (currentStepValue - currentValue)) / (currentStepTime - spentMin);
            return currentValue + ((nextStepValue - currentValue) / currentStepTime);
        }
        private void Read_Write()
        {
            object[] readVals = null;
            short[] qlty = null;
            int[] errs = null;
            FILETIME[] ft = null;

            if (opc.Read(curRegs, out readVals, out qlty, out ft, out errs) == true)
            {
                // スケール変換を入れる場合はここに入れる
                /*
                 Kinukuha nya ung current Temp Hum Sun Min
                 * 
                 */
                int curTmp = Convert.ToInt32(readVals[0]);
                int curHmd = Convert.ToInt32(readVals[1]);
                int curSun = Convert.ToInt32(readVals[2]);
                // MessageBox.Show(" curTmp: " + curTmp.ToString() + "curHmd: " + curHmd.ToString() + "curSun: " +curSun.ToString());

                int cstpTmp = Convert.ToInt32(tmps[curStp - 1].Text);
                int cstpHmd = Convert.ToInt32(hmds[curStp - 1].Text);
                int cstpSun = Convert.ToInt32(suns[curStp - 1].Text);

                int cstpTim = Convert.ToInt32(mins[curStp - 1].Text);
                // MessageBox.Show(" cstpTmp: " + cstpTmp.ToString() + "cstpHmd: " + cstpHmd.ToString() + "cstpSun: " + cstpSun.ToString() + "cstpTim: " + cstpTim.ToString());


                int tgtTmp = 0;
                int tgtHmd = 0;
                int tgtSun = 0;
                if (curStp < MAX_STEP && Convert.ToInt32(mins[curStp].Text) > 0)
                {
                    int nstpTmp = Convert.ToInt32(tmps[curStp].Text);
                    int nstpHmd = Convert.ToInt32(hmds[curStp].Text);
                    int nstpSun = Convert.ToInt32(suns[curStp].Text);
                    //MessageBox.Show(" nstpTmp: " + nstpTmp.ToString() + "nstpHmd: " + nstpHmd.ToString() + "nstpSun: " + nstpSun.ToString());

                    tgtTmp = GetTarget2StepValue(curTmp, cstpTmp, cstpTim, nstpTmp);
                    tgtHmd = GetTarget2StepValue(curHmd, cstpHmd, cstpTim, nstpHmd);
                    tgtSun = GetTarget2StepValue(curSun, cstpSun, cstpTim, nstpSun);
                    //MessageBox.Show("If tgtTmp: " + tgtTmp.ToString() + "tgtHmd: " + tgtHmd.ToString() + "tgtSun: " + tgtSun.ToString() + "cstpTim: " + cstpTim.ToString());
                    if (tgtTmp > nstpTmp && tgtHmd > nstpHmd && tgtSun > nstpSun)
                    {
                        foreach (Label l in stps)
                        {
                            l.BackColor = Color.White;
                        }
                        MessageBox.Show("Read step");
                        curStp++;
                        stps[curStp - 1].BackColor = Color.LightGreen;
                    }
                }
                else
                {
                    tgtTmp = cstpTmp;
                    tgtHmd = cstpHmd;
                    tgtSun = cstpSun;
                }

                object[] writeVals = new object[] {
                    tgtTmp, tgtHmd, tgtSun,
                };

                if (opc.Write(setRegs, writeVals, out errs) == false)
                {
                    Debug.Assert(false, "Opc.Write failed.");
                    return;
                }
            }
            else
            {
                Debug.Assert(false, "Opc.Read failed.");
                return;
            }
        }
        private void ImportBtn_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                PatternNameTxt.Text = Path.GetFileNameWithoutExtension(openFileDialog.FileName);
                using (TextFieldParser tfp = new TextFieldParser(openFileDialog.FileName))
                {
                    tfp.Delimiters = new string[] { "," };

                    List<string[]> csvLines = new List<string[]>(100);
                    while (!tfp.EndOfData)
                    {
                        //フィールドを読み込む
                        string[] fields = tfp.ReadFields();
                        csvLines.Add(fields);
                    }

                    if (csvLines.Count != 100)
                    {
                        MessageBox.Show("インポート・ファイルの内容が正しくありません。", 
                                "インポート・エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    for (int i=0; i<MAX_STEP; i++) 
                    {
                        if (csvLines[i].Length != 5)
                        {
                            MessageBox.Show("インポート・ファイルの内容が正しくありません。",
                                    "インポート・エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        tmps[i].Text = csvLines[i][1];
                        hmds[i].Text = csvLines[i][2];
                        suns[i].Text = csvLines[i][3];
                        mins[i].Text = csvLines[i][4];
                    }
                }
            }
        }

        private void ExportBtn_Click(object sender, EventArgs e)
        {
            saveFileDialog.FileName = PatternNameTxt.Text;
            if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < MAX_STEP; i++)
                {
                    sb.AppendLine(string.Format("{0},{1},{2},{3},{4}", i+1, tmps[i].Text, hmds[i].Text, suns[i].Text, mins[i].Text));
                }
                File.WriteAllText(saveFileDialog.FileName, sb.ToString());
            }
        }

        private void ProgramPattern_Load(object sender, EventArgs e)
        {
            if (opc.Connect(Properties.Settings.Default.NodeName, 
                            Properties.Settings.Default.ServerName) == false)
            {
                MessageBox.Show("PLC通信ソフト（OPCサーバー）との接続に失敗しました。",
                    "通信エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
        }

        private void ProgramPattern_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (opc.Disconnect() == false)
            {
                Debug.Assert(false, "Failed to disconnect to OPC server");
            }
        }

        private void mainTable_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
