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
        string errorTxt2 = "Cannot contain letters, symbols or be emptied";

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
        double spentMin = 0;


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
            TextBox tb = sender as TextBox;
            if (tb == null)
            {
                Debug.Assert(false, "テキストボックス認識失敗");
                return;
            }


            double ret = 0;
            if (double.TryParse(tb.Text, out ret))
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
            TextBox tb = sender as TextBox;
            if (tb == null)
            {
                Debug.Assert(false, "テキストボックス認識失敗");
                return;
            }


            double ret = 0;
            if (double.TryParse(tb.Text, out ret))
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

            TextBox tb = sender as TextBox;
            if (tb == null)
            {
                Debug.Assert(false, "テキストボックス認識失敗");
                return;
            }


            double ret = 0;
            if (double.TryParse(tb.Text, out ret))
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
            TextBox tb = sender as TextBox;
            if (tb == null)
            {
                Debug.Assert(false, "テキストボックス認識失敗");
                return;
            }


            double ret = 0;
            if (double.TryParse(tb.Text, out ret))
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
                Debug.WriteLine("Stop operation.");
                programEnd();
            }
            else
            {

                //search first for the end

                
                string sb = "";
                bool checkSkiptime = false, checkTime = true;

                int count = 99;
                for (; count >= 0; count--)
                {
                    //if end value has been found

                    if (mins[count].Text != "0")
                    {
                        checkTime = false;
                        for (; count >= 0; count--)
                        {
                            // if a zero value in between steps was found

                            if (mins[count].Text == "0")
                            {
                                checkSkiptime = true;

                                //stores step number

                                sb = (count + 1) + "  " + sb;
                            }
                        }
                        break;
                    }
                }

                //if their are skipped time

                if (checkSkiptime)
                {
                    MessageBox.Show("下記のステップ入力がスキップされています。修正して下さい。" + Environment.NewLine + sb,
                        "ステップ入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    programEnd();
                }
                else 
                { 
                    SetBtn.Text = "設定停止";

                    //if all the time were left empty (zero value)

                    if (checkTime)
                    {
                        MessageBox.Show("少なくとも１つのステップを入力してください。", 
                            "ステップ入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        programEnd();
                    }
                    else
                    {
                        //if passed all validations

                        Debug.WriteLine("Setting a new value.");
                        mainTable.Enabled = false;
                        int[] errs = null;

                        //stores first step value to OPC

                        object[] writeVals = new object[] {
                             Convert.ToDouble(tmps[0].Text),
                             Convert.ToDouble(hmds[0].Text),
                             Convert.ToDouble(suns[0].Text),                   
                         };
                        if (opc.Write(setRegs, writeVals, out errs) == false)
                        {
                            Debug.Assert(false, "Opc.Write failed.");
                            programEnd();
                        }
                        else
                        {
                            timer.Enabled = true;

                            //this will apply background to the step to be executed

                            foreach (Label l in stps)
                            {
                                l.BackColor = Color.White;
                            }
                            stps[0].BackColor = Color.LightGreen;
                        }
                    }
                }
            }
        }
        private void programEnd()
        {
            Debug.WriteLine("Operation finished!");
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
            StepProcess();
            Read_Write();
        }
        private void StepProcess()
        {
            //this finds for the next step to execute

            if (spentMin <= 0)
            {
                for (; curStp < MAX_STEP; curStp++)
                {
                    double min = 0;
                    if (double.TryParse(mins[curStp - 1].Text, out min))
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

                //this will apply background to the step to be executed

                foreach (Label l in stps)
                {
                    l.BackColor = Color.White;
                }
                stps[curStp - 1].BackColor = Color.LightGreen;
            }
        }

        private double GetTarget2StepValue(double currentValue, double currentStepValue, double currentStepTime, double nextStepValue)
        {
            double answer = currentValue + ((nextStepValue - currentStepValue) / currentStepTime);
            bool check = Set_Bool_Method(currentStepValue, currentValue, nextStepValue);
            if (check)
            {
                return nextStepValue;
            }
            else { 
                return answer;
            }
        }
        private void Read_Write()
        {
            spentMin++;
            object[] readVals = null;
            short[] qlty = null;
            int[] errs = null;
            FILETIME[] ft = null;

            if (opc.Read(curRegs, out readVals, out qlty, out ft, out errs) == true)
            {
                // スケール変換を入れる場合はここに入れる
                double curTmp = Convert.ToDouble(readVals[0]);
                double curHmd = Convert.ToDouble(readVals[1]);
                double curSun = Convert.ToDouble(readVals[2]);
                
                double cstpTmp = Convert.ToDouble(tmps[curStp - 1].Text);
                double cstpHmd = Convert.ToDouble(hmds[curStp - 1].Text);
                double cstpSun = Convert.ToDouble(suns[curStp - 1].Text);
                double cstpTim = Convert.ToDouble(mins[curStp - 1].Text);
                

                double tgtTmp = 0;
                double tgtHmd = 0;
                double tgtSun = 0;
                
                if (curStp < MAX_STEP && Convert.ToDouble(mins[curStp].Text) > 0)
                {
                    double nstpTmp = Math.Round(Convert.ToDouble(tmps[curStp].Text), 2);
                    double nstpHmd = Math.Round(Convert.ToDouble(hmds[curStp].Text), 2);
                    double nstpSun = Math.Round(Convert.ToDouble(suns[curStp].Text), 2);                    
                    tgtTmp = Math.Round(GetTarget2StepValue(curTmp, cstpTmp, cstpTim, nstpTmp), 2);
                    tgtHmd = Math.Round(GetTarget2StepValue(curHmd, cstpHmd, cstpTim, nstpHmd), 2);
                    tgtSun = Math.Round(GetTarget2StepValue(curSun, cstpSun, cstpTim, nstpSun), 2);
                    Debug.Write(spentMin + "       ");
                    if (spentMin == Convert.ToDouble(mins[curStp - 1].Text))
                    {
                        tgtTmp = nstpTmp;
                        tgtHmd = nstpHmd;
                        tgtSun = nstpSun;
                    }
                    if (tgtTmp == nstpTmp && tgtHmd == nstpHmd && tgtSun == nstpSun && (spentMin == Convert.ToDouble(mins[curStp - 1].Text)))
                    {
                        Debug.WriteLine("\n");
                        curStp++;
                        spentMin = 0;
                        foreach (Label l in stps)
                        {
                            l.BackColor = Color.White;
                        }
                        stps[curStp - 1].BackColor = Color.LightGreen;                        
                    }
                }
                else
                {
                    tgtTmp = cstpTmp;
                    tgtHmd = cstpHmd;
                    tgtSun = cstpSun;
                    Debug.Write(spentMin + "       ");
                    if (mins[curStp - 1].Text == spentMin.ToString())
                    {
                        Debug.WriteLine("\n");
                        programEnd();
                    }
                }
                object[] writeVals = new object[] { tgtTmp, tgtHmd, tgtSun, };
                if (opc.Write(setRegs, writeVals, out errs) == false)
                {
                    Debug.Assert(false, "Opc.Write failed.");
                    programEnd();
                }
            }
            else
            {
                Debug.Assert(false, "Opc.Read failed.");
                programEnd();
            }
        }
        private bool Set_Bool_Method(double currentStepValue,double currentValue, double nextStepValue){
            bool true_false;
            if (currentStepValue < nextStepValue)
            {
                true_false = (currentValue >= nextStepValue);
            }
            else if (currentStepValue > nextStepValue)
            {
                true_false = (currentValue <= nextStepValue);
            }
            else
            {
                true_false =  true;
            }
            return true_false;
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
