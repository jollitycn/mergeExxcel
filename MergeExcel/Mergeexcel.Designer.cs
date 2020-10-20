using CCWin.SkinControl;
using CJBasic.CJBasic.Helpers;
using JGNet.Common.Components;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace MergeExcel
{
    //}
    partial class Mergeexcel
    {
       
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码
        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Mergeexcel));
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.skinToolStrip1 = new CCWin.SkinControl.SkinToolStrip();
            this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.skinPanel1 = new CCWin.SkinControl.SkinPanel();
            this.dataMergedView1 = new CJBasic.Widget.DataMergedView();
            this.skinPanel2 = new CCWin.SkinControl.SkinPanel();
            this.label2 = new System.Windows.Forms.Label();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.button1 = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnOpen = new System.Windows.Forms.Button();
            this.label1 = new CCWin.SkinControl.SkinLabel();
            this.btnOpreat = new System.Windows.Forms.Button();
            this.openPath2 = new System.Windows.Forms.TextBox();
            this.openPath = new System.Windows.Forms.TextBox();
            this.btnOpen2 = new System.Windows.Forms.Button();
            this.savePath = new System.Windows.Forms.TextBox();
            this.skinToolStrip1.SuspendLayout();
            this.skinPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataMergedView1)).BeginInit();
            this.skinPanel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.saveFileDialog1_FileOk);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // skinToolStrip1
            // 
            this.skinToolStrip1.Arrow = System.Drawing.Color.Black;
            this.skinToolStrip1.AutoSize = false;
            this.skinToolStrip1.Back = System.Drawing.Color.White;
            this.skinToolStrip1.BackColor = System.Drawing.Color.Transparent;
            this.skinToolStrip1.BackRadius = 4;
            this.skinToolStrip1.BackRectangle = new System.Drawing.Rectangle(10, 10, 10, 10);
            this.skinToolStrip1.Base = System.Drawing.Color.Transparent;
            this.skinToolStrip1.BaseFore = System.Drawing.Color.Black;
            this.skinToolStrip1.BaseForeAnamorphosis = false;
            this.skinToolStrip1.BaseForeAnamorphosisBorder = 4;
            this.skinToolStrip1.BaseForeAnamorphosisColor = System.Drawing.Color.White;
            this.skinToolStrip1.BaseForeOffset = new System.Drawing.Point(0, 0);
            this.skinToolStrip1.BaseHoverFore = System.Drawing.Color.White;
            this.skinToolStrip1.BaseItemAnamorphosis = true;
            this.skinToolStrip1.BaseItemBorder = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(148)))), ((int)(((byte)(212)))));
            this.skinToolStrip1.BaseItemBorderShow = true;
            this.skinToolStrip1.BaseItemDown = ((System.Drawing.Image)(resources.GetObject("skinToolStrip1.BaseItemDown")));
            this.skinToolStrip1.BaseItemHover = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(148)))), ((int)(((byte)(212)))));
            this.skinToolStrip1.BaseItemMouse = ((System.Drawing.Image)(resources.GetObject("skinToolStrip1.BaseItemMouse")));
            this.skinToolStrip1.BaseItemNorml = null;
            this.skinToolStrip1.BaseItemPressed = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(148)))), ((int)(((byte)(212)))));
            this.skinToolStrip1.BaseItemRadius = 4;
            this.skinToolStrip1.BaseItemRadiusStyle = CCWin.SkinClass.RoundStyle.All;
            this.skinToolStrip1.BaseItemSplitter = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(148)))), ((int)(((byte)(212)))));
            this.skinToolStrip1.BindTabControl = null;
            this.skinToolStrip1.DropDownImageSeparator = System.Drawing.Color.FromArgb(((int)(((byte)(197)))), ((int)(((byte)(197)))), ((int)(((byte)(197)))));
            this.skinToolStrip1.Fore = System.Drawing.Color.Black;
            this.skinToolStrip1.GripMargin = new System.Windows.Forms.Padding(2, 2, 4, 2);
            this.skinToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.skinToolStrip1.HoverFore = System.Drawing.Color.White;
            this.skinToolStrip1.ImageScalingSize = new System.Drawing.Size(32, 32);
            this.skinToolStrip1.ItemAnamorphosis = true;
            this.skinToolStrip1.ItemBorder = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(148)))), ((int)(((byte)(212)))));
            this.skinToolStrip1.ItemBorderShow = true;
            this.skinToolStrip1.ItemHover = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(148)))), ((int)(((byte)(212)))));
            this.skinToolStrip1.ItemPressed = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(148)))), ((int)(((byte)(212)))));
            this.skinToolStrip1.ItemRadius = 4;
            this.skinToolStrip1.ItemRadiusStyle = CCWin.SkinClass.RoundStyle.All;
            this.skinToolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripButton1});
            this.skinToolStrip1.Location = new System.Drawing.Point(4, 28);
            this.skinToolStrip1.Name = "skinToolStrip1";
            this.skinToolStrip1.RadiusStyle = CCWin.SkinClass.RoundStyle.All;
            this.skinToolStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.System;
            this.skinToolStrip1.ShowItemToolTips = false;
            this.skinToolStrip1.Size = new System.Drawing.Size(739, 22);
            this.skinToolStrip1.SkinAllColor = true;
            this.skinToolStrip1.TabIndex = 14;
            this.skinToolStrip1.Text = "skinToolStrip1";
            this.skinToolStrip1.TitleAnamorphosis = true;
            this.skinToolStrip1.TitleColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(228)))), ((int)(((byte)(236)))));
            this.skinToolStrip1.TitleRadius = 4;
            this.skinToolStrip1.TitleRadiusStyle = CCWin.SkinClass.RoundStyle.All;
            // 
            // toolStripButton1
            // 
            this.toolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton1.Image")));
            this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton1.Name = "toolStripButton1";
            this.toolStripButton1.Size = new System.Drawing.Size(36, 19);
            this.toolStripButton1.Text = "toolStripButton1";
            // 
            // skinPanel1
            // 
            this.skinPanel1.AutoSize = true;
            this.skinPanel1.BackColor = System.Drawing.Color.Transparent;
            this.skinPanel1.Controls.Add(this.skinPanel2);
            this.skinPanel1.Controls.Add(this.dataMergedView1);
            this.skinPanel1.ControlState = CCWin.SkinClass.ControlState.Normal;
            this.skinPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.skinPanel1.DownBack = null;
            this.skinPanel1.Location = new System.Drawing.Point(4, 50);
            this.skinPanel1.MouseBack = null;
            this.skinPanel1.Name = "skinPanel1";
            this.skinPanel1.NormlBack = null;
            this.skinPanel1.Size = new System.Drawing.Size(739, 429);
            this.skinPanel1.TabIndex = 15;
            // 
            // dataMergedView1
            // 
            this.dataMergedView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataMergedView1.Location = new System.Drawing.Point(0, 404);
            this.dataMergedView1.MergeColumnHeaderBackColor = System.Drawing.SystemColors.Control;
            this.dataMergedView1.MergeColumnNames = ((System.Collections.Generic.List<string>)(resources.GetObject("dataMergedView1.MergeColumnNames")));
            this.dataMergedView1.Name = "dataMergedView1";
            this.dataMergedView1.RowTemplate.Height = 23;
            this.dataMergedView1.Size = new System.Drawing.Size(115, 25);
            this.dataMergedView1.TabIndex = 11;
            this.dataMergedView1.Visible = false;
            // 
            // skinPanel2
            // 
            this.skinPanel2.AutoSize = true;
            this.skinPanel2.BackColor = System.Drawing.Color.Transparent;
            this.skinPanel2.Controls.Add(this.label2);
            this.skinPanel2.Controls.Add(this.dateTimePicker1);
            this.skinPanel2.Controls.Add(this.button1);
            this.skinPanel2.Controls.Add(this.btnSave);
            this.skinPanel2.Controls.Add(this.btnOpen);
            this.skinPanel2.Controls.Add(this.label1);
            this.skinPanel2.Controls.Add(this.btnOpreat);
            this.skinPanel2.Controls.Add(this.openPath2);
            this.skinPanel2.Controls.Add(this.openPath);
            this.skinPanel2.Controls.Add(this.btnOpen2);
            this.skinPanel2.Controls.Add(this.savePath);
            this.skinPanel2.ControlState = CCWin.SkinClass.ControlState.Normal;
            this.skinPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.skinPanel2.DownBack = null;
            this.skinPanel2.Location = new System.Drawing.Point(0, 0);
            this.skinPanel2.MouseBack = null;
            this.skinPanel2.Name = "skinPanel2";
            this.skinPanel2.NormlBack = null;
            this.skinPanel2.Size = new System.Drawing.Size(739, 429);
            this.skinPanel2.TabIndex = 12;
            // 
            // label2
            // 
            this.label2.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(154, 96);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 12);
            this.label2.TabIndex = 13;
            this.label2.Text = "选择月份";
            this.label2.Visible = false;
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.dateTimePicker1.CustomFormat = "yyyy-MM";
            this.dateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePicker1.Location = new System.Drawing.Point(213, 96);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(401, 21);
            this.dateTimePicker1.TabIndex = 12;
            this.dateTimePicker1.Visible = false;
            // 
            // button1
            // 
            this.button1.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.button1.Location = new System.Drawing.Point(619, 173);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(76, 25);
            this.button1.TabIndex = 1;
            this.button1.Text = "打开文件";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnSave
            // 
            this.btnSave.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnSave.Location = new System.Drawing.Point(131, 210);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(76, 25);
            this.btnSave.TabIndex = 1;
            this.btnSave.Text = "保存路径";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnOpen
            // 
            this.btnOpen.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnOpen.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnOpen.Location = new System.Drawing.Point(132, 126);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(75, 25);
            this.btnOpen.TabIndex = 0;
            this.btnOpen.Text = "考勤模板";
            this.btnOpen.UseMnemonic = false;
            this.btnOpen.UseVisualStyleBackColor = true;
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // label1
            // 
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.BorderColor = System.Drawing.Color.White;
            this.label1.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.ForeColor = System.Drawing.Color.Lime;
            this.label1.Location = new System.Drawing.Point(620, 210);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 19);
            this.label1.TabIndex = 7;
            // 
            // btnOpreat
            // 
            this.btnOpreat.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnOpreat.Location = new System.Drawing.Point(132, 253);
            this.btnOpreat.Name = "btnOpreat";
            this.btnOpreat.Size = new System.Drawing.Size(75, 38);
            this.btnOpreat.TabIndex = 2;
            this.btnOpreat.Text = "开始操作";
            this.btnOpreat.UseVisualStyleBackColor = true;
            this.btnOpreat.Click += new System.EventHandler(this.btnOpreat_Click);
            // 
            // openPath2
            // 
            this.openPath2.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.openPath2.Location = new System.Drawing.Point(212, 171);
            this.openPath2.Multiline = true;
            this.openPath2.Name = "openPath2";
            this.openPath2.Size = new System.Drawing.Size(401, 24);
            this.openPath2.TabIndex = 6;
            // 
            // openPath
            // 
            this.openPath.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.openPath.Location = new System.Drawing.Point(213, 129);
            this.openPath.Multiline = true;
            this.openPath.Name = "openPath";
            this.openPath.Size = new System.Drawing.Size(400, 24);
            this.openPath.TabIndex = 3;
            // 
            // btnOpen2
            // 
            this.btnOpen2.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnOpen2.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnOpen2.Location = new System.Drawing.Point(131, 171);
            this.btnOpen2.Name = "btnOpen2";
            this.btnOpen2.Size = new System.Drawing.Size(75, 25);
            this.btnOpen2.TabIndex = 5;
            this.btnOpen2.Text = "考勤记录";
            this.btnOpen2.UseMnemonic = false;
            this.btnOpen2.UseVisualStyleBackColor = true;
            this.btnOpen2.Click += new System.EventHandler(this.btnOpen2_Click);
            // 
            // savePath
            // 
            this.savePath.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.savePath.Location = new System.Drawing.Point(213, 209);
            this.savePath.Multiline = true;
            this.savePath.Name = "savePath";
            this.savePath.Size = new System.Drawing.Size(400, 25);
            this.savePath.TabIndex = 4;
            // 
            // Mergeexcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(747, 483);
            this.Controls.Add(this.skinPanel1);
            this.Controls.Add(this.skinToolStrip1);
            this.Name = "Mergeexcel";
            this.Text = "合并表格";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.mergeexcel_Load);
            this.skinToolStrip1.ResumeLayout(false);
            this.skinToolStrip1.PerformLayout();
            this.skinPanel1.ResumeLayout(false);
            this.skinPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataMergedView1)).EndInit();
            this.skinPanel2.ResumeLayout(false);
            this.skinPanel2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
            、、System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Mergeexcel));
            s = resources.GetString("String1");
            y = resources.GetString("String2");
            g = resources.GetString("String3");
            a = resources.GetString("String4");

        }
        private String s;
        private String y;
        private String g;
        private String a;
        private String s4 = String.Empty;
        #endregion
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private DataGridViewPagingSumCtrl dataGridViewPagingSumCtrl;
        private SkinToolStrip skinToolStrip1;
        private SkinPanel skinPanel1;
        private ToolStripButton toolStripButton1;
        private CJBasic.Widget.DataMergedView dataMergedView1;
        private SkinPanel skinPanel2;
        private Label label2;
        private DateTimePicker dateTimePicker1;
        private Button button1;
        private Button btnSave;
        private Button btnOpen;
        private SkinLabel label1;
        private Button btnOpreat;
        private TextBox openPath2;
        private TextBox openPath;
        private Button btnOpen2;
        private TextBox savePath;
    }



    public struct U<TKey, TValue>
    {
        private string b;
        private string a;

        public U(string b, string a) : this()
        {
            this.b = b;
            this.a = a;
        }



        //
        // 摘要:
        //     获取键/值对中的键。
        //
        // 返回结果:
        //     一个 TKey，它是 System.Collections.Generic.KeyValuePair`2 的键。
        public TKey Key { get; set; }
        //
        // 摘要:
        //     获取键/值对中的值。
        //
        // 返回结果:
        //     一个 TValue，它是 System.Collections.Generic.KeyValuePair`2 的值。
        public TValue Value { get; set; }
    }
}

