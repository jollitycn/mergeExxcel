using CCWin.SkinControl;
using CJBasic.CJBasic.Helpers;
using JGNet.Common.Components;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace MergeExcel1
{
    //}
    partial class Mergeexcel1
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
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.skinPanel2 = new CCWin.SkinControl.SkinPanel();
            this.label1 = new CCWin.SkinControl.SkinLabel();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnOpen = new System.Windows.Forms.Button();
            this.btnOpreat = new System.Windows.Forms.Button();
            this.openPath2 = new System.Windows.Forms.TextBox();
            this.openPath = new System.Windows.Forms.TextBox();
            this.btnOpen2 = new System.Windows.Forms.Button();
            this.savePath = new System.Windows.Forms.TextBox();
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
            // skinPanel2
            // 
            this.skinPanel2.AutoSize = true;
            this.skinPanel2.BackColor = System.Drawing.Color.Transparent;
            this.skinPanel2.Controls.Add(this.label1);
            this.skinPanel2.Controls.Add(this.btnSave);
            this.skinPanel2.Controls.Add(this.btnOpen);
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
            this.skinPanel2.Size = new System.Drawing.Size(547, 307);
            this.skinPanel2.TabIndex = 16;
            // 
            // label1
            // 
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.BorderColor = System.Drawing.Color.White;
            this.label1.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.ForeColor = System.Drawing.Color.Lime;
            this.label1.Location = new System.Drawing.Point(140, 211);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 19);
            this.label1.TabIndex = 8;
            // 
            // btnSave
            // 
            this.btnSave.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnSave.Location = new System.Drawing.Point(43, 157);
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
            this.btnOpen.Location = new System.Drawing.Point(44, 73);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(75, 25);
            this.btnOpen.TabIndex = 0;
            this.btnOpen.Text = "考勤模板";
            this.btnOpen.UseMnemonic = false;
            this.btnOpen.UseVisualStyleBackColor = true;
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // btnOpreat
            // 
            this.btnOpreat.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnOpreat.Location = new System.Drawing.Point(44, 204);
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
            this.openPath2.Location = new System.Drawing.Point(124, 115);
            this.openPath2.Multiline = true;
            this.openPath2.Name = "openPath2";
            this.openPath2.Size = new System.Drawing.Size(401, 24);
            this.openPath2.TabIndex = 6;
            // 
            // openPath
            // 
            this.openPath.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.openPath.Location = new System.Drawing.Point(125, 74);
            this.openPath.Multiline = true;
            this.openPath.Name = "openPath";
            this.openPath.Size = new System.Drawing.Size(400, 24);
            this.openPath.TabIndex = 3;
            // 
            // btnOpen2
            // 
            this.btnOpen2.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnOpen2.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnOpen2.Location = new System.Drawing.Point(43, 114);
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
            this.savePath.Location = new System.Drawing.Point(125, 157);
            this.savePath.Multiline = true;
            this.savePath.Name = "savePath";
            this.savePath.Size = new System.Drawing.Size(400, 25);
            this.savePath.TabIndex = 4;
            // 
            // Mergeexcel1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(547, 307);
            this.Controls.Add(this.skinPanel2);
            this.Name = "Mergeexcel1";
            this.Text = "合并表格";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.mergeexcel_Load);
            this.skinPanel2.ResumeLayout(false);
            this.skinPanel2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

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
        private SkinPanel skinPanel2;
        private SkinLabel label1;
        private Button btnSave;
        private Button btnOpen;
        private Button btnOpreat;
        private TextBox openPath2;
        private TextBox openPath;
        private Button btnOpen2;
        private TextBox savePath;
        // private DataGridViewPagingSumCtrl dataGridViewPagingSumCtrl;
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

