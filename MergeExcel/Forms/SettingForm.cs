using CCWin;
using JGNet.Core.InteractEntity;
using CJBasic.Loggers;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;

using System.Text;

using System.Windows.Forms;

namespace JGNet.Common
{
    public partial class SettingForm : Common.BaseForm
    {
        private SettingsConfiguration configuration = new SettingsConfiguration();
        private Setting setting = null;
        public Setting result {
            get { return setting; }
            set { setting = value; }
        }

        public SettingForm(Setting setting = null)
        {
            InitializeComponent();
            this.setting = setting;
            this.Initialize();
        }

        private void Initialize()
        {
            //this.guideComboBox1.Initialize(GuideComboBoxInitializeType.Null, CommonGlobalCache.CurrentShopID);
        }

        private void BaseButton_OK_Click(object sender, EventArgs e)
        {
            //configuration.Settings = this.setting;
           // configuration.Save();
                this.DialogResult = DialogResult.OK;
        }

        private void BaseButton_Cancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }
         
    }


    public class SettingsConfiguration : CJBasic.Serialization.AgileConfiguration
    {
        public List< Setting> Settings { get; set; }
    }

    [Serializable]
    public class  Setting
    {
        public DateTime Start { get; set; }
        public DateTime End { get; set; } 
    }
}
