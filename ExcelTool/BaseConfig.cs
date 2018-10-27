using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelTool
{
   public class BaseConfig
    {
        public static string LoadConfigInfo(string aGroupName, string aFieldName,string pathName)
        {
            string configPath = @".\Config.ini";
            if (!File.Exists(configPath))
            {
                File.Create(configPath).Dispose();
            }

            EasyConfig.ConfigFile configFile = new EasyConfig.ConfigFile(configPath);
            string _fieldName = configFile[aGroupName][aFieldName].AsString();
            if(string.IsNullOrEmpty(_fieldName))
            {
                System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
                dialog.Description = "请选择"+ pathName;
                dialog.SelectedPath = Application.StartupPath;
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    configFile[aGroupName].WriteSetting(aFieldName, dialog.SelectedPath);
                    configFile.Save(configPath);
                    dialog.Dispose();
                    return configFile[aGroupName][aFieldName].AsString(); 
                }

            }
            return _fieldName;
        }
        public static string SetConfigInfo(string aGroupName, string aFieldName, string pathName)
        {
            string configPath = @".\Config.ini";
            if (!File.Exists(configPath))
            {
                File.Create(configPath).Dispose();
            }

            EasyConfig.ConfigFile configFile = new EasyConfig.ConfigFile(configPath);
            System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
            dialog.Description = "请选择" + pathName;
            dialog.SelectedPath = Application.StartupPath;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                configFile[aGroupName].WriteSetting(aFieldName, dialog.SelectedPath);
                configFile.Save(configPath);
                dialog.Dispose();
            }
            return configFile[aGroupName][aFieldName].AsString();
        }

    }
}
