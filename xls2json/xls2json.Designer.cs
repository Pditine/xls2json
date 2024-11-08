﻿using System.IO;
using System.Collections;
using System.Linq;
using System.Text.Json;
namespace xls2json
{
    partial class xls2json
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(xls2json));
            ConvertButton = new Button();
            CheckBoxList = new CheckedListBox();
            Link = new LinkLabel();
            LogBox = new RichTextBox();
            SuspendLayout();
            // 
            // ConvertButton
            // 
            resources.ApplyResources(ConvertButton, "ConvertButton");
            ConvertButton.Name = "ConvertButton";
            ConvertButton.UseVisualStyleBackColor = true;
            ConvertButton.Click += Convert_Click;
            // 
            // CheckBoxList
            // 
            CheckBoxList.FormattingEnabled = true;
            resources.ApplyResources(CheckBoxList, "CheckBoxList");
            CheckBoxList.Name = "CheckBoxList";
            CheckBoxList.SelectedIndexChanged += checkedListBox1_SelectedIndexChanged;
            // 
            // Link
            // 
            resources.ApplyResources(Link, "Link");
            Link.Name = "Link";
            Link.TabStop = true;
            Link.LinkClicked += LinkLabel;
            // 
            // LogBox
            // 
            LogBox.AcceptsTab = true;
            LogBox.ForeColor = SystemColors.WindowText;
            resources.ApplyResources(LogBox, "LogBox");
            LogBox.Name = "LogBox";
            LogBox.ReadOnly = true;
            // 
            // xls2json
            // 
            resources.ApplyResources(this, "$this");
            AutoScaleMode = AutoScaleMode.Font;
            Controls.Add(CheckBoxList);
            Controls.Add(LogBox);
            Controls.Add(Link);
            Controls.Add(ConvertButton);
            Name = "xls2json";
            Load += Xls2json_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private void PreLoadExcel()
        {
            var files =  Directory.GetFiles("/Excel");
            var fileList = new List<string>(files);
            //fileList.RemoveAll(f => f.)
        }

        private Button ConvertButton;
        private CheckedListBox CheckBoxList;
        private LinkLabel Link;
        private RichTextBox LogBox;

        //private bool NotExcel(string f)
        //{

        //}

    }
}
