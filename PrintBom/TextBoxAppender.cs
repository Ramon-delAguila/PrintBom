﻿using log4net.Appender;
using log4net.Core;
using System;
using System.Windows.Forms;

namespace PrintBom
{
    public class TextBoxAppender : AppenderSkeleton
    {
        private TextBox _textBox;

        public TextBoxAppender()
        {

        }

        protected override void Append(LoggingEvent loggingEvent)
        {
            if (_textBox != null)
            {
                _textBox.Do(t =>
                {
                    t.AppendText(String.Format("{0}{1}", loggingEvent.RenderedMessage, Environment.NewLine));
                });
            }
        }

        public TextBox TextBox
        {
            get { return _textBox; }
            set { _textBox = value; }
        }
    }

    //private TextBox _textBox;
    //public TextBox AppenderTextBox
    //{
    //    get
    //    {
    //        return _textBox;
    //    }
    //    set
    //    {
    //        _textBox = value;
    //    }
    //}
    //public string FormName { get; set; }
    //public string TextBoxName { get; set; }

    //private Control FindControlRecursive(Control root, string textBoxName)
    //{
    //    if (root.Name == textBoxName) return root;
    //    foreach (Control c in root.Controls)
    //    {
    //        Control t = FindControlRecursive(c, textBoxName);
    //        if (t != null) return t;
    //    }
    //    return null;
    //}

    //protected override void Append(log4net.Core.LoggingEvent loggingEvent)
    //{
    //    if (_textBox == null)
    //    {
    //        if (String.IsNullOrEmpty(FormName) ||
    //            String.IsNullOrEmpty(TextBoxName))
    //            return;

    //        Form form = Application.OpenForms[FormName];
    //        if (form == null)
    //            return;

    //        _textBox = (TextBox)FindControlRecursive(form, TextBoxName);
    //        if (_textBox == null)
    //            return;

    //        form.FormClosing += (s, e) => _textBox = null;
    //    }
    //    _textBox.BeginInvoke((MethodInvoker)delegate
    //    {
    //        _textBox.AppendText(RenderLoggingEvent(loggingEvent));
    //    });
    //}
}



