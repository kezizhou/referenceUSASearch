using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Controls;

namespace referenceUSASearch {
	public class TextBoxWriter : TextWriter {

		private TextBox txt;
		public TextBoxWriter(TextBox _txt) {
			txt = _txt;
		}

		public override void Write(char value) {
			txt.Text += value;
		}

		public override void Write(string value) {
			txt.Text += value;
		}

		public override Encoding Encoding {
			get { 
				return Encoding.Unicode;  
			}
		}
	}
}
