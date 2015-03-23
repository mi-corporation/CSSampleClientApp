using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

namespace CSSampleClientApp
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.MainMenu mainMenu1;
		private System.Windows.Forms.MenuItem menuItem1;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private System.Windows.Forms.MenuItem menuItem3;
		private System.Windows.Forms.MenuItem menuItem4;
		private System.Windows.Forms.MenuItem menuItem5;
		private System.Windows.Forms.MenuItem menuItem6;
		private System.Windows.Forms.ToolBar toolBar1;
		private System.Windows.Forms.ToolBarButton _btnPrevPage;
		private System.Windows.Forms.ToolBarButton _btnNextPage;
		private System.Windows.Forms.ImageList imageList1;
		private MiCo.MiForms.MiFormsComponent miFormsComponent1;
		private System.Windows.Forms.MenuItem mnuItmLoad;
		private System.Windows.Forms.MenuItem mnuItmPrint;
		private System.Windows.Forms.MenuItem mnuItmPrintPreview;
		private System.ComponentModel.IContainer components;

		public Form1()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			string strResourcePath = System.Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
			strResourcePath = System.IO.Path.Combine(strResourcePath, "Mi-Co");
			strResourcePath = System.IO.Path.Combine(strResourcePath, "Mi-Forms SDK");
			strResourcePath = System.IO.Path.Combine(strResourcePath, "res");

			this.miFormsComponent1.RecognitionResourcePath = strResourcePath;
			miFormsComponent1.PrintRequested += new MiCo.MiForms.MiFormsComponent.PrintRequestedEventHandler(mfc_PrintRequested);

			_printDoc.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(printDocument_PrintPage);
			_printDoc.QueryPageSettings += new System.Drawing.Printing.QueryPageSettingsEventHandler(printDocument_QueryPageSettings);

			_printDialog.Document = _printDoc;
			_printPreviewDialog.Document = _printDoc;

                        miFormsComponent1.LicensePath = MiCo.MiForms.License.DefaultLicensePath();
			bool bRes = this.miFormsComponent1.Init();
			if (!bRes)
			{
				MessageBox.Show ("Init error: " + this.miFormsComponent1.InitError);
			}
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Form1));
			this.mainMenu1 = new System.Windows.Forms.MainMenu();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this.mnuItmLoad = new System.Windows.Forms.MenuItem();
			this.mnuItmPrint = new System.Windows.Forms.MenuItem();
			this.mnuItmPrintPreview = new System.Windows.Forms.MenuItem();
			this.menuItem3 = new System.Windows.Forms.MenuItem();
			this.menuItem4 = new System.Windows.Forms.MenuItem();
			this.menuItem5 = new System.Windows.Forms.MenuItem();
			this.menuItem6 = new System.Windows.Forms.MenuItem();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.toolBar1 = new System.Windows.Forms.ToolBar();
			this._btnPrevPage = new System.Windows.Forms.ToolBarButton();
			this._btnNextPage = new System.Windows.Forms.ToolBarButton();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.miFormsComponent1 = new MiCo.MiForms.MiFormsComponent();
			this.SuspendLayout();
			// 
			// mainMenu1
			// 
			this.mainMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																					  this.menuItem1,
																					  this.menuItem3});
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 0;
			this.menuItem1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																					  this.mnuItmLoad,
																					  this.mnuItmPrint,
																					  this.mnuItmPrintPreview});
			this.menuItem1.Text = "&File";
			// 
			// mnuItmLoad
			// 
			this.mnuItmLoad.Index = 0;
			this.mnuItmLoad.Text = "&Load Form";
			this.mnuItmLoad.Click += new System.EventHandler(this.mnuItmLoad_Click);
			// 
			// mnuItmPrint
			// 
			this.mnuItmPrint.Enabled = false;
			this.mnuItmPrint.Index = 1;
			this.mnuItmPrint.Shortcut = System.Windows.Forms.Shortcut.CtrlP;
			this.mnuItmPrint.Text = "Print";
			this.mnuItmPrint.Click += new System.EventHandler(this.mnuItmPrint_Click);
			// 
			// mnuItmPrintPreview
			// 
			this.mnuItmPrintPreview.Enabled = false;
			this.mnuItmPrintPreview.Index = 2;
			this.mnuItmPrintPreview.Text = "Print Preview";
			this.mnuItmPrintPreview.Click += new System.EventHandler(this.mnuItmPrintPreview_Click);
			// 
			// menuItem3
			// 
			this.menuItem3.Index = 1;
			this.menuItem3.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																					  this.menuItem4,
																					  this.menuItem5,
																					  this.menuItem6});
			this.menuItem3.Text = "&Input";
			this.menuItem3.Popup += new System.EventHandler(this.menuItem3_Popup);
			// 
			// menuItem4
			// 
			this.menuItem4.Index = 0;
			this.menuItem4.RadioCheck = true;
			this.menuItem4.Text = "&Pen";
			this.menuItem4.Click += new System.EventHandler(this.menuItem4_Click);
			// 
			// menuItem5
			// 
			this.menuItem5.Index = 1;
			this.menuItem5.RadioCheck = true;
			this.menuItem5.Text = "&Eraser";
			this.menuItem5.Click += new System.EventHandler(this.menuItem5_Click);
			// 
			// menuItem6
			// 
			this.menuItem6.Index = 2;
			this.menuItem6.RadioCheck = true;
			this.menuItem6.Text = "&Keyboard";
			this.menuItem6.Click += new System.EventHandler(this.menuItem6_Click);
			// 
			// toolBar1
			// 
			this.toolBar1.Buttons.AddRange(new System.Windows.Forms.ToolBarButton[] {
																						this._btnPrevPage,
																						this._btnNextPage});
			this.toolBar1.DropDownArrows = true;
			this.toolBar1.ImageList = this.imageList1;
			this.toolBar1.Location = new System.Drawing.Point(0, 0);
			this.toolBar1.Name = "toolBar1";
			this.toolBar1.ShowToolTips = true;
			this.toolBar1.Size = new System.Drawing.Size(680, 28);
			this.toolBar1.TabIndex = 1;
			this.toolBar1.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(this.toolBar1_ButtonClick);
			// 
			// _btnPrevPage
			// 
			this._btnPrevPage.Enabled = false;
			this._btnPrevPage.ImageIndex = 0;
			// 
			// _btnNextPage
			// 
			this._btnNextPage.Enabled = false;
			this._btnNextPage.ImageIndex = 1;
			// 
			// imageList1
			// 
			this.imageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit;
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// miFormsComponent1
			// 
			this.miFormsComponent1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.miFormsComponent1.Location = new System.Drawing.Point(0, 28);
			this.miFormsComponent1.Name = "miFormsComponent1";
			this.miFormsComponent1.Size = new System.Drawing.Size(680, 538);
			this.miFormsComponent1.TabIndex = 0;
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(680, 566);
			this.Controls.Add(this.miFormsComponent1);
			this.Controls.Add(this.toolBar1);
			this.Menu = this.mainMenu1;
			this.Name = "Form1";
			this.Text = "Form1";
			this.Load += new System.EventHandler(this.Form1_Load);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new Form1());
		}

		private void Form1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void mnuItmLoad_Click(object sender, System.EventArgs e)
		{
			MiCo.MiForms.LoadedForm xloadedform;


			openFileDialog1.Filter = "Mi-Forms Templates (*.mft)|*.mft";

			if (openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				xloadedform = miFormsComponent1.LoadFormFromFile(openFileDialog1.FileName);

				if (xloadedform == null) return;

				// set the size of the component appropriately to display the page at 72 dpi
				MiCo.MiForms.FormPage xpage = (MiCo.MiForms.FormPage)xloadedform.Form.Pages[0];
				SetCanvasSize(xpage.Size.Width / 2.54, xpage.Size.Height / 2.54);

				miFormsComponent1.ActiveForm = xloadedform;

				enableToolbarButtons();
			}
		}

		private void SetCanvasSize(double fWidth, double fHeight)
		{
			int nBorderH = this.Size.Height - miFormsComponent1.Size.Height;
			int nBorderW = this.Size.Width - miFormsComponent1.Size.Width;																																																																	

			int nWidth = (int)(fWidth * 72.0 + (double)nBorderW);
			int nHeight = (int)(fHeight * 72.0 + (double)nBorderH);

			this.Size = new System.Drawing.Size(nWidth, nHeight);
		}

		private void menuItem4_Click(object sender, System.EventArgs e)
		{
			miFormsComponent1.InkEnabled = true;
			miFormsComponent1.EraseMode = false;
			miFormsComponent1.ShowFocus = false;
		}

		private void menuItem5_Click(object sender, System.EventArgs e)
		{
			miFormsComponent1.InkEnabled = true;
			miFormsComponent1.EraseMode = true;
			miFormsComponent1.ShowFocus = false;
		}

		private void menuItem6_Click(object sender, System.EventArgs e)
		{
			miFormsComponent1.InkEnabled = false;
			miFormsComponent1.EraseMode = false;
			miFormsComponent1.ShowFocus = true;
		}


		private void menuItem3_Popup(object sender, System.EventArgs e)
		{
			if (miFormsComponent1.EraseMode)
			{
				menuItem4.Checked = false;
				menuItem5.Checked = true;
				menuItem6.Checked = false;
			}
			else
			{
				if (miFormsComponent1.InkEnabled)
				{
					menuItem4.Checked = true;
					menuItem5.Checked = false;
					menuItem6.Checked = false;
				}
				else
				{
					menuItem4.Checked = false;
					menuItem5.Checked = false;
					menuItem6.Checked = true;
				}
			}		
		}

		private void toolBar1_ButtonClick(object sender, System.Windows.Forms.ToolBarButtonClickEventArgs e)
		{
			MiCo.MiForms.LoadedForm xlf = miFormsComponent1.ActiveForm;

			if (e.Button == this._btnPrevPage)
			{
				if (xlf.CurrentPage > 1)
				{
					miFormsComponent1.GotoPage(xlf, xlf.CurrentPage - 1);
				}
			}
			else if (e.Button == this._btnNextPage)
			{
				if (xlf.CurrentPage < xlf.Form.Pages.Count)
				{
					miFormsComponent1.GotoPage(xlf, xlf.CurrentPage + 1);
				}
			}

			enableToolbarButtons();
		}

		private void enableToolbarButtons()
		{
			MiCo.MiForms.LoadedForm xlf = miFormsComponent1.ActiveForm;

			if (xlf == null)
			{
				mnuItmPrint.Enabled = false;
				mnuItmPrintPreview.Enabled = false;
				return;
			}

			mnuItmPrint.Enabled = true;
			mnuItmPrintPreview.Enabled = true;

			_btnPrevPage.Enabled = (xlf.CurrentPage > 1);
			_btnNextPage.Enabled = (xlf.CurrentPage < xlf.Form.Pages.Count);
		}

		#region "Print Support"
		
		bool _bPrintPreview = false;
		bool _bPagesPrinted = false;
		bool _bOverrideRenderType = false;
		MiCo.MiForms.RenderType _xRenderType = MiCo.MiForms.RenderType.FieldValues;
		int _nPageToPrint = 0;
		PrintDialog _printDialog = new PrintDialog();
		PrintPreviewDialog _printPreviewDialog = new PrintPreviewDialog();
		System.Drawing.Printing.PrintDocument _printDoc = new System.Drawing.Printing.PrintDocument();

		/// <summary>
		/// Standard QueryPageSettings event handler
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void printDocument_QueryPageSettings(object sender, System.Drawing.Printing.QueryPageSettingsEventArgs e)
		{
			MiCo.MiForms.LoadedForm xlform = miFormsComponent1.ActiveForm;

			// call SignalPagePrint in the component to see if the client's script
			// code wants to cancel the print...
			while ((!miFormsComponent1.SignalPagePrint(xlform, _nPageToPrint, MiCo.MiForms.RenderType.FieldValues) && _nPageToPrint <= xlform.Form.Pages.Count))
			{
				_nPageToPrint++;
			}

			bool bTooBig = false;
			if (_nPageToPrint > xlform.Form.Pages.Count) 
			{
				bTooBig = true;
			}

			if (!bTooBig) 
			{
				_bPagesPrinted = true;
				MiCo.MiForms.FormPage xFormPage = (MiCo.MiForms.FormPage)xlform.Form.Pages[_nPageToPrint-1];
				if (xFormPage.Size.Width > xFormPage.Size.Height) 
				{
					e.PageSettings.Landscape = true;
				}
				else 
				{
					e.PageSettings.Landscape = false;
				}
			}
			else 
			{
				if (!_bPagesPrinted) 
				{
					e.Cancel = true;
				}
			}
		}
		
		/// <summary>
		/// Standard PrintPage event handler
		/// </summary>
		private void printDocument_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
		{
			System.Drawing.Graphics xGraphics;
			xGraphics = e.Graphics;

			System.Drawing.Bitmap xBitmap;

			// create a render engine that we'll use to render the given page
			MiCo.MiForms.RenderEngine xRender = new MiCo.MiForms.RenderEngine();

			if (_bOverrideRenderType)
			{
				// this is set via the script call RequestPrint();
				xRender.Type = _xRenderType;
				_bOverrideRenderType = false;
			}
			else
			{
				// here you could use a user preference if you wanted; we just hard-code
				// it to render field values
				xRender.Type = MiCo.MiForms.RenderType.FieldValues;
			}

			MiCo.MiForms.LoadedForm xlform = miFormsComponent1.ActiveForm;
			MiCo.MiForms.FormPage xFormPage = (MiCo.MiForms.FormPage)xlform.Form.Pages[_nPageToPrint-1];
			double nMaxResolution = 0;
			double nResolution = 0;
			foreach (MiCo.MiForms.BackgroundImage xBackground in xFormPage.BackgroundImages) 
			{
				if (xBackground.Resolution > nMaxResolution) 
				{
					nMaxResolution = xBackground.Resolution;
				}
			}
			nMaxResolution *= 2.54;
			if (xGraphics.DpiX < nMaxResolution) 
			{
				nResolution = xGraphics.DpiX;
			}
			else 
			{
				nResolution = nMaxResolution;
			}
			nResolution /= 2.54;
			xBitmap = xRender.RenderPage(xFormPage, nResolution);
			xGraphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
			xGraphics.DrawImage(xBitmap, 0, 0);

			if (_bPrintPreview) 
			{
				if (_nPageToPrint < xlform.Form.Pages.Count) 
				{
					_nPageToPrint++;
					e.HasMorePages = true;
				}
				else 
				{
					e.HasMorePages = false;
					_nPageToPrint = 1;
				}
			}
			else if (_nPageToPrint < _printDoc.PrinterSettings.ToPage) 
			{
				_nPageToPrint++;
				e.HasMorePages = true;
			}
			else 
			{
				e.HasMorePages = false;
				_nPageToPrint = 1;
			}

		}

		/// <summary>
		/// Handles the PrintRequested event; we handle this so that form scripts are able
		/// to request prints.
		/// </summary>
		private void mfc_PrintRequested(object sender, MiCo.MiForms.MiFormsComponent.PrintRequestedEventArgs e)
		{
			_bPagesPrinted = false;
			if (e.Form == null) 
			{
				return;
			}
			if (e.Page < 1 || e.Page > e.Form.Form.Pages.Count) 
			{
				return;
			}

			// we'll override the default render type so that the script author can
			// specify the desired render type
			_bOverrideRenderType = true;
			_xRenderType = e.RenderType;

			MiCo.MiForms.LoadedForm xlform = miFormsComponent1.ActiveForm;

			// fire off the print job
			_printDoc.DocumentName = xlform.Form.Name + " (" + xlform.Form.GetSessionDescriptor() + ")";
			_printDoc.PrintController = new System.Drawing.Printing.StandardPrintController();
			_printDoc.PrinterSettings.FromPage = e.Page;
			_printDoc.PrinterSettings.ToPage = e.Page;
			_nPageToPrint = _printDoc.PrinterSettings.FromPage;
			_printDoc.Print();

			e.Success = true;
		}


#endregion

		private void mnuItmPrintPreview_Click(object sender, System.EventArgs e)
		{
			MiCo.MiForms.LoadedForm xlform = miFormsComponent1.ActiveForm;

			_bPrintPreview = true;
			_nPageToPrint = 1;
			_printDoc.DocumentName = xlform.Form.Name + " (" + xlform.Form.GetSessionDescriptor() + ")";

			try
			{
				// call SignalFormPrint() so the form script has the chance to cancel the print
				if (miFormsComponent1.SignalFormPrint(xlform, 1, xlform.Form.Pages.Count, MiCo.MiForms.RenderType.FieldValues)) 
				{
					_printPreviewDialog.ShowDialog();
				}
			}
			catch(System.Exception ex)
			{
				System.Windows.Forms.MessageBox.Show(this, String.Format ("Print preview failure: {0}", ex.Message));
			}
		}

		private void mnuItmPrint_Click(object sender, System.EventArgs e)
		{
			_bPagesPrinted = false;
			_bPrintPreview = false;

			MiCo.MiForms.LoadedForm xlform = miFormsComponent1.ActiveForm;
			_printDoc.DocumentName = xlform.Form.Name + " (" + xlform.Form.GetSessionDescriptor() + ")";
			_printDialog.AllowSomePages = true;
			_printDialog.PrintToFile = true;
			_printDialog.PrinterSettings.FromPage = 1;
			_printDialog.PrinterSettings.ToPage = xlform.Form.Pages.Count;
			_printDialog.PrintToFile = false;
			try
			{
				if (_printDialog.ShowDialog() == DialogResult.OK)
				{
					MiCo.MiForms.RenderType xType = MiCo.MiForms.RenderType.FieldValues;

					// call SignalFormPrint() in case the client script wants to cancel the
					// print
					if (miFormsComponent1.SignalFormPrint(xlform, _printDialog.PrinterSettings.FromPage, _printDialog.PrinterSettings.ToPage, xType)) 
					{
						_nPageToPrint = _printDoc.PrinterSettings.FromPage;
						_printDoc.PrintController = new System.Drawing.Printing.StandardPrintController();
						_printDoc.Print();
					}
				}
			}
			catch (System.Exception ex)
			{
				System.Windows.Forms.MessageBox.Show(this, String.Format ("Print failed: {0}", ex.Message));
			}
		}



	}
}
